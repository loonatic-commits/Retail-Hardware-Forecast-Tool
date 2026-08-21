Attribute VB_Name = "modForecastSelect"
Option Explicit

' =============================================================
' Forecast selection
'
' Replaces the old FF-vs-BP-driven month selection from N+3 on,
' and overrides N..N+4 entirely when a model/month is supply
' constrained.
'
' Constraint is INFERRED per model AND per month:
'   availability(month) < demandSignal(month)
'
'   availability : running inventory pool - starts at the current
'                  month's Available Inventory (on hand today),
'                  adds each later month's incoming, and is
'                  reduced by whatever consensus consumes as the
'                  walk moves forward.
'   demandSignal : max(Consensus N-1, Field N-1)
'                  (flip CONSTRAINT_USE_FIELD_N1 in modConfig to
'                   use the current Field submission instead)
'
' Constrained path
'   N .. N+3   consensus = availability
'   N+4        consensus = availability + backorder,
'              or Field if Field exceeds that figure
'   N+5 on     falls through to the normal path
'
' Normal path
'   N .. N+2   unchanged - existing pass logic stands
'   N+3 on     consensus = Consensus N-1
'              (held even when Field varies by more than 15%;
'               variance only raises a flag, it never rewrites
'               the number)
' =============================================================

' Decision labels, also consumed by the info block
Public Const DEC_CONSTRAINED As String = "Constrained - availability"
Public Const DEC_CONSTRAINED_BO As String = "Constrained + backorder"
Public Const DEC_FIELD_OVERRIDE As String = "Field override (N+4)"
Public Const DEC_NORMAL_EXISTING As String = "Normal - existing logic"
Public Const DEC_CONSENSUS_HOLD As String = "Consensus N-1 hold"

Public Sub Run_Forecast_Selection()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)
    If curCol = 0 Then
        Debug.Print "ForecastSelect: no column matches the current month - pass aborted."
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        Dim decisions As Object
        Set decisions = ComputeDecisions(ws, b, months, curCol, True)
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "ForecastSelect failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Forecast Selection"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' -------------------------------------------------------------
' ComputeDecisions
'   Walks the months for one model and returns a Dictionary of
'     columnIndex -> decision label
'   When applyValues is True the chosen values are written to the
'   Opportunity Qty Final row. When False the walk is identical
'   but read-only, so the info block reports exactly what the
'   selection pass did.
'
'   The inventory pool is always reduced by the value sitting in
'   the Opportunity cell after the decision, so both modes
'   deplete the pool identically.
' -------------------------------------------------------------
Public Function ComputeDecisions(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                                 ByVal curCol As Long, ByVal applyValues As Boolean) As Object
    Dim decisions As Object
    Set decisions = CreateObject("Scripting.Dictionary")
    Set ComputeDecisions = decisions

    Dim oppRow As Long, invRow As Long, boRow As Long
    Dim ffRow As Long, ffN1Row As Long, consN1Row As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    invRow = GetKeyRow(block, KF_AVAIL_INV)
    boRow = GetKeyRow(block, KF_BACK_ORDER)
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    ffN1Row = GetKeyRow(block, KF_FIELD_FCST_N1)
    consN1Row = GetKeyRow(block, KF_CONSENSUS_N1)

    If oppRow = 0 Then
        Debug.Print "ForecastSelect: model '" & block("name") & "' has no Opportunity row - skipped."
        Exit Function
    End If

    ' Index of the current month within the months array
    Dim iCur As Long: iCur = -1
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        If CLng(months(i, 1)) = curCol Then
            iCur = i
            Exit For
        End If
    Next i
    If iCur < 0 Then Exit Function

    Dim pool As Double: pool = 0

    For i = iCur To UBound(months, 1)
        Dim col As Long, idx As Long
        col = CLng(months(i, 1))
        idx = i - iCur                      ' 0 = N, 1 = N+1, ...

        ' Availability: on-hand at N, plus incoming in later months
        If invRow > 0 Then
            pool = pool + SafeNum(ws.Cells(invRow, col).Value)
        End If
        Dim availNow As Double
        availNow = pool
        If availNow < 0 Then availNow = 0

        Dim consN1 As Double, ffN1 As Double, ff As Double, bo As Double
        If consN1Row > 0 Then consN1 = SafeNum(ws.Cells(consN1Row, col).Value)
        If ffN1Row > 0 Then ffN1 = SafeNum(ws.Cells(ffN1Row, col).Value)
        If ffRow > 0 Then ff = SafeNum(ws.Cells(ffRow, col).Value)
        If boRow > 0 Then bo = SafeNum(ws.Cells(boRow, col).Value)

        ' Demand signal for the constraint test
        Dim signal As Double
        signal = consN1
        If CONSTRAINT_USE_FIELD_N1 Then
            If ffN1 > signal Then signal = ffN1
        Else
            If ff > signal Then signal = ff
        End If

        ' A model with no inventory row cannot be assessed as constrained.
        Dim constrained As Boolean
        constrained = (invRow > 0) And (signal > 0) And (availNow < signal)

        Dim decided As Boolean, val As Double, label As String
        decided = False

        If constrained And idx <= OFFSET_AVAILABILITY_LAST Then
            val = availNow
            label = DEC_CONSTRAINED
            decided = True
        ElseIf constrained And idx = OFFSET_BACKORDER Then
            val = availNow + bo
            label = DEC_CONSTRAINED_BO
            If ff > val Then
                val = ff
                label = DEC_FIELD_OVERRIDE
            End If
            decided = True
        Else
            ' Normal path (also every month from N+5 on)
            If idx >= OFFSET_CONSENSUS_HOLD Then
                val = consN1
                label = DEC_CONSENSUS_HOLD
                decided = True
            Else
                label = DEC_NORMAL_EXISTING
                decided = False
            End If
        End If

        If decided And applyValues Then
            ws.Cells(oppRow, col).Value = val
        End If

        decisions(col) = label

        ' Deplete the pool by whatever the Opportunity row now holds.
        pool = pool - SafeNum(ws.Cells(oppRow, col).Value)
    Next i
End Function
