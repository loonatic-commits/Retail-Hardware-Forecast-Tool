Attribute VB_Name = "modFFvsBP"
Option Explicit

' =============================================================
' Pass 2 - FF vs BP
'
' Seeds the Opportunity Qty Final row from the current month
' forward. Past months are reserved for sell-in vs field
' forecast accuracy review and are never overwritten.
'
' This pass now owns N..N+2 only in practice - the forecast
' selection pass overwrites N+3 onward (and N..N+4 when the
' model is supply constrained).
'
'   FF > BP AND accuracy >= 0.80 -> write FF
'   FF > BP AND accuracy <  0.80 -> write BP (FF unreliable)
'   BP >= FF                     -> write BP
'   accuracy = -1 (ungradable)   -> treat as below threshold
'
' Writes values only. All cell fills were removed with the
' color-coding module.
' =============================================================

Public Sub Run_Module_2_FFvsBP()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)
    If curCol = 0 Then
        Debug.Print "Module 2: no column matches the current month - pass aborted to avoid overwriting past data."
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        SeedModelOpportunity ws, b, months, curCol
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 2 (FFvsBP) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 2 - FFvsBP"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SeedModelOpportunity(ws As Worksheet, ByVal block As Object, _
                                 ByVal months As Variant, ByVal curCol As Long)
    Dim ffRow As Long, bpRow As Long, oppRow As Long
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    bpRow = GetKeyRow(block, KF_BUSINESS_PLAN)
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)

    If ffRow = 0 Or bpRow = 0 Or oppRow = 0 Then
        Debug.Print "Module 2: model '" & block("name") & "' missing FF/BP/Opportunity row - skipped."
        Exit Sub
    End If

    Dim accuracy As Double
    accuracy = GetModelAccuracy(CStr(block("name")))
    Dim ffReliable As Boolean
    ffReliable = (accuracy >= ACCURACY_GREEN)

    Dim i As Long, col As Long
    Dim ff As Double, bp As Double

    For i = LBound(months, 1) To UBound(months, 1)
        col = CLng(months(i, 1))
        If col >= curCol Then
            ff = SafeNum(ws.Cells(ffRow, col).Value)
            bp = SafeNum(ws.Cells(bpRow, col).Value)

            If ff > bp And ffReliable Then
                ws.Cells(oppRow, col).Value = ff
            Else
                ws.Cells(oppRow, col).Value = bp
            End If
        End If
    Next i
End Sub
