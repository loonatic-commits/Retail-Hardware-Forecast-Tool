Attribute VB_Name = "modInfoBlock"
Option Explicit

' =============================================================
' Info block
'
' Writes a decisions-and-diagnostics table for every model in
' the columns immediately right of the last month column, with
' its own month header row so it reads straight across from the
' model block it describes.
'
' Rewritten in full on every run: the whole info area and every
' dissonance fill on the Opportunity rows are cleared first, so
' no stale rows or fills survive from a prior run.
'
' Per model:
'   Decision applied  which rule set each month's number
'   Field movement    months where Field moved materially
'                     against Field N-1 (15% test)
'   SOF reference     SOF for N..N+2, inline
'   Dissonance flags  three 15% tests, each naming the pair,
'                     the direction and the size of the gap
'
' Every variance is a percentage OF CONSENSUS - Consensus N-1
' is the denominator for all three tests. A month with no
' Consensus N-1 value cannot be tested and is left blank.
'
' Any month that raises a flag also gets a fill on that month's
' Opportunity Qty Final cell. That fill is the only conditional
' formatting this module writes.
' =============================================================

Private Const INFO_ROWS As Long = 8   ' header + months + 6 content rows

Public Sub Run_Info_Block()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)
    If curCol = 0 Then
        Debug.Print "InfoBlock: no column matches the current month - pass aborted."
        Exit Sub
    End If

    Dim nMonths As Long
    nMonths = UBound(months, 1) - LBound(months, 1) + 1

    Dim lastMonthCol As Long
    lastMonthCol = CLng(months(UBound(months, 1), 1))

    Dim infoCol As Long
    infoCol = lastMonthCol + INFO_BLOCK_COL_OFFSET

    ' --- Clear the whole info area and every prior dissonance fill ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    ws.Range(ws.Cells(1, infoCol), ws.Cells(lastRow + INFO_ROWS, infoCol + nMonths + 1)).Clear

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        ClearOpportunityFills ws, b, months
        WriteInfoForModel ws, b, months, curCol, infoCol
NextBlock:
        On Error GoTo Fail
    Next b

    ws.Columns(infoCol).ColumnWidth = 26

    Exit Sub

BlockFail:
    Debug.Print "InfoBlock failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Info Block"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ClearOpportunityFills(ws As Worksheet, ByVal block As Object, ByVal months As Variant)
    Dim oppRow As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    If oppRow = 0 Then Exit Sub
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        ws.Cells(oppRow, CLng(months(i, 1))).Interior.Pattern = xlNone
    Next i
End Sub

Private Sub WriteInfoForModel(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                              ByVal curCol As Long, ByVal infoCol As Long)
    Dim startRow As Long
    startRow = CLng(block("startRow"))

    Dim oppRow As Long, ffRow As Long, ffN1Row As Long, consN1Row As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    ffN1Row = GetKeyRow(block, KF_FIELD_FCST_N1)
    consN1Row = GetKeyRow(block, KF_CONSENSUS_N1)

    ' Replay the selection walk read-only to label each month.
    Dim decisions As Object
    Set decisions = ComputeDecisions(ws, block, months, curCol, False)

    Dim sofRow As Long
    sofRow = FindSofRowForModel(CStr(block("name")))

    ' --- Header rows ---
    ws.Cells(startRow, infoCol).Value = CStr(block("name")) & " - forecast info"
    ws.Cells(startRow, infoCol).Font.Bold = True

    Dim rTitle As Long: rTitle = startRow + 1
    ws.Cells(rTitle, infoCol).Value = "Key Figure"
    ws.Cells(rTitle, infoCol).Font.Bold = True

    Dim i As Long, c As Long
    For i = LBound(months, 1) To UBound(months, 1)
        c = infoCol + 1 + (i - LBound(months, 1))
        ws.Cells(rTitle, c).Value = Format(CDate(months(i, 2)), "yy-mmm")
        ws.Cells(rTitle, c).Font.Bold = True
        ws.Cells(rTitle, c).HorizontalAlignment = xlCenter
    Next i

    Dim rDecision As Long: rDecision = startRow + 2
    Dim rMovement As Long: rMovement = startRow + 3
    Dim rSof As Long: rSof = startRow + 4
    Dim rFlag1 As Long: rFlag1 = startRow + 5
    Dim rFlag2 As Long: rFlag2 = startRow + 6
    Dim rFlag3 As Long: rFlag3 = startRow + 7

    ws.Cells(rDecision, infoCol).Value = "Decision applied"
    ws.Cells(rMovement, infoCol).Value = "Field movement"
    ws.Cells(rSof, infoCol).Value = "SOF reference"
    ws.Cells(rFlag1, infoCol).Value = "Flag: Field N-1 vs Field"
    ws.Cells(rFlag2, infoCol).Value = "Flag: Field vs Consensus N-1"
    ws.Cells(rFlag3, infoCol).Value = "Flag: Consensus N-1 vs SOF"

    ' --- Per month ---
    Dim iCur As Long: iCur = -1
    For i = LBound(months, 1) To UBound(months, 1)
        If CLng(months(i, 1)) = curCol Then iCur = i
    Next i

    For i = LBound(months, 1) To UBound(months, 1)
        Dim col As Long, idx As Long
        col = CLng(months(i, 1))
        c = infoCol + 1 + (i - LBound(months, 1))
        idx = i - iCur

        ' Decision applied
        If decisions.Exists(col) Then
            ws.Cells(rDecision, c).Value = decisions(col)
        End If

        Dim ff As Double, ffN1 As Double, consN1 As Double
        If ffRow > 0 Then ff = SafeNum(ws.Cells(ffRow, col).Value) Else ff = 0
        If ffN1Row > 0 Then ffN1 = SafeNum(ws.Cells(ffN1Row, col).Value) Else ffN1 = 0
        If consN1Row > 0 Then consN1 = SafeNum(ws.Cells(consN1Row, col).Value) Else consN1 = 0

        ' SOF reference for N..N+2
        Dim sofVal As Double: sofVal = 0
        Dim haveSof As Boolean: haveSof = False
        If sofRow > 0 Then
            haveSof = TryGetSof(sofRow, CDate(months(i, 2)), sofVal)
        End If
        If haveSof And idx >= 0 And idx <= 2 Then
            ws.Cells(rSof, c).Value = sofVal
        End If

        ' All three tests use Consensus N-1 as the denominator.
        Dim flagged As Boolean: flagged = False
        If consN1 <> 0 Then
            Dim v1 As Double, v2 As Double, v3 As Double

            v1 = (ff - ffN1) / consN1
            If Abs(v1) > VARIANCE_DISSONANCE Then
                ws.Cells(rFlag1, c).Value = "Field " & SignedPct(v1) & " vs Field N-1"
                flagged = True
            End If

            v2 = (ff - consN1) / consN1
            If Abs(v2) > VARIANCE_DISSONANCE Then
                ws.Cells(rFlag2, c).Value = "Field " & SignedPct(v2) & " vs Cons N-1"
                flagged = True
            End If

            If haveSof Then
                v3 = (sofVal - consN1) / consN1
                If Abs(v3) > VARIANCE_DISSONANCE Then
                    ws.Cells(rFlag3, c).Value = "SOF " & SignedPct(v3) & " vs Cons N-1"
                    flagged = True
                End If
            End If

            ' Field movement uses the same 15% test on Field vs Field N-1.
            If Abs(v1) > VARIANCE_DISSONANCE Then
                ws.Cells(rMovement, c).Value = SignedPct(v1) & " vs FF N-1"
            End If
        End If

        If flagged And oppRow > 0 Then
            ApplyFillColor ws.Cells(oppRow, col), CLR_DISSONANCE
        End If
    Next i
End Sub

Private Function SignedPct(ByVal v As Double) As String
    Dim s As String
    s = Format(Abs(v) * 100, "0.0") & "%"
    If v >= 0 Then
        SignedPct = "+" & s
    Else
        SignedPct = "-" & s
    End If
End Function

' --- SOF lookup helpers -------------------------------------

Private Function FindSofRowForModel(ByVal modelName As String) As Long
    Dim wsSof As Worksheet
    On Error Resume Next
    Set wsSof = ThisWorkbook.Worksheets(SHEET_SOF)
    On Error GoTo 0
    If wsSof Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = wsSof.Cells(wsSof.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(wsSof.Cells(r, 1).Value)), modelName, vbTextCompare) = 0 Then
            FindSofRowForModel = r
            Exit Function
        End If
    Next r
End Function

Private Function TryGetSof(ByVal sofRow As Long, ByVal monthDate As Date, ByRef outVal As Double) As Boolean
    Dim wsSof As Worksheet
    On Error Resume Next
    Set wsSof = ThisWorkbook.Worksheets(SHEET_SOF)
    On Error GoTo 0
    If wsSof Is Nothing Then Exit Function

    Dim sofMonths As Variant
    sofMonths = GetSofMonthColumns(wsSof)
    If IsEmpty(sofMonths) Then Exit Function

    Dim j As Long
    For j = LBound(sofMonths, 1) To UBound(sofMonths, 1)
        If CDate(sofMonths(j, 2)) = monthDate Then
            outVal = SafeNum(wsSof.Cells(sofRow, CLng(sofMonths(j, 1))).Value)
            TryGetSof = True
            Exit Function
        End If
    Next j
End Function
