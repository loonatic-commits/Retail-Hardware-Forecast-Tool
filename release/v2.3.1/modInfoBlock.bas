Attribute VB_Name = "modInfoBlock"
Option Explicit

' =============================================================
' Info block - notes column
'
' Writes a plain-language notes column in the first free column
' after the month data, one section per model, starting on that
' model's first row so it reads straight across from the block
' it describes.
'
' This is a notes column, not a table: every line is a sentence
' naming the months it refers to. Read it top to bottom to see
' what the run decided and what still needs a human call.
'
' Per model:
'   Decision applied  which rule set each month's number,
'                     collapsed into month ranges
'   Field movement    months where Field moved materially
'                     against Field N-1 (15% test)
'   SOF reference     SOF for N..N+2, inline
'   Dissonance flags  three 15% tests, each naming the pair,
'                     the direction and the size of the gap
'
' Every variance is a percentage OF CONSENSUS - Consensus N-1
' is the denominator for all three tests. A month with no
' Consensus N-1 value cannot be tested and is skipped.
'
' Rewritten in full on every run: the notes column and every
' dissonance fill on the Opportunity rows are cleared first, so
' nothing stale survives from a prior run.
'
' Any month that raises a flag also gets a fill on that month's
' Opportunity Qty Final cell. That fill is the only conditional
' formatting this module writes.
' =============================================================

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

    Dim lastMonthCol As Long
    lastMonthCol = CLng(months(UBound(months, 1), 1))

    Dim infoCol As Long
    infoCol = lastMonthCol + INFO_BLOCK_COL_OFFSET

    ' --- Clear the notes column in full ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    ws.Range(ws.Cells(1, infoCol), ws.Cells(lastRow + 20, infoCol)).Clear

    ws.Cells(1, infoCol).Value = "Notes"
    ws.Cells(1, infoCol).Font.Bold = True

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        ClearOpportunityFills ws, b, months
        WriteNotesForModel ws, b, months, curCol, infoCol
NextBlock:
        On Error GoTo Fail
    Next b

    ws.Columns(infoCol).ColumnWidth = 70

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

Private Sub WriteNotesForModel(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                               ByVal curCol As Long, ByVal infoCol As Long)
    Dim startRow As Long, endRow As Long
    startRow = CLng(block("startRow"))
    endRow = CLng(block("endRow"))

    Dim oppRow As Long, ffRow As Long, ffN1Row As Long, consN1Row As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    ffN1Row = GetKeyRow(block, KF_FIELD_FCST_N1)
    consN1Row = GetKeyRow(block, KF_CONSENSUS_N1)

    ' Replay the selection walk read-only so the notes report
    ' exactly what the selection pass did.
    Dim decisions As Object
    Set decisions = ComputeDecisions(ws, block, months, curCol, False)

    Dim sofRow As Long
    sofRow = FindSofRowForModel(CStr(block("name")))

    Dim iCur As Long: iCur = -1
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        If CLng(months(i, 1)) = curCol Then iCur = i
    Next i
    If iCur < 0 Then Exit Sub

    Dim notes As Collection
    Set notes = New Collection

    ' ---------- Decision applied, collapsed into runs ----------
    Dim prevLabel As String, runStart As Long
    prevLabel = ""
    runStart = iCur
    For i = iCur To UBound(months, 1)
        Dim col As Long, lbl As String
        col = CLng(months(i, 1))
        If decisions.Exists(col) Then lbl = CStr(decisions(col)) Else lbl = ""

        If lbl <> prevLabel Then
            If Len(prevLabel) > 0 Then
                notes.Add FormatRun(months, runStart, i - 1, prevLabel)
            End If
            prevLabel = lbl
            runStart = i
        End If
    Next i
    If Len(prevLabel) > 0 Then
        notes.Add FormatRun(months, runStart, UBound(months, 1), prevLabel)
    End If

    ' ---------- Field movement + dissonance flags ----------
    Dim movement As String
    Dim flagLines As Collection
    Set flagLines = New Collection

    For i = iCur To UBound(months, 1)
        Dim c2 As Long, mLabel As String
        c2 = CLng(months(i, 1))
        mLabel = Format(CDate(months(i, 2)), "yy-mmm")

        Dim ff As Double, ffN1 As Double, consN1 As Double
        If ffRow > 0 Then ff = SafeNum(ws.Cells(ffRow, c2).Value) Else ff = 0
        If ffN1Row > 0 Then ffN1 = SafeNum(ws.Cells(ffN1Row, c2).Value) Else ffN1 = 0
        If consN1Row > 0 Then consN1 = SafeNum(ws.Cells(consN1Row, c2).Value) Else consN1 = 0

        Dim sofVal As Double, haveSof As Boolean
        haveSof = False
        If sofRow > 0 Then haveSof = TryGetSof(sofRow, CDate(months(i, 2)), sofVal)

        If consN1 <> 0 Then
            Dim v1 As Double, v2 As Double, v3 As Double
            Dim parts As String
            parts = ""

            ' All three tests measured as a percentage of Consensus N-1.
            v1 = (ff - ffN1) / consN1
            If Abs(v1) > VARIANCE_DISSONANCE Then
                parts = AppendPart(parts, "Field " & SignedPct(v1) & " vs Field N-1")
                movement = AppendList(movement, mLabel & " " & SignedPct(v1))
            End If

            v2 = (ff - consN1) / consN1
            If Abs(v2) > VARIANCE_DISSONANCE Then
                parts = AppendPart(parts, "Field " & SignedPct(v2) & " vs Consensus N-1")
            End If

            If haveSof Then
                v3 = (sofVal - consN1) / consN1
                If Abs(v3) > VARIANCE_DISSONANCE Then
                    parts = AppendPart(parts, "SOF " & SignedPct(v3) & " vs Consensus N-1")
                End If
            End If

            If Len(parts) > 0 Then
                flagLines.Add "FLAG " & mLabel & ": " & parts
                If oppRow > 0 Then ApplyFillColor ws.Cells(oppRow, c2), CLR_DISSONANCE
            End If
        End If
    Next i

    If Len(movement) > 0 Then
        notes.Add "Field movement vs Field N-1: " & movement
    Else
        notes.Add "Field movement vs Field N-1: none beyond " & _
                  Format(VARIANCE_DISSONANCE * 100, "0") & "%"
    End If

    ' ---------- SOF reference for N..N+2 ----------
    Dim sofLine As String
    sofLine = ""
    If sofRow > 0 Then
        For i = iCur To WorksheetFunction.Min(iCur + 2, UBound(months, 1))
            Dim sv As Double
            If TryGetSof(sofRow, CDate(months(i, 2)), sv) Then
                sofLine = AppendPipe(sofLine, Format(CDate(months(i, 2)), "yy-mmm") & " " & Format(sv, "#,##0"))
            End If
        Next i
    End If
    If Len(sofLine) > 0 Then
        notes.Add "SOF reference (N..N+2): " & sofLine
    Else
        notes.Add "SOF reference (N..N+2): not on SOF sheet"
    End If

    ' ---------- Flags last ----------
    Dim fl As Variant
    For Each fl In flagLines
        notes.Add CStr(fl)
    Next fl

    ' ---------- Write, bounded by the model's own rows ----------
    ws.Cells(startRow, infoCol).Value = CStr(block("name")) & " - notes"
    ws.Cells(startRow, infoCol).Font.Bold = True

    Dim capacity As Long
    capacity = endRow - startRow      ' rows available below the header line
    If capacity < 1 Then capacity = 1

    Dim writeCount As Long
    writeCount = notes.Count
    Dim truncated As Long
    truncated = 0
    If writeCount > capacity Then
        truncated = writeCount - (capacity - 1)
        writeCount = capacity - 1
        If writeCount < 0 Then writeCount = 0
    End If

    Dim k As Long
    For k = 1 To writeCount
        ws.Cells(startRow + k, infoCol).Value = notes(k)
    Next k

    If truncated > 0 Then
        ws.Cells(startRow + writeCount + 1, infoCol).Value = _
            "(+" & truncated & " more flag(s) - see Immediate window)"
        Dim t As Long
        For t = writeCount + 1 To notes.Count
            Debug.Print CStr(block("name")) & " | " & notes(t)
        Next t
    End If
End Sub

' -------------------------------------------------------------
' Helpers
' -------------------------------------------------------------

Private Function FormatRun(ByVal months As Variant, ByVal iFrom As Long, ByVal iTo As Long, _
                           ByVal label As String) As String
    Dim a As String, b As String
    a = Format(CDate(months(iFrom, 2)), "yy-mmm")
    b = Format(CDate(months(iTo, 2)), "yy-mmm")
    If iFrom = iTo Then
        FormatRun = a & ": " & label
    Else
        FormatRun = a & " to " & b & ": " & label
    End If
End Function

Private Function SignedPct(ByVal v As Double) As String
    Dim s As String
    s = Format(Abs(v) * 100, "0.0") & "%"
    If v >= 0 Then
        SignedPct = "+" & s
    Else
        SignedPct = "-" & s
    End If
End Function

Private Function AppendPart(ByVal existing As String, ByVal part As String) As String
    If Len(existing) = 0 Then
        AppendPart = part
    Else
        AppendPart = existing & "; " & part
    End If
End Function

Private Function AppendList(ByVal existing As String, ByVal part As String) As String
    If Len(existing) = 0 Then
        AppendList = part
    Else
        AppendList = existing & ", " & part
    End If
End Function

Private Function AppendPipe(ByVal existing As String, ByVal part As String) As String
    If Len(existing) = 0 Then
        AppendPipe = part
    Else
        AppendPipe = existing & " | " & part
    End If
End Function

' --- SOF lookup ---------------------------------------------

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
