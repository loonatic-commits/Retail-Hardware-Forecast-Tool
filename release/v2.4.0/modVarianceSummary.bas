Attribute VB_Name = "modVarianceSummary"
Option Explicit

' =============================================================
' modVarianceSummary - standalone, read-only
'
' Builds a summary sheet comparing Marketing Demand Forecast
' against Field Forecast over a FIXED Aug - Mar window, by
' model, in units and percent, with the reasoning behind each
' model's Opportunity numbers.
'
' STANDALONE. This does not require Run_All_Passes to have run,
' does not call any pass, and writes nothing to the Consensus
' sheet. It only reads. Manual edits to Opportunity are read as
' they stand and reported as manual.
'
' The "why" column is EVIDENCE BASED: for each month it compares
' the Opportunity value against every candidate source on the
' sheet and reports which one the number actually matches -
' availability, consensus N-1, field, business plan, run rate,
' SOF - or "manual" when it matches none of them. It describes
' what is in the cells, not what a pass would have written.
'
' Three tables on one sheet:
'   1  Summary by model   Aug-Mar totals, variance in units and
'                         percent, and where each model's
'                         Opportunity numbers came from
'   2  Monthly variance   units, model x month
'   3  Monthly variance   percent, model x month
'
' The window is fixed, not derived from today's date. To move
' it, edit the four window constants below.
'
' Run from the macro list as Build_Variance_Summary.
' =============================================================

Private Const SHEET_VARIANCE As String = "Mktg vs FF Variance"

' --- Fixed reporting window: Aug 2026 through Mar 2027 -------
Private Const VAR_START_YEAR As Long = 2026
Private Const VAR_START_MONTH As Long = 8      ' August
Private Const VAR_END_YEAR As Long = 2027
Private Const VAR_END_MONTH As Long = 3        ' March

' Which row on Consensus is the Marketing demand forecast.
' The sheet carries several Marketing rows - this is the current
' marketing submission, NOT the N-1 prior-period row and NOT the
' Marketing Opportunity row.
Private Const KF_MKTG_DEMAND As String = "Marketing Forecast Qty"

' Variance = Mktg - FF. True divides by FF (the baseline being
' compared against); False divides by Mktg.
Private Const VAR_PCT_OF_FF As Boolean = True

' Percent variance at or beyond this is called out as material.
Private Const VAR_MATERIAL As Double = 0.15

' Two values count as the same source when they are this close.
Private Const MATCH_TOLERANCE As Double = 0.5

' Highlight for material variances. Literal rather than a call
' to InitColors, so this module stands on its own.
Private Const VAR_HILITE As Long = 49407       ' RGB(255, 192, 0)

Public Sub Build_Variance_Summary()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then
        MsgBox "Could not parse Consensus month columns.", vbCritical, "Variance Summary"
        Exit Sub
    End If

    ' --- Resolve the fixed window against the sheet ---
    Dim iFrom As Long, iTo As Long
    iFrom = IndexOfMonth(months, VAR_START_YEAR, VAR_START_MONTH)
    iTo = IndexOfMonth(months, VAR_END_YEAR, VAR_END_MONTH)

    If iFrom < 0 Or iTo < 0 Then
        MsgBox "Could not find the reporting window on '" & SHEET_CONSENSUS & "'." & vbCrLf & vbCrLf & _
               "Looking for " & Format(DateSerial(VAR_START_YEAR, VAR_START_MONTH, 1), "yy-mmm") & _
               " through " & Format(DateSerial(VAR_END_YEAR, VAR_END_MONTH, 1), "yy-mmm") & "." & vbCrLf & _
               "Edit the window constants at the top of modVarianceSummary to move it.", _
               vbCritical, "Variance Summary"
        Exit Sub
    End If
    If iTo < iFrom Then
        MsgBox "The reporting window ends before it starts. Check the window constants.", _
               vbCritical, "Variance Summary"
        Exit Sub
    End If

    Dim nMonths As Long
    nMonths = iTo - iFrom + 1

    ' Where the availability walk starts. Current month if it is
    ' on the sheet, otherwise the start of the window.
    Dim poolStart As Long
    poolStart = IndexOfMonth(months, Year(Date), Month(Date))
    If poolStart < 0 Then poolStart = iFrom

    Dim rep As Worksheet
    Set rep = EnsureSheet(SHEET_VARIANCE)
    rep.Cells.Clear

    Dim windowLabel As String
    windowLabel = Format(CDate(months(iFrom, 2)), "mmm") & " - " & _
                  Format(CDate(months(iTo, 2)), "mmm")

    ' ---------- Title ----------
    rep.Cells(1, 1).Value = "Marketing Demand Forecast vs Field Forecast - " & windowLabel
    rep.Cells(1, 1).Font.Bold = True
    rep.Cells(1, 1).Font.Size = 14

    rep.Cells(2, 1).Value = "Variance = Marketing Demand Fcst - Field Forecast" & _
                            IIf(VAR_PCT_OF_FF, ", percent of Field Forecast", ", percent of Marketing") & _
                            ".  Built " & Format(Now, "yyyy-mm-dd hh:nn") & "."
    rep.Cells(3, 1).Value = "Read-only snapshot of the sheet as it stands. " & _
                            "Marketing row: '" & KF_MKTG_DEMAND & "'   |   " & _
                            "Field row: '" & KF_FIELD_FCST_FINAL & "'"
    rep.Cells(3, 1).Font.Italic = True

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    ' =========================================================
    ' Table 1 - summary by model
    ' =========================================================
    Dim r As Long
    r = 5
    rep.Cells(r, 1).Value = "1. Summary by model (" & windowLabel & " totals)"
    rep.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim hdrRow As Long: hdrRow = r
    rep.Cells(r, 1).Value = "Model"
    rep.Cells(r, 2).Value = "Mktg Demand Fcst"
    rep.Cells(r, 3).Value = "Field Forecast"
    rep.Cells(r, 4).Value = "Variance (units)"
    rep.Cells(r, 5).Value = "Variance (%)"
    rep.Cells(r, 6).Value = "Where the Opportunity numbers came from"
    StyleHeader rep.Range(rep.Cells(r, 1), rep.Cells(r, 6))
    r = r + 1

    Dim summaryFirst As Long: summaryFirst = r

    Dim b As Variant, modelName As String
    For Each b In blocks
        On Error GoTo BlockFail
        modelName = CStr(b("name"))

        Dim mktgRow As Long, ffRow As Long
        mktgRow = GetKeyRow(b, KF_MKTG_DEMAND)
        ffRow = GetKeyRow(b, KF_FIELD_FCST_FINAL)

        rep.Cells(r, 1).Value = modelName

        If mktgRow = 0 Or ffRow = 0 Then
            rep.Cells(r, 6).Value = "Cannot compare - missing row: " & _
                IIf(mktgRow = 0, "'" & KF_MKTG_DEMAND & "'", "") & _
                IIf(mktgRow = 0 And ffRow = 0, " and ", "") & _
                IIf(ffRow = 0, "'" & KF_FIELD_FCST_FINAL & "'", "")
            rep.Cells(r, 6).Font.Italic = True
            r = r + 1
            GoTo NextBlock
        End If

        Dim mktgTot As Double, ffTot As Double
        mktgTot = SumOver(ws, mktgRow, months, iFrom, iTo)
        ffTot = SumOver(ws, ffRow, months, iFrom, iTo)

        Dim varUnits As Double, varPct As Double, havePct As Boolean
        varUnits = mktgTot - ffTot
        havePct = ComputePct(mktgTot, ffTot, varPct)

        rep.Cells(r, 2).Value = mktgTot
        rep.Cells(r, 3).Value = ffTot
        rep.Cells(r, 4).Value = varUnits
        If havePct Then
            rep.Cells(r, 5).Value = varPct
            rep.Cells(r, 5).NumberFormat = "0.0%"
            If Abs(varPct) >= VAR_MATERIAL Then
                rep.Cells(r, 5).Interior.Color = VAR_HILITE
            End If
        Else
            rep.Cells(r, 5).Value = "n/a"
        End If

        rep.Cells(r, 6).Value = WhyText(ws, b, months, iFrom, iTo, poolStart, varPct, havePct)

        rep.Cells(r, 2).NumberFormat = "#,##0"
        rep.Cells(r, 3).NumberFormat = "#,##0"
        rep.Cells(r, 4).NumberFormat = "#,##0;[Red]-#,##0"

        r = r + 1
NextBlock:
        On Error GoTo Fail
    Next b

    Dim summaryLast As Long: summaryLast = r - 1

    ' ---------- Total line ----------
    If summaryLast >= summaryFirst Then
        rep.Cells(r, 1).Value = "TOTAL"
        rep.Cells(r, 2).Formula = "=SUM(B" & summaryFirst & ":B" & summaryLast & ")"
        rep.Cells(r, 3).Formula = "=SUM(C" & summaryFirst & ":C" & summaryLast & ")"
        rep.Cells(r, 4).Formula = "=B" & r & "-C" & r
        If VAR_PCT_OF_FF Then
            rep.Cells(r, 5).Formula = "=IF(C" & r & "=0,""n/a"",D" & r & "/C" & r & ")"
        Else
            rep.Cells(r, 5).Formula = "=IF(B" & r & "=0,""n/a"",D" & r & "/B" & r & ")"
        End If
        rep.Cells(r, 5).NumberFormat = "0.0%"
        rep.Cells(r, 2).NumberFormat = "#,##0"
        rep.Cells(r, 3).NumberFormat = "#,##0"
        rep.Cells(r, 4).NumberFormat = "#,##0;[Red]-#,##0"
        rep.Range(rep.Cells(r, 1), rep.Cells(r, 6)).Font.Bold = True
        rep.Range(rep.Cells(r, 1), rep.Cells(r, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
        r = r + 1
    End If

    ' =========================================================
    ' Table 2 - monthly variance in units
    ' Table 3 - monthly variance in percent
    ' =========================================================
    r = r + 2
    r = WriteMonthlyTable(rep, ws, blocks, months, iFrom, iTo, r, _
                          "2. Monthly variance (units) - Marketing minus Field", False)
    r = r + 2
    r = WriteMonthlyTable(rep, ws, blocks, months, iFrom, iTo, r, _
                          "3. Monthly variance (%) - " & _
                          IIf(VAR_PCT_OF_FF, "of Field Forecast", "of Marketing"), True)

    ' ---------- Cosmetics ----------
    rep.Columns("A:A").ColumnWidth = 18
    rep.Columns("B:E").AutoFit
    rep.Columns("F:F").ColumnWidth = 95
    rep.Rows(hdrRow).WrapText = True
    rep.Activate
    rep.Range("A1").Select

    MsgBox "Variance summary built for " & blocks.Count & " model(s) over " & _
           nMonths & " month(s) (" & windowLabel & ")." & vbCrLf & vbCrLf & _
           "Nothing on '" & SHEET_CONSENSUS & "' was changed." & vbCrLf & _
           "Sheet: '" & SHEET_VARIANCE & "'", vbInformation, "Variance Summary"
    Exit Sub

BlockFail:
    Debug.Print "VarianceSummary failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Build_Variance_Summary failed: " & Err.Description, vbCritical, "Variance Summary"
End Sub

' -------------------------------------------------------------
' Monthly table writer. asPct:=False writes unit variances,
' asPct:=True writes percentages.
' Returns the row after the table.
' -------------------------------------------------------------
Private Function WriteMonthlyTable(rep As Worksheet, ws As Worksheet, blocks As Collection, _
                                   ByVal months As Variant, ByVal iFrom As Long, ByVal iTo As Long, _
                                   ByVal startRow As Long, ByVal title As String, _
                                   ByVal asPct As Boolean) As Long
    Dim r As Long: r = startRow
    rep.Cells(r, 1).Value = title
    rep.Cells(r, 1).Font.Bold = True
    r = r + 1

    rep.Cells(r, 1).Value = "Model"
    Dim i As Long, c As Long
    For i = iFrom To iTo
        c = 2 + (i - iFrom)
        rep.Cells(r, c).Value = Format(CDate(months(i, 2)), "yy-mmm")
    Next i
    StyleHeader rep.Range(rep.Cells(r, 1), rep.Cells(r, 1 + (iTo - iFrom) + 1))
    r = r + 1

    Dim b As Variant
    For Each b In blocks
        Dim mktgRow As Long, ffRow As Long
        mktgRow = GetKeyRow(b, KF_MKTG_DEMAND)
        ffRow = GetKeyRow(b, KF_FIELD_FCST_FINAL)

        rep.Cells(r, 1).Value = CStr(b("name"))

        If mktgRow > 0 And ffRow > 0 Then
            For i = iFrom To iTo
                c = 2 + (i - iFrom)
                Dim col As Long: col = CLng(months(i, 1))
                Dim m As Double, f As Double
                m = SafeNum(ws.Cells(mktgRow, col).Value)
                f = SafeNum(ws.Cells(ffRow, col).Value)

                If asPct Then
                    Dim p As Double
                    If ComputePct(m, f, p) Then
                        rep.Cells(r, c).Value = p
                        rep.Cells(r, c).NumberFormat = "0.0%"
                        If Abs(p) >= VAR_MATERIAL Then
                            rep.Cells(r, c).Interior.Color = VAR_HILITE
                        End If
                    End If
                Else
                    rep.Cells(r, c).Value = m - f
                    rep.Cells(r, c).NumberFormat = "#,##0;[Red]-#,##0"
                End If
            Next i
        End If
        r = r + 1
    Next b

    WriteMonthlyTable = r
End Function

' -------------------------------------------------------------
' Where the Opportunity numbers came from.
'
' For each month in the window, compares the Opportunity value
' against every candidate source on the sheet and reports which
' one it matches. Consecutive months with the same answer are
' collapsed into ranges. Nothing here replays a pass - it reads
' the cells as they stand, so manual edits show up as manual.
' -------------------------------------------------------------
Private Function WhyText(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                         ByVal iFrom As Long, ByVal iTo As Long, ByVal poolStart As Long, _
                         ByVal varPct As Double, ByVal havePct As Boolean) As String
    Dim s As String

    Dim oppRow As Long, invRow As Long, boRow As Long, soFarRow As Long
    Dim ffRow As Long, consN1Row As Long, bpRow As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    invRow = GetKeyRow(block, KF_AVAIL_INV)
    boRow = GetKeyRow(block, KF_BACK_ORDER)
    soFarRow = GetKeyRow(block, KF_SO_FAR)
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    consN1Row = GetKeyRow(block, KF_CONSENSUS_N1)
    bpRow = GetKeyRow(block, KF_BUSINESS_PLAN)

    If oppRow = 0 Then
        WhyText = "no Opportunity Qty Final row on this model"
        Exit Function
    End If

    Dim sofRow As Long
    sofRow = FindSofRowForModel(CStr(block("name")))

    ' Walk the availability pool forward from poolStart so the
    ' availability comparison uses the same running-pool model
    ' the sheet is built on.
    Dim pool As Double: pool = 0
    Dim i As Long
    If poolStart < iFrom Then
        For i = poolStart To iFrom - 1
            If invRow > 0 Then pool = pool + SafeNum(ws.Cells(invRow, CLng(months(i, 1))).Value)
            pool = pool - SafeNum(ws.Cells(oppRow, CLng(months(i, 1))).Value)
        Next i
    End If

    Dim prevLabel As String, runStart As Long
    prevLabel = ""
    runStart = iFrom

    For i = iFrom To iTo
        Dim col As Long
        col = CLng(months(i, 1))

        If invRow > 0 Then pool = pool + SafeNum(ws.Cells(invRow, col).Value)
        Dim avail As Double
        avail = pool
        If avail < 0 Then avail = 0

        Dim oppVal As Variant
        oppVal = ws.Cells(oppRow, col).Value

        Dim lbl As String
        lbl = SourceLabel(ws, oppVal, col, avail, invRow, boRow, soFarRow, _
                          ffRow, consN1Row, bpRow, sofRow, months, i)

        If lbl <> prevLabel Then
            If Len(prevLabel) > 0 Then
                s = AppendSemi(s, RunText(months, runStart, i - 1, prevLabel))
            End If
            prevLabel = lbl
            runStart = i
        End If

        pool = pool - SafeNum(oppVal)
    Next i

    If Len(prevLabel) > 0 Then
        s = AppendSemi(s, RunText(months, runStart, iTo, prevLabel))
    End If

    ' Variance context
    If havePct Then
        If Abs(varPct) >= VAR_MATERIAL Then
            s = AppendSemi(s, "marketing sits " & Format(Abs(varPct), "0.0%") & " " & _
                              IIf(varPct > 0, "above", "below") & " field over the window, " & _
                              "past the " & Format(VAR_MATERIAL, "0%") & " bar - worth a look")
        Else
            s = AppendSemi(s, "marketing and field within " & Format(VAR_MATERIAL, "0%") & _
                              " over the window")
        End If
    End If

    WhyText = s
End Function

' -------------------------------------------------------------
' Which source does this Opportunity value equal?
' First match wins, most specific first.
' -------------------------------------------------------------
Private Function SourceLabel(ws As Worksheet, ByVal oppVal As Variant, ByVal col As Long, _
                             ByVal avail As Double, ByVal invRow As Long, ByVal boRow As Long, _
                             ByVal soFarRow As Long, ByVal ffRow As Long, ByVal consN1Row As Long, _
                             ByVal bpRow As Long, ByVal sofRow As Long, _
                             ByVal months As Variant, ByVal i As Long) As String
    If IsBlankOrZero(oppVal) And Not IsNumeric(oppVal) Then
        SourceLabel = "blank"
        Exit Function
    End If

    Dim v As Double
    v = SafeNum(oppVal)

    Dim bo As Double, soFar As Double
    If boRow > 0 Then bo = SafeNum(ws.Cells(boRow, col).Value)
    If soFarRow > 0 Then soFar = SafeNum(ws.Cells(soFarRow, col).Value)

    If soFarRow > 0 Then
        If soFar <> 0 And Near(v, soFar + bo) Then
            SourceLabel = "run rate (SO FAR + backorder)"
            Exit Function
        End If
    End If

    If invRow > 0 Then
        If bo <> 0 And Near(v, avail + bo) Then
            SourceLabel = "availability + backorder"
            Exit Function
        End If
        If Near(v, avail) Then
            SourceLabel = "held to availability"
            Exit Function
        End If
    End If

    Dim sofVal As Double
    If sofRow > 0 Then
        If TryGetSof(sofRow, CDate(months(i, 2)), sofVal) Then
            If Near(v, sofVal) Then
                SourceLabel = "SOF"
                Exit Function
            End If
        End If
    End If

    If consN1Row > 0 Then
        If Near(v, SafeNum(ws.Cells(consN1Row, col).Value)) Then
            SourceLabel = "consensus N-1 hold"
            Exit Function
        End If
    End If

    If ffRow > 0 Then
        If Near(v, SafeNum(ws.Cells(ffRow, col).Value)) Then
            SourceLabel = "field forecast"
            Exit Function
        End If
    End If

    If bpRow > 0 Then
        If Near(v, SafeNum(ws.Cells(bpRow, col).Value)) Then
            SourceLabel = "business plan"
            Exit Function
        End If
    End If

    SourceLabel = "manual - matches no source row"
End Function

Private Function Near(ByVal a As Double, ByVal b As Double) As Boolean
    Near = (Abs(a - b) <= MATCH_TOLERANCE)
End Function

' -------------------------------------------------------------
' Helpers
' -------------------------------------------------------------

Private Function IndexOfMonth(ByVal months As Variant, ByVal y As Long, ByVal m As Long) As Long
    IndexOfMonth = -1
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        Dim d As Date: d = CDate(months(i, 2))
        If Year(d) = y And Month(d) = m Then
            IndexOfMonth = i
            Exit Function
        End If
    Next i
End Function

Private Function SumOver(ws As Worksheet, ByVal rowNum As Long, ByVal months As Variant, _
                         ByVal iFrom As Long, ByVal iTo As Long) As Double
    Dim t As Double, i As Long
    For i = iFrom To iTo
        t = t + SafeNum(ws.Cells(rowNum, CLng(months(i, 1))).Value)
    Next i
    SumOver = t
End Function

Private Function ComputePct(ByVal mktg As Double, ByVal ff As Double, ByRef outPct As Double) As Boolean
    Dim denom As Double
    If VAR_PCT_OF_FF Then denom = ff Else denom = mktg
    If denom = 0 Then Exit Function
    outPct = (mktg - ff) / denom
    ComputePct = True
End Function

Private Function RunText(ByVal months As Variant, ByVal iFrom As Long, ByVal iTo As Long, _
                         ByVal label As String) As String
    Dim a As String, b As String
    a = Format(CDate(months(iFrom, 2)), "yy-mmm")
    b = Format(CDate(months(iTo, 2)), "yy-mmm")
    If iFrom = iTo Then
        RunText = a & " " & label
    Else
        RunText = a & " to " & b & " " & label
    End If
End Function

Private Function AppendSemi(ByVal existing As String, ByVal part As String) As String
    If Len(existing) = 0 Then
        AppendSemi = part
    Else
        AppendSemi = existing & "; " & part
    End If
End Function

Private Sub StyleHeader(rng As Range)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(217, 217, 217)
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = name
    End If
    Set EnsureSheet = w
End Function

' --- SOF lookup, local so this module needs no other pass -----

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

' Reads the SOF value for one model row and one month. Returns
' False when there is no SOF sheet or no column for that month.
Private Function TryGetSof(ByVal sofRow As Long, ByVal target As Date, _
                           ByRef outVal As Double) As Boolean
    If sofRow <= 0 Then Exit Function

    Dim wsSof As Worksheet
    On Error Resume Next
    Set wsSof = ThisWorkbook.Worksheets(SHEET_SOF)
    On Error GoTo 0
    If wsSof Is Nothing Then Exit Function

    Dim lastCol As Long
    lastCol = wsSof.Cells(1, wsSof.Columns.Count).End(xlToLeft).Column

    Dim c As Long, d As Date
    For c = 2 To lastCol
        If TryParseMonth(wsSof.Cells(1, c).Value, d) Then
            If DateSerial(Year(d), Month(d), 1) = DateSerial(Year(target), Month(target), 1) Then
                outVal = SafeNum(wsSof.Cells(sofRow, c).Value)
                TryGetSof = True
                Exit Function
            End If
        End If
    Next c
End Function
