Attribute VB_Name = "modVarianceSummary"
Option Explicit

' =============================================================
' modVarianceSummary - standalone pass
'
' Builds a summary sheet comparing Marketing Demand Forecast
' against Field Forecast over a FIXED Aug - Mar window, by
' model, in units and percent, with the reasoning behind each
' model's numbers.
'
' Three tables on one sheet:
'   1  Summary by model   Aug-Mar totals, variance in units and
'                         percent, and why each model's
'                         Opportunity landed where it did
'   2  Monthly variance   units, model x month
'   3  Monthly variance   percent, model x month
'
' The window is fixed, not derived from today's date. To move
' it, edit the four window constants below.
'
' Run AFTER Run_All_Passes. The "why" column replays the same
' decision walk the selection pass used, so it only reflects
' reality once that pass has run.
'
' Run from the macro list as Build_Variance_Summary. Not part of
' Run_All_Passes.
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
' Marketing Opportunity row. Change this one string if the
' intended row is a different one.
Private Const KF_MKTG_DEMAND As String = "Marketing Forecast Qty"

' Variance = Mktg - FF. True divides by FF (the baseline being
' compared against); False divides by Mktg.
Private Const VAR_PCT_OF_FF As Boolean = True

' Percent variance at or beyond this is called out as material.
Private Const VAR_MATERIAL As Double = 0.15

Public Sub Build_Variance_Summary()
    On Error GoTo Fail

    InitColors

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

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)   ' only used to replay the decision walk

    Dim nMonths As Long
    nMonths = iTo - iFrom + 1

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
    rep.Cells(3, 1).Value = "Marketing row: '" & KF_MKTG_DEMAND & "'   |   " & _
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
    rep.Cells(r, 6).Value = "Why the Opportunity landed where it did"
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
                rep.Cells(r, 5).Interior.Color = CLR_DISSONANCE
            End If
        Else
            rep.Cells(r, 5).Value = "n/a"
        End If

        rep.Cells(r, 6).Value = WhyText(ws, b, months, curCol, iFrom, iTo, varPct, havePct)

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
    rep.Columns("F:F").ColumnWidth = 90
    rep.Rows(hdrRow).WrapText = True
    rep.Activate
    rep.Range("A1").Select

    MsgBox "Variance summary built for " & blocks.Count & " model(s) over " & _
           nMonths & " month(s) (" & windowLabel & ")." & vbCrLf & vbCrLf & _
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
                            rep.Cells(r, c).Interior.Color = CLR_DISSONANCE
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
' Why this model's Opportunity landed where it did.
'
' Replays the selection pass read-only and collapses the decision
' labels into month ranges, then adds a note on the size of the
' Marketing-vs-Field gap.
' -------------------------------------------------------------
Private Function WhyText(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                         ByVal curCol As Long, ByVal iFrom As Long, ByVal iTo As Long, _
                         ByVal varPct As Double, ByVal havePct As Boolean) As String
    Dim s As String

    If curCol > 0 Then
        Dim decisions As Object
        On Error Resume Next
        Set decisions = ComputeDecisions(ws, block, months, curCol, False)
        On Error GoTo 0

        If Not decisions Is Nothing Then
            Dim prevLabel As String, runStart As Long
            prevLabel = ""
            runStart = iFrom
            Dim i As Long
            For i = iFrom To iTo
                Dim col As Long, lbl As String
                col = CLng(months(i, 1))
                If decisions.Exists(col) Then lbl = CStr(decisions(col)) Else lbl = ""
                If lbl <> prevLabel Then
                    If Len(prevLabel) > 0 Then
                        s = AppendSemi(s, RunText(months, runStart, i - 1, prevLabel))
                    End If
                    prevLabel = lbl
                    runStart = i
                End If
            Next i
            If Len(prevLabel) > 0 Then
                s = AppendSemi(s, RunText(months, runStart, iTo, prevLabel))
            End If
        End If
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
' Helpers
' -------------------------------------------------------------

' Index into the months array for a given year/month, or -1.
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

' Variance percent with the configured denominator.
' Returns False when the denominator is zero.
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
        RunText = a & " " & LCase$(label)
    Else
        RunText = a & " to " & b & " " & LCase$(label)
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
