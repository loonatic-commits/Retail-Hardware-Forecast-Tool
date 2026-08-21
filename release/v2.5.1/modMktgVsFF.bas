Attribute VB_Name = "modMktgVsFF"
Option Explicit

' =============================================================
' modMktgVsFF - Opportunity Qty Final vs Field Forecast
'
' Replaces modVarianceSummary. DELETE that module - if both are
' loaded you may run the old one by mistake.
'
' Entry points:
'   Build_Mktg_vs_FF          the summary sheet
'   Show_Month_Headers        diagnostic - dumps exactly what
'                             this module reads from row 1
'
' STANDALONE. Requires no pass to have run, calls no pass, and
' writes nothing to the Consensus sheet. It only reads.
'
' It also parses month headers ITSELF rather than calling
' modUtilities, so a stale copy of that module cannot affect it.
' The parser is year-aware: "27-Mar" is March 2027. VBA's own
' CDate reads "27-Mar" as 27 March of the CURRENT year, which is
' the trap this avoids.
'
' The "why" column is EVIDENCE BASED: for each month it compares
' the Opportunity value against every candidate source on the
' sheet and reports which one the number actually matches -
' availability, consensus N-1, field, business plan, run rate,
' SOF - or "manual" when it matches none. It describes what is
' in the cells, not what a pass would have written, so hand
' edits are reported as hand edits.
'
' Three tables on one sheet:
'   1  Summary by model   Aug-Mar totals, variance in units and
'                         percent, and where each model's
'                         Opportunity numbers came from
'   2  Monthly variance   units, model x month
'   3  Monthly variance   percent, model x month
' =============================================================

Private Const SHEET_OUT As String = "Mktg vs FF Variance"

' --- Fixed reporting window: Aug 2026 through Mar 2027 -------
Private Const WIN_START_YEAR As Long = 2026
Private Const WIN_START_MONTH As Long = 8      ' August
Private Const WIN_END_YEAR As Long = 2027
Private Const WIN_END_MONTH As Long = 3        ' March

' The two rows being compared.
' KF_LEFT is the demand forecast we publish - the Opportunity
' Qty Final row - and it is measured against Field Forecast
' Qty Final. Both names come from column B of Consensus.
Private Const KF_LEFT As String = "Opportunity Qty Final"

' Variance = Opportunity - FF. True divides by FF, False by Opportunity.
Private Const PCT_OF_FF As Boolean = True

' Percent variance at or beyond this is called out as material.
Private Const MATERIAL As Double = 0.15

' Two values count as the same source when they are this close.
Private Const TOLERANCE As Double = 0.5

Private Const HILITE As Long = 49407           ' RGB(255, 192, 0)
Private Const HDR_GREY As Long = 14277081      ' RGB(217, 217, 217)

' =============================================================
' Diagnostic - run this first if the window cannot be found
' =============================================================
Public Sub Show_Month_Headers()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim report As String, shown As Long
    report = "Row 1 of '" & SHEET_CONSENSUS & "', columns C onward:" & vbCrLf & vbCrLf

    Debug.Print String(60, "=")
    Debug.Print "Month header diagnostic"
    Debug.Print String(60, "=")

    Dim c As Long
    For c = 3 To lastCol
        Dim v As Variant
        v = ws.Cells(1, c).Value

        Dim raw As String
        If IsEmpty(v) Then raw = "(empty)" Else raw = CStr(v)

        Dim d As Date, ok As Boolean
        ok = ParseMonth(v, d)

        Dim line As String
        line = ColLetter(c) & "  raw=[" & raw & "]  type=" & TypeName(v) & _
               "  -> " & IIf(ok, Format(d, "mmm yyyy"), "NOT PARSED")
        Debug.Print "  " & line

        If shown < 26 Then
            report = report & line & vbCrLf
            shown = shown + 1
        End If
    Next c

    report = report & vbCrLf & "Looking for " & _
             Format(DateSerial(WIN_START_YEAR, WIN_START_MONTH, 1), "mmm yyyy") & _
             " through " & Format(DateSerial(WIN_END_YEAR, WIN_END_MONTH, 1), "mmm yyyy") & "."
    report = report & vbCrLf & "Full list is in the Immediate window (Ctrl+G)."

    MsgBox report, vbInformation, "Month header diagnostic"
End Sub

' =============================================================
' Main
' =============================================================
Public Sub Build_Mktg_vs_FF()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = MonthCols(ws)
    If IsEmpty(months) Then
        MsgBox "No month headers could be parsed from row 1 of '" & SHEET_CONSENSUS & "'." & vbCrLf & _
               "Run Show_Month_Headers to see what is there.", vbCritical, "Mktg vs FF"
        Exit Sub
    End If

    Dim iFrom As Long, iTo As Long
    iFrom = IndexOfMonth(months, WIN_START_YEAR, WIN_START_MONTH)
    iTo = IndexOfMonth(months, WIN_END_YEAR, WIN_END_MONTH)

    If iFrom < 0 Or iTo < 0 Then
        MsgBox "Could not find the reporting window on '" & SHEET_CONSENSUS & "'." & vbCrLf & vbCrLf & _
               "Looking for  " & Format(DateSerial(WIN_START_YEAR, WIN_START_MONTH, 1), "mmm yyyy") & _
               "  through  " & Format(DateSerial(WIN_END_YEAR, WIN_END_MONTH, 1), "mmm yyyy") & vbCrLf & _
               "  " & Format(DateSerial(WIN_START_YEAR, WIN_START_MONTH, 1), "mmm yyyy") & ": " & _
               IIf(iFrom < 0, "NOT FOUND", "found") & vbCrLf & _
               "  " & Format(DateSerial(WIN_END_YEAR, WIN_END_MONTH, 1), "mmm yyyy") & ": " & _
               IIf(iTo < 0, "NOT FOUND", "found") & vbCrLf & vbCrLf & _
               "Months actually found on the sheet:" & vbCrLf & _
               MonthListText(months) & vbCrLf & vbCrLf & _
               "Either edit the window constants at the top of modMktgVsFF, " & _
               "or run Show_Month_Headers for a column-by-column breakdown.", _
               vbCritical, "Mktg vs FF"
        Exit Sub
    End If
    If iTo < iFrom Then
        MsgBox "The reporting window ends before it starts. Check the window constants.", _
               vbCritical, "Mktg vs FF"
        Exit Sub
    End If

    Dim nMonths As Long
    nMonths = iTo - iFrom + 1

    ' Where the availability walk starts: current month if present,
    ' otherwise the start of the window.
    Dim poolStart As Long
    poolStart = IndexOfMonth(months, Year(Date), Month(Date))
    If poolStart < 0 Then poolStart = iFrom

    Dim rep As Worksheet
    Set rep = EnsureSheet(SHEET_OUT)
    rep.Cells.Clear

    Dim winLabel As String
    winLabel = Format(CDate(months(iFrom, 2)), "mmm") & " - " & _
               Format(CDate(months(iTo, 2)), "mmm")

    rep.Cells(1, 1).Value = "Opportunity Qty Final vs Field Forecast - " & winLabel
    rep.Cells(1, 1).Font.Bold = True
    rep.Cells(1, 1).Font.Size = 14

    rep.Cells(2, 1).Value = "Variance = Opportunity Qty Final - Field Forecast" & _
                            IIf(PCT_OF_FF, ", percent of Field Forecast", ", percent of Opportunity") & _
                            ".  Built " & Format(Now, "yyyy-mm-dd hh:nn") & "."
    rep.Cells(3, 1).Value = "Read-only snapshot of the sheet as it stands.  " & _
                            "Comparing '" & KF_LEFT & "'  against  '" & KF_FIELD_FCST_FINAL & "'"
    rep.Cells(3, 1).Font.Italic = True

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    ' ---------------- Table 1 ----------------
    Dim r As Long
    r = 5
    rep.Cells(r, 1).Value = "1. Summary by model (" & winLabel & " totals)"
    rep.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim hdrRow As Long: hdrRow = r
    rep.Cells(r, 1).Value = "Model"
    rep.Cells(r, 2).Value = "Opportunity Qty Final"
    rep.Cells(r, 3).Value = "Field Forecast"
    rep.Cells(r, 4).Value = "Variance (units)"
    rep.Cells(r, 5).Value = "Variance (%)"
    rep.Cells(r, 6).Value = "Where the Opportunity numbers came from"
    StyleHeader rep.Range(rep.Cells(r, 1), rep.Cells(r, 6))
    r = r + 1

    Dim first As Long: first = r

    Dim b As Variant, modelName As String
    For Each b In blocks
        On Error GoTo BlockFail
        modelName = CStr(b("name"))

        Dim leftRow As Long, ffRow As Long
        leftRow = GetKeyRow(b, KF_LEFT)
        ffRow = GetKeyRow(b, KF_FIELD_FCST_FINAL)

        rep.Cells(r, 1).Value = modelName

        If leftRow = 0 Or ffRow = 0 Then
            rep.Cells(r, 6).Value = "Cannot compare - missing row: " & _
                IIf(leftRow = 0, "'" & KF_LEFT & "'", "") & _
                IIf(leftRow = 0 And ffRow = 0, " and ", "") & _
                IIf(ffRow = 0, "'" & KF_FIELD_FCST_FINAL & "'", "")
            rep.Cells(r, 6).Font.Italic = True
            r = r + 1
            GoTo NextBlock
        End If

        Dim lTot As Double, fTot As Double
        lTot = SumOver(ws, leftRow, months, iFrom, iTo)
        fTot = SumOver(ws, ffRow, months, iFrom, iTo)

        Dim vUnits As Double, vPct As Double, havePct As Boolean
        vUnits = lTot - fTot
        havePct = Pct(lTot, fTot, vPct)

        rep.Cells(r, 2).Value = lTot
        rep.Cells(r, 3).Value = fTot
        rep.Cells(r, 4).Value = vUnits
        If havePct Then
            rep.Cells(r, 5).Value = vPct
            rep.Cells(r, 5).NumberFormat = "0.0%"
            If Abs(vPct) >= MATERIAL Then rep.Cells(r, 5).Interior.Color = HILITE
        Else
            rep.Cells(r, 5).Value = "n/a"
        End If

        rep.Cells(r, 6).Value = WhyText(ws, b, months, iFrom, iTo, poolStart, vPct, havePct)

        rep.Cells(r, 2).NumberFormat = "#,##0"
        rep.Cells(r, 3).NumberFormat = "#,##0"
        rep.Cells(r, 4).NumberFormat = "#,##0;[Red]-#,##0"

        r = r + 1
NextBlock:
        On Error GoTo Fail
    Next b

    Dim last As Long: last = r - 1

    If last >= first Then
        rep.Cells(r, 1).Value = "TOTAL"
        rep.Cells(r, 2).Formula = "=SUM(B" & first & ":B" & last & ")"
        rep.Cells(r, 3).Formula = "=SUM(C" & first & ":C" & last & ")"
        rep.Cells(r, 4).Formula = "=B" & r & "-C" & r
        If PCT_OF_FF Then
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

    ' ---------------- Tables 2 and 3 ----------------
    r = r + 2
    r = MonthlyTable(rep, ws, blocks, months, iFrom, iTo, r, _
                     "2. Monthly variance (units) - Opportunity minus Field", False)
    r = r + 2
    r = MonthlyTable(rep, ws, blocks, months, iFrom, iTo, r, _
                     "3. Monthly variance (%) - " & _
                     IIf(PCT_OF_FF, "of Field Forecast", "of Opportunity"), True)

    rep.Columns("A:A").ColumnWidth = 18
    rep.Columns("B:E").AutoFit
    rep.Columns("F:F").ColumnWidth = 95
    rep.Rows(hdrRow).WrapText = True
    rep.Activate
    rep.Range("A1").Select

    MsgBox "Built for " & blocks.Count & " model(s) over " & nMonths & _
           " month(s) (" & winLabel & ")." & vbCrLf & vbCrLf & _
           "Nothing on '" & SHEET_CONSENSUS & "' was changed." & vbCrLf & _
           "Sheet: '" & SHEET_OUT & "'", vbInformation, "Mktg vs FF"
    Exit Sub

BlockFail:
    Debug.Print "MktgVsFF failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Build_Mktg_vs_FF failed: " & Err.Description, vbCritical, "Mktg vs FF"
End Sub

' =============================================================
' Month parsing - local, year aware
' =============================================================

' Returns a 2-D array (1 To n, 1 To 2): column index, month date.
Private Function MonthCols(ws As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then
        MonthCols = Empty
        Exit Function
    End If

    Dim tmp() As Variant
    ReDim tmp(1 To lastCol, 1 To 2)
    Dim n As Long: n = 0

    Dim c As Long, d As Date
    For c = 3 To lastCol
        If ParseMonth(ws.Cells(1, c).Value, d) Then
            n = n + 1
            tmp(n, 1) = c
            tmp(n, 2) = DateSerial(Year(d), Month(d), 1)
        End If
    Next c

    If n = 0 Then
        MonthCols = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        out(i, 1) = tmp(i, 1)
        out(i, 2) = tmp(i, 2)
    Next i
    MonthCols = out
End Function

' Year-aware month parser.
'
' The numeric half of a two-part header is always the YEAR:
' "26-Aug" is August 2026 and "27-Mar" is March 2027. VBA's own
' IsDate/CDate treat "27-Mar" as 27 March of the current year,
' which silently moves every header past a year boundary into
' the wrong year, so text never goes near CDate until the
' two-part forms have been ruled out.
Public Function ParseMonth(ByVal v As Variant, ByRef outDate As Date) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    ' Genuine Date cell - unambiguous.
    If VarType(v) = vbDate Then
        outDate = CDate(v)
        ParseMonth = True
        Exit Function
    End If

    ' A bare date serial stored as a number.
    If IsNumeric(v) And Not VarType(v) = vbString Then
        Dim n As Double
        n = CDbl(v)
        If n >= 20000 And n <= 80000 Then
            outDate = CDate(n)
            ParseMonth = True
            Exit Function
        End If
    End If

    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr$(160), " ")      ' non-breaking space
    s = Replace(s, Chr$(9), " ")        ' tab
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    ' Two-part forms: "26-Aug", "Aug-27", "MAY 2026", "26/Aug".
    Dim sep As String
    sep = ""
    If InStr(s, "-") > 0 Then
        sep = "-"
    ElseIf InStr(s, "/") > 0 Then
        sep = "/"
    ElseIf InStr(s, " ") > 0 Then
        sep = " "
    End If

    If Len(sep) > 0 Then
        Dim parts() As String
        parts = Split(s, sep)
        If UBound(parts) - LBound(parts) = 1 Then
            Dim a As String, b As String
            a = Trim$(parts(LBound(parts)))
            b = Trim$(parts(LBound(parts) + 1))
            Dim yr As Long, mo As Long
            yr = 0: mo = 0
            If IsNumeric(a) And Not IsNumeric(b) Then
                yr = FullYear(CLng(Val(a)))
                mo = MonthNum(b)
            ElseIf IsNumeric(b) And Not IsNumeric(a) Then
                yr = FullYear(CLng(Val(b)))
                mo = MonthNum(a)
            ElseIf IsNumeric(a) And IsNumeric(b) Then
                ' "2026-08" or "08-2026"
                If CLng(Val(a)) > 12 Then
                    yr = FullYear(CLng(Val(a)))
                    mo = CLng(Val(b))
                Else
                    mo = CLng(Val(a))
                    yr = FullYear(CLng(Val(b)))
                End If
            End If
            If yr > 0 And mo >= 1 And mo <= 12 Then
                outDate = DateSerial(yr, mo, 1)
                ParseMonth = True
                Exit Function
            End If
        End If
    End If

    ' Single token that is just a month name is not enough on its
    ' own - it would need a year we do not have. Fall back to
    ' CDate for anything else, e.g. "August 2026".
    On Error Resume Next
    Dim d As Date
    d = CDate(s)
    If Err.Number = 0 Then
        On Error GoTo 0
        outDate = d
        ParseMonth = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function FullYear(ByVal y As Long) As Long
    If y < 0 Then
        FullYear = 0
    ElseIf y < 100 Then
        FullYear = 2000 + y
    Else
        FullYear = y
    End If
End Function

Private Function MonthNum(ByVal s As String) As Long
    Select Case UCase$(Left$(Trim$(s), 3))
        Case "JAN": MonthNum = 1
        Case "FEB": MonthNum = 2
        Case "MAR": MonthNum = 3
        Case "APR": MonthNum = 4
        Case "MAY": MonthNum = 5
        Case "JUN": MonthNum = 6
        Case "JUL": MonthNum = 7
        Case "AUG": MonthNum = 8
        Case "SEP": MonthNum = 9
        Case "OCT": MonthNum = 10
        Case "NOV": MonthNum = 11
        Case "DEC": MonthNum = 12
        Case Else: MonthNum = 0
    End Select
End Function

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

Private Function MonthListText(ByVal months As Variant) As String
    Dim s As String, i As Long, count As Long
    For i = LBound(months, 1) To UBound(months, 1)
        If count > 0 Then s = s & ", "
        If count > 0 And count Mod 6 = 0 Then s = s & vbCrLf
        s = s & Format(CDate(months(i, 2)), "mmm yyyy")
        count = count + 1
    Next i
    If Len(s) = 0 Then s = "(none)"
    MonthListText = s
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim s As String
    s = Cells(1, c).Address(True, False)
    ColLetter = Left$(s, InStr(s, "$") - 1)
End Function

' =============================================================
' Monthly tables
' =============================================================
Private Function MonthlyTable(rep As Worksheet, ws As Worksheet, blocks As Collection, _
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
        Dim leftRow As Long, ffRow As Long
        leftRow = GetKeyRow(b, KF_LEFT)
        ffRow = GetKeyRow(b, KF_FIELD_FCST_FINAL)

        rep.Cells(r, 1).Value = CStr(b("name"))

        If leftRow > 0 And ffRow > 0 Then
            For i = iFrom To iTo
                c = 2 + (i - iFrom)
                Dim col As Long: col = CLng(months(i, 1))
                Dim m As Double, f As Double
                m = SafeNum(ws.Cells(leftRow, col).Value)
                f = SafeNum(ws.Cells(ffRow, col).Value)

                If asPct Then
                    Dim p As Double
                    If Pct(m, f, p) Then
                        rep.Cells(r, c).Value = p
                        rep.Cells(r, c).NumberFormat = "0.0%"
                        If Abs(p) >= MATERIAL Then rep.Cells(r, c).Interior.Color = HILITE
                    End If
                Else
                    rep.Cells(r, c).Value = m - f
                    rep.Cells(r, c).NumberFormat = "#,##0;[Red]-#,##0"
                End If
            Next i
        End If
        r = r + 1
    Next b

    MonthlyTable = r
End Function

' =============================================================
' Where the Opportunity numbers came from
' =============================================================
Private Function WhyText(ws As Worksheet, ByVal block As Object, ByVal months As Variant, _
                         ByVal iFrom As Long, ByVal iTo As Long, ByVal poolStart As Long, _
                         ByVal vPct As Double, ByVal havePct As Boolean) As String
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
    sofRow = SofRowFor(CStr(block("name")))

    ' Walk the availability pool forward from poolStart.
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
                          ffRow, consN1Row, bpRow, sofRow, CDate(months(i, 2)))

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

    If havePct Then
        If Abs(vPct) >= MATERIAL Then
            s = AppendSemi(s, "opportunity sits " & Format(Abs(vPct), "0.0%") & " " & _
                              IIf(vPct > 0, "above", "below") & " field over the window, " & _
                              "past the " & Format(MATERIAL, "0%") & " bar - worth a look")
        Else
            s = AppendSemi(s, "opportunity and field within " & Format(MATERIAL, "0%") & _
                              " over the window")
        End If
    End If

    WhyText = s
End Function

' First match wins, most specific first.
Private Function SourceLabel(ws As Worksheet, ByVal oppVal As Variant, ByVal col As Long, _
                             ByVal avail As Double, ByVal invRow As Long, ByVal boRow As Long, _
                             ByVal soFarRow As Long, ByVal ffRow As Long, ByVal consN1Row As Long, _
                             ByVal bpRow As Long, ByVal sofRow As Long, _
                             ByVal monthDate As Date) As String
    If IsEmpty(oppVal) Or IsNull(oppVal) Then
        SourceLabel = "blank"
        Exit Function
    End If
    If Not IsNumeric(oppVal) Then
        If Len(Trim$(CStr(oppVal))) = 0 Then
            SourceLabel = "blank"
            Exit Function
        End If
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
        If SofValue(sofRow, monthDate, sofVal) Then
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
    Near = (Abs(a - b) <= TOLERANCE)
End Function

' =============================================================
' SOF lookup - local
' =============================================================
Private Function SofRowFor(ByVal modelName As String) As Long
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
            SofRowFor = r
            Exit Function
        End If
    Next r
End Function

Private Function SofValue(ByVal sofRow As Long, ByVal target As Date, ByRef outVal As Double) As Boolean
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
        If ParseMonth(wsSof.Cells(1, c).Value, d) Then
            If DateSerial(Year(d), Month(d), 1) = DateSerial(Year(target), Month(target), 1) Then
                outVal = SafeNum(wsSof.Cells(sofRow, c).Value)
                SofValue = True
                Exit Function
            End If
        End If
    Next c
End Function

' =============================================================
' Small helpers
' =============================================================
Private Function SumOver(ws As Worksheet, ByVal rowNum As Long, ByVal months As Variant, _
                         ByVal iFrom As Long, ByVal iTo As Long) As Double
    Dim t As Double, i As Long
    For i = iFrom To iTo
        t = t + SafeNum(ws.Cells(rowNum, CLng(months(i, 1))).Value)
    Next i
    SumOver = t
End Function

Private Function Pct(ByVal leftVal As Double, ByVal ff As Double, ByRef outPct As Double) As Boolean
    Dim denom As Double
    If PCT_OF_FF Then denom = ff Else denom = leftVal
    If denom = 0 Then Exit Function
    outPct = (leftVal - ff) / denom
    Pct = True
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
    rng.Interior.Color = HDR_GREY
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
