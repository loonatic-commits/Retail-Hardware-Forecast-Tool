Attribute VB_Name = "modUtilities"
Option Explicit

' =============================================================
' modUtilities - shared helpers for all passes
' =============================================================

' -------------------------------------------------------------
' FindModelBlocks
'   Scans Column A of the Consensus sheet for unique non-blank
'   model names. Each model owns the rows from its name row
'   through the row immediately before the next model name.
'
'   Returns a Collection of Dictionaries. Each dictionary has:
'     name      : model name (string)
'     startRow  : first row of the block (the model name row)
'     endRow    : last row of the block
'     keyRows   : Dictionary keyed by KF label -> row number
' -------------------------------------------------------------
Public Function FindModelBlocks() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    Dim currentName As String
    Dim currentStart As Long
    currentName = ""
    currentStart = 0

    For r = 2 To lastRow
        Dim aVal As String
        aVal = Trim$(CStr(ws.Cells(r, "A").Value))
        If Len(aVal) > 0 Then
            If Len(currentName) > 0 Then
                result.Add BuildBlock(ws, currentName, currentStart, r - 1)
            End If
            currentName = aVal
            currentStart = r
        End If
    Next r

    If Len(currentName) > 0 Then
        result.Add BuildBlock(ws, currentName, currentStart, lastRow)
    End If

    Set FindModelBlocks = result
End Function

Private Function BuildBlock(ws As Worksheet, modelName As String, startRow As Long, endRow As Long) As Object
    Dim block As Object
    Set block = CreateObject("Scripting.Dictionary")
    block("name") = modelName
    block("startRow") = startRow
    block("endRow") = endRow

    Dim keyRows As Object
    Set keyRows = CreateObject("Scripting.Dictionary")
    keyRows.CompareMode = 1   ' TextCompare - case insensitive

    Dim r As Long, label As String
    For r = startRow To endRow
        label = Trim$(CStr(ws.Cells(r, "B").Value))
        If Len(label) > 0 Then
            ' Last occurrence wins if duplicates appear (e.g. two "Marketing" rows)
            ' so callers should use the more specific labels when both exist.
            If Not keyRows.Exists(label) Then
                keyRows.Add label, r
            Else
                ' Allow second-occurrence lookup under a suffixed key.
                keyRows(label & "#2") = r
            End If
        End If
    Next r

    Set block("keyRows") = keyRows
    Set BuildBlock = block
End Function

' -------------------------------------------------------------
' GetKeyRow
'   Convenience accessor. Returns 0 if the key figure isn't found.
' -------------------------------------------------------------
Public Function GetKeyRow(block As Object, keyFigure As String) As Long
    Dim keyRows As Object
    Set keyRows = block("keyRows")
    If keyRows.Exists(keyFigure) Then
        GetKeyRow = keyRows(keyFigure)
    Else
        GetKeyRow = 0
    End If
End Function

' -------------------------------------------------------------
' GetMonthColumns
'   Parses row-1 headers as dates from Column C onward.
'   Returns a 2-D array sized (1 To n, 1 To 2):
'     (i, 1) = column index
'     (i, 2) = month date (first of the month)
' -------------------------------------------------------------
Public Function GetMonthColumns(ws As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim tmp() As Variant
    ReDim tmp(1 To lastCol, 1 To 2)
    Dim n As Long: n = 0

    Dim c As Long, v As Variant, d As Date
    For c = 3 To lastCol
        v = ws.Cells(1, c).Value
        If TryParseMonth(v, d) Then
            n = n + 1
            tmp(n, 1) = c
            tmp(n, 2) = DateSerial(Year(d), Month(d), 1)
        End If
    Next c

    If n = 0 Then
        GetMonthColumns = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        out(i, 1) = tmp(i, 1)
        out(i, 2) = tmp(i, 2)
    Next i
    GetMonthColumns = out
End Function

' -------------------------------------------------------------
' TryParseMonth
'   Attempts to interpret a header value as a calendar month.
'   Accepts true Date cells and strings like "26-Mar", "Mar 26",
'   "MAR 2026", "2026-03", etc.
' -------------------------------------------------------------
Public Function TryParseMonth(v As Variant, ByRef outDate As Date) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsDate(v) Then
        outDate = CDate(v)
        TryParseMonth = True
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' Try direct cast
    On Error Resume Next
    Dim d As Date
    d = CDate(s)
    If Err.Number = 0 Then
        On Error GoTo 0
        outDate = d
        TryParseMonth = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' Parse forms like "26-Mar" or "Mar-26" or "MAR 2026"
    Dim parts() As String
    Dim sep As String
    sep = " "
    If InStr(s, "-") > 0 Then sep = "-"
    parts = Split(s, sep)
    If UBound(parts) - LBound(parts) = 1 Then
        Dim a As String, b As String
        a = Trim$(parts(LBound(parts)))
        b = Trim$(parts(LBound(parts) + 1))
        Dim yr As Long, mo As Long
        If IsNumeric(a) And Not IsNumeric(b) Then
            yr = NormalizeYear(CLng(a))
            mo = MonthFromName(b)
        ElseIf IsNumeric(b) And Not IsNumeric(a) Then
            yr = NormalizeYear(CLng(b))
            mo = MonthFromName(a)
        End If
        If yr > 0 And mo > 0 Then
            outDate = DateSerial(yr, mo, 1)
            TryParseMonth = True
        End If
    End If
End Function

Private Function NormalizeYear(y As Long) As Long
    If y < 100 Then
        NormalizeYear = 2000 + y
    Else
        NormalizeYear = y
    End If
End Function

Private Function MonthFromName(s As String) As Long
    Dim u As String
    u = UCase$(Left$(s, 3))
    Select Case u
        Case "JAN": MonthFromName = 1
        Case "FEB": MonthFromName = 2
        Case "MAR": MonthFromName = 3
        Case "APR": MonthFromName = 4
        Case "MAY": MonthFromName = 5
        Case "JUN": MonthFromName = 6
        Case "JUL": MonthFromName = 7
        Case "AUG": MonthFromName = 8
        Case "SEP": MonthFromName = 9
        Case "OCT": MonthFromName = 10
        Case "NOV": MonthFromName = 11
        Case "DEC": MonthFromName = 12
        Case Else: MonthFromName = 0
    End Select
End Function

' -------------------------------------------------------------
' CurrentMonthColumn
'   Returns the column index whose parsed header matches the
'   current calendar month on the given sheet. 0 if not found.
' -------------------------------------------------------------
Public Function CurrentMonthColumn(ws As Worksheet) As Long
    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Function

    Dim target As Date
    target = DateSerial(Year(Date), Month(Date), 1)

    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        If CDate(months(i, 2)) = target Then
            CurrentMonthColumn = CLng(months(i, 1))
            Exit Function
        End If
    Next i
End Function

' -------------------------------------------------------------
' CalcAccuracy
'   Returns 1 - mean(abs(forecast - actual)/actual) over paired
'   entries where actual is non-zero numeric and forecast is
'   numeric. Returns ACCURACY_UNGRADABLE (-1) when fewer than 2
'   valid pairs exist.
' -------------------------------------------------------------
Public Function CalcAccuracy(forecasts As Variant, actuals As Variant) As Double
    Dim lb As Long, ub As Long
    lb = LBound(forecasts)
    ub = UBound(forecasts)
    If LBound(actuals) <> lb Or UBound(actuals) <> ub Then
        CalcAccuracy = ACCURACY_UNGRADABLE
        Exit Function
    End If

    Dim i As Long, n As Long
    Dim sumErr As Double
    sumErr = 0
    n = 0

    For i = lb To ub
        If IsNumeric(forecasts(i)) And IsNumeric(actuals(i)) Then
            Dim a As Double, f As Double
            a = CDbl(actuals(i))
            f = CDbl(forecasts(i))
            If a <> 0 Then
                sumErr = sumErr + Abs(f - a) / Abs(a)
                n = n + 1
            End If
        End If
    Next i

    If n < 2 Then
        CalcAccuracy = ACCURACY_UNGRADABLE
    Else
        CalcAccuracy = 1# - (sumErr / n)
    End If
End Function

' -------------------------------------------------------------
' ApplyFillColor / ApplyFontColor
'   Single-cell wrappers that null-check the range first.
' -------------------------------------------------------------
Public Sub ApplyFillColor(rng As Range, colorVal As Long)
    If rng Is Nothing Then Exit Sub
    rng.Interior.Color = colorVal
End Sub

Public Sub ApplyFontColor(rng As Range, colorVal As Long)
    If rng Is Nothing Then Exit Sub
    rng.Font.Color = colorVal
End Sub

Public Function GetFillColor(rng As Range) As Variant
    If rng Is Nothing Then
        GetFillColor = Empty
        Exit Function
    End If
    GetFillColor = rng.Interior.Color
End Function

Public Sub ClearFill(rng As Range)
    If rng Is Nothing Then Exit Sub
    rng.Interior.Pattern = xlNone
End Sub

' -------------------------------------------------------------
' Grading cache - hidden sheet keyed by model name
'   Column A: model name
'   Column B: accuracy score
' -------------------------------------------------------------
Public Sub EnsureGradingCacheSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_GRADING_CACHE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_GRADING_CACHE
        ws.Cells(1, 1).Value = "Model"
        ws.Cells(1, 2).Value = "Accuracy"
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

Public Sub SetModelAccuracy(modelName As String, value As Double)
    EnsureGradingCacheSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_GRADING_CACHE)
    Dim r As Long
    r = FindCacheRow(ws, modelName)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = modelName
    End If
    ws.Cells(r, 2).Value = value
End Sub

Public Function GetModelAccuracy(modelName As String) As Double
    EnsureGradingCacheSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_GRADING_CACHE)
    Dim r As Long
    r = FindCacheRow(ws, modelName)
    If r = 0 Then
        GetModelAccuracy = ACCURACY_UNGRADABLE
        Exit Function
    End If
    Dim v As Variant
    v = ws.Cells(r, 2).Value
    If IsNumeric(v) Then
        GetModelAccuracy = CDbl(v)
    Else
        GetModelAccuracy = ACCURACY_UNGRADABLE
    End If
End Function

Private Function FindCacheRow(ws As Worksheet, modelName As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).Value), modelName, vbTextCompare) = 0 Then
            FindCacheRow = r
            Exit Function
        End If
    Next r
End Function

' -------------------------------------------------------------
' SafeNum
'   Returns CDbl(v) if numeric, else returnDefault.
' -------------------------------------------------------------
Public Function SafeNum(v As Variant, Optional returnDefault As Double = 0) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = returnDefault
    End If
End Function

Public Function IsBlankOrZero(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsBlankOrZero = True
        Exit Function
    End If
    If Not IsNumeric(v) Then
        IsBlankOrZero = (Len(Trim$(CStr(v))) = 0)
        Exit Function
    End If
    IsBlankOrZero = (CDbl(v) = 0)
End Function
