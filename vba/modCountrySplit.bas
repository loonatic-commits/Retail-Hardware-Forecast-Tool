Attribute VB_Name = "modCountrySplit"
Option Explicit

' =============================================================
' modCountrySplit - standalone pass
'
' Splits the Consensus sheet's Opportunity Qty Final
' (North America total) into USA and Canada on the
' 'Country Split' sheet.
'
'   - Canada values are read AS-IS from the Country Split
'     sheet and never modified.
'   - USA values are computed and written:
'         USA = NA Opportunity - Canada
'
' Country Split layout assumed:
'   Col A  Model Desc
'   Col B  Business Segment Desc  (must include "CANADA TOTAL"
'                                  and "USA TOTAL" per model)
'   Col C  Key Figure             (free text; only one row per
'                                  segment per model is updated -
'                                  the first match found)
'   Col D+ Month headers (parsed the same way as Consensus)
'
' Updates only the current calendar month and forward, matching
' the established "don't touch past months" rule for forecasts.
'
' Run from the macro list as Run_Country_Split. Not part of
' Run_All_Passes.
' =============================================================

Public Const SHEET_COUNTRY_SPLIT As String = "Country Split"
Private Const SEGMENT_CANADA As String = "CANADA TOTAL"
Private Const SEGMENT_USA As String = "USA TOTAL"
Private Const COUNTRY_MONTH_START_COL As Long = 4

Public Sub Run_Country_Split()
    On Error GoTo Fail

    InitColors

    Dim wsCon As Worksheet
    Set wsCon = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim wsCS As Worksheet
    On Error Resume Next
    Set wsCS = ThisWorkbook.Worksheets(SHEET_COUNTRY_SPLIT)
    On Error GoTo 0
    If wsCS Is Nothing Then
        MsgBox "Missing sheet: '" & SHEET_COUNTRY_SPLIT & "'.", vbCritical, "Country Split"
        Exit Sub
    End If

    Dim conMonths As Variant
    conMonths = GetMonthColumns(wsCon)
    If IsEmpty(conMonths) Then
        MsgBox "Could not parse Consensus month columns.", vbCritical, "Country Split"
        Exit Sub
    End If

    Dim csMonths As Variant
    csMonths = GetCountrySplitMonthColumns(wsCS)
    If IsEmpty(csMonths) Then
        MsgBox "Could not parse '" & SHEET_COUNTRY_SPLIT & "' month columns " & _
               "(expected from column D onward).", vbCritical, "Country Split"
        Exit Sub
    End If

    Dim curCol As Long
    curCol = CurrentMonthColumn(wsCon)
    If curCol = 0 Then
        MsgBox "Could not find a column matching today's calendar month on '" & _
               SHEET_CONSENSUS & "'.", vbCritical, "Country Split"
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim updated As Long, missingBoth As Long, missingCA As Long, missingUS As Long, missingOpp As Long

    Dim b As Variant
    Dim modelName As String
    For Each b In blocks
        On Error GoTo BlockFail
        modelName = CStr(b("name"))

        Dim oppRow As Long
        oppRow = GetKeyRow(b, KF_OPPORTUNITY)
        If oppRow = 0 Then
            missingOpp = missingOpp + 1
            Debug.Print "CountrySplit: '" & modelName & "' has no Opportunity Qty Final row - skipped."
            GoTo NextBlock
        End If

        Dim caRow As Long, usRow As Long
        caRow = FindCountryRow(wsCS, modelName, SEGMENT_CANADA)
        usRow = FindCountryRow(wsCS, modelName, SEGMENT_USA)

        If caRow = 0 And usRow = 0 Then
            missingBoth = missingBoth + 1
            Debug.Print "CountrySplit: '" & modelName & "' not on '" & SHEET_COUNTRY_SPLIT & "' - skipped."
            GoTo NextBlock
        End If
        If caRow = 0 Then
            missingCA = missingCA + 1
            Debug.Print "CountrySplit: '" & modelName & "' missing CANADA TOTAL row - skipped."
            GoTo NextBlock
        End If
        If usRow = 0 Then
            missingUS = missingUS + 1
            Debug.Print "CountrySplit: '" & modelName & "' missing USA TOTAL row - skipped."
            GoTo NextBlock
        End If

        Dim i As Long
        For i = LBound(conMonths, 1) To UBound(conMonths, 1)
            Dim conCol As Long, conDate As Date
            conCol = CLng(conMonths(i, 1))
            conDate = CDate(conMonths(i, 2))
            If conCol >= curCol Then
                Dim csCol As Long: csCol = MatchMonthCol(csMonths, conDate)
                If csCol > 0 Then
                    Dim naVal As Double, caVal As Double
                    naVal = SafeNum(wsCon.Cells(oppRow, conCol).Value)
                    caVal = SafeNum(wsCS.Cells(caRow, csCol).Value)
                    wsCS.Cells(usRow, csCol).Value = naVal - caVal
                End If
            End If
        Next i

        updated = updated + 1
NextBlock:
        On Error GoTo Fail
    Next b

    Dim msg As String
    msg = updated & " model(s) split into US/CA on '" & SHEET_COUNTRY_SPLIT & "'."
    If missingOpp > 0 Then msg = msg & vbCrLf & missingOpp & " missing Opportunity Qty Final row."
    If missingBoth > 0 Then msg = msg & vbCrLf & missingBoth & " not on Country Split sheet."
    If missingCA > 0 Then msg = msg & vbCrLf & missingCA & " missing CANADA TOTAL row."
    If missingUS > 0 Then msg = msg & vbCrLf & missingUS & " missing USA TOTAL row."
    MsgBox msg, vbInformation, "Country Split"
    Exit Sub

BlockFail:
    Debug.Print "CountrySplit failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Run_Country_Split failed: " & Err.Description, vbCritical, "Country Split"
End Sub

Private Function FindCountryRow(ws As Worksheet, ByVal modelName As String, ByVal segment As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, "A").Value)), modelName, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, "B").Value)), segment, vbTextCompare) = 0 Then
                FindCountryRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function MatchMonthCol(csMonths As Variant, ByVal target As Date) As Long
    Dim j As Long
    For j = LBound(csMonths, 1) To UBound(csMonths, 1)
        If CDate(csMonths(j, 2)) = target Then
            MatchMonthCol = CLng(csMonths(j, 1))
            Exit Function
        End If
    Next j
End Function

Private Function GetCountrySplitMonthColumns(ws As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim tmp() As Variant
    ReDim tmp(1 To lastCol, 1 To 2)
    Dim n As Long: n = 0

    Dim c As Long, v As Variant, d As Date
    For c = COUNTRY_MONTH_START_COL To lastCol
        v = ws.Cells(1, c).Value
        If TryParseMonth(v, d) Then
            n = n + 1
            tmp(n, 1) = c
            tmp(n, 2) = DateSerial(Year(d), Month(d), 1)
        End If
    Next c

    If n = 0 Then
        GetCountrySplitMonthColumns = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        out(i, 1) = tmp(i, 1)
        out(i, 2) = tmp(i, 2)
    Next i
    GetCountrySplitMonthColumns = out
End Function
