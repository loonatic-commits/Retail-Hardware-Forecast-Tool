Attribute VB_Name = "modCountrySplit"
Option Explicit

' =============================================================
' modCountrySplit - standalone pass
'
' Splits the Consensus sheet's Opportunity Qty Final
' (North America total) into USA and Canada on the
' 'Country Split' sheet.
'
'   - USA values are computed and written:
'         USA = NA Opportunity - Canada
'   - CANADA + USA always adds back to the NA total.
'
' CURRENT MONTH ONLY: if Canada's SO FAR is running at or above
' the Canada forecast, Canada is already tracking to beat its
' number, so an uprated figure is used instead:
'
'         Canada effective = Canada SO FAR x 1.20
'
' "at or above" means SO FAR >= Canada forecast x SOFAR_CLOSE_RATIO,
' so a SO FAR sitting just under the forecast still counts as
' close. Every other month is untouched by this rule.
'
' In an uplifted month the Canada row is RESTATED to the uprated
' figure, so the two halves still reconcile to the NA total. That
' is the only case where this pass writes to a Canada cell.
'
' Country Split layout assumed:
'   Col A  Model Desc
'   Col B  Business Segment Desc  ("CANADA TOTAL" / "USA TOTAL")
'   Col C  Key Figure             (a row whose key figure reads
'                                  "SO FAR" is treated as the
'                                  run-rate row; any other row for
'                                  that segment is the forecast row)
'   Col D+ Month headers (parsed the same way as Consensus)
'
' Updates only the current calendar month and forward, matching
' the established "don't touch past months" rule.
'
' Run from the macro list as Run_Country_Split. Not part of
' Run_All_Passes.
' =============================================================

Public Const SHEET_COUNTRY_SPLIT As String = "Country Split"
Private Const SEGMENT_CANADA As String = "CANADA TOTAL"
Private Const SEGMENT_USA As String = "USA TOTAL"
Private Const COUNTRY_MONTH_START_COL As Long = 4

' Canada SO FAR handling, current month only
Private Const SOFAR_CLOSE_RATIO As Double = 0.9   ' >= 90% of forecast counts as "close"
Private Const SOFAR_UPLIFT As Double = 0.2        ' add 20% over SO FAR
' Restate the Canada row when the uplift fires, so CANADA + USA
' always reconciles to the NA total. Set False to leave the Canada
' row at its original forecast (the halves then will not add up).
Private Const WRITE_UPLIFT_TO_CANADA As Boolean = True

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

    Dim updated As Long, uplifted As Long
    Dim missingBoth As Long, missingCA As Long, missingUS As Long, missingOpp As Long

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

        Dim caRow As Long, usRow As Long, caSoFarRow As Long
        caRow = FindCountryRow(wsCS, modelName, SEGMENT_CANADA, True)
        usRow = FindCountryRow(wsCS, modelName, SEGMENT_USA, True)
        caSoFarRow = FindCountryRow(wsCS, modelName, SEGMENT_CANADA, False)

        If caRow = 0 And usRow = 0 Then
            missingBoth = missingBoth + 1
            Debug.Print "CountrySplit: '" & modelName & "' not on '" & SHEET_COUNTRY_SPLIT & "' - skipped."
            GoTo NextBlock
        End If
        If caRow = 0 Then
            missingCA = missingCA + 1
            Debug.Print "CountrySplit: '" & modelName & "' missing CANADA TOTAL forecast row - skipped."
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
                    Dim naVal As Double, caVal As Double, caEffective As Double
                    naVal = SafeNum(wsCon.Cells(oppRow, conCol).Value)
                    caVal = SafeNum(wsCS.Cells(caRow, csCol).Value)
                    caEffective = caVal

                    ' Current month only: uprate off Canada's run rate when
                    ' SO FAR is already at or above the Canada forecast.
                    If conCol = curCol And caSoFarRow > 0 Then
                        Dim caSoFar As Double
                        caSoFar = SafeNum(wsCS.Cells(caSoFarRow, csCol).Value)
                        If caSoFar > 0 Then
                            If caSoFar >= caVal * SOFAR_CLOSE_RATIO Then
                                caEffective = caSoFar * (1 + SOFAR_UPLIFT)
                                uplifted = uplifted + 1
                                Debug.Print "CountrySplit: '" & modelName & "' current month uplift - " & _
                                            "CA fcst " & Format(caVal, "#,##0") & _
                                            ", SO FAR " & Format(caSoFar, "#,##0") & _
                                            " -> CA restated to " & Format(caEffective, "#,##0")
                                If WRITE_UPLIFT_TO_CANADA Then
                                    wsCS.Cells(caRow, csCol).Value = caEffective
                                End If
                            End If
                        End If
                    End If

                    wsCS.Cells(usRow, csCol).Value = naVal - caEffective
                End If
            End If
        Next i

        updated = updated + 1
NextBlock:
        On Error GoTo Fail
    Next b

    Dim msg As String
    msg = updated & " model(s) split into US/CA on '" & SHEET_COUNTRY_SPLIT & "'."
    If uplifted > 0 Then
        msg = msg & vbCrLf & uplifted & " had a Canada SO FAR uplift at the current month."
    End If
    If missingOpp > 0 Then msg = msg & vbCrLf & missingOpp & " missing Opportunity Qty Final row."
    If missingBoth > 0 Then msg = msg & vbCrLf & missingBoth & " not on Country Split sheet."
    If missingCA > 0 Then msg = msg & vbCrLf & missingCA & " missing CANADA TOTAL forecast row."
    If missingUS > 0 Then msg = msg & vbCrLf & missingUS & " missing USA TOTAL row."
    MsgBox msg, vbInformation, "Country Split"
    Exit Sub

BlockFail:
    Debug.Print "CountrySplit failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Run_Country_Split failed: " & Err.Description, vbCritical, "Country Split"
End Sub

' -------------------------------------------------------------
' FindCountryRow
'   Matches on model (col A) + segment (col B), then uses the
'   Key Figure (col C) to tell the forecast row apart from the
'   SO FAR row - a model now has more than one row per segment.
'
'   wantForecast = True  -> first row whose key figure is NOT SO FAR
'   wantForecast = False -> first row whose key figure IS SO FAR
' -------------------------------------------------------------
Private Function FindCountryRow(ws As Worksheet, ByVal modelName As String, _
                                ByVal segment As String, ByVal wantForecast As Boolean) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, "A").Value)), modelName, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, "B").Value)), segment, vbTextCompare) = 0 Then
                Dim isSoFar As Boolean
                isSoFar = IsSoFarKeyFigure(CStr(ws.Cells(r, "C").Value))
                If wantForecast <> isSoFar Then
                    FindCountryRow = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function IsSoFarKeyFigure(ByVal keyFigure As String) As Boolean
    IsSoFarKeyFigure = (InStr(1, Trim$(keyFigure), "SO FAR", vbTextCompare) > 0)
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
