Attribute VB_Name = "modFieldGrading"
Option Explicit

' =============================================================
' Pass 1 - Field Grading
'
' For each model:
'   1. Find the 3 most recently completed months: rightmost
'      columns where Sell-In Quantity is non-blank and non-zero.
'   2. Pair (Field Forecast Qty Final, Sell-In Quantity) over
'      those months.
'   3. Compute accuracy = 1 - mean(|F-A|/|A|). Fewer than 2 valid
'      pairs returns the sentinel -1 (ungradable).
'   4. Color the FF Qty Final label cell (Column B):
'        >= 0.80   green
'        >= 0.60   yellow
'        <  0.60   red
'        sentinel  leave uncolored (clear any prior fill)
'   5. Persist the accuracy via SetModelAccuracy for Pass 2.
' =============================================================

Public Sub Run_Module_1_FieldGrading()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    EnsureGradingCacheSheet

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        GradeModel ws, b
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 1 (FieldGrading) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 1 - FieldGrading"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GradeModel(ws As Worksheet, ByVal block As Object)
    Dim ffRow As Long, siRow As Long
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    siRow = GetKeyRow(block, KF_SELL_IN)

    If ffRow = 0 Or siRow = 0 Then
        Debug.Print "Module 1: model '" & block("name") & "' missing FF Final or Sell-In row - skipped."
        Exit Sub
    End If

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    ' Walk right-to-left collecting up to ACCURACY_LOOKBACK_MONTHS
    ' months whose Sell-In is non-blank and non-zero.
    Dim picked() As Long
    ReDim picked(1 To ACCURACY_LOOKBACK_MONTHS)
    Dim k As Long: k = 0

    Dim i As Long
    For i = UBound(months, 1) To LBound(months, 1) Step -1
        Dim col As Long
        col = CLng(months(i, 1))
        If Not IsBlankOrZero(ws.Cells(siRow, col).Value) Then
            k = k + 1
            picked(k) = col
            If k = ACCURACY_LOOKBACK_MONTHS Then Exit For
        End If
    Next i

    Dim labelCell As Range
    Set labelCell = ws.Cells(ffRow, "B")

    If k < 2 Then
        ' Ungradable - clear any prior fill, record sentinel.
        ClearFill labelCell
        SetModelAccuracy CStr(block("name")), ACCURACY_UNGRADABLE
        Exit Sub
    End If

    Dim forecasts() As Variant, actuals() As Variant
    ReDim forecasts(1 To k)
    ReDim actuals(1 To k)
    Dim j As Long
    For j = 1 To k
        forecasts(j) = ws.Cells(ffRow, picked(j)).Value
        actuals(j) = ws.Cells(siRow, picked(j)).Value
    Next j

    Dim score As Double
    score = CalcAccuracy(forecasts, actuals)

    If score = ACCURACY_UNGRADABLE Then
        ClearFill labelCell
    ElseIf score >= ACCURACY_GREEN Then
        ApplyFillColor labelCell, CLR_FIELD_GREEN
    ElseIf score >= ACCURACY_YELLOW Then
        ApplyFillColor labelCell, CLR_FIELD_YELLOW
    Else
        ApplyFillColor labelCell, CLR_FIELD_RED
    End If

    SetModelAccuracy CStr(block("name")), score
End Sub
