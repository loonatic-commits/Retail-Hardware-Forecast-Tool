Attribute VB_Name = "modDeficitReport"
Option Explicit

' =============================================================
' modDeficitReport - excess inventory report (all SKUs)
'
' Month-dynamic. No hardcoded month constants - everything is
' derived from today's calendar date, so the same file works
' every month without editing.
'
' Comparison month: the most recent completed month
'                   = current calendar month - 1
'   For each model:
'     deficit = max(Consensus Fcst Qty Final N-1
'                   - Sell-In Qty (Gross), 0)
'
' Excess projection months: current month and next month.
'   pool  = Available Inventory @ current   (on-hand today)
'   pool -= Opportunity @ current
'   pool += Available Inventory @ current+1 (incoming)
'   pool -= Opportunity @ current+1
'   excess = max(pool, 0)
'
' Coverage:
'   if deficit = 0  : covered = 0, coverage_pct = 100% (nothing to cover)
'   else            : covered = min(excess, deficit)
'                     coverage_pct = covered / deficit
'
' Row coloring: 100% green, partial yellow, 0% red
'
' Sheet name includes the comparison month (e.g.
' "Apr Deficit Report", "May Deficit Report"), so prior
' months' reports are preserved across runs.
' =============================================================

Public Sub Build_Deficit_Report()
    On Error GoTo Fail

    InitColors

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then
        MsgBox "Could not parse any month columns on '" & SHEET_CONSENSUS & "'.", vbCritical
        Exit Sub
    End If

    ' --- Derive the three target dates from today ---
    Dim today As Date: today = Date
    Dim curDate As Date: curDate = DateSerial(Year(today), Month(today), 1)
    Dim prevDate As Date: prevDate = DateSerial(Year(today), Month(today) - 1, 1)
    Dim nextDate As Date: nextDate = DateSerial(Year(today), Month(today) + 1, 1)

    Dim prevCol As Long: prevCol = FindMonthCol(months, Year(prevDate), Month(prevDate))
    Dim curCol As Long: curCol = FindMonthCol(months, Year(curDate), Month(curDate))
    Dim nextCol As Long: nextCol = FindMonthCol(months, Year(nextDate), Month(nextDate))

    If prevCol = 0 Or curCol = 0 Or nextCol = 0 Then
        MsgBox "Could not find one or more of the required month columns " & _
               "(" & Format(prevDate, "yy-mmm") & ", " & _
               Format(curDate, "yy-mmm") & ", " & _
               Format(nextDate, "yy-mmm") & ") on '" & SHEET_CONSENSUS & "'.", vbCritical
        Exit Sub
    End If

    Dim prevLabel As String, curLabel As String, nextLabel As String
    prevLabel = Format(prevDate, "yy-mmm")
    curLabel = Format(curDate, "yy-mmm")
    nextLabel = Format(nextDate, "yy-mmm")

    Dim sheetName As String
    sheetName = Format(prevDate, "mmm") & " Deficit Report"

    Dim rep As Worksheet
    Set rep = EnsureReportSheet(sheetName)
    rep.Cells.Clear

    rep.Cells(1, 1).Value = "Model"
    rep.Cells(1, 2).Value = "Sell-In (Gross) " & prevLabel
    rep.Cells(1, 3).Value = "Consensus Fcst Final N-1 " & prevLabel
    rep.Cells(1, 4).Value = "Deficit"
    rep.Cells(1, 5).Value = "Avail Inv on hand (" & curLabel & ")"
    rep.Cells(1, 6).Value = "Opportunity " & curLabel
    rep.Cells(1, 7).Value = "Incoming Inv " & nextLabel
    rep.Cells(1, 8).Value = "Opportunity " & nextLabel
    rep.Cells(1, 9).Value = "Excess after " & curLabel & "+" & nextLabel
    rep.Cells(1, 10).Value = "Deficit covered"
    rep.Cells(1, 11).Value = "Coverage %"
    rep.Range("A1:K1").Font.Bold = True
    rep.Range("A1:K1").Interior.Color = RGB(217, 217, 217)
    rep.Range("A1:K1").WrapText = True

    Dim outRow As Long: outRow = 2
    Dim deficitCount As Long: deficitCount = 0

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant
    For Each b In blocks
        On Error GoTo BlockFail
        Dim modelName As String
        modelName = CStr(b("name"))

        Dim sellInRow As Long, consensusRow As Long, invRow As Long, oppRow As Long
        sellInRow = GetKeyRow(b, KF_SELL_IN)
        consensusRow = GetKeyRow(b, "Consensus Fcst Qty Final N-1")
        invRow = GetKeyRow(b, KF_AVAIL_INV)
        oppRow = GetKeyRow(b, KF_OPPORTUNITY)

        Dim sellIn As Double, consensus As Double
        sellIn = 0: consensus = 0
        If sellInRow > 0 Then sellIn = SafeNum(ws.Cells(sellInRow, prevCol).Value)
        If consensusRow > 0 Then consensus = SafeNum(ws.Cells(consensusRow, prevCol).Value)

        Dim deficit As Double
        deficit = consensus - sellIn
        If deficit < 0 Then deficit = 0
        If deficit > 0 Then deficitCount = deficitCount + 1

        Dim invCur As Double, oppCur As Double, invNext As Double, oppNext As Double
        If invRow > 0 Then invCur = SafeNum(ws.Cells(invRow, curCol).Value)
        If invRow > 0 Then invNext = SafeNum(ws.Cells(invRow, nextCol).Value)
        If oppRow > 0 Then oppCur = SafeNum(ws.Cells(oppRow, curCol).Value)
        If oppRow > 0 Then oppNext = SafeNum(ws.Cells(oppRow, nextCol).Value)

        Dim pool As Double
        pool = invCur - oppCur + invNext - oppNext

        Dim excess As Double
        excess = pool
        If excess < 0 Then excess = 0

        Dim covered As Double, pct As Double
        If deficit = 0 Then
            covered = 0
            pct = 1#
        Else
            covered = excess
            If covered > deficit Then covered = deficit
            pct = covered / deficit
        End If

        rep.Cells(outRow, 1).Value = modelName
        rep.Cells(outRow, 2).Value = sellIn
        rep.Cells(outRow, 3).Value = consensus
        rep.Cells(outRow, 4).Value = deficit
        rep.Cells(outRow, 5).Value = invCur
        rep.Cells(outRow, 6).Value = oppCur
        rep.Cells(outRow, 7).Value = invNext
        rep.Cells(outRow, 8).Value = oppNext
        rep.Cells(outRow, 9).Value = pool
        rep.Cells(outRow, 10).Value = covered
        rep.Cells(outRow, 11).Value = pct
        rep.Cells(outRow, 11).NumberFormat = "0.0%"

        Dim rowRng As Range
        Set rowRng = rep.Range(rep.Cells(outRow, 1), rep.Cells(outRow, 11))

        If pct >= 1 Then
            rowRng.Interior.Color = CLR_FIELD_GREEN
        ElseIf pct <= 0 Then
            rowRng.Interior.Color = CLR_FIELD_RED
        Else
            rowRng.Interior.Color = CLR_FIELD_YELLOW
        End If

        outRow = outRow + 1
NextBlock:
        On Error GoTo Fail
    Next b

    rep.Columns("A:K").AutoFit
    rep.Rows(1).RowHeight = 30
    rep.Activate
    rep.Range("A1").Select

    MsgBox (outRow - 2) & " model(s) reported for " & prevLabel & ". " & _
           deficitCount & " had a deficit.", vbInformation, sheetName
    Exit Sub

BlockFail:
    Debug.Print "Deficit report failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Build_Deficit_Report failed: " & Err.Description, vbCritical
End Sub

' Backwards-compatible alias so anything that calls the old name still works.
Public Sub Build_April_Deficit_Report()
    Build_Deficit_Report
End Sub

Private Function FindMonthCol(months As Variant, ByVal y As Long, ByVal m As Long) As Long
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        Dim d As Date: d = CDate(months(i, 2))
        If Year(d) = y And Month(d) = m Then
            FindMonthCol = CLng(months(i, 1))
            Exit Function
        End If
    Next i
End Function

Private Function EnsureReportSheet(ByVal name As String) As Worksheet
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = name
    End If
    Set EnsureReportSheet = w
End Function
