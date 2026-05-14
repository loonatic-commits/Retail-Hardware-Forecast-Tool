Attribute VB_Name = "modDeficitReport"
Option Explicit

' =============================================================
' modDeficitReport - excess inventory report (all SKUs)
'
' For every model:
'   deficit = max(Consensus Fcst Qty Final N-1 - Sell-In Qty Gross, 0)
'             at 26-Apr
'
' Then project remaining inventory after fulfilling 26-May and
' 26-Jun Opportunity Qty Final values, and report how much of
' the Apr deficit (if any) the excess can cover.
'
' Excess inventory model (matches the agreed inventory logic):
'   pool  = Available Inventory @ 26-May (on-hand today)
'   pool -= Opportunity @ 26-May
'   pool += Available Inventory @ 26-Jun  (incoming)
'   pool -= Opportunity @ 26-Jun
'   excess = max(pool, 0)
'
' Coverage:
'   if deficit = 0  : covered = 0, coverage_pct = 100% (nothing to cover)
'   else            : covered = min(excess, deficit)
'                     coverage_pct = covered / deficit
'
' Row coloring:
'   100% (incl. no deficit) green
'   partial                 yellow
'   0%                      red
' =============================================================

Private Const REPORT_SHEET As String = "Apr Deficit Report"

Public Sub Build_April_Deficit_Report()
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

    Dim aprCol As Long: aprCol = FindMonthCol(months, 2026, 4)
    Dim mayCol As Long: mayCol = FindMonthCol(months, 2026, 5)
    Dim junCol As Long: junCol = FindMonthCol(months, 2026, 6)
    If aprCol = 0 Or mayCol = 0 Or junCol = 0 Then
        MsgBox "Could not find one or more of 26-Apr / 26-May / 26-Jun columns.", vbCritical
        Exit Sub
    End If

    Dim rep As Worksheet
    Set rep = EnsureReportSheet()
    rep.Cells.Clear

    rep.Cells(1, 1).Value = "Model"
    rep.Cells(1, 2).Value = "Sell-In (Gross) 26-Apr"
    rep.Cells(1, 3).Value = "Consensus Fcst Final N-1 26-Apr"
    rep.Cells(1, 4).Value = "Deficit"
    rep.Cells(1, 5).Value = "Avail Inv on hand (26-May)"
    rep.Cells(1, 6).Value = "Opportunity 26-May"
    rep.Cells(1, 7).Value = "Incoming Inv 26-Jun"
    rep.Cells(1, 8).Value = "Opportunity 26-Jun"
    rep.Cells(1, 9).Value = "Excess after May+Jun"
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
        If sellInRow > 0 Then sellIn = SafeNum(ws.Cells(sellInRow, aprCol).Value)
        If consensusRow > 0 Then consensus = SafeNum(ws.Cells(consensusRow, aprCol).Value)

        Dim deficit As Double
        deficit = consensus - sellIn
        If deficit < 0 Then deficit = 0
        If deficit > 0 Then deficitCount = deficitCount + 1

        Dim invMay As Double, oppMay As Double, invJun As Double, oppJun As Double
        If invRow > 0 Then invMay = SafeNum(ws.Cells(invRow, mayCol).Value)
        If invRow > 0 Then invJun = SafeNum(ws.Cells(invRow, junCol).Value)
        If oppRow > 0 Then oppMay = SafeNum(ws.Cells(oppRow, mayCol).Value)
        If oppRow > 0 Then oppJun = SafeNum(ws.Cells(oppRow, junCol).Value)

        Dim pool As Double
        pool = invMay - oppMay + invJun - oppJun

        Dim excess As Double
        excess = pool
        If excess < 0 Then excess = 0

        Dim covered As Double, pct As Double
        If deficit = 0 Then
            covered = 0
            pct = 1#                ' no deficit -> fully covered by definition
        Else
            covered = excess
            If covered > deficit Then covered = deficit
            pct = covered / deficit
        End If

        rep.Cells(outRow, 1).Value = modelName
        rep.Cells(outRow, 2).Value = sellIn
        rep.Cells(outRow, 3).Value = consensus
        rep.Cells(outRow, 4).Value = deficit
        rep.Cells(outRow, 5).Value = invMay
        rep.Cells(outRow, 6).Value = oppMay
        rep.Cells(outRow, 7).Value = invJun
        rep.Cells(outRow, 8).Value = oppJun
        rep.Cells(outRow, 9).Value = pool        ' signed projection
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

    MsgBox (outRow - 2) & " model(s) reported. " & deficitCount & " had a 26-Apr deficit.", _
           vbInformation, "Apr Deficit Report"
    Exit Sub

BlockFail:
    Debug.Print "Deficit report failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Build_April_Deficit_Report failed: " & Err.Description, vbCritical
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

Private Function EnsureReportSheet() As Worksheet
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = REPORT_SHEET
    End If
    Set EnsureReportSheet = w
End Function
