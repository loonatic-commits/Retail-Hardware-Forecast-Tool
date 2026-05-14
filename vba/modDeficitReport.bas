Attribute VB_Name = "modDeficitReport"
Option Explicit

' =============================================================
' modDeficitReport - one-off report
'
' For the 26-Apr column, find every model where
'     Sell-In Quantity (Gross)  <  Consensus Fcst Qty Final N-1
' compute the deficit, and write a report sheet listing the
' model, both values, the deficit, the model's available
' inventory, and a Coverable / Inventory short status.
'
' Status row coloring:
'   Coverable        green  (deficit <= available inventory)
'   Inventory short  red    (deficit > available inventory)
'
' "Available inventory" here means the value in the
' Available Inventory (Manual) row at the *current month*
' column - that is, what we have on hand today, available to
' cover the historical shortfall.
' =============================================================

Private Const REPORT_SHEET As String = "Apr Deficit Report"
Private Const TARGET_YEAR As Long = 2026
Private Const TARGET_MONTH As Long = 4   ' April

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

    ' Locate the 26-Apr column.
    Dim aprCol As Long: aprCol = 0
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        Dim d As Date: d = CDate(months(i, 2))
        If Year(d) = TARGET_YEAR And Month(d) = TARGET_MONTH Then
            aprCol = CLng(months(i, 1))
            Exit For
        End If
    Next i
    If aprCol = 0 Then
        MsgBox "Could not find a 26-Apr column on '" & SHEET_CONSENSUS & "'.", vbCritical
        Exit Sub
    End If

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)   ' may be 0; only used for inventory readout

    Dim rep As Worksheet
    Set rep = EnsureReportSheet()
    rep.Cells.Clear

    rep.Cells(1, 1).Value = "Model"
    rep.Cells(1, 2).Value = "Sell-In (Gross) 26-Apr"
    rep.Cells(1, 3).Value = "Consensus Fcst Final N-1 26-Apr"
    rep.Cells(1, 4).Value = "Deficit"
    rep.Cells(1, 5).Value = "Available Inventory (current month)"
    rep.Cells(1, 6).Value = "Status"
    rep.Range("A1:F1").Font.Bold = True
    rep.Range("A1:F1").Interior.Color = RGB(217, 217, 217)

    Dim outRow As Long: outRow = 2

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant
    For Each b In blocks
        On Error GoTo BlockFail
        Dim modelName As String
        modelName = CStr(b("name"))

        Dim sellInRow As Long, consensusRow As Long, invRow As Long
        sellInRow = GetKeyRow(b, KF_SELL_IN)                            ' "Sell-In Quantity (Gross)"
        consensusRow = GetKeyRow(b, "Consensus Fcst Qty Final N-1")
        invRow = GetKeyRow(b, KF_AVAIL_INV)

        If sellInRow > 0 And consensusRow > 0 Then
            Dim sellIn As Double, consensus As Double
            sellIn = SafeNum(ws.Cells(sellInRow, aprCol).Value)
            consensus = SafeNum(ws.Cells(consensusRow, aprCol).Value)

            If sellIn < consensus Then
                Dim deficit As Double
                deficit = consensus - sellIn

                Dim inv As Double
                If invRow > 0 And curCol > 0 Then
                    inv = SafeNum(ws.Cells(invRow, curCol).Value)
                Else
                    inv = 0
                End If

                rep.Cells(outRow, 1).Value = modelName
                rep.Cells(outRow, 2).Value = sellIn
                rep.Cells(outRow, 3).Value = consensus
                rep.Cells(outRow, 4).Value = deficit
                rep.Cells(outRow, 5).Value = inv

                Dim rowRng As Range
                Set rowRng = rep.Range(rep.Cells(outRow, 1), rep.Cells(outRow, 6))

                If deficit > inv Then
                    rep.Cells(outRow, 6).Value = "Inventory short"
                    rowRng.Interior.Color = CLR_FIELD_RED      ' soft red
                Else
                    rep.Cells(outRow, 6).Value = "Coverable"
                    rowRng.Interior.Color = CLR_FIELD_GREEN    ' soft green
                End If

                outRow = outRow + 1
            End If
        End If
NextBlock:
        On Error GoTo Fail
    Next b

    rep.Columns("A:F").AutoFit
    rep.Activate
    rep.Range("A1").Select

    MsgBox (outRow - 2) & " model(s) had a 26-Apr deficit. Report on sheet '" & REPORT_SHEET & "'.", _
           vbInformation, "Apr Deficit Report"
    Exit Sub

BlockFail:
    Debug.Print "Deficit report failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    MsgBox "Build_April_Deficit_Report failed: " & Err.Description, vbCritical
End Sub

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
