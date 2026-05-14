Attribute VB_Name = "modInventoryPass"
Option Explicit

' =============================================================
' Pass 5 - Inventory Pass
'
' Available Inventory (Manual) row layout:
'   - The current month's cell is the cumulative inventory we
'     currently have on hand.
'   - Future months' cells are INCOMING inventory arriving in
'     that month (added to the running pool).
'
' Walk forward starting at the current month. At each month:
'   1. Add that month's row value to the pool (current month
'      adds the starting balance; future months add incoming).
'   2. If pool < that month's Opportunity, flag the cell red.
'   3. Subtract demand from the pool (can go negative; we keep
'      tracking so later months stay flagged appropriately).
'
' Past months are never touched. This pass changes font color
' only - fills from Passes 2/3/4 are preserved.
' =============================================================

Public Sub Run_Module_5_InventoryPass()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)
    If curCol = 0 Then
        Debug.Print "Module 5: no current month column - aborted."
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        FlagInventoryForModel ws, b, months, curCol
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 5 (InventoryPass) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 5 - InventoryPass"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub FlagInventoryForModel(ws As Worksheet, ByVal block As Object, _
                                  ByVal months As Variant, ByVal curCol As Long)
    Dim oppRow As Long, invRow As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    invRow = GetKeyRow(block, KF_AVAIL_INV)
    If oppRow = 0 Or invRow = 0 Then
        Debug.Print "Module 5: model '" & block("name") & "' missing Opportunity or Available Inventory row - skipped."
        Exit Sub
    End If

    Dim pool As Double: pool = 0
    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        Dim col As Long
        col = CLng(months(i, 1))
        If col >= curCol Then
            ' Add starting balance (current month) or incoming (future months).
            pool = pool + SafeNum(ws.Cells(invRow, col).Value)

            Dim demand As Double
            demand = SafeNum(ws.Cells(oppRow, col).Value)

            If pool < demand Then
                ApplyFontColor ws.Cells(oppRow, col), CLR_INV_CONSTRAINED_TEXT
            End If

            pool = pool - demand
        End If
    Next i
End Sub
