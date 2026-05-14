Attribute VB_Name = "modInventoryPass"
Option Explicit

' =============================================================
' Pass 5 - Inventory Pass
'
' Available Inventory (Manual) is a single starting pool that
' depletes month over month. Walk forward chronologically; the
' first month where cumulative Opportunity Qty Final demand
' exceeds the pool, and every month after it, gets red font on
' its Opportunity cell. Cell fills from earlier passes are
' preserved - this pass touches font color only.
' =============================================================

Public Sub Run_Module_5_InventoryPass()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim months As Variant
    months = GetMonthColumns(ws)
    If IsEmpty(months) Then Exit Sub

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        FlagInventoryForModel ws, b, months
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

Private Sub FlagInventoryForModel(ws As Worksheet, block As Object, months As Variant)
    Dim oppRow As Long, invRow As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    invRow = GetKeyRow(block, KF_AVAIL_INV)
    If oppRow = 0 Or invRow = 0 Then
        Debug.Print "Module 5: model '" & block("name") & "' missing Opportunity or Available Inventory row - skipped."
        Exit Sub
    End If

    Dim pool As Double
    pool = FirstNonBlankNumeric(ws, invRow, months)
    If pool <= 0 Then Exit Sub

    Dim cumulative As Double: cumulative = 0
    Dim exceeded As Boolean: exceeded = False

    Dim i As Long
    For i = LBound(months, 1) To UBound(months, 1)
        Dim col As Long
        col = CLng(months(i, 1))
        cumulative = cumulative + SafeNum(ws.Cells(oppRow, col).Value)
        If Not exceeded Then
            If cumulative > pool Then exceeded = True
        End If
        If exceeded Then
            ApplyFontColor ws.Cells(oppRow, col), CLR_INV_CONSTRAINED_TEXT
        End If
    Next i
End Sub

Private Function FirstNonBlankNumeric(ws As Worksheet, row As Long, months As Variant) As Double
    Dim i As Long, col As Long, v As Variant
    For i = LBound(months, 1) To UBound(months, 1)
        col = CLng(months(i, 1))
        v = ws.Cells(row, col).Value
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CDbl(v) <> 0 Then
                FirstNonBlankNumeric = CDbl(v)
                Exit Function
            End If
        End If
    Next i
End Function
