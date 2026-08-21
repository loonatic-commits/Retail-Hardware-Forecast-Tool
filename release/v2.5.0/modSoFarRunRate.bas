Attribute VB_Name = "modSoFarRunRate"
Option Explicit

' =============================================================
' Pass 4 - SO FAR + Back Order Run Rate
'
' Current calendar month only. The actual run rate signal is
'   SO FAR  +  Back Order Qty
' - units already sold plus orders already received but not yet
' shipped.
'
' Writes values only. All cell fills were removed with the
' color-coding module.
' =============================================================

Public Sub Run_Module_4_SoFarRunRate()
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    Dim curCol As Long
    curCol = CurrentMonthColumn(ws)
    If curCol = 0 Then
        Debug.Print "Module 4: no column matches the current month - pass skipped entirely."
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        ApplyRunRateToModel ws, b, curCol
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 4 (SoFarRunRate) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 4 - SoFarRunRate"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ApplyRunRateToModel(ws As Worksheet, ByVal block As Object, ByVal curCol As Long)
    Dim soFarRow As Long, oppRow As Long, boRow As Long
    soFarRow = GetKeyRow(block, KF_SO_FAR)
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    boRow = GetKeyRow(block, KF_BACK_ORDER)

    If soFarRow = 0 Or oppRow = 0 Then
        Debug.Print "Module 4: model '" & block("name") & "' missing SO FAR or Opportunity row - skipped."
        Exit Sub
    End If

    Dim soFar As Double, backOrder As Double, runRate As Double
    soFar = SafeNum(ws.Cells(soFarRow, curCol).Value)
    If boRow > 0 Then backOrder = SafeNum(ws.Cells(boRow, curCol).Value)
    runRate = soFar + backOrder

    If runRate = 0 Then Exit Sub

    ws.Cells(oppRow, curCol).Value = runRate
End Sub
