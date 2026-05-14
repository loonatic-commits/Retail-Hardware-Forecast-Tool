Attribute VB_Name = "modSoFarRunRate"
Option Explicit

' =============================================================
' Pass 4 - SO FAR + Back Order Run Rate
'
' Current calendar month only. The "actual run rate" signal is
'   SO FAR  +  Back Order Qty
' Together they represent units already sold plus orders already
' received but not yet shipped - the freshest demand signal we
' have for the current month.
'
'   - If the combined value is blank or zero, skip the model.
'   - deviation = (run_rate - existing) / existing
'   - Overwrite Opportunity with the combined value
'   - deviation >  SO_FAR_VARIANCE  -> CLR_SOFAR_UPGRADE
'   - deviation < -SO_FAR_VARIANCE  -> CLR_SOFAR_DOWNGRADE
'   - within +/-SO_FAR_VARIANCE     -> preserve existing fill
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
    boRow = GetKeyRow(block, KF_BACK_ORDER)   ' optional - 0 if missing

    If soFarRow = 0 Or oppRow = 0 Then
        Debug.Print "Module 4: model '" & block("name") & "' missing SO FAR or Opportunity row - skipped."
        Exit Sub
    End If

    Dim soFar As Double, backOrder As Double, runRate As Double
    soFar = SafeNum(ws.Cells(soFarRow, curCol).Value)
    If boRow > 0 Then
        backOrder = SafeNum(ws.Cells(boRow, curCol).Value)
    Else
        backOrder = 0
    End If
    runRate = soFar + backOrder

    If runRate = 0 Then Exit Sub   ' nothing meaningful to write

    Dim oppCell As Range
    Set oppCell = ws.Cells(oppRow, curCol)

    Dim existing As Double
    existing = SafeNum(oppCell.Value)

    Dim priorFill As Variant
    priorFill = GetFillColor(oppCell)

    oppCell.Value = runRate

    If existing = 0 Then
        ' Can't compute deviation - leave prior fill as-is.
        Exit Sub
    End If

    Dim deviation As Double
    deviation = (runRate - existing) / existing
    If deviation > SO_FAR_VARIANCE Then
        ApplyFillColor oppCell, CLR_SOFAR_UPGRADE
    ElseIf deviation < -SO_FAR_VARIANCE Then
        ApplyFillColor oppCell, CLR_SOFAR_DOWNGRADE
    Else
        If IsNumeric(priorFill) Then ApplyFillColor oppCell, CLng(priorFill)
    End If
End Sub
