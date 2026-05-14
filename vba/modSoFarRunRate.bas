Attribute VB_Name = "modSoFarRunRate"
Option Explicit

' =============================================================
' Pass 4 - SO FAR Run Rate
'
' Current calendar month only. SO FAR is the actual in-month
' run rate and is fresher than SOF, so it replaces the
' Opportunity Qty Final for the current month.
'
'   - If SO FAR is blank or zero, skip the model.
'   - deviation = (so_far - existing) / existing
'   - Overwrite Opportunity with SO FAR value
'   - deviation >  0.05  -> CLR_SOFAR_UPGRADE
'   - deviation < -0.05  -> CLR_SOFAR_DOWNGRADE
'   - within +/-5%       -> preserve existing fill (Pass 3 color)
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
    Dim soFarRow As Long, oppRow As Long
    soFarRow = GetKeyRow(block, KF_SO_FAR)
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    If soFarRow = 0 Or oppRow = 0 Then
        Debug.Print "Module 4: model '" & block("name") & "' missing SO FAR or Opportunity row - skipped."
        Exit Sub
    End If

    Dim soFarVal As Variant
    soFarVal = ws.Cells(soFarRow, curCol).Value
    If IsBlankOrZero(soFarVal) Then Exit Sub

    Dim oppCell As Range
    Set oppCell = ws.Cells(oppRow, curCol)

    Dim existing As Double, soFar As Double
    existing = SafeNum(oppCell.Value)
    soFar = CDbl(soFarVal)

    Dim priorFill As Variant
    priorFill = GetFillColor(oppCell)

    oppCell.Value = soFar

    If existing = 0 Then
        ' Can't compute deviation - leave prior fill as-is.
        Exit Sub
    End If

    Dim deviation As Double
    deviation = (soFar - existing) / existing
    If deviation > SO_FAR_VARIANCE Then
        ApplyFillColor oppCell, CLR_SOFAR_UPGRADE
    ElseIf deviation < -SO_FAR_VARIANCE Then
        ApplyFillColor oppCell, CLR_SOFAR_DOWNGRADE
    Else
        If IsNumeric(priorFill) Then ApplyFillColor oppCell, CLng(priorFill)
    End If
End Sub
