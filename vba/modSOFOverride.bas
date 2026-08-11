Attribute VB_Name = "modSOFOverride"
Option Explicit

' =============================================================
' Pass 3 - SOF Override
'
' Overwrites Opportunity with SOF for matched future months.
' Past months are never written.
'
' Writes values only. All cell fills were removed with the
' color-coding module.
' =============================================================

Public Sub Run_Module_3_SOFOverride()
    On Error GoTo Fail

    Dim wsCon As Worksheet, wsSof As Worksheet
    Set wsCon = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Set wsSof = ThisWorkbook.Worksheets(SHEET_SOF)

    Dim conMonths As Variant, sofMonths As Variant
    conMonths = GetMonthColumns(wsCon)
    sofMonths = GetSofMonthColumns(wsSof)
    If IsEmpty(conMonths) Or IsEmpty(sofMonths) Then Exit Sub

    Dim curCol As Long
    curCol = CurrentMonthColumn(wsCon)
    If curCol = 0 Then
        Debug.Print "Module 3: no current month column - aborted."
        Exit Sub
    End If

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        OverrideModelFromSOF wsCon, wsSof, b, conMonths, sofMonths, curCol
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 3 (SOFOverride) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 3 - SOFOverride"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub OverrideModelFromSOF(wsCon As Worksheet, wsSof As Worksheet, ByVal block As Object, _
                                 ByVal conMonths As Variant, ByVal sofMonths As Variant, _
                                 ByVal curCol As Long)
    Dim oppRow As Long
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)
    If oppRow = 0 Then
        Debug.Print "Module 3: model '" & block("name") & "' missing Opportunity row - skipped."
        Exit Sub
    End If

    Dim sofRow As Long
    sofRow = FindSofRow(wsSof, CStr(block("name")))
    If sofRow = 0 Then
        Debug.Print "Module 3: model '" & block("name") & "' not found on SOF sheet - skipped."
        Exit Sub
    End If

    Dim i As Long, j As Long
    For i = LBound(sofMonths, 1) To UBound(sofMonths, 1)
        Dim sofCol As Long, sofDate As Date
        sofCol = CLng(sofMonths(i, 1))
        sofDate = CDate(sofMonths(i, 2))

        For j = LBound(conMonths, 1) To UBound(conMonths, 1)
            If CDate(conMonths(j, 2)) = sofDate Then
                Dim conCol As Long
                conCol = CLng(conMonths(j, 1))

                ' Refuse to overwrite past months.
                If conCol < curCol Then Exit For

                wsCon.Cells(oppRow, conCol).Value = SafeNum(wsSof.Cells(sofRow, sofCol).Value)
                Exit For
            End If
        Next j
    Next i
End Sub

Private Function FindSofRow(wsSof As Worksheet, ByVal modelName As String) As Long
    Dim lastRow As Long
    lastRow = wsSof.Cells(wsSof.Rows.Count, "A").End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(wsSof.Cells(r, 1).Value)), modelName, vbTextCompare) = 0 Then
            FindSofRow = r
            Exit Function
        End If
    Next r
End Function
