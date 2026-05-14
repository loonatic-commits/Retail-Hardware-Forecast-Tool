Attribute VB_Name = "modSOFOverride"
Option Explicit

' =============================================================
' Pass 3 - SOF Override
'
' For each model, look up on the SOF sheet (exact name match).
' Match SOF column headers to Consensus column headers by date.
' For each matched month:
'   - Read current Opportunity (the Pass 2 seed)
'   - Read SOF value
'   - deviation = (opportunity - sof) / sof
'   - Overwrite Opportunity cell with SOF value
'   - Color:
'       deviation >  0.20  -> CLR_SOF_HIGH  (Pass 2 was 20%+ above SOF)
'       deviation < -0.20  -> CLR_SOF_LOW   (Pass 2 was 20%+ below SOF)
'       within +/-20%      -> preserve Pass 2 fill
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

    Dim blocks As Collection
    Set blocks = FindModelBlocks()

    Dim b As Variant, modelName As String
    For Each b In blocks
        modelName = b("name")
        On Error GoTo BlockFail
        OverrideModelFromSOF wsCon, wsSof, b, conMonths, sofMonths
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

Private Sub OverrideModelFromSOF(wsCon As Worksheet, wsSof As Worksheet, block As Object, _
                                 conMonths As Variant, sofMonths As Variant)
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

                Dim oppCell As Range
                Set oppCell = wsCon.Cells(oppRow, conCol)

                Dim sofVal As Double, oppVal As Double
                sofVal = SafeNum(wsSof.Cells(sofRow, sofCol).Value)
                oppVal = SafeNum(oppCell.Value)

                Dim priorFill As Variant
                priorFill = GetFillColor(oppCell)

                ' Overwrite first
                oppCell.Value = sofVal

                If sofVal = 0 Then
                    ' Can't compute deviation - preserve prior fill (already set)
                    ' (no-op)
                Else
                    Dim deviation As Double
                    deviation = (oppVal - sofVal) / sofVal
                    If deviation > SOF_VARIANCE Then
                        ApplyFillColor oppCell, CLR_SOF_HIGH
                    ElseIf deviation < -SOF_VARIANCE Then
                        ApplyFillColor oppCell, CLR_SOF_LOW
                    Else
                        ' Preserve Pass 2 fill - restore explicitly in case
                        ' the value assignment cleared formatting.
                        If IsNumeric(priorFill) Then
                            ApplyFillColor oppCell, CLng(priorFill)
                        End If
                    End If
                End If

                Exit For
            End If
        Next j
    Next i
End Sub

Private Function FindSofRow(wsSof As Worksheet, modelName As String) As Long
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
