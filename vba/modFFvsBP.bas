Attribute VB_Name = "modFFvsBP"
Option Explicit

' =============================================================
' Pass 2 - FF vs BP
'
' Seeds the Opportunity Qty Final row for all 12 months by
' choosing between Field Forecast Qty Final and Business Plan,
' gated by the accuracy score persisted in Pass 1.
'
'   FF > BP AND accuracy >= 0.80 -> write FF, light blue fill
'   FF > BP AND accuracy <  0.80 -> write BP, soft red fill (FF unreliable)
'   BP >= FF                     -> write BP, light purple fill
'   accuracy = -1 (ungradable)   -> treat as below threshold (BP)
' =============================================================

Public Sub Run_Module_2_FFvsBP()
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
        SeedModelOpportunity ws, b, months
NextBlock:
        On Error GoTo Fail
    Next b

    Exit Sub

BlockFail:
    Debug.Print "Module 2 (FFvsBP) failed on model " & modelName & ": " & Err.Description
    Resume NextBlock

Fail:
    Err.Source = "Module 2 - FFvsBP"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SeedModelOpportunity(ws As Worksheet, block As Object, months As Variant)
    Dim ffRow As Long, bpRow As Long, oppRow As Long
    ffRow = GetKeyRow(block, KF_FIELD_FCST_FINAL)
    bpRow = GetKeyRow(block, KF_BUSINESS_PLAN)
    oppRow = GetKeyRow(block, KF_OPPORTUNITY)

    If ffRow = 0 Or bpRow = 0 Or oppRow = 0 Then
        Debug.Print "Module 2: model '" & block("name") & "' missing FF/BP/Opportunity row - skipped."
        Exit Sub
    End If

    Dim accuracy As Double
    accuracy = GetModelAccuracy(CStr(block("name")))
    Dim ffReliable As Boolean
    ffReliable = (accuracy >= ACCURACY_GREEN)   ' sentinel -1 falls through as unreliable

    Dim i As Long, col As Long
    Dim ff As Double, bp As Double
    Dim oppCell As Range

    For i = LBound(months, 1) To UBound(months, 1)
        col = CLng(months(i, 1))
        ff = SafeNum(ws.Cells(ffRow, col).Value)
        bp = SafeNum(ws.Cells(bpRow, col).Value)
        Set oppCell = ws.Cells(oppRow, col)

        If ff > bp Then
            If ffReliable Then
                oppCell.Value = ff
                ApplyFillColor oppCell, CLR_FFvsBP_FF_USED
            Else
                oppCell.Value = bp
                ApplyFillColor oppCell, CLR_FFvsBP_BP_FALLBACK
            End If
        Else
            ' BP >= FF
            oppCell.Value = bp
            ApplyFillColor oppCell, CLR_FFvsBP_BP_USED
        End If
    Next i
End Sub
