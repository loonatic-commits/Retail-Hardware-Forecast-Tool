Attribute VB_Name = "modTests"
Option Explicit

' =============================================================
' modTests - synthetic workbook builder and assertions
'
' Build_Test_Workbook   creates fresh Consensus + SOF sheets
'                       with two models and hand-computed
'                       expected outputs.
'
' Run_All_Tests         runs each pass against that data and
'                       asserts Opportunity values and fills.
'
' Coverage:
'   ModelHi : high FF accuracy, FF > BP   (Pass 2 picks FF)
'             also has SOF deviation > 20% in one direction
'             and SO FAR for current month diverges meaningfully
'             from post-SOF value (Pass 4 upgrade)
'   ModelLo : low FF accuracy, FF > BP    (Pass 2 falls back to BP)
'             SOF deviation in the other direction
'             cumulative Opportunity exceeds Available Inventory
'             partway through the year
'   ModelBp : BP > FF                     (Pass 2 picks BP regardless)
' =============================================================

Private Const TEST_FAILS_COUNTER As String = "_TestFails"

Public Sub Build_Test_Workbook()
    InitColors

    Dim ws As Worksheet, sof As Worksheet
    Set ws = EnsureSheet(SHEET_CONSENSUS)
    Set sof = EnsureSheet(SHEET_SOF)
    ws.Cells.Clear
    sof.Cells.Clear

    ' --- Consensus column headers: 12 rolling months centered on today ---
    ws.Cells(1, 1).Value = "Model"
    ws.Cells(1, 2).Value = "Key Figure"
    Dim startDate As Date
    startDate = DateSerial(Year(Date), Month(Date) - 3, 1)   ' 3 months back so today is column index 6 (col F)
    Dim i As Long, c As Long
    For i = 0 To 11
        c = 3 + i
        ws.Cells(1, c).Value = DateSerial(Year(startDate), Month(startDate) + i, 1)
        ws.Cells(1, c).NumberFormat = "yy-mmm"
    Next i

    Dim curCol As Long: curCol = 3 + 3   ' the 4th of 12 = current month

    ' --- Three models, each block 13 rows ---
    WriteModelBlock ws, 2,  "ModelHi", curCol, hiAccuracy:=True, ffGreaterBp:=True, _
                     inventoryPool:=1000000, sofValues:=Array(1000, 1000, 1000), _
                     soFarCurrent:=2500
    WriteModelBlock ws, 16, "ModelLo", curCol, hiAccuracy:=False, ffGreaterBp:=True, _
                     inventoryPool:=500, sofValues:=Array(100, 100, 100), _
                     soFarCurrent:=0
    WriteModelBlock ws, 30, "ModelBp", curCol, hiAccuracy:=True, ffGreaterBp:=False, _
                     inventoryPool:=1000000, sofValues:=Array(0, 0, 0), _
                     soFarCurrent:=0

    ' --- SOF sheet: model in col A, three months (current+1, +2, +3) ---
    sof.Cells(1, 1).Value = "Model"
    sof.Cells(1, 2).Value = ws.Cells(1, curCol + 1).Value
    sof.Cells(1, 3).Value = ws.Cells(1, curCol + 2).Value
    sof.Cells(1, 4).Value = ws.Cells(1, curCol + 3).Value
    sof.Cells(1, 2).NumberFormat = "yy-mmm"
    sof.Cells(1, 3).NumberFormat = "yy-mmm"
    sof.Cells(1, 4).NumberFormat = "yy-mmm"

    sof.Cells(2, 1).Value = "ModelHi"
    sof.Cells(2, 2).Value = 1000: sof.Cells(2, 3).Value = 1000: sof.Cells(2, 4).Value = 1000
    sof.Cells(3, 1).Value = "ModelLo"
    sof.Cells(3, 2).Value = 100:  sof.Cells(3, 3).Value = 100:  sof.Cells(3, 4).Value = 100
    sof.Cells(4, 1).Value = "ModelBp"
    sof.Cells(4, 2).Value = 200:  sof.Cells(4, 3).Value = 200:  sof.Cells(4, 4).Value = 200
End Sub

Private Sub WriteModelBlock(ws As Worksheet, startRow As Long, modelName As String, _
                            curCol As Long, hiAccuracy As Boolean, ffGreaterBp As Boolean, _
                            inventoryPool As Double, sofValues As Variant, _
                            soFarCurrent As Double)
    ws.Cells(startRow + 0, 1).Value = modelName
    ws.Cells(startRow + 0, 2).Value = KF_BUSINESS_PLAN
    ws.Cells(startRow + 1, 2).Value = KF_SELL_IN
    ws.Cells(startRow + 2, 2).Value = KF_BACK_ORDER
    ws.Cells(startRow + 3, 2).Value = KF_SO_FAR
    ws.Cells(startRow + 4, 2).Value = KF_FIELD_FCST_RAW
    ws.Cells(startRow + 5, 2).Value = KF_FIELD_FCST_FINAL
    ws.Cells(startRow + 6, 2).Value = KF_MARKETING_PRIMARY
    ws.Cells(startRow + 7, 2).Value = KF_CONSENSUS
    ws.Cells(startRow + 8, 2).Value = KF_MARKETING_SECONDARY
    ws.Cells(startRow + 9, 2).Value = KF_OPPORTUNITY
    ws.Cells(startRow + 10, 2).Value = KF_MARKETING_OPP_FCST
    ws.Cells(startRow + 11, 2).Value = KF_ACTUALS_PLUS
    ws.Cells(startRow + 12, 2).Value = KF_AVAIL_INV

    Dim bpRow As Long: bpRow = startRow + 0
    Dim siRow As Long: siRow = startRow + 1
    Dim sfRow As Long: sfRow = startRow + 3
    Dim ffRow As Long: ffRow = startRow + 5
    Dim invRow As Long: invRow = startRow + 12

    Dim bp As Double, ff As Double
    If ffGreaterBp Then
        bp = 800: ff = 2000
    Else
        bp = 2000: ff = 800
    End If

    ' Past 3 months (cols curCol-3, curCol-2, curCol-1): set Sell-In + FF
    Dim i As Long
    For i = 1 To 3
        Dim past As Long: past = curCol - i
        If past >= 3 Then
            If hiAccuracy Then
                ws.Cells(siRow, past).Value = 1000
                ws.Cells(ffRow, past).Value = 1010   ' ~1% error
            Else
                ws.Cells(siRow, past).Value = 1000
                ws.Cells(ffRow, past).Value = 2000   ' 100% error
            End If
        End If
    Next i

    ' All 12 months: write BP and FF
    For i = 0 To 11
        Dim col As Long: col = 3 + i
        ws.Cells(bpRow, col).Value = bp
        ' Don't overwrite FF for past months that we already set with grading values
        If ws.Cells(ffRow, col).Value = "" Or IsEmpty(ws.Cells(ffRow, col).Value) Then
            ws.Cells(ffRow, col).Value = ff
        End If
    Next i

    ' Current-month SO FAR
    If soFarCurrent <> 0 Then
        ws.Cells(sfRow, curCol).Value = soFarCurrent
    End If

    ' Available Inventory in first month column
    ws.Cells(invRow, 3).Value = inventoryPool
End Sub

Private Function EnsureSheet(name As String) As Worksheet
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add
        w.Name = name
    End If
    Set EnsureSheet = w
End Function

' -------------------------------------------------------------
' Run_All_Tests
'   Builds the synthetic workbook, runs each pass, asserts.
'   Output goes to the Immediate window.
' -------------------------------------------------------------
Public Sub Run_All_Tests()
    Dim fails As Long: fails = 0
    Debug.Print String(60, "=")
    Debug.Print "Demand Forecasting Suite - Test Run"
    Debug.Print String(60, "=")

    Build_Test_Workbook
    InitColors

    Run_Module_1_FieldGrading
    fails = fails + AssertField_Pass1()

    Run_Module_2_FFvsBP
    fails = fails + AssertField_Pass2()

    Run_Module_3_SOFOverride
    fails = fails + AssertField_Pass3()

    Run_Module_4_SoFarRunRate
    fails = fails + AssertField_Pass4()

    Run_Module_5_InventoryPass
    fails = fails + AssertField_Pass5()

    Run_Module_6_Legend

    Debug.Print String(60, "-")
    If fails = 0 Then
        Debug.Print "ALL TESTS PASSED"
    Else
        Debug.Print fails & " TEST(S) FAILED"
    End If
End Sub

' --- Pass 1 assertions: FF Final label cell color reflects accuracy grade ---
Private Function AssertField_Pass1() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim blocks As Collection: Set blocks = FindModelBlocks()
    Dim fails As Long
    Dim b As Variant
    For Each b In blocks
        Dim ffRow As Long: ffRow = GetKeyRow(b, KF_FIELD_FCST_FINAL)
        Dim labelCell As Range: Set labelCell = ws.Cells(ffRow, "B")
        Dim modelName As String: modelName = CStr(b("name"))
        Select Case modelName
            Case "ModelHi", "ModelBp"
                fails = fails + AssertEq("Pass1." & modelName & ".labelFill", labelCell.Interior.Color, CLR_FIELD_GREEN)
            Case "ModelLo"
                fails = fails + AssertEq("Pass1." & modelName & ".labelFill", labelCell.Interior.Color, CLR_FIELD_RED)
        End Select
    Next b
    AssertField_Pass1 = fails
End Function

' --- Pass 2 assertions: Opportunity values and fills per branch ---
Private Function AssertField_Pass2() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim blocks As Collection: Set blocks = FindModelBlocks()
    Dim months As Variant: months = GetMonthColumns(ws)
    Dim fails As Long

    ' Check a future month - one well past the SOF override window so Pass 3 doesn't touch it.
    Dim col As Long: col = CLng(months(UBound(months, 1), 1))

    Dim b As Variant
    For Each b In blocks
        Dim oppRow As Long: oppRow = GetKeyRow(b, KF_OPPORTUNITY)
        Dim oppCell As Range: Set oppCell = ws.Cells(oppRow, col)
        Dim modelName As String: modelName = CStr(b("name"))
        Select Case modelName
            Case "ModelHi"
                fails = fails + AssertEq("Pass2.ModelHi.value", oppCell.Value, 2000)
                fails = fails + AssertEq("Pass2.ModelHi.fill", oppCell.Interior.Color, CLR_FFvsBP_FF_USED)
            Case "ModelLo"
                fails = fails + AssertEq("Pass2.ModelLo.value", oppCell.Value, 800)
                fails = fails + AssertEq("Pass2.ModelLo.fill", oppCell.Interior.Color, CLR_FFvsBP_BP_FALLBACK)
            Case "ModelBp"
                fails = fails + AssertEq("Pass2.ModelBp.value", oppCell.Value, 2000)
                fails = fails + AssertEq("Pass2.ModelBp.fill", oppCell.Interior.Color, CLR_FFvsBP_BP_USED)
        End Select
    Next b
    AssertField_Pass2 = fails
End Function

' --- Pass 3 assertions: SOF override values + deviation fills ---
Private Function AssertField_Pass3() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim blocks As Collection: Set blocks = FindModelBlocks()
    Dim months As Variant: months = GetMonthColumns(ws)
    Dim fails As Long

    Dim curCol As Long: curCol = CurrentMonthColumn(ws)
    Dim sof1 As Long: sof1 = curCol + 1
    Dim sof2 As Long: sof2 = curCol + 2
    Dim sof3 As Long: sof3 = curCol + 3

    Dim b As Variant
    For Each b In blocks
        Dim oppRow As Long: oppRow = GetKeyRow(b, KF_OPPORTUNITY)
        Dim modelName As String: modelName = CStr(b("name"))
        Select Case modelName
            Case "ModelHi"
                ' Pre-override Opp was 2000, SOF is 1000 -> deviation > 0.20 -> CLR_SOF_HIGH
                fails = fails + AssertEq("Pass3.ModelHi.sof1.value", ws.Cells(oppRow, sof1).Value, 1000)
                fails = fails + AssertEq("Pass3.ModelHi.sof1.fill", ws.Cells(oppRow, sof1).Interior.Color, CLR_SOF_HIGH)
                fails = fails + AssertEq("Pass3.ModelHi.sof2.value", ws.Cells(oppRow, sof2).Value, 1000)
                fails = fails + AssertEq("Pass3.ModelHi.sof3.value", ws.Cells(oppRow, sof3).Value, 1000)
            Case "ModelLo"
                ' Pre-override Opp was 800 (BP), SOF is 100 -> deviation way > 0.20 -> CLR_SOF_HIGH
                fails = fails + AssertEq("Pass3.ModelLo.sof1.value", ws.Cells(oppRow, sof1).Value, 100)
                fails = fails + AssertEq("Pass3.ModelLo.sof1.fill", ws.Cells(oppRow, sof1).Interior.Color, CLR_SOF_HIGH)
            Case "ModelBp"
                ' Pre-override Opp was 2000 (BP), SOF is 200 -> deviation > 0.20 -> CLR_SOF_HIGH
                fails = fails + AssertEq("Pass3.ModelBp.sof1.value", ws.Cells(oppRow, sof1).Value, 200)
                fails = fails + AssertEq("Pass3.ModelBp.sof1.fill", ws.Cells(oppRow, sof1).Interior.Color, CLR_SOF_HIGH)
        End Select
    Next b
    AssertField_Pass3 = fails
End Function

' --- Pass 4 assertions: SO FAR override on current month ---
Private Function AssertField_Pass4() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim blocks As Collection: Set blocks = FindModelBlocks()
    Dim fails As Long
    Dim curCol As Long: curCol = CurrentMonthColumn(ws)

    Dim b As Variant
    For Each b In blocks
        Dim oppRow As Long: oppRow = GetKeyRow(b, KF_OPPORTUNITY)
        Dim oppCell As Range: Set oppCell = ws.Cells(oppRow, curCol)
        Dim modelName As String: modelName = CStr(b("name"))
        Select Case modelName
            Case "ModelHi"
                ' Pass 2 wrote FF=2000 in curCol, Pass 3 doesn't touch curCol (SOF is curCol+1..+3),
                ' so Pass 4 sees existing=2000 vs SO FAR=2500 -> deviation=0.25 -> upgrade
                fails = fails + AssertEq("Pass4.ModelHi.value", oppCell.Value, 2500)
                fails = fails + AssertEq("Pass4.ModelHi.fill", oppCell.Interior.Color, CLR_SOFAR_UPGRADE)
            Case Else
                ' Other models have no SO FAR -> Pass 4 skips, Pass 2 value stays
        End Select
    Next b
    AssertField_Pass4 = fails
End Function

' --- Pass 5 assertions: inventory exhaustion paints red font from breach onward ---
Private Function AssertField_Pass5() As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim blocks As Collection: Set blocks = FindModelBlocks()
    Dim months As Variant: months = GetMonthColumns(ws)
    Dim fails As Long

    Dim b As Variant
    For Each b In blocks
        Dim modelName As String: modelName = CStr(b("name"))
        If modelName = "ModelLo" Then
            ' Pool = 500; Opportunity values are mostly 800 (BP), with three SOF months at 100.
            ' Even the first month alone (800) exceeds 500 -> first month and every month after gets red font.
            Dim oppRow As Long: oppRow = GetKeyRow(b, KF_OPPORTUNITY)
            Dim firstCol As Long: firstCol = CLng(months(LBound(months, 1), 1))
            Dim lastCol As Long: lastCol = CLng(months(UBound(months, 1), 1))
            fails = fails + AssertEq("Pass5.ModelLo.firstMonth.font", ws.Cells(oppRow, firstCol).Font.Color, CLR_INV_CONSTRAINED_TEXT)
            fails = fails + AssertEq("Pass5.ModelLo.lastMonth.font", ws.Cells(oppRow, lastCol).Font.Color, CLR_INV_CONSTRAINED_TEXT)
        End If
    Next b
    AssertField_Pass5 = fails
End Function

Private Function AssertEq(label As String, actual As Variant, expected As Variant) As Long
    Dim ok As Boolean
    If IsNumeric(actual) And IsNumeric(expected) Then
        ok = (CDbl(actual) = CDbl(expected))
    Else
        ok = (CStr(actual) = CStr(expected))
    End If
    If ok Then
        Debug.Print "  PASS  " & label & " = " & CStr(actual)
        AssertEq = 0
    Else
        Debug.Print "  FAIL  " & label & "  expected=" & CStr(expected) & "  actual=" & CStr(actual)
        AssertEq = 1
    End If
End Function
