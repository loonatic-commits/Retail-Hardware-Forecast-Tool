Attribute VB_Name = "modTests"
Option Explicit

' =============================================================
' modTests - synthetic workbook builder and assertions
'
' Build_Test_Workbook  creates fresh Consensus + SOF sheets with
'                      four models and hand-computed expected
'                      outputs for the current pass pipeline.
'
' Run_All_Tests        runs every pass against that data and
'                      asserts values, fills, fonts and notes.
'
' WARNING: these routines CLEAR the Consensus and SOF sheets.
' Run them in a scratch workbook, never on live data.
'
' Layout built by Build_Test_Workbook:
'   15 month columns, C..Q. Three completed months (C,D,E) so
'   Pass 1 has something to grade, then the current month at F
'   and eleven forward months G..Q. That puts N+4 at column J,
'   which the constrained path needs.
'
'   Month index:  F=N  G=N+1  H=N+2  I=N+3  J=N+4  K=N+5 ...
'
' Models and the branches they cover:
'   MDL-CONSTR      constrained. N..N+3 = availability,
'                   N+4 = availability + backorder (Field lower)
'   MDL-CONSTR-FLD  constrained, but Field at N+4 exceeds
'                   availability + backorder, so Field wins
'   MDL-NORMAL      never constrained. N..N+2 left to Passes
'                   2-4, N+3 onward = Consensus N-1
'   MDL-FLAG        never constrained, low field accuracy, and
'                   variances that trip the dissonance tests
' =============================================================

Private Const R_BP As Long = 0
Private Const R_SELLIN As Long = 1
Private Const R_BO As Long = 2
Private Const R_SOFAR As Long = 3
Private Const R_FF As Long = 4
Private Const R_FF_N1 As Long = 5
Private Const R_MKT_N1 As Long = 6
Private Const R_CONS_N1 As Long = 7
Private Const R_MKT As Long = 8
Private Const R_OPP As Long = 9
Private Const R_MKT_OPP As Long = 10
Private Const R_ACTUALS As Long = 11
Private Const R_INV As Long = 12

Private Const BLOCK_ROWS As Long = 13
Private Const N_MONTHS As Long = 15
Private Const CUR_COL As Long = 6      ' column F - the current month
Private Const FIRST_COL As Long = 3    ' column C

Private mFails As Long

' =============================================================
' Builder
' =============================================================

Public Sub Build_Test_Workbook()
    InitColors

    Dim ws As Worksheet, sof As Worksheet
    Set ws = EnsureSheet(SHEET_CONSENSUS)
    Set sof = EnsureSheet(SHEET_SOF)
    ws.Cells.Clear
    sof.Cells.Clear

    ' --- Month headers: current month - 3, then 15 months ---
    ws.Cells(1, 1).Value = "Model Desc"
    ws.Cells(1, 2).Value = "Key Figure"

    Dim startDate As Date
    startDate = DateSerial(Year(Date), Month(Date) - 3, 1)

    Dim i As Long
    For i = 0 To N_MONTHS - 1
        With ws.Cells(1, FIRST_COL + i)
            .Value = DateSerial(Year(startDate), Month(startDate) + i, 1)
            .NumberFormat = "yy-mmm"
        End With
    Next i

    ' --- Model blocks ---
    Build_Constrained ws, 2, "MDL-CONSTR", fieldAtN4:=200
    Build_Constrained ws, 2 + BLOCK_ROWS, "MDL-CONSTR-FLD", fieldAtN4:=5000
    Build_Normal ws, 2 + BLOCK_ROWS * 2
    Build_Flag ws, 2 + BLOCK_ROWS * 3

    ' --- SOF sheet: three months starting at the current month ---
    sof.Cells(1, 1).Value = "Model"
    Dim k As Long
    For k = 0 To 2
        With sof.Cells(1, 2 + k)
            .Value = ws.Cells(1, CUR_COL + k).Value
            .NumberFormat = "yy-mmm"
        End With
    Next k

    WriteSofRow sof, 2, "MDL-CONSTR", 1000, 1000, 1000
    WriteSofRow sof, 3, "MDL-CONSTR-FLD", 1000, 1000, 1000
    WriteSofRow sof, 4, "MDL-NORMAL", 1000, 1100, 1200
    WriteSofRow sof, 5, "MDL-FLAG", 1500, 1000, 1000
End Sub

' -------------------------------------------------------------
' Constrained model
'   Inventory: 100 on hand at N, nothing incoming after.
'   Demand signal max(Consensus N-1 1000, Field N-1 500) = 1000,
'   so every month is constrained.
'
'   Expected after the selection pass:
'     N   = 100   (availability)
'     N+1 = 0     (pool exhausted)
'     N+2 = 0
'     N+3 = 0
'     N+4 = 0 + backorder 250, unless Field beats it
'     N+5 on = Consensus N-1 = 1000 (normal path takes over)
' -------------------------------------------------------------
Private Sub Build_Constrained(ws As Worksheet, ByVal startRow As Long, _
                              ByVal modelName As String, ByVal fieldAtN4 As Double)
    WriteLabels ws, startRow, modelName
    WriteGradingHistory ws, startRow, accurate:=True

    Dim i As Long, col As Long
    For i = 0 To N_MONTHS - 1
        col = FIRST_COL + i
        If col >= CUR_COL Then
            ws.Cells(startRow + R_BP, col).Value = 900
            ws.Cells(startRow + R_FF, col).Value = 200
            ws.Cells(startRow + R_FF_N1, col).Value = 500
            ws.Cells(startRow + R_CONS_N1, col).Value = 1000
        End If
    Next i

    ' Field at N+4 drives the backorder-vs-Field branch
    ws.Cells(startRow + R_FF, CUR_COL + 4).Value = fieldAtN4
    ws.Cells(startRow + R_BO, CUR_COL + 4).Value = 250

    ' On hand today only - nothing incoming
    ws.Cells(startRow + R_INV, CUR_COL).Value = 100
End Sub

' -------------------------------------------------------------
' Normal model
'   Inventory is enormous, so it is never constrained.
'
'   Expected:
'     N   = SO FAR 300 + backorder 50 = 350   (Pass 4)
'     N+1 = SOF 1100                          (Pass 3)
'     N+2 = SOF 1200                          (Pass 3)
'     N+3 on = Consensus N-1 = 900            (selection)
' -------------------------------------------------------------
Private Sub Build_Normal(ws As Worksheet, ByVal startRow As Long)
    WriteLabels ws, startRow, "MDL-NORMAL"
    WriteGradingHistory ws, startRow, accurate:=True

    Dim i As Long, col As Long
    For i = 0 To N_MONTHS - 1
        col = FIRST_COL + i
        If col >= CUR_COL Then
            ws.Cells(startRow + R_BP, col).Value = 700
            ws.Cells(startRow + R_FF, col).Value = 750
            ws.Cells(startRow + R_FF_N1, col).Value = 800
            ws.Cells(startRow + R_CONS_N1, col).Value = 900
        End If
    Next i

    ws.Cells(startRow + R_SOFAR, CUR_COL).Value = 300
    ws.Cells(startRow + R_BO, CUR_COL).Value = 50
    ws.Cells(startRow + R_INV, CUR_COL).Value = 1000000
End Sub

' -------------------------------------------------------------
' Flag model
'   Never constrained. Field accuracy is poor, so Pass 1 grades
'   it red. Variances are tuned to trip the dissonance tests:
'
'     N     SOF 1500 vs Consensus N-1 1000 = +50%  -> SOF flag
'     N+5   Field 1300 vs Field N-1 1000   = +30%  -> field flag
'           Field 1300 vs Consensus N-1    = +30%  -> consensus flag
' -------------------------------------------------------------
Private Sub Build_Flag(ws As Worksheet, ByVal startRow As Long)
    WriteLabels ws, startRow, "MDL-FLAG"
    WriteGradingHistory ws, startRow, accurate:=False

    Dim i As Long, col As Long
    For i = 0 To N_MONTHS - 1
        col = FIRST_COL + i
        If col >= CUR_COL Then
            ws.Cells(startRow + R_BP, col).Value = 1000
            ws.Cells(startRow + R_FF, col).Value = 1000
            ws.Cells(startRow + R_FF_N1, col).Value = 1000
            ws.Cells(startRow + R_CONS_N1, col).Value = 1000
        End If
    Next i

    ws.Cells(startRow + R_FF, CUR_COL + 5).Value = 1300
    ws.Cells(startRow + R_INV, CUR_COL).Value = 1000000
End Sub

' -------------------------------------------------------------
' Shared builder helpers
' -------------------------------------------------------------

Private Sub WriteLabels(ws As Worksheet, ByVal startRow As Long, ByVal modelName As String)
    Dim labels As Variant
    labels = Array(KF_BUSINESS_PLAN, KF_SELL_IN, KF_BACK_ORDER, KF_SO_FAR, _
                   KF_FIELD_FCST_FINAL, KF_FIELD_FCST_N1, "Marketing Fcst Qty N-1", _
                   KF_CONSENSUS_N1, KF_MARKETING_SECONDARY, KF_OPPORTUNITY, _
                   KF_MARKETING_OPP_FCST, KF_ACTUALS_PLUS, KF_AVAIL_INV)

    Dim r As Long
    For r = 0 To BLOCK_ROWS - 1
        ' The live sheet repeats the model name on every row of a block.
        ws.Cells(startRow + r, 1).Value = modelName
        ws.Cells(startRow + r, 2).Value = labels(r)
    Next r
End Sub

' Three completed months of Sell-In and Field Forecast so Pass 1
' has a gradable history. accurate:=True lands around 99%,
' accurate:=False lands at 0%.
Private Sub WriteGradingHistory(ws As Worksheet, ByVal startRow As Long, ByVal accurate As Boolean)
    Dim i As Long, col As Long
    For i = 0 To 2
        col = FIRST_COL + i
        ws.Cells(startRow + R_SELLIN, col).Value = 1000
        If accurate Then
            ws.Cells(startRow + R_FF, col).Value = 1010
        Else
            ws.Cells(startRow + R_FF, col).Value = 2000
        End If
    Next i
End Sub

Private Sub WriteSofRow(sof As Worksheet, ByVal r As Long, ByVal modelName As String, _
                        ByVal v1 As Double, ByVal v2 As Double, ByVal v3 As Double)
    sof.Cells(r, 1).Value = modelName
    sof.Cells(r, 2).Value = v1
    sof.Cells(r, 3).Value = v2
    sof.Cells(r, 4).Value = v3
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
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

' =============================================================
' Runner
' =============================================================

Public Sub Run_All_Tests()
    mFails = 0

    Debug.Print String(64, "=")
    Debug.Print "Demand Forecasting Suite - Test Run - " & Format(Date, "yyyy-mm-dd")
    Debug.Print String(64, "=")

    Build_Test_Workbook
    InitColors

    Run_Module_1_FieldGrading
    Check_Pass1

    Run_Module_2_FFvsBP
    Run_Module_3_SOFOverride
    Run_Module_4_SoFarRunRate
    Check_EarlyMonths

    Run_Forecast_Selection
    Check_Selection

    Run_Module_5_InventoryPass
    Check_Inventory

    Run_Info_Block
    Check_Notes
    Check_DissonanceFills

    Debug.Print String(64, "-")
    If mFails = 0 Then
        Debug.Print "ALL TESTS PASSED"
    Else
        Debug.Print mFails & " TEST(S) FAILED"
    End If
End Sub

' --- Pass 1: accuracy grade on the FF Qty Final label cell ----
Private Sub Check_Pass1()
    Debug.Print "-- Pass 1: field grading"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)

    AssertLong "MDL-CONSTR label grade", LabelFill(ws, "MDL-CONSTR"), CLR_FIELD_GREEN
    AssertLong "MDL-NORMAL label grade", LabelFill(ws, "MDL-NORMAL"), CLR_FIELD_GREEN
    AssertLong "MDL-FLAG label grade", LabelFill(ws, "MDL-FLAG"), CLR_FIELD_RED
End Sub

Private Function LabelFill(ws As Worksheet, ByVal modelName As String) As Long
    Dim b As Object
    Set b = BlockFor(modelName)
    If b Is Nothing Then Exit Function
    LabelFill = ws.Cells(GetKeyRow(b, KF_FIELD_FCST_FINAL), "B").Interior.Color
End Function

' --- Passes 2-4 own N..N+2 on the normal path ----------------
Private Sub Check_EarlyMonths()
    Debug.Print "-- Passes 2-4: N..N+2 before selection"

    ' SO FAR 300 + back order 50
    AssertNum "MDL-NORMAL N = SO FAR + backorder", OppAt("MDL-NORMAL", 0), 350
    AssertNum "MDL-NORMAL N+1 = SOF", OppAt("MDL-NORMAL", 1), 1100
    AssertNum "MDL-NORMAL N+2 = SOF", OppAt("MDL-NORMAL", 2), 1200
End Sub

' --- Selection pass ------------------------------------------
Private Sub Check_Selection()
    Debug.Print "-- Selection: constrained and normal paths"

    ' Constrained: availability for N..N+3, then availability + backorder
    AssertNum "MDL-CONSTR N   = availability", OppAt("MDL-CONSTR", 0), 100
    AssertNum "MDL-CONSTR N+1 = availability", OppAt("MDL-CONSTR", 1), 0
    AssertNum "MDL-CONSTR N+2 = availability", OppAt("MDL-CONSTR", 2), 0
    AssertNum "MDL-CONSTR N+3 = availability", OppAt("MDL-CONSTR", 3), 0
    AssertNum "MDL-CONSTR N+4 = avail + backorder", OppAt("MDL-CONSTR", 4), 250
    AssertNum "MDL-CONSTR N+5 = Consensus N-1", OppAt("MDL-CONSTR", 5), 1000

    ' Same model shape, but Field at N+4 beats availability + backorder
    AssertNum "MDL-CONSTR-FLD N+4 = Field override", OppAt("MDL-CONSTR-FLD", 4), 5000
    AssertNum "MDL-CONSTR-FLD N+3 = availability", OppAt("MDL-CONSTR-FLD", 3), 0

    ' Normal: N..N+2 untouched by selection, N+3 onward held at Consensus N-1
    AssertNum "MDL-NORMAL N   untouched", OppAt("MDL-NORMAL", 0), 350
    AssertNum "MDL-NORMAL N+1 untouched", OppAt("MDL-NORMAL", 1), 1100
    AssertNum "MDL-NORMAL N+2 untouched", OppAt("MDL-NORMAL", 2), 1200
    AssertNum "MDL-NORMAL N+3 = Consensus N-1", OppAt("MDL-NORMAL", 3), 900
    AssertNum "MDL-NORMAL N+8 = Consensus N-1", OppAt("MDL-NORMAL", 8), 900

    ' Variance never rewrites the number - N+5 holds at Consensus N-1
    ' even though Field is 30% away from it.
    AssertNum "MDL-FLAG N+5 holds despite 30% field gap", OppAt("MDL-FLAG", 5), 1000
End Sub

' --- Inventory font ------------------------------------------
Private Sub Check_Inventory()
    Debug.Print "-- Inventory: constrained font"

    ' Pool is exhausted by N+4, where demand is 250.
    AssertLong "MDL-CONSTR N+4 red font", OppCell("MDL-CONSTR", 4).Font.Color, CLR_INV_CONSTRAINED_TEXT
    AssertLong "MDL-CONSTR N+5 red font", OppCell("MDL-CONSTR", 5).Font.Color, CLR_INV_CONSTRAINED_TEXT

    ' A model with a huge pool is never marked.
    AssertNotLong "MDL-NORMAL N+3 not red", OppCell("MDL-NORMAL", 3).Font.Color, CLR_INV_CONSTRAINED_TEXT
End Sub

' --- Notes column --------------------------------------------
Private Sub Check_Notes()
    Debug.Print "-- Info block: notes column"

    Dim constrNotes As String, normalNotes As String, flagNotes As String
    constrNotes = NotesFor("MDL-CONSTR")
    normalNotes = NotesFor("MDL-CONSTR-FLD")
    flagNotes = NotesFor("MDL-FLAG")

    AssertContains "MDL-CONSTR notes name the availability run", constrNotes, DEC_CONSTRAINED
    AssertContains "MDL-CONSTR notes name the backorder month", constrNotes, DEC_CONSTRAINED_BO
    AssertContains "MDL-CONSTR notes name the consensus hold", constrNotes, DEC_CONSENSUS_HOLD
    AssertContains "MDL-CONSTR-FLD notes name the field override", normalNotes, DEC_FIELD_OVERRIDE

    AssertContains "MDL-FLAG notes carry a flag line", flagNotes, "FLAG"
    AssertContains "MDL-FLAG notes cite Consensus N-1", flagNotes, "vs Consensus N-1"
    AssertContains "MDL-FLAG notes carry an SOF reference", flagNotes, "SOF reference"
End Sub

Private Function NotesFor(ByVal modelName As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim b As Object: Set b = BlockFor(modelName)
    If b Is Nothing Then Exit Function

    Dim months As Variant: months = GetMonthColumns(ws)
    Dim infoCol As Long
    infoCol = CLng(months(UBound(months, 1), 1)) + INFO_BLOCK_COL_OFFSET

    Dim s As String, r As Long
    For r = CLng(b("startRow")) To CLng(b("endRow"))
        s = s & CStr(ws.Cells(r, infoCol).Value) & vbLf
    Next r
    NotesFor = s
End Function

' --- Dissonance fills ----------------------------------------
Private Sub Check_DissonanceFills()
    Debug.Print "-- Info block: dissonance fills"

    ' SOF 1500 vs Consensus N-1 1000 at N = +50%
    AssertLong "MDL-FLAG N filled (SOF gap)", OppCell("MDL-FLAG", 0).Interior.Color, CLR_DISSONANCE
    ' Field 1300 vs both baselines at N+5 = +30%
    AssertLong "MDL-FLAG N+5 filled (field gap)", OppCell("MDL-FLAG", 5).Interior.Color, CLR_DISSONANCE
    ' Nothing moves at N+6, so no fill
    AssertNotLong "MDL-FLAG N+6 unfilled", OppCell("MDL-FLAG", 6).Interior.Color, CLR_DISSONANCE
End Sub

' =============================================================
' Lookup + assertion helpers
' =============================================================

Private Function BlockFor(ByVal modelName As String) As Object
    Dim blocks As Collection
    Set blocks = FindModelBlocks()
    Dim b As Variant
    For Each b In blocks
        If StrComp(CStr(b("name")), modelName, vbTextCompare) = 0 Then
            Set BlockFor = b
            Exit Function
        End If
    Next b
End Function

' offset 0 = current month (N)
Private Function OppCell(ByVal modelName As String, ByVal offset As Long) As Range
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    Dim b As Object: Set b = BlockFor(modelName)
    If b Is Nothing Then Exit Function
    Set OppCell = ws.Cells(GetKeyRow(b, KF_OPPORTUNITY), CUR_COL + offset)
End Function

Private Function OppAt(ByVal modelName As String, ByVal offset As Long) As Double
    Dim c As Range
    Set c = OppCell(modelName, offset)
    If c Is Nothing Then Exit Function
    OppAt = SafeNum(c.Value)
End Function

Private Sub AssertNum(ByVal label As String, ByVal actual As Double, ByVal expected As Double)
    If Abs(actual - expected) < 0.000001 Then
        Debug.Print "  PASS  " & label & " = " & Format(actual, "#,##0.##")
    Else
        Debug.Print "  FAIL  " & label & "  expected=" & Format(expected, "#,##0.##") & _
                    "  actual=" & Format(actual, "#,##0.##")
        mFails = mFails + 1
    End If
End Sub

Private Sub AssertLong(ByVal label As String, ByVal actual As Long, ByVal expected As Long)
    If actual = expected Then
        Debug.Print "  PASS  " & label
    Else
        Debug.Print "  FAIL  " & label & "  expected=" & expected & "  actual=" & actual
        mFails = mFails + 1
    End If
End Sub

Private Sub AssertNotLong(ByVal label As String, ByVal actual As Long, ByVal notExpected As Long)
    If actual <> notExpected Then
        Debug.Print "  PASS  " & label
    Else
        Debug.Print "  FAIL  " & label & "  did not expect " & notExpected
        mFails = mFails + 1
    End If
End Sub

Private Sub AssertContains(ByVal label As String, ByVal haystack As String, ByVal needle As String)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        Debug.Print "  PASS  " & label
    Else
        Debug.Print "  FAIL  " & label & "  missing: " & needle
        mFails = mFails + 1
    End If
End Sub
