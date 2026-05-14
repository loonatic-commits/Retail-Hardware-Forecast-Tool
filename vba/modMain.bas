Attribute VB_Name = "modMain"
Option Explicit

' =============================================================
' modMain - orchestrator
'   Run_All_Passes runs the six passes in order, with a
'   preflight check first to fail fast on missing structure.
' =============================================================

Public Sub Run_All_Passes()
    Dim prevCalc As XlCalculation
    Dim screenWasOn As Boolean
    Dim eventsWereOn As Boolean
    Dim restored As Boolean
    restored = False

    On Error GoTo Fail

    InitColors

    ' --- Preflight check: bail before mutating anything if the
    '     workbook isn't shaped the way the passes expect. ---
    Dim issues As Collection
    Set issues = PreflightCheck()
    If issues.Count > 0 Then
        MsgBox "Cannot run passes - preflight failed:" & vbCrLf & vbCrLf & _
               FormatIssues(issues), vbCritical, "Demand Forecasting Suite"
        Exit Sub
    End If

    screenWasOn = Application.ScreenUpdating
    eventsWereOn = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Run_Module_1_FieldGrading
    Run_Module_2_FFvsBP
    Run_Module_3_SOFOverride
    Run_Module_4_SoFarRunRate
    Run_Module_5_InventoryPass
    Run_Module_6_Legend

    Application.ScreenUpdating = screenWasOn
    Application.EnableEvents = eventsWereOn
    Application.Calculation = prevCalc
    restored = True

    MsgBox "All passes complete.", vbInformation, "Demand Forecasting Suite"
    Exit Sub

Fail:
    If Not restored Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End If
    MsgBox "Failed in: " & Err.Source & vbCrLf & Err.Description, _
           vbCritical, "Demand Forecasting Suite"
End Sub

' -------------------------------------------------------------
' PreflightCheck
'   Validates that the workbook has everything the passes need
'   before any mutation happens. Returns a Collection of
'   human-readable issue strings; empty means OK to run.
'
'   Checks:
'     - Consensus and SOF sheets exist
'     - Consensus has parseable month headers (>= 1)
'     - At least one model block in Consensus
'     - Each block has the required key figures
'     - SOF has model rows and parseable month headers
' -------------------------------------------------------------
Public Function PreflightCheck() As Collection
    Dim issues As New Collection

    Dim wsCon As Worksheet
    On Error Resume Next
    Set wsCon = ThisWorkbook.Worksheets(SHEET_CONSENSUS)
    On Error GoTo 0
    If wsCon Is Nothing Then
        issues.Add "Missing sheet: '" & SHEET_CONSENSUS & "'."
        Set PreflightCheck = issues
        Exit Function
    End If

    Dim wsSof As Worksheet
    On Error Resume Next
    Set wsSof = ThisWorkbook.Worksheets(SHEET_SOF)
    On Error GoTo 0
    If wsSof Is Nothing Then
        issues.Add "Missing sheet: '" & SHEET_SOF & "'."
    End If

    Dim months As Variant
    months = GetMonthColumns(wsCon)
    If IsEmpty(months) Then
        issues.Add "Consensus row 1 has no parseable month headers in columns C onward."
    End If

    Dim blocks As Collection
    On Error Resume Next
    Set blocks = FindModelBlocks()
    On Error GoTo 0
    If blocks Is Nothing Or blocks.Count = 0 Then
        issues.Add "No model blocks found in Consensus column A."
        Set PreflightCheck = issues
        Exit Function
    End If

    Dim requiredKFs As Variant
    requiredKFs = Array(KF_BUSINESS_PLAN, KF_SELL_IN, KF_SO_FAR, _
                        KF_FIELD_FCST_FINAL, KF_OPPORTUNITY, KF_AVAIL_INV)

    Dim b As Variant, modelName As String, i As Long
    For Each b In blocks
        modelName = b("name")
        For i = LBound(requiredKFs) To UBound(requiredKFs)
            If GetKeyRow(b, CStr(requiredKFs(i))) = 0 Then
                issues.Add "Model '" & modelName & "' is missing key figure row '" & _
                           CStr(requiredKFs(i)) & "'."
            End If
        Next i
    Next b

    If Not (wsSof Is Nothing) Then
        Dim sofMonths As Variant
        sofMonths = GetSofMonthColumns(wsSof)
        If IsEmpty(sofMonths) Then
            issues.Add "SOF sheet has no parseable month headers in columns B-D (row 1)."
        End If
        Dim sofLastRow As Long
        sofLastRow = wsSof.Cells(wsSof.Rows.Count, "A").End(xlUp).Row
        If sofLastRow < 2 Then
            issues.Add "SOF sheet has no model rows in column A."
        End If
    End If

    Set PreflightCheck = issues
End Function

' -------------------------------------------------------------
' GetSofMonthColumns
'   Parses SOF row-1 headers from column B onward. Returns the
'   same shape as modUtilities.GetMonthColumns.
' -------------------------------------------------------------
Public Function GetSofMonthColumns(ws As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim tmp() As Variant
    ReDim tmp(1 To lastCol, 1 To 2)
    Dim n As Long: n = 0

    Dim c As Long, v As Variant, d As Date
    For c = 2 To lastCol
        v = ws.Cells(1, c).Value
        If TryParseMonth(v, d) Then
            n = n + 1
            tmp(n, 1) = c
            tmp(n, 2) = DateSerial(Year(d), Month(d), 1)
        End If
    Next c

    If n = 0 Then
        GetSofMonthColumns = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        out(i, 1) = tmp(i, 1)
        out(i, 2) = tmp(i, 2)
    Next i
    GetSofMonthColumns = out
End Function

Private Function FormatIssues(issues As Collection) As String
    Dim s As String, v As Variant
    For Each v In issues
        s = s & "  - " & CStr(v) & vbCrLf
    Next v
    FormatIssues = s
End Function

' --- Pass stubs ---------------------------------------------
' These exist so Run_All_Passes can be compiled and run even
' before each pass module is filled in. Each one will be
' replaced by its real implementation in modFieldGrading,
' modFFvsBP, etc. The stubs below just no-op.
' VBA resolves to the implementation in the dedicated module
' once it's imported, since module-level subs share a global
' namespace. If both exist VBA will refuse to compile, so
' DELETE these stubs before importing the real modules.
' -------------------------------------------------------------
#Const STUBS_ENABLED = True
#If STUBS_ENABLED Then
Public Sub Run_Module_1_FieldGrading_STUB(): Debug.Print "Pass 1 stub": End Sub
Public Sub Run_Module_2_FFvsBP_STUB(): Debug.Print "Pass 2 stub": End Sub
Public Sub Run_Module_3_SOFOverride_STUB(): Debug.Print "Pass 3 stub": End Sub
Public Sub Run_Module_4_SoFarRunRate_STUB(): Debug.Print "Pass 4 stub": End Sub
Public Sub Run_Module_5_InventoryPass_STUB(): Debug.Print "Pass 5 stub": End Sub
Public Sub Run_Module_6_Legend_STUB(): Debug.Print "Pass 6 stub": End Sub
#End If
