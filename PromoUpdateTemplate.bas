Attribute VB_Name = "PromoUpdateTemplate"
Option Explicit

'=====================================================================================
' PromoUpdateTemplate.bas
'
' Builds the SAP IBP "Promo Update Template" (US only) from four input sheets in the
' active workbook:
'
'     Promo Grid            product x week promo net prices
'     CY Template           last year's submitted template (channels, policy price, %)
'     NY Template           next year's blank template (header row = output layout)
'     Full Funding Events   event x model funding windows
'     Config                Setting / Value pairs
'
' Writes two sheets, rebuilt from scratch on every run:
'
'     Promo Update Output   the upload file
'     Run Log               validation exceptions + a run summary
'
' The input sheets are never modified.  Entry point: GeneratePromoUpdateTemplate.
'
'-------------------------------------------------------------------------------------
' ASSUMPTIONS
'  1. Sheet names are matched case-insensitively and ignoring surrounding whitespace.
'     Promo Grid, CY Template and NY Template are required; Full Funding Events and
'     Config are optional (their absence is logged and defaults are used).
'  2. Header rows are row 1 on every input sheet.  Columns are located by header text
'     through an alias table; no column letter or index is ever hardcoded.
'  3. "Pricing Policy FY" is NOT treated as "Pricing Policy ID".  If the CY template
'     has no Pricing Policy ID column, the Config default is written and logged.
'  4. A week column on the Promo Grid is any header that parses to a date.  A WB column
'     on the NY template is any unmapped header that parses as "WB MON-DD YYYY" (or a
'     bare "MON-DD YYYY").  Grid weeks are matched to WB columns by exact date.
'  5. Weeks run Sunday-to-Saturday: a block's calendar span is its first week's date
'     through its last week's date + 6 days.  "Full Period" therefore ends on the
'     event's End Week + 6 days.
'  6. A promo block is a run of consecutive week columns carrying the same price
'     (tolerance 0.0001).  A grid week with no matching WB column ends the run.
'  7. Contribution % is averaged per product x channel over CY rows with a positive
'     discount, excluding last year's fully funded rows (Accn Cont. = 0 with a
'     populated Special Promo Start Date).  Fallback order: product x channel average,
'     then channel average across all products, then 0 (both fallbacks are logged).
'  8. A channel is a Full Funding account when the Config account token appears in
'     either its Forecast Customer ID or its Forecast Customer Description.
'  9. An event Model token matches a product when it appears in the Product ID or in
'     the Model Desc (case-insensitive substring), so ES-C380W finds WORKFORCE ES-C380W.
' 10. Policy Price comes from the CY row for that product x channel; if it is missing
'     or zero the Promo Grid MSRP is substituted and logged.  MSRP is the grid MSRP.
' 11. Country, Financial Year, Key Figure and Pricing Policy ID always come from Config.
'     Financial Year is written into a text-formatted column so a value such as "2028"
'     survives the upload exactly as Config spells it.
' 12. Currency is rounded half-up to 2 decimals.  Special promo dates are written as
'     text in the Config date format so the blank placeholder can share the column.
' 13. Output is values only, never formulas.  NY columns that are neither recognised
'     nor WB week columns are left blank.
'=====================================================================================

'--- constants ------------------------------------------------------------------------
Private Const OUT_SHEET As String = "Promo Update Output"
Private Const LOG_SHEET As String = "Run Log"
Private Const KSEP As String = "|"
Private Const FLD As String = ";;"
Private Const EPS As Double = 0.0000001
Private Const PRICE_TOL As Double = 0.0001

'--- run state ------------------------------------------------------------------------
Private mStage As String
Private mT0 As Double

'--- raw sheet arrays -----------------------------------------------------------------
Private mGrid() As Variant
Private mGridRows As Long
Private mGridCols As Long
Private mCY() As Variant
Private mCYRows As Long
Private mCYCols As Long
Private mNY() As Variant
Private mNYCols As Long
Private mEvt() As Variant
Private mEvtRows As Long
Private mEvtCols As Long

'--- header maps ----------------------------------------------------------------------
Private mAlias As Object
Private mGridCol As Object
Private mCYCol As Object
Private mNYCol As Object
Private mEvtCol As Object

'--- resolved column positions --------------------------------------------------------
Private mGCProd As Long
Private mGCModel As Long
Private mGCMSRP As Long

Private mCCProd As Long
Private mCCCust As Long
Private mCCCustD As Long
Private mCCClass As Long
Private mCCModel As Long
Private mCCCountry As Long
Private mCCPP As Long
Private mCCNet As Long
Private mCCAccn As Long
Private mCCSpStart As Long
Private mCCMSRP As Long
Private mCCPolID As Long

Private mECName As Long
Private mECStart As Long
Private mECEnd As Long
Private mECMode As Long
Private mECModel As Long

'--- config ---------------------------------------------------------------------------
Private mCfg As Object
Private mCfgFinYear As String
Private mCfgKeyFig As String
Private mCfgCountry As String
Private mCfgFundDays As Long
Private mCfgFundAcct As String
Private mCfgPolicyID As String
Private mCfgBlank As String
Private mCfgDateFmt As String

'--- NY week calendar -----------------------------------------------------------------
Private mNYWkByDate As Object

'--- Promo Grid week calendar (chronological) -----------------------------------------
Private mWkDate() As Double
Private mWkGridCol() As Long
Private mWkNYCol() As Long
Private mWkCap() As String
Private mWkCount As Long

'--- CY derived -----------------------------------------------------------------------
Private mChanList As Object
Private mChanCust As Object
Private mChanDesc As Object
Private mChanClass As Object
Private mChanModel As Object
Private mChanCountry As Object
Private mChanPP As Object
Private mChanMSRP As Object
Private mChanPolID As Object
Private mPctSum As Object
Private mPctCnt As Object
Private mChPctSum As Object
Private mChPctCnt As Object
Private mOnce As Object

'--- events ---------------------------------------------------------------------------
Private mEvName() As String
Private mEvStart() As Double
Private mEvEnd() As Double
Private mEvMode() As String
Private mEvModels() As String
Private mEvCount As Long
Private mEvIdx As Object

'--- funding windows / segments for the block in hand ---------------------------------
Private mWinS() As Double
Private mWinE() As Double
Private mWinN() As String
Private mWinCount As Long
Private mSegS() As Double
Private mSegE() As Double
Private mSegN() As String
Private mSegF() As Boolean
Private mSegCount As Long

'--- output buffer --------------------------------------------------------------------
Private mOut() As Variant
Private mOutTmp() As Variant
Private mOutRow As Long
Private mOutCap As Long
Private mOutCols As Long
Private mFundedRows As Long

'--- log buffer -----------------------------------------------------------------------
Private mLog() As Variant
Private mLogTmp() As Variant
Private mLogRow As Long
Private mLogCap As Long


'=====================================================================================
' ENTRY POINT
'=====================================================================================
Public Sub GeneratePromoUpdateTemplate()
    Dim wb As Workbook
    Dim savedCalc As Long
    Dim savedScreen As Boolean
    Dim savedEvents As Boolean
    Dim savedAlerts As Boolean
    Dim stateSaved As Boolean
    Dim msg As String

    On Error GoTo Fail
    mStage = "starting up"
    mT0 = Timer
    stateSaved = False

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Err.Raise vbObjectError + 513, , "There is no active workbook."
    End If

    savedCalc = Application.Calculation
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    stateSaved = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    mStage = "initialising"
    InitState

    mStage = "reading the Config sheet"
    ReadConfig wb

    mStage = "reading the Promo Grid"
    ReadGrid wb

    mStage = "reading the NY Template header"
    ReadNY wb

    mStage = "matching Promo Grid weeks to the NY week calendar"
    BuildWeekMap

    mStage = "reading the CY Template"
    ReadCY wb

    mStage = "reading Full Funding Events"
    ReadEvents wb

    mStage = "building the output rows"
    BuildOutput

    mStage = "writing the Promo Update Output sheet"
    WriteOutput wb

    mStage = "writing the Run Log sheet"
    WriteLog wb

    mStage = "finishing up"
    Application.Calculation = savedCalc
    Application.EnableEvents = savedEvents
    Application.DisplayAlerts = savedAlerts
    Application.ScreenUpdating = savedScreen

    msg = "Promo Update Template built." & vbCrLf & vbCrLf
    msg = msg & "Rows written:       " & mOutRow & vbCrLf
    msg = msg & "Fully funded rows:  " & mFundedRows & vbCrLf
    msg = msg & "Events loaded:      " & mEvCount & vbCrLf
    msg = msg & "Exceptions logged:  " & mLogRow & vbCrLf
    msg = msg & "Run time:           " & Format$(Timer - mT0, "0.0") & " sec"
    MsgBox msg, vbInformation, "Promo Update Template"
    Exit Sub

Fail:
    msg = "Promo Update Template generation failed while " & mStage & "." & vbCrLf & vbCrLf
    msg = msg & "Error " & Err.Number & ": " & Err.Description
    If stateSaved Then
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
        Application.DisplayAlerts = savedAlerts
        Application.ScreenUpdating = savedScreen
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
    MsgBox msg, vbCritical, "Promo Update Template"
End Sub


'=====================================================================================
' SETUP
'=====================================================================================
Private Sub InitState()
    Set mAlias = CreateObject("Scripting.Dictionary")
    Set mGridCol = CreateObject("Scripting.Dictionary")
    Set mCYCol = CreateObject("Scripting.Dictionary")
    Set mNYCol = CreateObject("Scripting.Dictionary")
    Set mEvtCol = CreateObject("Scripting.Dictionary")
    Set mNYWkByDate = CreateObject("Scripting.Dictionary")
    Set mChanList = CreateObject("Scripting.Dictionary")
    Set mChanCust = CreateObject("Scripting.Dictionary")
    Set mChanDesc = CreateObject("Scripting.Dictionary")
    Set mChanClass = CreateObject("Scripting.Dictionary")
    Set mChanModel = CreateObject("Scripting.Dictionary")
    Set mChanCountry = CreateObject("Scripting.Dictionary")
    Set mChanPP = CreateObject("Scripting.Dictionary")
    Set mChanMSRP = CreateObject("Scripting.Dictionary")
    Set mChanPolID = CreateObject("Scripting.Dictionary")
    Set mPctSum = CreateObject("Scripting.Dictionary")
    Set mPctCnt = CreateObject("Scripting.Dictionary")
    Set mChPctSum = CreateObject("Scripting.Dictionary")
    Set mChPctCnt = CreateObject("Scripting.Dictionary")
    Set mOnce = CreateObject("Scripting.Dictionary")
    Set mEvIdx = CreateObject("Scripting.Dictionary")

    mWkCount = 0
    mEvCount = 0
    mOutRow = 0
    mOutCap = 0
    mOutCols = 0
    mFundedRows = 0
    mLogRow = 0
    mLogCap = 256
    ReDim mLog(1 To mLogCap, 1 To 5)

    BuildAliasMap
End Sub

Private Sub BuildAliasMap()
    AddAlias "Country (Region)", "country;region;country region;country/region;ctry"
    AddAlias "Product Class Desc", "product class;product class description;prod class desc;product class desc."
    AddAlias "Product ID", "product;product no;product number;material;material no;sku;item;item id"
    AddAlias "Model Desc", "model;model description;model desc.;model name;models"
    AddAlias "Forecast Customer Description", "customer description;customer desc;forecast customer desc;forecast cust desc;customer name"
    AddAlias "Forecast Customer ID", "customer id;customer;forecast customer;forecast cust id;cust id;channel;account"
    AddAlias "Pricing Policy ID", "pricing policy;policy id;pricing policy id."
    AddAlias "Pricing Policy FY", "policy fy;pricing policy fy."
    AddAlias "Promo Note", "note;notes;promo notes;promo note."
    AddAlias "Policy End Date", "policy end;policy end dt"
    AddAlias "Special Promo Start Date", "special promo start;promo start date;promo start;spec promo start date"
    AddAlias "Special Promo End Date", "special promo end;promo end date;promo end;spec promo end date"
    AddAlias "MSRP", "msrp price;list price;msrp."
    AddAlias "Policy Price", "policy;policy price."
    AddAlias "Financial Year", "fy;fiscal year;financial yr"
    AddAlias "Epson Cont.", "epson ir;epson cont;epson contribution;epson;epson cont ."
    AddAlias "Accn Cont.", "accn cont;accn;account cont;account cont.;accn contribution;account contribution;acct cont"
    AddAlias "Net", "net price;net promo price;promo net price"
    AddAlias "Key Figure", "keyfigure;key figure id;key fig"
    AddAlias "Start Week", "start wk;startweek;start;start date;event start;event start week"
    AddAlias "End Week", "end wk;endweek;end;end date;event end;event end week"
    AddAlias "Event Name", "event;event desc;event description;promo event"
    AddAlias "Funding Mode", "mode;funding type;funding"
    AddAlias "Setting", "settings;name;parameter;key"
    AddAlias "Value", "values;setting value"
End Sub

Private Sub AddAlias(ByVal canon As String, ByVal aliasList As String)
    Dim parts As Variant
    Dim i As Long
    Dim k As String

    k = NormHdr(canon)
    If Len(k) > 0 Then
        If Not mAlias.Exists(k) Then mAlias.Add k, canon
    End If
    If Len(aliasList) = 0 Then Exit Sub
    parts = Split(aliasList, ";")
    For i = LBound(parts) To UBound(parts)
        k = NormHdr(CStr(parts(i)))
        If Len(k) > 0 Then
            If Not mAlias.Exists(k) Then mAlias.Add k, canon
        End If
    Next i
End Sub

'--- normalise a header caption: lower case, letters/digits/spaces only, single spaces
Private Function NormHdr(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim t As String
    Dim out As String

    t = LCase$(Trim$(s))
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "a" Then
            If ch <= "z" Then
                out = out & ch
            Else
                out = out & " "
            End If
        ElseIf ch >= "0" Then
            If ch <= "9" Then
                out = out & ch
            Else
                out = out & " "
            End If
        Else
            out = out & " "
        End If
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormHdr = Trim$(out)
End Function

'--- one header cell from whichever module-level array the caller names
Private Function HdrCell(ByVal sheetKey As String, ByVal c As Long) As Variant
    Select Case sheetKey
        Case "GRID": HdrCell = mGrid(1, c)
        Case "CY": HdrCell = mCY(1, c)
        Case "NY": HdrCell = mNY(1, c)
        Case "EVT": HdrCell = mEvt(1, c)
        Case Else: HdrCell = Empty
    End Select
End Function

Private Function HdrCount(ByVal sheetKey As String) As Long
    Select Case sheetKey
        Case "GRID": HdrCount = mGridCols
        Case "CY": HdrCount = mCYCols
        Case "NY": HdrCount = mNYCols
        Case "EVT": HdrCount = mEvtCols
        Case Else: HdrCount = 0
    End Select
End Function

Private Sub MapHeaderRow(ByVal sheetKey As String, ByVal d As Object)
    Dim c As Long
    Dim n As Long
    Dim k As String
    Dim canon As String

    n = HdrCount(sheetKey)
    For c = 1 To n
        k = NormHdr(CStr2(HdrCell(sheetKey, c)))
        If Len(k) > 0 Then
            If mAlias.Exists(k) Then
                canon = CStr(mAlias(k))
                If Not d.Exists(canon) Then d.Add canon, c
            End If
        End If
    Next c
End Sub

Private Function ReqCol(ByVal d As Object, ByVal canon As String, ByVal sheetNm As String) As Long
    If Not d.Exists(canon) Then
        Err.Raise vbObjectError + 514, , "The '" & sheetNm & "' sheet has no '" & canon & "' column. Columns are located by header text, so check the spelling in row 1."
    End If
    ReqCol = CLng(d(canon))
End Function

Private Function OptCol(ByVal d As Object, ByVal canon As String) As Long
    If d.Exists(canon) Then
        OptCol = CLng(d(canon))
    Else
        OptCol = 0
    End If
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Dim target As String

    target = NormHdr(nm)
    For Each ws In wb.Worksheets
        If NormHdr(ws.Name) = target Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

Private Function LastRowOf(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        LastRowOf = 0
    Else
        LastRowOf = f.Row
    End If
End Function

Private Function LastColOf(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        LastColOf = 0
    Else
        LastColOf = f.Column
    End If
End Function


'=====================================================================================
' READERS
'=====================================================================================
Private Sub ReadConfig(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim arr As Variant
    Dim lr As Long
    Dim r As Long
    Dim k As String

    Set mCfg = CreateObject("Scripting.Dictionary")
    Set ws = FindSheet(wb, "Config")

    If ws Is Nothing Then
        LogIssue "Config sheet missing", "", "", "", "No 'Config' sheet was found; built-in defaults were used for every setting."
    Else
        lr = LastRowOf(ws)
        If lr >= 1 Then
            arr = ws.Range(ws.Cells(1, 1), ws.Cells(lr, 2)).Value2
            For r = 1 To UBound(arr, 1)
                k = NormHdr(CStr2(arr(r, 1)))
                If Len(k) > 0 Then
                    If Not mCfg.Exists(k) Then mCfg.Add k, CStr2(arr(r, 2))
                End If
            Next r
        End If
    End If

    mCfgFinYear = CfgVal("Financial Year", "")
    mCfgKeyFig = CfgVal("Key Figure", "")
    mCfgCountry = CfgVal("Country", "US")
    mCfgPolicyID = CfgVal("Pricing Policy ID", "UP")
    mCfgBlank = CfgVal("Blank Placeholder", "(None)")
    mCfgDateFmt = CfgVal("Date Format", "MM/DD/YYYY")
    mCfgFundAcct = UCase$(CfgVal("Full Funding Accounts", "AMAZON"))

    mCfgFundDays = 3
    k = CfgVal("Full Funding Days", "")
    If Len(k) > 0 Then
        If IsNumeric(k) Then mCfgFundDays = CLng(CDbl(k))
    End If
    If mCfgFundDays < 1 Then
        LogIssue "Config value out of range", "", "", "", "'Full Funding Days' was " & mCfgFundDays & "; 3 was used instead."
        mCfgFundDays = 3
    End If

    If Len(mCfgDateFmt) = 0 Then mCfgDateFmt = "MM/DD/YYYY"
    If Len(mCfgBlank) = 0 Then mCfgBlank = "(None)"
    If Len(mCfgCountry) = 0 Then mCfgCountry = "US"
    If Len(mCfgKeyFig) = 0 Then
        LogIssue "Config setting missing", "", "", "", "'Key Figure' is not set in Config; the output Key Figure column was left blank."
    End If
End Sub

Private Function CfgVal(ByVal nm As String, ByVal dflt As String) As String
    Dim k As String
    k = NormHdr(nm)
    CfgVal = dflt
    If mCfg Is Nothing Then Exit Function
    If mCfg.Exists(k) Then
        If Len(Trim$(CStr(mCfg(k)))) > 0 Then CfgVal = Trim$(CStr(mCfg(k)))
    End If
End Function

Private Sub ReadGrid(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lr As Long
    Dim lc As Long
    Dim c As Long
    Dim d As Double
    Dim cap As String

    Set ws = FindSheet(wb, "Promo Grid")
    If ws Is Nothing Then
        Err.Raise vbObjectError + 515, , "The 'Promo Grid' sheet was not found."
    End If

    lr = LastRowOf(ws)
    lc = LastColOf(ws)
    If lr < 2 Then
        Err.Raise vbObjectError + 516, , "The 'Promo Grid' sheet has a header row but no data rows."
    End If
    If lc < 2 Then
        Err.Raise vbObjectError + 517, , "The 'Promo Grid' sheet has fewer than two columns."
    End If

    mGrid = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    mGridRows = lr
    mGridCols = lc

    MapHeaderRow "GRID", mGridCol
    mGCProd = ReqCol(mGridCol, "Product ID", "Promo Grid")
    mGCModel = OptCol(mGridCol, "Model Desc")
    mGCMSRP = OptCol(mGridCol, "MSRP")
    If mGCModel = 0 Then
        LogIssue "Optional column missing", "", "", "", "The Promo Grid has no 'Model Desc' column; the CY Template model description was used for event matching."
    End If
    If mGCMSRP = 0 Then
        LogIssue "Optional column missing", "", "", "", "The Promo Grid has no 'MSRP' column; the CY Template MSRP was used instead."
    End If

    ReDim mWkDate(1 To mGridCols)
    ReDim mWkGridCol(1 To mGridCols)
    ReDim mWkNYCol(1 To mGridCols)
    ReDim mWkCap(1 To mGridCols)
    mWkCount = 0

    For c = 1 To mGridCols
        If c <> mGCProd Then
            If c <> mGCModel Then
                If c <> mGCMSRP Then
                    cap = CStr2(HdrCell("GRID", c))
                    d = ToDate(mGrid(1, c))
                    If d > 0 Then
                        mWkCount = mWkCount + 1
                        mWkDate(mWkCount) = d
                        mWkGridCol(mWkCount) = c
                        mWkNYCol(mWkCount) = 0
                        mWkCap(mWkCount) = cap
                    End If
                End If
            End If
        End If
    Next c

    If mWkCount = 0 Then
        Err.Raise vbObjectError + 518, , "No week columns were recognised on the 'Promo Grid'. Week headers must be real dates (for example 4/4/2027)."
    End If
    SortWeeks
End Sub

'--- insertion sort of the parallel week arrays into date order
Private Sub SortWeeks()
    Dim i As Long
    Dim j As Long
    Dim dd As Double
    Dim gc As Long
    Dim nc As Long
    Dim cp As String

    For i = 2 To mWkCount
        dd = mWkDate(i)
        gc = mWkGridCol(i)
        nc = mWkNYCol(i)
        cp = mWkCap(i)
        j = i - 1
        Do While j >= 1
            If mWkDate(j) <= dd Then Exit Do
            mWkDate(j + 1) = mWkDate(j)
            mWkGridCol(j + 1) = mWkGridCol(j)
            mWkNYCol(j + 1) = mWkNYCol(j)
            mWkCap(j + 1) = mWkCap(j)
            j = j - 1
        Loop
        mWkDate(j + 1) = dd
        mWkGridCol(j + 1) = gc
        mWkNYCol(j + 1) = nc
        mWkCap(j + 1) = cp
    Next i
End Sub

Private Sub ReadNY(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lc As Long
    Dim c As Long
    Dim d As Double
    Dim cap As String
    Dim k As String
    Dim wkCols As Long

    Set ws = FindSheet(wb, "NY Template")
    If ws Is Nothing Then
        Err.Raise vbObjectError + 519, , "The 'NY Template' sheet was not found."
    End If

    lc = LastColOf(ws)
    If lc < 2 Then
        Err.Raise vbObjectError + 520, , "The 'NY Template' header row has fewer than two columns."
    End If

    mNY = ws.Range(ws.Cells(1, 1), ws.Cells(1, lc)).Value2
    mNYCols = lc
    MapHeaderRow "NY", mNYCol

    ReqCol mNYCol, "Product ID", "NY Template"

    wkCols = 0
    For c = 1 To mNYCols
        cap = CStr2(mNY(1, c))
        If Len(Trim$(cap)) > 0 Then
            k = NormHdr(cap)
            If Not mAlias.Exists(k) Then
                d = ParseWBCaption(cap)
                If d > 0 Then
                    wkCols = wkCols + 1
                    If Not mNYWkByDate.Exists(CStr(CLng(d))) Then
                        mNYWkByDate.Add CStr(CLng(d)), c
                    End If
                End If
            End If
        End If
    Next c

    If wkCols = 0 Then
        Err.Raise vbObjectError + 521, , "No 'WB' week columns were recognised on the 'NY Template'. Captions must look like 'WB MAR-28 2027'."
    End If

    If Len(mCfgFinYear) = 0 Then
        LogIssue "Config setting missing", "", "", "", "'Financial Year' is not set in Config; the output Financial Year column was left blank."
    End If
End Sub

'--- link each Promo Grid week to its NY WB column, by parsed date
Private Sub BuildWeekMap()
    Dim i As Long
    Dim k As String
    Dim c As Long

    For i = 1 To mWkCount
        k = CStr(CLng(mWkDate(i)))
        If mNYWkByDate.Exists(k) Then
            c = CLng(mNYWkByDate(k))
            mWkNYCol(i) = c
            mWkCap(i) = CStr2(mNY(1, c))
        Else
            mWkNYCol(i) = 0
            LogIssue "Grid week not in NY calendar", "", "", FmtD(mWkDate(i)), "The Promo Grid has a week column for " & FmtD(mWkDate(i)) & " but the NY Template has no matching 'WB' column; that week was skipped."
        End If
    Next i
End Sub

Private Sub ReadCY(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lr As Long
    Dim lc As Long
    Dim r As Long
    Dim prod As String
    Dim cust As String
    Dim custD As String
    Dim key As String
    Dim pp As Double
    Dim nt As Double
    Dim ac As Double
    Dim disc As Double
    Dim pct As Double
    Dim skipRow As Boolean

    Set ws = FindSheet(wb, "CY Template")
    If ws Is Nothing Then
        Err.Raise vbObjectError + 522, , "The 'CY Template' sheet was not found."
    End If

    lr = LastRowOf(ws)
    lc = LastColOf(ws)
    If lr < 2 Then
        Err.Raise vbObjectError + 523, , "The 'CY Template' sheet has a header row but no data rows."
    End If
    If lc < 2 Then
        Err.Raise vbObjectError + 524, , "The 'CY Template' sheet has fewer than two columns."
    End If

    mCY = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    mCYRows = lr
    mCYCols = lc
    MapHeaderRow "CY", mCYCol

    mCCProd = ReqCol(mCYCol, "Product ID", "CY Template")
    mCCPP = ReqCol(mCYCol, "Policy Price", "CY Template")
    mCCNet = ReqCol(mCYCol, "Net", "CY Template")
    mCCAccn = ReqCol(mCYCol, "Accn Cont.", "CY Template")

    mCCCust = OptCol(mCYCol, "Forecast Customer ID")
    mCCCustD = OptCol(mCYCol, "Forecast Customer Description")
    mCCClass = OptCol(mCYCol, "Product Class Desc")
    mCCModel = OptCol(mCYCol, "Model Desc")
    mCCCountry = OptCol(mCYCol, "Country (Region)")
    mCCSpStart = OptCol(mCYCol, "Special Promo Start Date")
    mCCMSRP = OptCol(mCYCol, "MSRP")
    mCCPolID = OptCol(mCYCol, "Pricing Policy ID")

    If mCCCust = 0 Then
        If mCCCustD = 0 Then
            Err.Raise vbObjectError + 525, , "The 'CY Template' sheet has neither a 'Forecast Customer ID' nor a 'Forecast Customer Description' column, so channels cannot be identified."
        End If
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'Forecast Customer ID' column; the Forecast Customer Description was used as the channel key."
    End If
    If mCCCustD = 0 Then
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'Forecast Customer Description' column; the Forecast Customer ID was written into the description column."
    End If
    If mCCClass = 0 Then
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'Product Class Desc' column; the output Product Class Desc was left blank."
    End If
    If mCCSpStart = 0 Then
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'Special Promo Start Date' column; last year's fully funded rows could not be excluded from the contribution % average."
    End If
    If mCCMSRP = 0 Then
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'MSRP' column; the Promo Grid MSRP was used."
    End If
    If mCCPolID = 0 Then
        LogIssue "Optional column missing", "", "", "", "The CY Template has no 'Pricing Policy ID' column ('Pricing Policy FY' is deliberately not treated as the same field); the Config default '" & mCfgPolicyID & "' was written into every output row."
    End If

    For r = 2 To mCYRows
        prod = UCase$(Trim$(CStr2(mCY(r, mCCProd))))
        If Len(prod) > 0 Then
            cust = ""
            custD = ""
            If mCCCust > 0 Then cust = Trim$(CStr2(mCY(r, mCCCust)))
            If mCCCustD > 0 Then custD = Trim$(CStr2(mCY(r, mCCCustD)))
            If Len(cust) = 0 Then cust = custD
            If Len(custD) = 0 Then custD = cust

            If Len(cust) > 0 Then
                key = prod & KSEP & UCase$(cust)

                If Not mChanList.Exists(prod) Then
                    mChanList.Add prod, ""
                End If
                If Not mChanCust.Exists(key) Then
                    mChanCust.Add key, cust
                    mChanDesc.Add key, custD
                    mChanClass.Add key, ""
                    mChanModel.Add key, ""
                    mChanCountry.Add key, ""
                    mChanPP.Add key, CDbl(0)
                    mChanMSRP.Add key, CDbl(0)
                    mChanPolID.Add key, ""
                    mChanList(prod) = JoinTok(CStr(mChanList(prod)), cust)
                End If

                If mCCClass > 0 Then
                    If Len(CStr(mChanClass(key))) = 0 Then mChanClass(key) = Trim$(CStr2(mCY(r, mCCClass)))
                End If
                If mCCModel > 0 Then
                    If Len(CStr(mChanModel(key))) = 0 Then mChanModel(key) = Trim$(CStr2(mCY(r, mCCModel)))
                End If
                If mCCCountry > 0 Then
                    If Len(CStr(mChanCountry(key))) = 0 Then mChanCountry(key) = Trim$(CStr2(mCY(r, mCCCountry)))
                End If
                If mCCPolID > 0 Then
                    If Len(CStr(mChanPolID(key))) = 0 Then mChanPolID(key) = Trim$(CStr2(mCY(r, mCCPolID)))
                End If

                pp = ToNum(mCY(r, mCCPP))
                nt = ToNum(mCY(r, mCCNet))
                ac = ToNum(mCY(r, mCCAccn))

                If pp > 0 Then
                    If CDbl(mChanPP(key)) <= 0 Then mChanPP(key) = pp
                End If
                If mCCMSRP > 0 Then
                    If CDbl(mChanMSRP(key)) <= 0 Then mChanMSRP(key) = ToNum(mCY(r, mCCMSRP))
                End If

                disc = pp - nt
                If disc > EPS Then
                    skipRow = False
                    If Abs(ac) < EPS Then
                        If mCCSpStart > 0 Then
                            If HasVal(mCY(r, mCCSpStart)) Then skipRow = True
                        End If
                    End If
                    If Not skipRow Then
                        pct = ac / disc
                        If mPctCnt.Exists(key) Then
                            mPctSum(key) = CDbl(mPctSum(key)) + pct
                            mPctCnt(key) = CLng(mPctCnt(key)) + 1
                        Else
                            mPctSum.Add key, pct
                            mPctCnt.Add key, CLng(1)
                        End If
                        If mChPctCnt.Exists(UCase$(cust)) Then
                            mChPctSum(UCase$(cust)) = CDbl(mChPctSum(UCase$(cust))) + pct
                            mChPctCnt(UCase$(cust)) = CLng(mChPctCnt(UCase$(cust))) + 1
                        Else
                            mChPctSum.Add UCase$(cust), pct
                            mChPctCnt.Add UCase$(cust), CLng(1)
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Sub ReadEvents(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lr As Long
    Dim lc As Long
    Dim r As Long
    Dim nm As String
    Dim ky As String
    Dim md As String
    Dim idx As Long
    Dim ds As Double
    Dim de As Double

    mEvCount = 0
    Set ws = FindSheet(wb, "Full Funding Events")
    If ws Is Nothing Then
        LogIssue "Full Funding Events sheet missing", "", "", "", "No 'Full Funding Events' sheet was found; every row was written with the normal contribution split."
        Exit Sub
    End If

    lr = LastRowOf(ws)
    lc = LastColOf(ws)
    If lr < 2 Then
        LogIssue "No funding events", "", "", "", "The 'Full Funding Events' sheet has no data rows; every row was written with the normal contribution split."
        Exit Sub
    End If
    If lc < 2 Then
        LogIssue "No funding events", "", "", "", "The 'Full Funding Events' sheet has fewer than two columns; it was ignored."
        Exit Sub
    End If

    mEvt = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    mEvtRows = lr
    mEvtCols = lc
    MapHeaderRow "EVT", mEvtCol

    mECName = ReqCol(mEvtCol, "Event Name", "Full Funding Events")
    mECModel = ReqCol(mEvtCol, "Model Desc", "Full Funding Events")
    mECStart = ReqCol(mEvtCol, "Start Week", "Full Funding Events")
    mECEnd = OptCol(mEvtCol, "End Week")
    mECMode = OptCol(mEvtCol, "Funding Mode")

    If mECEnd = 0 Then
        LogIssue "Optional column missing", "", "", "", "The Full Funding Events sheet has no 'End Week' column; 'Full Period' events were treated as one week long."
    End If
    If mECMode = 0 Then
        LogIssue "Optional column missing", "", "", "", "The Full Funding Events sheet has no 'Funding Mode' column; every event was treated as 'Full Period'."
    End If

    ReDim mEvName(1 To mEvtRows)
    ReDim mEvStart(1 To mEvtRows)
    ReDim mEvEnd(1 To mEvtRows)
    ReDim mEvMode(1 To mEvtRows)
    ReDim mEvModels(1 To mEvtRows)

    For r = 2 To mEvtRows
        nm = Trim$(CStr2(mEvt(r, mECName)))
        If Len(nm) > 0 Then
            ky = UCase$(nm)
            If mEvIdx.Exists(ky) Then
                idx = CLng(mEvIdx(ky))
            Else
                mEvCount = mEvCount + 1
                idx = mEvCount
                mEvIdx.Add ky, idx
                mEvName(idx) = nm
                mEvModels(idx) = ""
                ds = ToDate(mEvt(r, mECStart))
                de = 0
                If mECEnd > 0 Then de = ToDate(mEvt(r, mECEnd))
                If de <= 0 Then de = ds
                mEvStart(idx) = ds
                mEvEnd(idx) = de
                If mECMode > 0 Then
                    mEvMode(idx) = Trim$(CStr2(mEvt(r, mECMode)))
                Else
                    mEvMode(idx) = "Full Period"
                End If
                If ds <= 0 Then
                    LogIssue "Event start week unreadable", "", "", "", "Event '" & nm & "' has no readable Start Week and was ignored."
                End If
            End If
            md = Trim$(CStr2(mEvt(r, mECModel)))
            If Len(md) > 0 Then
                mEvModels(idx) = JoinTok(mEvModels(idx), md)
            End If
        End If
    Next r
End Sub


'=====================================================================================
' OUTPUT CONSTRUCTION
'=====================================================================================
Private Sub BuildOutput()
    Dim r As Long
    Dim i As Long
    Dim j As Long
    Dim prod As String
    Dim pk As String
    Dim model As String
    Dim gridMSRP As Double
    Dim price As Double
    Dim usable As Boolean
    Dim ok As Boolean

    mOutCols = mNYCols
    mOutCap = 512
    ReDim mOut(1 To mOutCap, 1 To mOutCols)
    mOutRow = 0
    mFundedRows = 0

    For r = 2 To mGridRows
        prod = Trim$(CStr2(mGrid(r, mGCProd)))
        If Len(prod) > 0 Then
            pk = UCase$(prod)
            model = ""
            If mGCModel > 0 Then model = Trim$(CStr2(mGrid(r, mGCModel)))
            gridMSRP = 0
            If mGCMSRP > 0 Then gridMSRP = ToNum(mGrid(r, mGCMSRP))

            If Not mChanList.Exists(pk) Then
                LogIssue "No CY channel rows", prod, "", "", "'" & prod & "' appears on the Promo Grid but has no rows on the CY Template, so no channel could be assigned. No output rows were written for it."
            Else
                i = 1
                Do While i <= mWkCount
                    usable = False
                    price = 0
                    If mWkNYCol(i) > 0 Then
                        If HasNum(mGrid(r, mWkGridCol(i))) Then
                            price = ToNum(mGrid(r, mWkGridCol(i)))
                            If price > 0 Then usable = True
                        End If
                    End If

                    If usable Then
                        j = i
                        Do While j < mWkCount
                            ok = False
                            If mWkNYCol(j + 1) > 0 Then
                                If HasNum(mGrid(r, mWkGridCol(j + 1))) Then
                                    If Abs(ToNum(mGrid(r, mWkGridCol(j + 1))) - price) < PRICE_TOL Then ok = True
                                End If
                            End If
                            If Not ok Then Exit Do
                            j = j + 1
                        Loop
                        EmitBlock r, prod, model, gridMSRP, price, i, j
                        i = j + 1
                    Else
                        i = i + 1
                    End If
                Loop
            End If
        End If
    Next r
End Sub

'--- write every channel row for one contiguous promo block (grid rows i1..i2)
Private Sub EmitBlock(ByVal gr As Long, ByVal prod As String, ByVal model As String, ByVal gridMSRP As Double, ByVal price As Double, ByVal i1 As Long, ByVal i2 As Long)
    Dim pk As String
    Dim custs As Variant
    Dim ci As Long
    Dim cust As String
    Dim key As String
    Dim useModel As String
    Dim useMSRP As Double
    Dim pp As Double
    Dim pct As Double
    Dim isFF As Boolean
    Dim blockStart As Double
    Dim blockEnd As Double
    Dim s As Long
    Dim w As Long
    Dim disc As Double
    Dim accn As Double
    Dim epson As Double
    Dim note As String
    Dim sdText As String
    Dim edText As String
    Dim polID As String
    Dim wholeBlock As Boolean
    Dim wkLabel As String

    pk = UCase$(prod)
    blockStart = mWkDate(i1)
    blockEnd = mWkDate(i2) + 6
    wkLabel = FmtD(blockStart) & " - " & FmtD(blockEnd)

    custs = Split(CStr(mChanList(pk)), FLD)
    For ci = LBound(custs) To UBound(custs)
        cust = CStr(custs(ci))
        If Len(cust) > 0 Then
            key = pk & KSEP & UCase$(cust)

            useModel = model
            If Len(useModel) = 0 Then useModel = CStr(mChanModel(key))

            useMSRP = gridMSRP
            If useMSRP <= 0 Then useMSRP = CDbl(mChanMSRP(key))

            pp = CDbl(mChanPP(key))
            If pp <= 0 Then
                pp = gridMSRP
                LogOnce "PPMISS" & key, "Policy Price missing", prod, cust, "", "The CY Template carries no positive Policy Price for this product x channel; the Promo Grid MSRP (" & Format$(gridMSRP, "0.00") & ") was used."
            End If
            If gridMSRP > 0 Then
                If Abs(pp - gridMSRP) > 0.005 Then
                    LogOnce "PPMSRP" & key, "Policy Price <> grid MSRP", prod, cust, "", "CY Policy Price " & Format$(pp, "0.00") & " does not equal the Promo Grid MSRP " & Format$(gridMSRP, "0.00") & ". The CY Policy Price was used for the contribution split."
                End If
            End If
            If price >= pp - EPS Then
                LogOnce "PGTP" & key & wkLabel, "Promo price >= policy price", prod, cust, wkLabel, "Promo net price " & Format$(price, "0.00") & " is not below the policy price " & Format$(pp, "0.00") & "; the discount for these weeks is zero or negative."
            End If

            pct = GetPct(prod, cust)
            isFF = IsFundingAccount(cust, CStr(mChanDesc(key)))

            polID = mCfgPolicyID
            If mCCPolID > 0 Then
                If Len(CStr(mChanPolID(key))) > 0 Then polID = CStr(mChanPolID(key))
            End If

            BuildWindows prod, useModel, blockStart, blockEnd, isFF
            BuildSegments blockStart, blockEnd

            For s = 1 To mSegCount
                mOutRow = mOutRow + 1
                EnsureOutCapacity mOutRow

                disc = R2(pp - price)
                If mSegF(s) Then
                    accn = 0
                    epson = disc
                    note = mSegN(s)
                    mFundedRows = mFundedRows + 1
                Else
                    accn = R2(pct * disc)
                    epson = R2(disc - accn)
                    note = mCfgBlank
                End If

                If epson < -EPS Then
                    LogIssue "Negative Epson contribution", prod, cust, FmtD(mSegS(s)) & " - " & FmtD(mSegE(s)), "Epson Cont. computed as " & Format$(epson, "0.00") & " (policy price " & Format$(pp, "0.00") & ", net " & Format$(price, "0.00") & ", contribution % " & Format$(pct, "0.0%") & ")."
                End If

                wholeBlock = False
                If mSegS(s) <= blockStart + EPS Then
                    If mSegE(s) >= blockEnd - EPS Then wholeBlock = True
                End If
                If wholeBlock Then
                    sdText = mCfgBlank
                    edText = mCfgBlank
                Else
                    sdText = Format$(CDate(mSegS(s)), mCfgDateFmt)
                    edText = Format$(CDate(mSegE(s)), mCfgDateFmt)
                End If

                SetOut "Country (Region)", mCfgCountry
                SetOut "Product Class Desc", CStr(mChanClass(key))
                SetOut "Product ID", prod
                SetOut "Model Desc", useModel
                SetOut "Forecast Customer Description", CStr(mChanDesc(key))
                SetOut "Forecast Customer ID", CStr(mChanCust(key))
                SetOut "Pricing Policy ID", polID
                SetOut "Promo Note", note
                SetOut "Special Promo Start Date", sdText
                SetOut "Special Promo End Date", edText
                SetOut "MSRP", R2(useMSRP)
                SetOut "Policy Price", R2(pp)
                SetOut "Financial Year", mCfgFinYear
                SetOut "Epson Cont.", epson
                SetOut "Accn Cont.", accn
                SetOut "Net", R2(price)
                SetOut "Key Figure", mCfgKeyFig
                SetOut "Start Week", mWkCap(i1)

                For w = i1 To i2
                    mOut(mOutRow, mWkNYCol(w)) = R2(price)
                Next w
            Next s
        End If
    Next ci
End Sub

'--- fully funded windows overlapping [bs, be] for this product, merged and clipped
Private Sub BuildWindows(ByVal prod As String, ByVal model As String, ByVal bs As Double, ByVal be As Double, ByVal ffAccount As Boolean)
    Dim e As Long
    Dim fs As Double
    Dim fe As Double

    mWinCount = 0
    ReDim mWinS(1 To 8)
    ReDim mWinE(1 To 8)
    ReDim mWinN(1 To 8)

    If Not ffAccount Then Exit Sub
    If mEvCount = 0 Then Exit Sub

    ReDim mWinS(1 To mEvCount)
    ReDim mWinE(1 To mEvCount)
    ReDim mWinN(1 To mEvCount)

    For e = 1 To mEvCount
        If mEvStart(e) > 0 Then
            If EventCoversModel(e, prod, model) Then
                fs = mEvStart(e)
                If InStr(1, mEvMode(e), "FIRST", vbTextCompare) > 0 Then
                    fe = fs + mCfgFundDays - 1
                Else
                    fe = mEvEnd(e) + 6
                End If
                If fs < bs Then fs = bs
                If fe > be Then fe = be
                If fs <= fe Then
                    mWinCount = mWinCount + 1
                    mWinS(mWinCount) = fs
                    mWinE(mWinCount) = fe
                    mWinN(mWinCount) = mEvName(e)
                End If
            End If
        End If
    Next e

    SortMergeWindows
End Sub

'--- sort the window list by start date, then merge overlapping / abutting windows
Private Sub SortMergeWindows()
    Dim i As Long
    Dim j As Long
    Dim ds As Double
    Dim de As Double
    Dim nn As String
    Dim k As Long

    If mWinCount < 2 Then Exit Sub

    For i = 2 To mWinCount
        ds = mWinS(i)
        de = mWinE(i)
        nn = mWinN(i)
        j = i - 1
        Do While j >= 1
            If mWinS(j) <= ds Then Exit Do
            mWinS(j + 1) = mWinS(j)
            mWinE(j + 1) = mWinE(j)
            mWinN(j + 1) = mWinN(j)
            j = j - 1
        Loop
        mWinS(j + 1) = ds
        mWinE(j + 1) = de
        mWinN(j + 1) = nn
    Next i

    k = 1
    For i = 2 To mWinCount
        If mWinS(i) <= mWinE(k) + 1 Then
            If mWinE(i) > mWinE(k) Then mWinE(k) = mWinE(i)
            If InStr(1, mWinN(k), mWinN(i), vbTextCompare) = 0 Then
                mWinN(k) = mWinN(k) & " / " & mWinN(i)
            End If
        Else
            k = k + 1
            mWinS(k) = mWinS(i)
            mWinE(k) = mWinE(i)
            mWinN(k) = mWinN(i)
        End If
    Next i
    mWinCount = k
End Sub

'--- split [bs, be] into funded and normal segments covering every day exactly once
Private Sub BuildSegments(ByVal bs As Double, ByVal be As Double)
    Dim w As Long
    Dim cur As Double
    Dim cap As Long

    cap = mWinCount * 2 + 2
    ReDim mSegS(1 To cap)
    ReDim mSegE(1 To cap)
    ReDim mSegN(1 To cap)
    ReDim mSegF(1 To cap)
    mSegCount = 0
    cur = bs

    For w = 1 To mWinCount
        If mWinS(w) > cur Then
            mSegCount = mSegCount + 1
            mSegS(mSegCount) = cur
            mSegE(mSegCount) = mWinS(w) - 1
            mSegN(mSegCount) = ""
            mSegF(mSegCount) = False
        End If
        mSegCount = mSegCount + 1
        mSegS(mSegCount) = mWinS(w)
        mSegE(mSegCount) = mWinE(w)
        mSegN(mSegCount) = mWinN(w)
        mSegF(mSegCount) = True
        cur = mWinE(w) + 1
    Next w

    If cur <= be Then
        mSegCount = mSegCount + 1
        mSegS(mSegCount) = cur
        mSegE(mSegCount) = be
        mSegN(mSegCount) = ""
        mSegF(mSegCount) = False
    End If

    If mSegCount = 0 Then
        mSegCount = 1
        mSegS(1) = bs
        mSegE(1) = be
        mSegN(1) = ""
        mSegF(1) = False
    End If
End Sub

Private Function EventCoversModel(ByVal e As Long, ByVal prod As String, ByVal model As String) As Boolean
    Dim toks As Variant
    Dim i As Long
    Dim t As String

    EventCoversModel = False
    If Len(mEvModels(e)) = 0 Then Exit Function
    toks = Split(mEvModels(e), FLD)
    For i = LBound(toks) To UBound(toks)
        t = Trim$(CStr(toks(i)))
        If Len(t) > 0 Then
            If Len(prod) > 0 Then
                If InStr(1, prod, t, vbTextCompare) > 0 Then
                    EventCoversModel = True
                    Exit Function
                End If
            End If
            If Len(model) > 0 Then
                If InStr(1, model, t, vbTextCompare) > 0 Then
                    EventCoversModel = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function IsFundingAccount(ByVal cust As String, ByVal custDesc As String) As Boolean
    Dim toks As Variant
    Dim i As Long
    Dim t As String

    IsFundingAccount = False
    If Len(mCfgFundAcct) = 0 Then Exit Function
    toks = Split(mCfgFundAcct, ";")
    For i = LBound(toks) To UBound(toks)
        t = Trim$(CStr(toks(i)))
        If Len(t) > 0 Then
            If Len(cust) > 0 Then
                If InStr(1, cust, t, vbTextCompare) > 0 Then
                    IsFundingAccount = True
                    Exit Function
                End If
            End If
            If Len(custDesc) > 0 Then
                If InStr(1, custDesc, t, vbTextCompare) > 0 Then
                    IsFundingAccount = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

'--- contribution % carried forward from last year, with logged fallbacks
Private Function GetPct(ByVal prod As String, ByVal cust As String) As Double
    Dim key As String
    Dim ck As String
    Dim v As Double

    key = UCase$(prod) & KSEP & UCase$(cust)
    ck = UCase$(cust)
    GetPct = 0

    If mPctCnt.Exists(key) Then
        If CLng(mPctCnt(key)) > 0 Then
            GetPct = CDbl(mPctSum(key)) / CLng(mPctCnt(key))
            Exit Function
        End If
    End If

    If mChPctCnt.Exists(ck) Then
        If CLng(mChPctCnt(ck)) > 0 Then
            v = CDbl(mChPctSum(ck)) / CLng(mChPctCnt(ck))
            GetPct = v
            LogOnce "PCTFB" & key, "Contribution % fallback", prod, cust, "", "No usable CY rows for this product x channel; used the channel average across all products (" & Format$(v, "0.0%") & ")."
            Exit Function
        End If
    End If

    LogOnce "PCTZERO" & key, "Contribution % is zero", prod, cust, "", "No usable CY rows for this product x channel and no channel average was available; Accn Cont. was written as 0 and Epson Cont. carries the whole discount."
End Function

Private Sub SetOut(ByVal canon As String, ByVal v As Variant)
    Dim c As Long
    If Not mNYCol.Exists(canon) Then Exit Sub
    c = CLng(mNYCol(canon))
    If c < 1 Then Exit Sub
    If c > mOutCols Then Exit Sub
    mOut(mOutRow, c) = v
End Sub


'=====================================================================================
' BUFFERS AND LOGGING
'=====================================================================================
Private Sub EnsureOutCapacity(ByVal needed As Long)
    Dim i As Long
    Dim j As Long
    Dim newCap As Long

    If needed <= mOutCap Then Exit Sub
    newCap = mOutCap * 2
    If newCap < needed Then newCap = needed + 512

    ReDim mOutTmp(1 To newCap, 1 To mOutCols)
    For i = 1 To mOutCap
        For j = 1 To mOutCols
            mOutTmp(i, j) = mOut(i, j)
        Next j
    Next i
    mOut = mOutTmp
    Erase mOutTmp
    mOutCap = newCap
End Sub

Private Sub EnsureLogCapacity(ByVal needed As Long)
    Dim i As Long
    Dim j As Long
    Dim newCap As Long

    If needed <= mLogCap Then Exit Sub
    newCap = mLogCap * 2
    If newCap < needed Then newCap = needed + 256

    ReDim mLogTmp(1 To newCap, 1 To 5)
    For i = 1 To mLogCap
        For j = 1 To 5
            mLogTmp(i, j) = mLog(i, j)
        Next j
    Next i
    mLog = mLogTmp
    Erase mLogTmp
    mLogCap = newCap
End Sub

Private Sub LogIssue(ByVal issue As String, ByVal prod As String, ByVal chan As String, ByVal wk As String, ByVal detail As String)
    mLogRow = mLogRow + 1
    EnsureLogCapacity mLogRow
    mLog(mLogRow, 1) = issue
    mLog(mLogRow, 2) = prod
    mLog(mLogRow, 3) = chan
    mLog(mLogRow, 4) = wk
    mLog(mLogRow, 5) = detail
End Sub

Private Sub LogOnce(ByVal onceKey As String, ByVal issue As String, ByVal prod As String, ByVal chan As String, ByVal wk As String, ByVal detail As String)
    If mOnce.Exists(onceKey) Then Exit Sub
    mOnce.Add onceKey, 1
    LogIssue issue, prod, chan, wk, detail
End Sub


'=====================================================================================
' SHEET WRITERS
'=====================================================================================
Private Sub WriteOutput(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As String
    Dim d As Double
    Dim cur As Variant
    Dim ci As Long

    KillSheet wb, OUT_SHEET
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = OUT_SHEET

    '--- number formats first, so text dates stay text and never coerce to serials
    cur = Array("MSRP", "Policy Price", "Epson Cont.", "Accn Cont.", "Net")
    For ci = LBound(cur) To UBound(cur)
        If mNYCol.Exists(CStr(cur(ci))) Then
            ws.Columns(CLng(mNYCol(CStr(cur(ci))))).NumberFormat = "0.00"
        End If
    Next ci
    If mNYCol.Exists("Special Promo Start Date") Then
        ws.Columns(CLng(mNYCol("Special Promo Start Date"))).NumberFormat = "@"
    End If
    If mNYCol.Exists("Special Promo End Date") Then
        ws.Columns(CLng(mNYCol("Special Promo End Date"))).NumberFormat = "@"
    End If
    If mNYCol.Exists("Financial Year") Then
        ws.Columns(CLng(mNYCol("Financial Year"))).NumberFormat = "@"
    End If
    If mNYCol.Exists("Start Week") Then
        ws.Columns(CLng(mNYCol("Start Week"))).NumberFormat = "@"
    End If
    For c = 1 To mNYCols
        k = NormHdr(CStr2(mNY(1, c)))
        If Not mAlias.Exists(k) Then
            d = ParseWBCaption(CStr2(mNY(1, c)))
            If d > 0 Then ws.Columns(c).NumberFormat = "0.00"
        End If
    Next c

    '--- header row, copied verbatim from the NY Template
    For c = 1 To mNYCols
        ws.Cells(1, c).Value = CStr2(mNY(1, c))
    Next c
    ws.Rows(1).Font.Bold = True

    '--- values only, one array dump
    If mOutRow > 0 Then
        ReDim mOutTmp(1 To mOutRow, 1 To mOutCols)
        For i = 1 To mOutRow
            For j = 1 To mOutCols
                mOutTmp(i, j) = mOut(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + mOutRow, mOutCols)).Value = mOutTmp
        Erase mOutTmp
    End If

    '--- cosmetic only; never let a window-state quirk fail a good run
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim r As Long

    KillSheet wb, LOG_SHEET
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = LOG_SHEET

    ws.Cells(1, 1).Value = "Issue"
    ws.Cells(1, 2).Value = "Product ID"
    ws.Cells(1, 3).Value = "Channel"
    ws.Cells(1, 4).Value = "Week"
    ws.Cells(1, 5).Value = "Detail"
    ws.Rows(1).Font.Bold = True

    If mLogRow > 0 Then
        ReDim mLogTmp(1 To mLogRow, 1 To 5)
        For i = 1 To mLogRow
            For j = 1 To 5
                mLogTmp(i, j) = mLog(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + mLogRow, 5)).Value = mLogTmp
        Erase mLogTmp
    Else
        ws.Cells(2, 1).Value = "No exceptions were raised on this run."
    End If

    r = 3 + mLogRow
    If mLogRow = 0 Then r = 4

    ws.Cells(r, 1).Value = "RUN SUMMARY"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r + 1, 1).Value = "Rows written"
    ws.Cells(r + 1, 2).Value = mOutRow
    ws.Cells(r + 2, 1).Value = "Fully funded rows"
    ws.Cells(r + 2, 2).Value = mFundedRows
    ws.Cells(r + 3, 1).Value = "Events loaded"
    ws.Cells(r + 3, 2).Value = mEvCount
    ws.Cells(r + 4, 1).Value = "Exceptions logged"
    ws.Cells(r + 4, 2).Value = mLogRow
    ws.Cells(r + 5, 1).Value = "Run time (seconds)"
    ws.Cells(r + 5, 2).Value = R2(Timer - mT0)
    ws.Cells(r + 6, 1).Value = "Generated"
    ws.Cells(r + 6, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r + 7, 1).Value = "Products on grid"
    ws.Cells(r + 7, 2).Value = mGridRows - 1
    ws.Cells(r + 8, 1).Value = "Grid weeks mapped"
    ws.Cells(r + 8, 2).Value = mWkCount

    ws.Columns(1).ColumnWidth = 32
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 26
    ws.Columns(4).ColumnWidth = 24
    ws.Columns(5).ColumnWidth = 110
    ws.Columns(5).WrapText = False
    On Error Resume Next
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Private Sub KillSheet(ByVal wb As Workbook, ByVal nm As String)
    Dim ws As Worksheet

    Set ws = FindSheet(wb, nm)
    If ws Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub


'=====================================================================================
' UTILITIES
'=====================================================================================

'--- half-up rounding to cents (VBA's Round is banker's rounding)
Private Function R2(ByVal v As Double) As Double
    Dim sgn As Double

    sgn = 1
    If v < 0 Then sgn = -1
    R2 = sgn * CDbl(Int(CDec(Abs(v)) * 100 + 0.5)) / 100
End Function

'--- IsNumeric says True for Empty, so every numeric read goes through here
Private Function ToNum(ByVal v As Variant) As Double
    ToNum = 0
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If VarType(v) = vbBoolean Then Exit Function
    If IsNumeric(v) Then ToNum = CDbl(v)
End Function

Private Function HasNum(ByVal v As Variant) As Boolean
    HasNum = False
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If VarType(v) = vbBoolean Then Exit Function
    If IsNumeric(v) Then HasNum = True
End Function

Private Function HasVal(ByVal v As Variant) As Boolean
    HasVal = False
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If Len(Trim$(CStr(v))) > 0 Then HasVal = True
End Function

Private Function CStr2(ByVal v As Variant) As String
    CStr2 = ""
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    CStr2 = CStr(v)
End Function

'--- accepts a serial, a real date or a date-like string; 0 means "not a date"
Private Function ToDate(ByVal v As Variant) As Double
    Dim s As String

    ToDate = 0
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function

    If VarType(v) = vbDate Then
        ToDate = CDbl(CDate(v))
        Exit Function
    End If

    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        If Len(s) = 0 Then Exit Function
        If IsDate(s) Then ToDate = CDbl(CDate(s))
        If ToDate = 0 Then ToDate = ParseWBCaption(s)
        Exit Function
    End If

    If VarType(v) = vbBoolean Then Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 Then
            If CDbl(v) < 80000 Then ToDate = CDbl(Int(CDbl(v)))
        End If
    End If
End Function

'--- "WB MAR-28 2027" (or a bare "MAR-28 2027") to a date serial; 0 means no match
Private Function ParseWBCaption(ByVal s As String) As Double
    Dim t As String
    Dim parts As Variant
    Dim mo As Long
    Dim dy As Long
    Dim yr As Long

    ParseWBCaption = 0
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    If UCase$(Left$(t, 2)) = "WB" Then
        If Len(t) > 2 Then
            If InStr("-_/ .", Mid$(t, 3, 1)) > 0 Then t = Trim$(Mid$(t, 4))
        End If
    End If

    t = Replace(t, "-", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, ",", " ")
    t = Replace(t, ".", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)
    If Len(t) = 0 Then Exit Function

    parts = Split(t, " ")
    If UBound(parts) < 2 Then Exit Function

    mo = MonthNum(CStr(parts(0)))
    If mo = 0 Then Exit Function
    If Not IsNumeric(CStr(parts(1))) Then Exit Function
    If Not IsNumeric(CStr(parts(2))) Then Exit Function

    dy = CLng(CDbl(CStr(parts(1))))
    yr = CLng(CDbl(CStr(parts(2))))
    If dy < 1 Then Exit Function
    If dy > 31 Then Exit Function
    If yr < 1900 Then Exit Function
    If yr > 2199 Then Exit Function

    ParseWBCaption = CDbl(DateSerial(yr, mo, dy))
End Function

Private Function MonthNum(ByVal s As String) As Long
    Dim t As String

    MonthNum = 0
    t = UCase$(Trim$(s))
    If Len(t) < 3 Then Exit Function
    t = Left$(t, 3)
    Select Case t
        Case "JAN": MonthNum = 1
        Case "FEB": MonthNum = 2
        Case "MAR": MonthNum = 3
        Case "APR": MonthNum = 4
        Case "MAY": MonthNum = 5
        Case "JUN": MonthNum = 6
        Case "JUL": MonthNum = 7
        Case "AUG": MonthNum = 8
        Case "SEP": MonthNum = 9
        Case "OCT": MonthNum = 10
        Case "NOV": MonthNum = 11
        Case "DEC": MonthNum = 12
        Case Else: MonthNum = 0
    End Select
End Function

Private Function FmtD(ByVal d As Double) As String
    If d <= 0 Then
        FmtD = ""
    Else
        FmtD = Format$(CDate(d), mCfgDateFmt)
    End If
End Function

Private Function JoinTok(ByVal cur As String, ByVal add As String) As String
    If Len(cur) = 0 Then
        JoinTok = add
    Else
        JoinTok = cur & FLD & add
    End If
End Function
