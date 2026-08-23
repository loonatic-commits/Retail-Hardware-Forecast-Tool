Attribute VB_Name = "PromoUpdateTemplateTests"
Option Explicit

'=====================================================================================
' PromoUpdateTemplateTests.bas
'
' Verifies that the account contribution percentage carried forward from the CY
' Template is consistent across models within a channel.
'
' The premise: a channel negotiates one funding split and applies it to every model it
' carries, so two models on the same channel should show the same Accn Cont. share of
' the total discount.  A model that disagrees points at a data error in the CY Template
' - a mistyped Accn Cont., a wrong Policy Price, or a promo row that belongs to another
' channel - and it would be carried straight into next year's template.
'
' Amazon is exempt: its funding structure changed, so its models legitimately differ.
' The exempt list is Config 'Pct Test Exempt Accounts', or Config 'Full Funding
' Accounts' when that is absent, or AMAZON when neither is set.
'
' This module is deliberately standalone - it re-reads the CY Template itself rather
' than calling into PromoUpdateTemplate, so a bug in the generator cannot hide a bug in
' the data.  It reads only; it never writes to an input sheet.
'
' Entry point: TestAccountContributionConsistency.
'
' Result sheet 'Pct Consistency Test' is deleted and rebuilt on every run:
'   CHANNEL SUMMARY   one row per channel, with the spread across its models
'   PRODUCT DETAIL    one row per product x channel, with its % and deviation
'=====================================================================================

'--- constants ------------------------------------------------------------------------
Private Const RESULT_SHEET As String = "Pct Consistency Test"
Private Const DEFAULT_TOL As Double = 0.005
Private Const TFLD As String = ";;"

'--- run state ------------------------------------------------------------------------
Private tStage As String
Private tTol As Double
Private tExempt As String

'--- CY template ----------------------------------------------------------------------
Private tCY() As Variant
Private tCYRows As Long
Private tCYCols As Long
Private tAlias As Object
Private tCol As Object
Private tCfg As Object

'--- resolved column positions --------------------------------------------------------
Private tcProd As Long
Private tcModel As Long
Private tcCust As Long
Private tcCustD As Long
Private tcPP As Long
Private tcNet As Long
Private tcAccn As Long
Private tcSpStart As Long

'--- derived --------------------------------------------------------------------------
Private tPctSum As Object
Private tPctCnt As Object
Private tModelOf As Object
Private tCustName As Object
Private tChanProds As Object

'--- median helper --------------------------------------------------------------------
Private tSortV() As Double
Private tSortN As Long

'--- result buffers -------------------------------------------------------------------
Private tSum() As Variant
Private tSumRow As Long
Private tDet() As Variant
Private tDetRow As Long
Private tDetCap As Long
Private tDetTmp() As Variant

'--- counters -------------------------------------------------------------------------
Private tChanTested As Long
Private tChanPassed As Long
Private tChanFailed As Long
Private tChanExempt As Long
Private tOutliers As Long
Private tNoData As Long


'=====================================================================================
' ENTRY POINT
'=====================================================================================
Public Sub TestAccountContributionConsistency()
    Dim wb As Workbook
    Dim savedCalc As Long
    Dim savedScreen As Boolean
    Dim savedEvents As Boolean
    Dim stateSaved As Boolean
    Dim msg As String

    On Error GoTo Fail
    tStage = "starting up"
    stateSaved = False

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Err.Raise vbObjectError + 601, , "There is no active workbook."
    End If

    savedCalc = Application.Calculation
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    tStage = "initialising"
    TestInit

    tStage = "reading the Config sheet"
    TestReadConfig wb

    tStage = "reading the CY Template"
    TestReadCY wb

    tStage = "comparing contribution percentages across models"
    TestCompare

    tStage = "writing the Pct Consistency Test sheet"
    TestWriteResults wb

    Application.Calculation = savedCalc
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen

    If tChanFailed = 0 Then
        msg = "PASS - every channel's models share one contribution %." & vbCrLf & vbCrLf
    Else
        msg = "FAIL - " & tChanFailed & " channel(s) disagree across models." & vbCrLf & vbCrLf
    End If
    msg = msg & "Channels tested:    " & tChanTested & vbCrLf
    msg = msg & "Channels passed:    " & tChanPassed & vbCrLf
    msg = msg & "Channels failed:    " & tChanFailed & vbCrLf
    msg = msg & "Channels exempt:    " & tChanExempt & vbCrLf
    msg = msg & "Outlier models:     " & tOutliers & vbCrLf
    msg = msg & "Models with no data:" & tNoData & vbCrLf
    msg = msg & "Tolerance:          " & Format$(tTol, "0.000%") & vbCrLf & vbCrLf
    msg = msg & "See the '" & RESULT_SHEET & "' sheet."

    If tChanFailed = 0 Then
        MsgBox msg, vbInformation, "Contribution % consistency"
    Else
        MsgBox msg, vbExclamation, "Contribution % consistency"
    End If
    Exit Sub

Fail:
    msg = "The contribution % test failed while " & tStage & "." & vbCrLf & vbCrLf
    msg = msg & "Error " & Err.Number & ": " & Err.Description
    If stateSaved Then
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
        Application.ScreenUpdating = savedScreen
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    MsgBox msg, vbCritical, "Contribution % consistency"
End Sub


'=====================================================================================
' SETUP
'=====================================================================================
Private Sub TestInit()
    Set tAlias = CreateObject("Scripting.Dictionary")
    Set tCol = CreateObject("Scripting.Dictionary")
    Set tPctSum = CreateObject("Scripting.Dictionary")
    Set tPctCnt = CreateObject("Scripting.Dictionary")
    Set tModelOf = CreateObject("Scripting.Dictionary")
    Set tCustName = CreateObject("Scripting.Dictionary")
    Set tChanProds = CreateObject("Scripting.Dictionary")

    tSumRow = 0
    tDetRow = 0
    tDetCap = 512
    ReDim tDet(1 To tDetCap, 1 To 7)

    tChanTested = 0
    tChanPassed = 0
    tChanFailed = 0
    tChanExempt = 0
    tOutliers = 0
    tNoData = 0

    TestAddAlias "Product ID", "product;product no;product number;material;sku;item;item id"
    TestAddAlias "Model Desc", "model;model description;model desc.;model name"
    TestAddAlias "Forecast Customer ID", "customer id;customer;forecast customer;cust id;channel;account"
    TestAddAlias "Forecast Customer Description", "customer description;customer desc;forecast customer desc;customer name"
    TestAddAlias "Policy Price", "policy;policy price."
    TestAddAlias "Net", "net price;net promo price;promo net price"
    TestAddAlias "Accn Cont.", "accn cont;accn;account cont;account cont.;accn contribution;account contribution;acct cont"
    TestAddAlias "Special Promo Start Date", "special promo start;promo start date;promo start"
End Sub

Private Sub TestAddAlias(ByVal canon As String, ByVal aliasList As String)
    Dim parts As Variant
    Dim i As Long
    Dim k As String

    k = TestNorm(canon)
    If Len(k) > 0 Then
        If Not tAlias.Exists(k) Then tAlias.Add k, canon
    End If
    If Len(aliasList) = 0 Then Exit Sub
    parts = Split(aliasList, ";")
    For i = LBound(parts) To UBound(parts)
        k = TestNorm(CStr(parts(i)))
        If Len(k) > 0 Then
            If Not tAlias.Exists(k) Then tAlias.Add k, canon
        End If
    Next i
End Sub

Private Sub TestReadConfig(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim arr As Variant
    Dim lr As Long
    Dim r As Long
    Dim k As String
    Dim v As String

    Set tCfg = CreateObject("Scripting.Dictionary")
    Set ws = TestFindSheet(wb, "Config")
    If Not ws Is Nothing Then
        lr = TestLastRow(ws)
        If lr >= 1 Then
            arr = ws.Range(ws.Cells(1, 1), ws.Cells(lr, 2)).Value2
            For r = 1 To UBound(arr, 1)
                k = TestNorm(TestStr(arr(r, 1)))
                If Len(k) > 0 Then
                    If Not tCfg.Exists(k) Then tCfg.Add k, TestStr(arr(r, 2))
                End If
            Next r
        End If
    End If

    tTol = DEFAULT_TOL
    v = TestCfg("Pct Test Tolerance", "")
    If Len(v) > 0 Then
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then tTol = CDbl(v)
        End If
    End If

    tExempt = TestCfg("Pct Test Exempt Accounts", "")
    If Len(tExempt) = 0 Then tExempt = TestCfg("Full Funding Accounts", "")
    If Len(tExempt) = 0 Then tExempt = "AMAZON"
    tExempt = UCase$(tExempt)
End Sub

Private Function TestCfg(ByVal nm As String, ByVal dflt As String) As String
    Dim k As String
    k = TestNorm(nm)
    TestCfg = dflt
    If tCfg Is Nothing Then Exit Function
    If tCfg.Exists(k) Then
        If Len(Trim$(CStr(tCfg(k)))) > 0 Then TestCfg = Trim$(CStr(tCfg(k)))
    End If
End Function


'=====================================================================================
' READ AND MEASURE
'=====================================================================================
Private Sub TestReadCY(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lr As Long
    Dim lc As Long
    Dim c As Long
    Dim r As Long
    Dim k As String
    Dim canon As String
    Dim prod As String
    Dim cust As String
    Dim custD As String
    Dim ck As String
    Dim key As String
    Dim pp As Double
    Dim nt As Double
    Dim ac As Double
    Dim disc As Double
    Dim skipRow As Boolean

    Set ws = TestFindSheet(wb, "CY Template")
    If ws Is Nothing Then
        Err.Raise vbObjectError + 602, , "The 'CY Template' sheet was not found."
    End If

    lr = TestLastRow(ws)
    lc = TestLastCol(ws)
    If lr < 2 Then
        Err.Raise vbObjectError + 603, , "The 'CY Template' sheet has no data rows."
    End If
    If lc < 2 Then
        Err.Raise vbObjectError + 604, , "The 'CY Template' sheet has fewer than two columns."
    End If

    tCY = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    tCYRows = lr
    tCYCols = lc

    For c = 1 To tCYCols
        k = TestNorm(TestStr(tCY(1, c)))
        If Len(k) > 0 Then
            If tAlias.Exists(k) Then
                canon = CStr(tAlias(k))
                If Not tCol.Exists(canon) Then tCol.Add canon, c
            End If
        End If
    Next c

    tcProd = TestReq("Product ID")
    tcPP = TestReq("Policy Price")
    tcNet = TestReq("Net")
    tcAccn = TestReq("Accn Cont.")
    tcModel = TestOpt("Model Desc")
    tcCust = TestOpt("Forecast Customer ID")
    tcCustD = TestOpt("Forecast Customer Description")
    tcSpStart = TestOpt("Special Promo Start Date")

    If tcCust = 0 Then
        If tcCustD = 0 Then
            Err.Raise vbObjectError + 605, , "The 'CY Template' sheet has no customer ID or description column, so channels cannot be identified."
        End If
    End If

    For r = 2 To tCYRows
        prod = UCase$(Trim$(TestStr(tCY(r, tcProd))))
        If Len(prod) > 0 Then
            cust = ""
            custD = ""
            If tcCust > 0 Then cust = Trim$(TestStr(tCY(r, tcCust)))
            If tcCustD > 0 Then custD = Trim$(TestStr(tCY(r, tcCustD)))
            If Len(cust) = 0 Then cust = custD
            If Len(custD) = 0 Then custD = cust

            If Len(cust) > 0 Then
                ck = UCase$(cust)
                key = prod & "|" & ck

                If Not tCustName.Exists(ck) Then tCustName.Add ck, custD
                If Not tChanProds.Exists(ck) Then tChanProds.Add ck, ""
                If InStr(1, TFLD & CStr(tChanProds(ck)) & TFLD, TFLD & prod & TFLD, vbTextCompare) = 0 Then
                    If Len(CStr(tChanProds(ck))) = 0 Then
                        tChanProds(ck) = prod
                    Else
                        tChanProds(ck) = CStr(tChanProds(ck)) & TFLD & prod
                    End If
                End If
                If tcModel > 0 Then
                    If Not tModelOf.Exists(prod) Then tModelOf.Add prod, Trim$(TestStr(tCY(r, tcModel)))
                End If

                pp = TestNum(tCY(r, tcPP))
                nt = TestNum(tCY(r, tcNet))
                ac = TestNum(tCY(r, tcAccn))
                disc = pp - nt

                '--- same exclusions the generator applies: no discount, and last year's
                '--- fully funded rows (Accn = 0 with a populated special promo start)
                If disc > 0.0000001 Then
                    skipRow = False
                    If Abs(ac) < 0.0000001 Then
                        If tcSpStart > 0 Then
                            If TestHasVal(tCY(r, tcSpStart)) Then skipRow = True
                        End If
                    End If
                    If Not skipRow Then
                        If tPctCnt.Exists(key) Then
                            tPctSum(key) = CDbl(tPctSum(key)) + ac / disc
                            tPctCnt(key) = CLng(tPctCnt(key)) + 1
                        Else
                            tPctSum.Add key, ac / disc
                            tPctCnt.Add key, CLng(1)
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Sub TestCompare()
    Dim chans As Variant
    Dim prods As Variant
    Dim ci As Long
    Dim pi As Long
    Dim ck As String
    Dim prod As String
    Dim key As String
    Dim pct As Double
    Dim med As Double
    Dim lo As Double
    Dim hi As Double
    Dim withData As Long
    Dim exempt As Boolean
    Dim verdict As String
    Dim status As String
    Dim dev As Double

    If tChanProds.Count = 0 Then
        Err.Raise vbObjectError + 607, , "No product x channel rows could be read from the 'CY Template', so there is nothing to compare."
    End If

    chans = tChanProds.Keys
    ReDim tSum(1 To tChanProds.Count, 1 To 6)

    For ci = LBound(chans) To UBound(chans)
        ck = CStr(chans(ci))
        prods = Split(CStr(tChanProds(ck)), TFLD)
        exempt = TestIsExempt(ck, CStr(tCustName(ck)))

        '--- collect this channel's per-model percentages
        tSortN = 0
        ReDim tSortV(1 To UBound(prods) - LBound(prods) + 1)
        For pi = LBound(prods) To UBound(prods)
            prod = CStr(prods(pi))
            If Len(prod) > 0 Then
                key = prod & "|" & ck
                If tPctCnt.Exists(key) Then
                    If CLng(tPctCnt(key)) > 0 Then
                        tSortN = tSortN + 1
                        tSortV(tSortN) = CDbl(tPctSum(key)) / CLng(tPctCnt(key))
                    End If
                End If
            End If
        Next pi

        withData = tSortN
        lo = 0
        hi = 0
        med = 0
        If withData > 0 Then
            TestSortValues
            lo = tSortV(1)
            hi = tSortV(tSortN)
            If tSortN Mod 2 = 1 Then
                med = tSortV((tSortN + 1) \ 2)
            Else
                med = (tSortV(tSortN \ 2) + tSortV(tSortN \ 2 + 1)) / 2
            End If
        End If

        If exempt Then
            verdict = "EXEMPT"
            tChanExempt = tChanExempt + 1
        ElseIf withData = 0 Then
            verdict = "NO DATA"
        ElseIf withData = 1 Then
            verdict = "PASS (single model)"
            tChanTested = tChanTested + 1
            tChanPassed = tChanPassed + 1
        ElseIf hi - lo <= tTol Then
            verdict = "PASS"
            tChanTested = tChanTested + 1
            tChanPassed = tChanPassed + 1
        Else
            verdict = "FAIL"
            tChanTested = tChanTested + 1
            tChanFailed = tChanFailed + 1
        End If

        tSumRow = tSumRow + 1
        tSum(tSumRow, 1) = CStr(tCustName(ck))
        tSum(tSumRow, 2) = withData
        tSum(tSumRow, 3) = lo
        tSum(tSumRow, 4) = hi
        tSum(tSumRow, 5) = hi - lo
        tSum(tSumRow, 6) = verdict

        '--- one detail row per model on this channel
        For pi = LBound(prods) To UBound(prods)
            prod = CStr(prods(pi))
            If Len(prod) > 0 Then
                key = prod & "|" & ck
                pct = 0
                status = "NO DATA"
                dev = 0
                If tPctCnt.Exists(key) Then
                    If CLng(tPctCnt(key)) > 0 Then
                        pct = CDbl(tPctSum(key)) / CLng(tPctCnt(key))
                        dev = pct - med
                        If exempt Then
                            status = "EXEMPT"
                        ElseIf Abs(dev) > tTol Then
                            status = "OUTLIER"
                            tOutliers = tOutliers + 1
                        Else
                            status = "OK"
                        End If
                    End If
                End If
                If status = "NO DATA" Then tNoData = tNoData + 1

                tDetRow = tDetRow + 1
                TestGrowDetail tDetRow
                tDet(tDetRow, 1) = CStr(tCustName(ck))
                tDet(tDetRow, 2) = prod
                If tModelOf.Exists(prod) Then
                    tDet(tDetRow, 3) = CStr(tModelOf(prod))
                Else
                    tDet(tDetRow, 3) = ""
                End If
                If tPctCnt.Exists(key) Then
                    tDet(tDetRow, 4) = CLng(tPctCnt(key))
                Else
                    tDet(tDetRow, 4) = 0
                End If
                If status = "NO DATA" Then
                    tDet(tDetRow, 5) = ""
                    tDet(tDetRow, 6) = ""
                Else
                    tDet(tDetRow, 5) = pct
                    tDet(tDetRow, 6) = dev
                End If
                tDet(tDetRow, 7) = status
            End If
        Next pi
    Next ci
End Sub

'--- insertion sort of the current channel's percentages
Private Sub TestSortValues()
    Dim i As Long
    Dim j As Long
    Dim v As Double

    For i = 2 To tSortN
        v = tSortV(i)
        j = i - 1
        Do While j >= 1
            If tSortV(j) <= v Then Exit Do
            tSortV(j + 1) = tSortV(j)
            j = j - 1
        Loop
        tSortV(j + 1) = v
    Next i
End Sub

Private Function TestIsExempt(ByVal cust As String, ByVal custDesc As String) As Boolean
    Dim toks As Variant
    Dim i As Long
    Dim t As String

    TestIsExempt = False
    If Len(tExempt) = 0 Then Exit Function
    toks = Split(tExempt, ";")
    For i = LBound(toks) To UBound(toks)
        t = Trim$(CStr(toks(i)))
        If Len(t) > 0 Then
            If Len(cust) > 0 Then
                If InStr(1, cust, t, vbTextCompare) > 0 Then
                    TestIsExempt = True
                    Exit Function
                End If
            End If
            If Len(custDesc) > 0 Then
                If InStr(1, custDesc, t, vbTextCompare) > 0 Then
                    TestIsExempt = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Sub TestGrowDetail(ByVal needed As Long)
    Dim i As Long
    Dim j As Long
    Dim newCap As Long

    If needed <= tDetCap Then Exit Sub
    newCap = tDetCap * 2
    If newCap < needed Then newCap = needed + 512
    ReDim tDetTmp(1 To newCap, 1 To 7)
    For i = 1 To tDetCap
        For j = 1 To 7
            tDetTmp(i, j) = tDet(i, j)
        Next j
    Next i
    tDet = tDetTmp
    Erase tDetTmp
    tDetCap = newCap
End Sub


'=====================================================================================
' RESULT SHEET
'=====================================================================================
Private Sub TestWriteResults(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim old As Worksheet
    Dim i As Long
    Dim j As Long
    Dim r As Long

    Set old = TestFindSheet(wb, RESULT_SHEET)
    If Not old Is Nothing Then
        Application.DisplayAlerts = False
        old.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = RESULT_SHEET

    ws.Cells(1, 1).Value = "Account contribution % consistency test"
    ws.Cells(1, 1).Font.Bold = True
    If tChanFailed = 0 Then
        ws.Cells(1, 3).Value = "PASS"
    Else
        ws.Cells(1, 3).Value = "FAIL (" & tChanFailed & " channel(s))"
    End If
    ws.Cells(1, 3).Font.Bold = True

    ws.Cells(2, 1).Value = "Tolerance"
    ws.Cells(2, 2).Value = Format$(tTol, "0.000%")
    ws.Cells(3, 1).Value = "Exempt accounts"
    ws.Cells(3, 2).Value = tExempt
    ws.Cells(4, 1).Value = "Run"
    ws.Cells(4, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    r = 6
    ws.Cells(r, 1).Value = "CHANNEL SUMMARY"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "Channel"
    ws.Cells(r, 2).Value = "Models with data"
    ws.Cells(r, 3).Value = "Min %"
    ws.Cells(r, 4).Value = "Max %"
    ws.Cells(r, 5).Value = "Spread"
    ws.Cells(r, 6).Value = "Result"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Font.Bold = True

    If tSumRow > 0 Then
        ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + tSumRow, 5)).NumberFormat = "0.000%"
        For i = 1 To tSumRow
            For j = 1 To 6
                ws.Cells(r + i, j).Value = tSum(i, j)
            Next j
        Next i
    End If

    r = r + tSumRow + 2
    ws.Cells(r, 1).Value = "PRODUCT DETAIL"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "Channel"
    ws.Cells(r, 2).Value = "Product ID"
    ws.Cells(r, 3).Value = "Model Desc"
    ws.Cells(r, 4).Value = "CY rows used"
    ws.Cells(r, 5).Value = "Contribution %"
    ws.Cells(r, 6).Value = "Deviation from channel median"
    ws.Cells(r, 7).Value = "Status"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True

    If tDetRow > 0 Then
        ws.Range(ws.Cells(r + 1, 5), ws.Cells(r + tDetRow, 6)).NumberFormat = "0.000%"
        ReDim tDetTmp(1 To tDetRow, 1 To 7)
        For i = 1 To tDetRow
            For j = 1 To 7
                tDetTmp(i, j) = tDet(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + tDetRow, 7)).Value = tDetTmp
        Erase tDetTmp
    End If

    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 16
    ws.Columns(6).ColumnWidth = 28
    ws.Columns(7).ColumnWidth = 12

    On Error Resume Next
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub


'=====================================================================================
' UTILITIES
'=====================================================================================
Private Function TestReq(ByVal canon As String) As Long
    If Not tCol.Exists(canon) Then
        Err.Raise vbObjectError + 606, , "The 'CY Template' sheet has no '" & canon & "' column."
    End If
    TestReq = CLng(tCol(canon))
End Function

Private Function TestOpt(ByVal canon As String) As Long
    If tCol.Exists(canon) Then
        TestOpt = CLng(tCol(canon))
    Else
        TestOpt = 0
    End If
End Function

Private Function TestNorm(ByVal s As String) As String
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
    TestNorm = Trim$(out)
End Function

Private Function TestStr(ByVal v As Variant) As String
    TestStr = ""
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    TestStr = CStr(v)
End Function

Private Function TestNum(ByVal v As Variant) As Double
    TestNum = 0
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If VarType(v) = vbBoolean Then Exit Function
    If IsNumeric(v) Then TestNum = CDbl(v)
End Function

Private Function TestHasVal(ByVal v As Variant) As Boolean
    TestHasVal = False
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If IsError(v) Then Exit Function
    If Len(Trim$(CStr(v))) > 0 Then TestHasVal = True
End Function

Private Function TestFindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Dim target As String

    target = TestNorm(nm)
    For Each ws In wb.Worksheets
        If TestNorm(ws.Name) = target Then
            Set TestFindSheet = ws
            Exit Function
        End If
    Next ws
    Set TestFindSheet = Nothing
End Function

Private Function TestLastRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        TestLastRow = 0
    Else
        TestLastRow = f.Row
    End If
End Function

Private Function TestLastCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        TestLastCol = 0
    Else
        TestLastCol = f.Column
    End If
End Function
