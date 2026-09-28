```vba
Option Explicit

'============================================================
' ENTRY POINT - MAIN CONTROL MENU
'============================================================
Sub MainMenu()
    Dim frm As frmReportMenu
    Set frm = New frmReportMenu

    frm.lstOptions.Clear
    frm.lstOptions.AddItem "Generate Factoring Report (> 4,999 EUR)"
    frm.lstOptions.AddItem "Generate Individual Company Code Tabs (> 4,999 EUR)"
    frm.lstOptions.AddItem "Reset Tracker (clears reports, history, CC tabs)"
    frm.lstOptions.AddItem "Export a Single Report to a New File"
    frm.lstOptions.AddItem "Export All Reports to a Folder"

    frm.Show

    If frm.UserCancelled Then
        Unload frm
        Exit Sub
    End If

    Dim sel As Long
    sel = frm.SelectedIndex
    Unload frm

    Select Case sel
        Case 0: GenerateSummaryReport
        Case 1: GenerateCompanyCodeTabs
        Case 2: ResetTracker
        Case 3: ExportGeneratedReport
        Case 4: ExportAllGeneratedReports
    End Select
End Sub

'============================================================
' HELPER - CHECK IF CC IS IN THE FACTORING SCOPE
'============================================================
Function IsFactoringCC(cc As String) As Boolean
    Dim ccUpper As String
    ccUpper = UCase(Trim(cc))
    Select Case ccUpper
        Case "BNL", "ISG", "SFR", "IUK", "AMI", "ITA", "VNL"
            IsFactoringCC = True
        Case Else
            IsFactoringCC = False
    End Select
End Function

'============================================================
' HELPER - PROTECTED MASTER DATA SHEETS FOR SYSTEM RESET
'============================================================
Function IsProtectedSheet(sheetName As String) As Boolean
    Select Case UCase(Trim(sheetName))
        Case "BW_SUMMARY", "FACTORINGPARTNERS", "ATRADIUSPARTNERS", "RESPONSIBLE", "HOME", "ICONLIBRARY", "GUIDE"
            IsProtectedSheet = True
        Case Else
            IsProtectedSheet = False
    End Select
End Function

'============================================================
' NORMALIZE CUSTOMER NUMBER - STRIPS LEADING ZEROS
'============================================================
Function NormalizeCustNo(v As Variant) As String
    Dim s As String
    s = Trim(CStr(v))
    If IsNumeric(s) And s <> "" Then
        NormalizeCustNo = CStr(CLng(s))
    Else
        NormalizeCustNo = s
    End If
End Function

'============================================================
' SAFELY CONVERT A CELL VALUE TO DOUBLE - RETURNS 0 IF NOT NUMERIC
'============================================================
Function SafeDouble(v As Variant) As Double
    If IsNumeric(v) And Trim(CStr(v)) <> "" Then
        SafeDouble = CDbl(v)
    Else
        SafeDouble = 0
    End If
End Function

'============================================================
' CHECK IF A CELL HOLDS A NON-NUMERIC / MIXED-CURRENCY MARKER
'============================================================
Function IsMixedCurrency(v As Variant) As Boolean
    Dim s As String
    s = UCase(Trim(CStr(v)))
    If s = "" Then
        IsMixedCurrency = False
    ElseIf IsNumeric(v) Then
        IsMixedCurrency = False
    Else
        IsMixedCurrency = True
    End If
End Function

'============================================================
' FILL DOWN COMPANY CODE (WITH ORPHAN ROW SELF-HEALING)
'============================================================
Sub FillDownCompanyCode(ws As Worksheet, lastRow As Long)
    Dim i As Long, lastCC As String
    lastCC = ""

    For i = 3 To lastRow
        If ws.Cells(i, "A").Value <> "" Then
            lastCC = ws.Cells(i, "A").Value
        ElseIf ws.Cells(i, "C").Value <> "" And ws.Cells(i, "C").Value <> "Result" Then
            ws.Cells(i, "A").Value = lastCC
        End If
    Next i

    Dim custCC As Object
    Set custCC = CreateObject("Scripting.Dictionary")
    Dim custKey As String

    For i = 3 To lastRow
        custKey = Trim(ws.Cells(i, "C").Value)
        If custKey <> "" And custKey <> "Result" And ws.Cells(i, "A").Value <> "" Then
            If Not custCC.Exists(custKey) Then
                custCC(custKey) = ws.Cells(i, "A").Value
            End If
        End If
    Next i

    For i = 3 To lastRow
        custKey = Trim(ws.Cells(i, "C").Value)
        If ws.Cells(i, "A").Value = "" And custKey <> "" And custKey <> "Result" Then
            If custCC.Exists(custKey) Then
                ws.Cells(i, "A").Value = custCC(custKey)
            End If
        End If
    Next i
End Sub

'============================================================
' CREATE A DISPOSABLE HIDDEN COPY OF BW_SUMMARY TO WORK ON
'============================================================
Function CreateTempSourceCopy(wsOriginal As Worksheet) As Worksheet
    Dim wsTemp As Worksheet
    Dim lastR As Long, lastC As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("TempBW_Working").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = "TempBW_Working"
    wsTemp.Visible = xlSheetHidden

    lastR = wsOriginal.Cells(wsOriginal.Rows.Count, "C").End(xlUp).Row
    lastC = wsOriginal.Cells(3, wsOriginal.Columns.Count).End(xlToLeft).Column
    If lastC < 15 Then lastC = 15

    wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(lastR, lastC)).Value = _
        wsOriginal.Range(wsOriginal.Cells(1, 1), wsOriginal.Cells(lastR, lastC)).Value

    wsTemp.AutoFilterMode = False
    wsTemp.Rows.Hidden = False

    Set CreateTempSourceCopy = wsTemp
End Function

'============================================================
' GET OR CREATE THE PERSISTENT FACTORING REPORT SHEET
'============================================================
Function GetOrCreateReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Factoring Report")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Factoring Report"
        ws.Outline.SummaryRow = xlAbove
    End If
    Set GetOrCreateReportSheet = ws
End Function

'============================================================
' FIND WHERE THE NEXT WEEKLY BLOCK SHOULD START
'============================================================
Function GetNextBlockStartRow(ws As Worksheet) As Long
    Dim lastUsed As Long
    lastUsed = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastUsed = 1 And ws.Cells(1, "A").Value = "" Then
        GetNextBlockStartRow = 1
    Else
        GetNextBlockStartRow = lastUsed + 3
    End If
End Function

'============================================================
' OPTION 1 - FACTORING REPORT (> 4,999 EUR)
'============================================================
Sub GenerateSummaryReport()
    Dim wsSrcOriginal As Worksheet, wsSrc As Worksheet, wsRep As Worksheet, wsArch As Worksheet
    Dim wsFactor As Worksheet, wsAtradius As Worksheet
    Dim lastRow As Long, i As Long, r As Long, blockStart As Long, headerRow As Long
    Dim threshold As Double: threshold = 4999
    Dim totalCustomers As Long, totalExposure As Double, newThisWeek As Long
    Dim custKey As String
    Dim prevList As Object
    Set prevList = CreateObject("Scripting.Dictionary")

    Dim factorSet As Object, atradiusSet As Object
    Set factorSet = CreateObject("Scripting.Dictionary")
    Set atradiusSet = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanFail

    On Error Resume Next
    Set wsSrcOriginal = ThisWorkbook.Sheets("BW_Summary")
    On Error GoTo 0
    If wsSrcOriginal Is Nothing Then
        MsgBox "Sheet 'BW_Summary' not found.", vbCritical
        Exit Sub
    End If

    Set wsSrc = CreateTempSourceCopy(wsSrcOriginal)

    On Error Resume Next
    Set wsFactor = ThisWorkbook.Sheets("FactoringPartners")
    On Error GoTo 0
    If Not wsFactor Is Nothing Then
        Dim factLast As Long, fj As Long, factCustOnly As String, factKey As String
        factLast = wsFactor.Cells(wsFactor.Rows.Count, "A").End(xlUp).Row
        For fj = 2 To factLast
            factCustOnly = NormalizeCustNo(wsFactor.Cells(fj, "A").Value)
            If factCustOnly <> "" Then
                factorSet(factCustOnly) = True
                factKey = factCustOnly & "|" & UCase(Trim(wsFactor.Cells(fj, "B").Value))
                factorSet(factKey) = True
            End If
        Next fj
    End If

    On Error Resume Next
    Set wsAtradius = ThisWorkbook.Sheets("AtradiusPartners")
    On Error GoTo 0
    If Not wsAtradius Is Nothing Then
        Dim atrLast As Long, aj As Long, atrKey As String, atrCustOnly As String
        atrLast = wsAtradius.Cells(wsAtradius.Rows.Count, "A").End(xlUp).Row
        For aj = 2 To atrLast
            atrCustOnly = NormalizeCustNo(wsAtradius.Cells(aj, "A").Value)
            If atrCustOnly <> "" Then
                atradiusSet(atrCustOnly) = True
                atrKey = atrCustOnly & "|" & UCase(Trim(wsAtradius.Cells(aj, "B").Value))
                atradiusSet(atrKey) = True
            End If
        Next aj
    End If

    On Error Resume Next
    Set wsArch = ThisWorkbook.Sheets("RiskArchive")
    On Error GoTo 0
    If Not wsArch Is Nothing Then
        Dim archLast As Long, j As Long
        archLast = wsArch.Cells(wsArch.Rows.Count, "A").End(xlUp).Row
        For j = 2 To archLast
            prevList(wsArch.Cells(j, "A").Value & "|" & wsArch.Cells(j, "B").Value) = True
        Next j
    End If

    Dim lastRowA As Long, lastRowC As Long
    lastRowA = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowC = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowC)

    FillDownCompanyCode wsSrc, lastRow
    wsSrc.Range("A4:O" & lastRow).Sort Key1:=wsSrc.Range("A4"), Order1:=xlAscending, Header:=xlNo

    Set wsRep = GetOrCreateReportSheet()
    blockStart = GetNextBlockStartRow(wsRep)

    Dim weekNum As Integer
    weekNum = Application.WorksheetFunction.weekNum(Date, 21)
    With wsRep
        .Range(.Cells(blockStart, "A"), .Cells(blockStart, "J")).Merge
        .Cells(blockStart, "A").Value = "WEEK " & weekNum & " - Report generated: " & Format(Now, "dd/mm/yyyy hh:nn")
        .Cells(blockStart, "A").Font.Bold = True
        .Cells(blockStart, "A").Font.Size = 12
        .Cells(blockStart, "A").Interior.Color = RGB(44, 62, 80)
        .Cells(blockStart, "A").Font.Color = RGB(236, 240, 241)
        .Cells(blockStart, "A").HorizontalAlignment = xlCenter
        .Rows(blockStart).RowHeight = 22
        .Rows(blockStart).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Rows(blockStart).Borders(xlEdgeTop).Weight = xlThick
    End With

    headerRow = blockStart + 3
    With wsRep
        .Range(.Cells(headerRow, "A"), .Cells(headerRow, "J")).Value = _
            Array("Company Code", "Customer", "Customer Name", "Overdue 60-89", "Overdue 90-179", "Overdue in Risk", "Total Overdue", "Status", "Factoring", "Atradius")
        .Range(.Cells(headerRow, "A"), .Cells(headerRow, "J")).Font.Bold = True
        .Range(.Cells(headerRow, "A"), .Cells(headerRow, "J")).Interior.Color = RGB(68, 84, 106)
        .Range(.Cells(headerRow, "A"), .Cells(headerRow, "J")).Font.Color = RGB(255, 255, 255)
    End With

    r = headerRow + 1
    totalCustomers = 0: totalExposure = 0: newThisWeek = 0

    Dim currentCC As String, ccStartRow As Long
    Dim ccTotal60 As Double, ccTotal90 As Double, ccTotalRisk As Double, ccTotalAll As Double, ccCustCount As Long
    currentCC = ""

    For i = 4 To lastRow + 1
        Dim ccVal As String, custVal As String, custName As String
        Dim d60 As Double, d90 As Double, totAmt As Double

        If i <= lastRow Then
            ccVal = wsSrc.Cells(i, "A").Value
            custVal = wsSrc.Cells(i, "C").Value
            custName = wsSrc.Cells(i, "D").Value
            d60 = SafeDouble(wsSrc.Cells(i, "K").Value)
            d90 = SafeDouble(wsSrc.Cells(i, "L").Value)
            totAmt = SafeDouble(wsSrc.Cells(i, "O").Value)
        Else
            ccVal = "###END###"
        End If

        If custVal = "Result" Then GoTo ContinueLoop

        If ccVal <> "###END###" And ccVal <> "" Then
            If Not IsFactoringCC(ccVal) Then GoTo ContinueLoop
        End If

        If ccVal <> currentCC Then
            If currentCC <> "" And ccCustCount > 0 Then
                wsRep.Rows(ccStartRow).Insert Shift:=xlDown
                With wsRep
                    .Cells(ccStartRow, "A").Value = currentCC
                    .Cells(ccStartRow, "C").Value = ccCustCount & " customer(s) at risk"
                    .Cells(ccStartRow, "F").Value = ccTotalRisk
                    .Cells(ccStartRow, "G").Value = ccTotalAll
                    With .Range(.Cells(ccStartRow, "A"), .Cells(ccStartRow, "J"))
                        .Font.Bold = True
                        .Font.Color = RGB(0, 0, 0)
                        .Interior.Color = RGB(213, 219, 226)
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).Weight = xlThick
                        .Borders(xlEdgeTop).Color = RGB(52, 73, 94)
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeBottom).Color = RGB(52, 73, 94)
                    End With
                    .Rows(ccStartRow).RowHeight = 18
                End With

                AddSleekEmailButton wsRep, ccStartRow, currentCC

                r = r + 1
                If r - 1 >= ccStartRow + 1 Then
                    wsRep.Rows(ccStartRow + 1 & ":" & r - 1).Group
                End If
            End If
            currentCC = ccVal
            ccStartRow = r
            ccTotal60 = 0: ccTotal90 = 0: ccTotalRisk = 0: ccTotalAll = 0: ccCustCount = 0
        End If

        If ccVal <> "###END###" Then
            If (d60 > threshold Or d90 > threshold) And custVal <> "" Then
                wsRep.Cells(r, "A").Value = ccVal
                wsRep.Cells(r, "B").Value = custVal
                wsRep.Cells(r, "C").Value = "     " & custName

                If IsMixedCurrency(wsSrc.Cells(i, "L").Value) Or IsMixedCurrency(wsSrc.Cells(i, "O").Value) Then
                    wsRep.Cells(r, "C").Value = wsRep.Cells(r, "C").Value & "  MIXED CURRENCY"
                    wsRep.Cells(r, "C").Font.Italic = True
                    wsRep.Cells(r, "C").Font.Color = RGB(230, 81, 0)
                End If

                wsRep.Cells(r, "D").Value = d60
                wsRep.Cells(r, "E").Value = d90
                wsRep.Cells(r, "F").Value = d60 + d90
                wsRep.Cells(r, "G").Value = totAmt

                custKey = ccVal & "|" & custVal
                If prevList.Exists(custKey) Then
                    wsRep.Cells(r, "H").Value = "Existing"
                Else
                    wsRep.Cells(r, "H").Value = "NEW"
                    newThisWeek = newThisWeek + 1
                End If

                Dim custNorm As String, ccKey As String
                custNorm = NormalizeCustNo(custVal)
                ccKey = custNorm & "|" & UCase(Trim(ccVal))

                If factorSet.Exists(custNorm) Or factorSet.Exists(ccKey) Then
                    wsRep.Cells(r, "I").Value = "FACTOR"
                    wsRep.Cells(r, "I").Font.Bold = True
                    wsRep.Cells(r, "I").Font.Color = RGB(13, 71, 161)
                    wsRep.Cells(r, "I").HorizontalAlignment = xlCenter
                End If

                If atradiusSet.Exists(custNorm) Or atradiusSet.Exists(ccKey) Then
                    wsRep.Cells(r, "J").Value = "ATRADIUS"
                    wsRep.Cells(r, "J").Font.Bold = True
                    wsRep.Cells(r, "J").Font.Color = RGB(56, 142, 60)
                    wsRep.Cells(r, "J").HorizontalAlignment = xlCenter
                End If

                ccTotal60 = ccTotal60 + d60
                ccTotal90 = ccTotal90 + d90
                ccTotalRisk = ccTotalRisk + (d60 + d90)
                ccTotalAll = ccTotalAll + totAmt
                ccCustCount = ccCustCount + 1
                totalCustomers = totalCustomers + 1
                totalExposure = totalExposure + (d60 + d90)
                r = r + 1
            End If
        End If
ContinueLoop:
    Next i

    Dim lastDataRow As Long
    lastDataRow = r - 1

    With wsRep
        .Columns("A:J").AutoFit
        .Columns("K").ColumnWidth = 16
        If lastDataRow >= headerRow + 1 Then
            .Range("D" & headerRow + 1 & ":D" & lastDataRow).Interior.Color = RGB(255, 235, 156)
            .Range("E" & headerRow + 1 & ":E" & lastDataRow).Interior.Color = RGB(255, 199, 206)
            .Range("F" & headerRow + 1 & ":F" & lastDataRow).Interior.Color = RGB(224, 236, 255)
            .Range("D" & headerRow + 1 & ":G" & lastDataRow).NumberFormat = "#,##0.00"
            With .Range("A" & headerRow & ":J" & lastDataRow).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(200, 200, 200)
            End With
            .Range("A" & headerRow & ":J" & lastDataRow).BorderAround _
                Weight:=xlMedium, Color:=RGB(52, 73, 94)
        End If
    End With

    Dim k As Long
    For k = headerRow + 1 To lastDataRow
        If wsRep.Cells(k, "B").Value <> "" Then
            wsRep.Range(wsRep.Cells(k, "A"), wsRep.Cells(k, "J")).Font.Bold = False
            If wsRep.Cells(k, "H").Value = "NEW" Then
                wsRep.Cells(k, "H").Font.Bold = True
                wsRep.Cells(k, "H").Font.Color = RGB(198, 40, 40)
            End If
        End If
    Next k

    BuildKPIDashboardAt wsRep, blockStart + 1, totalCustomers, totalExposure, newThisWeek
    SaveRiskArchive wsRep, headerRow + 1, lastDataRow

    If blockStart > 1 Then
        wsRep.Rows(blockStart & ":" & (lastDataRow + 2)).Cut
        wsRep.Rows("1:1").Insert Shift:=xlDown
    End If

    Application.DisplayAlerts = False
    wsSrc.Delete
    Application.DisplayAlerts = True

    EnsureTabOrder

    MsgBox totalCustomers & " factoring customers found exceeding threshold." & vbNewLine & _
           "Total risk exposure (60-179 Days): " & Format(totalExposure, "#,##0.00") & " EUR" & vbNewLine & _
           newThisWeek & " new risk customers this week.", vbInformation, "Factoring Report Complete"
    Exit Sub

CleanFail:
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsSrc Is Nothing Then wsSrc.Delete
    Application.DisplayAlerts = True
    MsgBox "An error occurred while generating the report: " & Err.Description, vbCritical, "Process Error"
End Sub

'============================================================
' OPTION 2 - DETAILED CC RISK TABS (> 4,999 EUR)
'============================================================
Sub GenerateCompanyCodeTabs()
    Dim wsSrcOriginal As Worksheet, wsSrc As Worksheet, wsNew As Worksheet
    Dim wsFactor As Worksheet, wsAtradius As Worksheet, wsArch As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim threshold As Double: threshold = 4999
    Dim uniqueCC As New Collection
    Dim cc As Variant
    Dim noRiskList As String, tabDate As String
    Dim hasActiveRisk As Boolean

    Dim prevList As Object
    Set prevList = CreateObject("Scripting.Dictionary")
    Dim factorSet As Object, atradiusSet As Object
    Set factorSet = CreateObject("Scripting.Dictionary")
    Set atradiusSet = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanFail

    On Error Resume Next
    Set wsSrcOriginal = ThisWorkbook.Sheets("BW_Summary")
    On Error GoTo 0
    If wsSrcOriginal Is Nothing Then
        MsgBox "Sheet 'BW_Summary' not found.", vbCritical
        Exit Sub
    End If

    Set wsSrc = CreateTempSourceCopy(wsSrcOriginal)

    On Error Resume Next
    Set wsFactor = ThisWorkbook.Sheets("FactoringPartners")
    On Error GoTo 0
    If Not wsFactor Is Nothing Then
        Dim factLast As Long, fj As Long, factCustOnly As String, factKey As String
        factLast = wsFactor.Cells(wsFactor.Rows.Count, "A").End(xlUp).Row
        For fj = 2 To factLast
            factCustOnly = NormalizeCustNo(wsFactor.Cells(fj, "A").Value)
            If factCustOnly <> "" Then
                factorSet(factCustOnly) = True
                factKey = factCustOnly & "|" & UCase(Trim(wsFactor.Cells(fj, "B").Value))
                factorSet(factKey) = True
            End If
        Next fj
    End If

    On Error Resume Next
    Set wsAtradius = ThisWorkbook.Sheets("AtradiusPartners")
    On Error GoTo 0
    If Not wsAtradius Is Nothing Then
        Dim atrLast As Long, aj As Long, atrKey As String, atrCustOnly As String
        atrLast = wsAtradius.Cells(wsAtradius.Rows.Count, "A").End(xlUp).Row
        For aj = 2 To atrLast
            atrCustOnly = NormalizeCustNo(wsAtradius.Cells(aj, "A").Value)
            If atrCustOnly <> "" Then
                atradiusSet(atrCustOnly) = True
                atrKey = atrCustOnly & "|" & UCase(Trim(wsAtradius.Cells(aj, "B").Value))
                atradiusSet(atrKey) = True
            End If
        Next aj
    End If

    On Error Resume Next
    Set wsArch = ThisWorkbook.Sheets("RiskArchive")
    On Error GoTo 0
    If Not wsArch Is Nothing Then
        Dim archLast As Long, ajArch As Long
        archLast = wsArch.Cells(wsArch.Rows.Count, "A").End(xlUp).Row
        For ajArch = 2 To archLast
            prevList(wsArch.Cells(ajArch, "A").Value & "|" & wsArch.Cells(ajArch, "B").Value) = True
        Next ajArch
    End If

    Dim lastRowA As Long, lastRowC As Long
    lastRowA = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowC = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowC)

    FillDownCompanyCode wsSrc, lastRow

    On Error Resume Next
    For i = 3 To lastRow
        Dim ccName As String
        ccName = UCase(Trim(wsSrc.Cells(i, "A").Value))
        If ccName <> "" And ccName <> "RESULT" And ccName <> "COMPANY CODE" Then
            uniqueCC.Add ccName, CStr(ccName)
        End If
    Next i
    On Error GoTo 0

    tabDate = Format(Date, "dd.mm.yyyy")
    noRiskList = ""

    For Each cc In uniqueCC
        hasActiveRisk = False

        For i = 4 To lastRow
            If UCase(Trim(wsSrc.Cells(i, "A").Value)) = UCase(Trim(cc)) And wsSrc.Cells(i, "C").Value <> "Result" Then
                If (SafeDouble(wsSrc.Cells(i, "K").Value) > threshold) Or (SafeDouble(wsSrc.Cells(i, "L").Value) > threshold) Then
                    hasActiveRisk = True
                    Exit For
                End If
            End If
        Next i

        If Not hasActiveRisk Then
            noRiskList = noRiskList & cc & ", "
        Else
            Dim sheetName As String
            sheetName = Left("R AGING_" & cc & "_" & tabDate, 31)

            Application.DisplayAlerts = False
            On Error Resume Next
            ThisWorkbook.Sheets(sheetName).Delete
            On Error GoTo 0
            Application.DisplayAlerts = True

            Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsNew.Name = sheetName
            wsNew.Outline.SummaryRow = xlAbove

            Dim ccCustCount As Long, ccNewCount As Long
            Dim ccTotal60 As Double, ccTotal90 As Double, ccTotalRisk As Double, ccTotalAll As Double
            ccCustCount = 0: ccNewCount = 0
            ccTotal60 = 0: ccTotal90 = 0: ccTotalRisk = 0: ccTotalAll = 0

            For i = 4 To lastRow
                If UCase(Trim(wsSrc.Cells(i, "A").Value)) = UCase(Trim(cc)) And wsSrc.Cells(i, "C").Value <> "Result" Then
                    Dim d60Check As Double, d90Check As Double, totCheck As Double
                    d60Check = SafeDouble(wsSrc.Cells(i, "K").Value)
                    d90Check = SafeDouble(wsSrc.Cells(i, "L").Value)
                    totCheck = SafeDouble(wsSrc.Cells(i, "O").Value)

                    If (d60Check > threshold) Or (d90Check > threshold) Then
                        ccTotal60 = ccTotal60 + d60Check
                        ccTotal90 = ccTotal90 + d90Check
                        ccTotalRisk = ccTotalRisk + (d60Check + d90Check)
                        ccTotalAll = ccTotalAll + totCheck
                        ccCustCount = ccCustCount + 1

                        If Not prevList.Exists(cc & "|" & wsSrc.Cells(i, "C").Value) Then
                            ccNewCount = ccNewCount + 1
                        End If
                    End If
                End If
            Next i

            Dim weekNum As Integer
            weekNum = Application.WorksheetFunction.weekNum(Date, 21)
            With wsNew
                .Range("A1:J1").Merge
                .Cells(1, "A").Value = "WEEK " & weekNum & " - " & cc & " Risk Sheet generated: " & Format(Now, "dd/mm/yyyy hh:nn")
                .Cells(1, "A").Font.Bold = True
                .Cells(1, "A").Font.Size = 12
                .Cells(1, "A").Interior.Color = RGB(44, 62, 80)
                .Cells(1, "A").Font.Color = RGB(236, 240, 241)
                .Cells(1, "A").HorizontalAlignment = xlCenter
                .Rows(1).RowHeight = 22
                .Rows(1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Rows(1).Borders(xlEdgeTop).Weight = xlThick
            End With

            BuildKPIDashboardAt wsNew, 2, ccCustCount, ccTotalRisk, ccNewCount

            With wsNew
                .Range("A5:J5").Value = Array("Company Code", "Customer", "Customer Name", "Overdue 60-89", "Overdue 90-179", "Overdue in Risk", "Total Overdue", "Status", "Factoring", "Atradius")
                .Range("A5:J5").Font.Bold = True
                .Range("A5:J5").Interior.Color = RGB(68, 84, 106)
                .Range("A5:J5").Font.Color = RGB(255, 255, 255)
            End With

            With wsNew
                .Cells(6, "A").Value = cc
                .Cells(6, "C").Value = ccCustCount & " customer(s) at risk"
                .Cells(6, "F").Value = ccTotalRisk
                .Cells(6, "G").Value = ccTotalAll
                With .Range("A6:J6")
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .Interior.Color = RGB(213, 219, 226)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlThick
                    .Borders(xlEdgeTop).Color = RGB(52, 73, 94)
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).Color = RGB(52, 73, 94)
                End With
                .Rows(6).RowHeight = 18
            End With

            r = 7
            For i = 4 To lastRow
                If UCase(Trim(wsSrc.Cells(i, "A").Value)) = UCase(Trim(cc)) And wsSrc.Cells(i, "C").Value <> "Result" Then
                    Dim d60 As Double, d90 As Double, totAmt As Double
                    d60 = SafeDouble(wsSrc.Cells(i, "K").Value)
                    d90 = SafeDouble(wsSrc.Cells(i, "L").Value)
                    totAmt = SafeDouble(wsSrc.Cells(i, "O").Value)

                    If (d60 > threshold) Or (d90 > threshold) Then
                        wsNew.Cells(r, 1).Value = cc
                        wsNew.Cells(r, 2).Value = wsSrc.Cells(i, "C").Value
                        wsNew.Cells(r, 3).Value = "     " & wsSrc.Cells(i, "D").Value

                        If IsMixedCurrency(wsSrc.Cells(i, "L").Value) Or IsMixedCurrency(wsSrc.Cells(i, "O").Value) Then
                            wsNew.Cells(r, 3).Value = wsNew.Cells(r, 3).Value & "  MIXED CURRENCY"
                            wsNew.Cells(r, 3).Font.Italic = True
                            wsNew.Cells(r, 3).Font.Color = RGB(230, 81, 0)
                        End If

                        wsNew.Cells(r, 4).Value = d60
                        wsNew.Cells(r, 5).Value = d90
                        wsNew.Cells(r, 6).Value = d60 + d90
                        wsNew.Cells(r, 7).Value = totAmt

                        Dim custNoStr As String, custKeyStr As String
                        custNoStr = wsSrc.Cells(i, "C").Value
                        custKeyStr = cc & "|" & custNoStr

                        If prevList.Exists(custKeyStr) Then
                            wsNew.Cells(r, 8).Value = "Existing"
                        Else
                            wsNew.Cells(r, 8).Value = "NEW"
                        End If

                        Dim custNorm As String, ccKey As String
                        custNorm = NormalizeCustNo(custNoStr)
                        ccKey = custNorm & "|" & UCase(Trim(cc))

                        If factorSet.Exists(custNorm) Or factorSet.Exists(ccKey) Then
                            wsNew.Cells(r, 9).Value = "FACTOR"
                            wsNew.Cells(r, 9).Font.Bold = True
                            wsNew.Cells(r, 9).Font.Color = RGB(13, 71, 161)
                            wsNew.Cells(r, 9).HorizontalAlignment = xlCenter
                        End If

                        If atradiusSet.Exists(custNorm) Or atradiusSet.Exists(ccKey) Then
                            wsNew.Cells(r, 10).Value = "ATRADIUS"
                            wsNew.Cells(r, 10).Font.Bold = True
                            wsNew.Cells(r, 10).Font.Color = RGB(56, 142, 60)
                            wsNew.Cells(r, 10).HorizontalAlignment = xlCenter
                        End If
                        r = r + 1
                    End If
                End If
            Next i

            With wsNew
                .Columns("A:J").AutoFit
                .Range("D7:D" & (r - 1)).Interior.Color = RGB(255, 235, 156)
                .Range("E7:E" & (r - 1)).Interior.Color = RGB(255, 199, 206)
                .Range("F7:F" & (r - 1)).Interior.Color = RGB(224, 236, 255)
                .Range("D7:G" & (r - 1)).NumberFormat = "#,##0.00"
                With .Range("A5:J" & (r - 1)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(200, 200, 200)
                End With
                .Range("A5:J" & (r - 1)).BorderAround Weight:=xlMedium, Color:=RGB(52, 73, 94)

                Dim rowIdx As Long
                For rowIdx = 7 To r - 1
                    .Range(.Cells(rowIdx, "A"), .Cells(rowIdx, "J")).Font.Bold = False
                    If .Cells(rowIdx, "H").Value = "NEW" Then
                        .Cells(rowIdx, "H").Font.Bold = True
                        .Cells(rowIdx, "H").Font.Color = RGB(198, 40, 40)
                    End If
                Next rowIdx

                If r > 7 Then
                    .Rows("7:" & (r - 1)).Group
                End If
            End With
        End If
    Next cc

    Application.DisplayAlerts = False
    wsSrc.Delete
    Application.DisplayAlerts = True

    EnsureTabOrder

    If noRiskList <> "" Then
        noRiskList = Left(noRiskList, Len(noRiskList) - 2)
        MsgBox "Company Code Tabs generation complete!" & vbNewLine & vbNewLine & _
               "The following Company Codes did not meet the " & Format(threshold, "#,##0") & " EUR threshold:" & vbNewLine & vbNewLine & _
               noRiskList, vbInformation, "Process Complete"
    Else
        MsgBox "Detailed Company Code Tabs generated for all entities!", vbInformation, "Process Complete"
    End If
    Exit Sub

CleanFail:
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsSrc Is Nothing Then wsSrc.Delete
    Application.DisplayAlerts = True
    MsgBox "An error occurred while generating individual tabs: " & Err.Description, vbCritical, "Process Error"
End Sub

'============================================================
' SHARED BUILDER - DRAWS FULL AGING CONTENT ONTO ANY GIVEN SHEET
' (Used for direct-to-file generation; no tab left behind)
'============================================================
Sub BuildFullAgingContent(wsNew As Worksheet, wsSrc As Worksheet, cc As String, lastRow As Long)
    Dim i As Long, r As Long
    Dim mixCols As Variant, mc As Variant, foundMix As Boolean
    mixCols = Array("G", "H", "I", "J", "K", "L", "M", "N", "O")

    With wsNew
        .Range("A1:L1").Merge
        .Cells(1, "A").Value = "AGING REPORT - " & cc & " - Generated: " & Format(Now, "dd/mm/yyyy hh:nn")
        .Cells(1, "A").Font.Bold = True
        .Cells(1, "A").Font.Size = 12
        .Cells(1, "A").Interior.Color = RGB(44, 62, 80)
        .Cells(1, "A").Font.Color = RGB(236, 240, 241)
        .Cells(1, "A").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 22
        .Rows(1).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Rows(1).Borders(xlEdgeTop).Weight = xlThick
    End With

    With wsNew
        .Range("A3:L3").Value = Array("Company Code", "Customer", "Customer Name", "Current", "Overdue 1-29", _
            "Overdue 1-7", "Overdue 8-29", "Overdue 60-89", "Overdue 90-179", "Overdue 180-359", "Overdue >359", "Total Overdue")
        .Range("A3:L3").Font.Bold = True
        .Range("A3:L3").Interior.Color = RGB(68, 84, 106)
        .Range("A3:L3").Font.Color = RGB(255, 255, 255)
    End With

    Dim sumCurrent As Double, sum129 As Double, sum17 As Double, sum829 As Double
    Dim sum6089 As Double, sum90179 As Double, sum180359 As Double, sumOver359 As Double, sumTotal As Double
    Dim ccCustCount As Long
    sumCurrent = 0: sum129 = 0: sum17 = 0: sum829 = 0
    sum6089 = 0: sum90179 = 0: sum180359 = 0: sumOver359 = 0: sumTotal = 0
    ccCustCount = 0

    For i = 4 To lastRow
        If UCase(Trim(wsSrc.Cells(i, "A").Value)) = UCase(Trim(cc)) And wsSrc.Cells(i, "C").Value <> "Result" And Trim(wsSrc.Cells(i, "C").Value) <> "" Then
            sumCurrent = sumCurrent + SafeDouble(wsSrc.Cells(i, "G").Value)
            sum129 = sum129 + SafeDouble(wsSrc.Cells(i, "H").Value)
            sum17 = sum17 + SafeDouble(wsSrc.Cells(i, "I").Value)
            sum829 = sum829 + SafeDouble(wsSrc.Cells(i, "J").Value)
            sum6089 = sum6089 + SafeDouble(wsSrc.Cells(i, "K").Value)
            sum90179 = sum90179 + SafeDouble(wsSrc.Cells(i, "L").Value)
            sum180359 = sum180359 + SafeDouble(wsSrc.Cells(i, "M").Value)
            sumOver359 = sumOver359 + SafeDouble(wsSrc.Cells(i, "N").Value)
            sumTotal = sumTotal + SafeDouble(wsSrc.Cells(i, "O").Value)
            ccCustCount = ccCustCount + 1
        End If
    Next i

    With wsNew
        .Cells(4, "A").Value = cc
        .Cells(4, "C").Value = ccCustCount & " customer(s)"
        .Cells(4, "D").Value = sumCurrent
        .Cells(4, "E").Value = sum129
        .Cells(4, "F").Value = sum17
        .Cells(4, "G").Value = sum829
        .Cells(4, "H").Value = sum6089
        .Cells(4, "I").Value = sum90179
        .Cells(4, "J").Value = sum180359
        .Cells(4, "K").Value = sumOver359
        .Cells(4, "L").Value = sumTotal
        With .Range("A4:L4")
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = RGB(213, 219, 226)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeTop).Color = RGB(52, 73, 94)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeBottom).Color = RGB(52, 73, 94)
        End With
        .Rows(4).RowHeight = 18
    End With

    r = 5
    For i = 4 To lastRow
        If UCase(Trim(wsSrc.Cells(i, "A").Value)) = UCase(Trim(cc)) And wsSrc.Cells(i, "C").Value <> "Result" And Trim(wsSrc.Cells(i, "C").Value) <> "" Then
            wsNew.Cells(r, 1).Value = cc
            wsNew.Cells(r, 2).Value = wsSrc.Cells(i, "C").Value
            wsNew.Cells(r, 3).Value = "     " & wsSrc.Cells(i, "D").Value

            foundMix = False
            For Each mc In mixCols
                If IsMixedCurrency(wsSrc.Cells(i, CStr(mc)).Value) Then foundMix = True
            Next mc
            If foundMix Then
                wsNew.Cells(r, 3).Value = wsNew.Cells(r, 3).Value & "  MIXED CURRENCY"
                wsNew.Cells(r, 3).Font.Italic = True
                wsNew.Cells(r, 3).Font.Color = RGB(230, 81, 0)
            End If

            wsNew.Cells(r, 4).Value = SafeDouble(wsSrc.Cells(i, "G").Value)
            wsNew.Cells(r, 5).Value = SafeDouble(wsSrc.Cells(i, "H").Value)
            wsNew.Cells(r, 6).Value = SafeDouble(wsSrc.Cells(i, "I").Value)
            wsNew.Cells(r, 7).Value = SafeDouble(wsSrc.Cells(i, "J").Value)
            wsNew.Cells(r, 8).Value = SafeDouble(wsSrc.Cells(i, "K").Value)
            wsNew.Cells(r, 9).Value = SafeDouble(wsSrc.Cells(i, "L").Value)
            wsNew.Cells(r, 10).Value = SafeDouble(wsSrc.Cells(i, "M").Value)
            wsNew.Cells(r, 11).Value = SafeDouble(wsSrc.Cells(i, "N").Value)
            wsNew.Cells(r, 12).Value = SafeDouble(wsSrc.Cells(i, "O").Value)
            r = r + 1
        End If
    Next i

    With wsNew
        .Columns("A:L").AutoFit
        If r - 1 >= 5 Then
            .Range("D5:L" & (r - 1)).NumberFormat = "#,##0.00"
            With .Range("A3:L" & (r - 1)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(200, 200, 200)
            End With
            .Range("A3:L" & (r - 1)).BorderAround Weight:=xlMedium, Color:=RGB(52, 73, 94)

            Dim rowIdx As Long
            For rowIdx = 5 To r - 1
                .Range(.Cells(rowIdx, "A"), .Cells(rowIdx, "L")).Font.Bold = False
            Next rowIdx

            .Outline.SummaryRow = xlAbove
            .Rows("5:" & (r - 1)).Group
        End If
    End With
End Sub

'============================================================
' FULL AGING REPORT - DIRECT TO FILES, NO TABS LEFT IN WORKBOOK
'============================================================
Sub GenerateFullAgingReportFiles()
    Dim wsSrcOriginal As Worksheet, wsSrc As Worksheet
    Dim lastRow As Long, i As Long
    Dim uniqueCC As New Collection
    Dim cc As Variant
    Dim folderPath As String
    Dim savedCount As Long

    On Error GoTo CleanFail

    On Error Resume Next
    Set wsSrcOriginal = ThisWorkbook.Sheets("BW_Summary")
    On Error GoTo 0
    If wsSrcOriginal Is Nothing Then
        MsgBox "Sheet 'BW_Summary' not found.", vbCritical
        Exit Sub
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select a folder to save the Full Aging Reports"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Set wsSrc = CreateTempSourceCopy(wsSrcOriginal)

    Dim lastRowA As Long, lastRowC As Long
    lastRowA = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowC = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowC)

    FillDownCompanyCode wsSrc, lastRow

    On Error Resume Next
    For i = 3 To lastRow
        Dim ccName As String
        ccName = UCase(Trim(wsSrc.Cells(i, "A").Value))
        If ccName <> "" And ccName <> "RESULT" And ccName <> "COMPANY CODE" Then
            uniqueCC.Add ccName, CStr(ccName)
        End If
    Next i
    On Error GoTo 0

    Dim tabDate As String
    tabDate = Format(Date, "dd.mm.yyyy")
    savedCount = 0

    Application.ScreenUpdating = False

    Dim frmProg As frmProgress
    Set frmProg = New frmProgress
    frmProg.Show vbModeless

    Dim totalCC As Long, ccCounter As Long
    totalCC = uniqueCC.Count
    ccCounter = 0

    For Each cc In uniqueCC
        ccCounter = ccCounter + 1
        frmProg.UpdateProgress ccCounter / totalCC, "Generating " & cc & " (" & ccCounter & " of " & totalCC & ")"
        DoEvents

        Dim newWB As Workbook
        Dim wsNew As Worksheet

        Set newWB = Workbooks.Add(xlWBATWorksheet)
        Set wsNew = newWB.Sheets(1)
        wsNew.Name = Left("F AGING_" & cc, 31)

        BuildFullAgingContent wsNew, wsSrc, CStr(cc), lastRow

        Dim fileName As String
        fileName = "F AGING_" & cc & "_" & tabDate & ".xlsx"
        fileName = Replace(fileName, "/", ".")
        fileName = Replace(fileName, "\", ".")
        fileName = Replace(fileName, ":", ".")

        On Error Resume Next
        Kill folderPath & fileName
        On Error GoTo 0

        frmProg.Hide
        DoEvents

        Dim fullSavePath As String
        fullSavePath = folderPath & fileName

        On Error Resume Next
        newWB.SaveAs Filename:=fullSavePath, FileFormat:=51
        On Error GoTo 0

        newWB.Close SaveChanges:=False

        frmProg.Show vbModeless
        DoEvents

        savedCount = savedCount + 1
    Next cc

    Unload frmProg
    Application.ScreenUpdating = True

    Application.DisplayAlerts = False
    wsSrc.Delete
    Application.DisplayAlerts = True

    MsgBox savedCount & " Full Aging report(s) saved successfully to:" & vbNewLine & folderPath, vbInformation, "Full Aging Reports Complete"
    Exit Sub

CleanFail:
    On Error Resume Next
    Unload frmProg
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsSrc Is Nothing Then wsSrc.Delete
    Application.DisplayAlerts = True
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Process Error"
End Sub

'============================================================
' SINGLE FULL AGING REPORT - DIRECT SAVE AS, NO TAB LEFT IN WORKBOOK
'============================================================
Sub GenerateSingleReportToFile()
    Dim wsSrcOriginal As Worksheet, wsSrc As Worksheet
    Dim lastRow As Long, i As Long
    Dim uniqueCC As New Collection
    Dim cc As Variant

    On Error GoTo CleanFail

    On Error Resume Next
    Set wsSrcOriginal = ThisWorkbook.Sheets("BW_Summary")
    On Error GoTo 0
    If wsSrcOriginal Is Nothing Then
        MsgBox "Sheet 'BW_Summary' not found.", vbCritical
        Exit Sub
    End If

    Set wsSrc = CreateTempSourceCopy(wsSrcOriginal)

    Dim lastRowA As Long, lastRowC As Long
    lastRowA = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowC = wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowC)

    FillDownCompanyCode wsSrc, lastRow

    On Error Resume Next
    For i = 3 To lastRow
        Dim ccName As String
        ccName = UCase(Trim(wsSrc.Cells(i, "A").Value))
        If ccName <> "" And ccName <> "RESULT" And ccName <> "COMPANY CODE" Then
            uniqueCC.Add ccName, CStr(ccName)
        End If
    Next i
    On Error GoTo 0

    If uniqueCC.Count = 0 Then
        Application.DisplayAlerts = False
        wsSrc.Delete
        Application.DisplayAlerts = True
        MsgBox "No Company Codes found in BW_Summary.", vbExclamation
        Exit Sub
    End If

    Dim frm As frmReportMenu
    Set frm = New frmReportMenu
    frm.lblTitle.Caption = "Select a Company Code:"

    frm.lstOptions.Clear
    For Each cc In uniqueCC
        frm.lstOptions.AddItem cc
    Next cc

    frm.Show

    If frm.UserCancelled Then
        Unload frm
        Application.DisplayAlerts = False
        wsSrc.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If

    Dim selectedCC As String
    selectedCC = frm.lstOptions.List(frm.SelectedIndex)
    Unload frm

    Dim newWB As Workbook
    Dim wsNew As Worksheet
    Set newWB = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = newWB.Sheets(1)
    wsNew.Name = Left("F AGING_" & selectedCC, 31)

    BuildFullAgingContent wsNew, wsSrc, selectedCC, lastRow

    Application.DisplayAlerts = False
    wsSrc.Delete
    Application.DisplayAlerts = True

    Dim defaultPath As String, defaultName As String
    defaultPath = Environ("USERPROFILE") & "\Desktop\"
    defaultName = defaultPath & "F AGING_" & selectedCC & "_" & Format(Date, "dd.mm.yyyy") & ".xlsx"

    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultName, FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

    If savePath = False Then
        newWB.Close SaveChanges:=False
        Exit Sub
    End If

    newWB.SaveAs savePath
    newWB.Close SaveChanges:=False

    MsgBox "Report for " & selectedCC & " saved successfully.", vbInformation, "Report Complete"
    Exit Sub

CleanFail:
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsSrc Is Nothing Then wsSrc.Delete
    Application.DisplayAlerts = True
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Process Error"
End Sub

'============================================================
' KPI DASHBOARD TILE - PLACED AT A GIVEN ROW
'============================================================
Sub BuildKPIDashboardAt(ws As Worksheet, topRow As Long, totalCustomers As Long, totalExposure As Double, newThisWeek As Long)
    Dim r1 As Long, r2 As Long
    r1 = topRow: r2 = topRow + 1

    With ws
        .Range(.Cells(r1, "A"), .Cells(r2, "B")).Merge
        .Cells(r1, "A").Value = "Customers at Risk" & vbNewLine & totalCustomers
        .Cells(r1, "A").HorizontalAlignment = xlCenter
        .Cells(r1, "A").VerticalAlignment = xlCenter
        .Cells(r1, "A").Font.Bold = True
        .Cells(r1, "A").Font.Size = 14
        If totalCustomers > 10 Then
            .Range(.Cells(r1, "A"), .Cells(r2, "B")).Interior.Color = RGB(255, 205, 210)
            .Cells(r1, "A").Font.Color = RGB(183, 28, 28)
        ElseIf totalCustomers > 0 Then
            .Range(.Cells(r1, "A"), .Cells(r2, "B")).Interior.Color = RGB(255, 236, 179)
            .Cells(r1, "A").Font.Color = RGB(230, 81, 0)
        Else
            .Range(.Cells(r1, "A"), .Cells(r2, "B")).Interior.Color = RGB(200, 230, 201)
            .Cells(r1, "A").Font.Color = RGB(27, 94, 32)
        End If

        .Range(.Cells(r1, "C"), .Cells(r2, "D")).Merge
        .Cells(r1, "C").Value = "Total Exposure EUR" & vbNewLine & Format(totalExposure, "#,##0.00")
        .Cells(r1, "C").HorizontalAlignment = xlCenter
        .Cells(r1, "C").VerticalAlignment = xlCenter
        .Cells(r1, "C").Font.Bold = True
        .Cells(r1, "C").Font.Size = 14
        If totalExposure > 100000 Then
            .Range(.Cells(r1, "C"), .Cells(r2, "D")).Interior.Color = RGB(255, 205, 210)
            .Cells(r1, "C").Font.Color = RGB(183, 28, 28)
        ElseIf totalExposure > 25000 Then
            .Range(.Cells(r1, "C"), .Cells(r2, "D")).Interior.Color = RGB(255, 236, 179)
            .Cells(r1, "C").Font.Color = RGB(230, 81, 0)
        Else
            .Range(.Cells(r1, "C"), .Cells(r2, "D")).Interior.Color = RGB(200, 230, 201)
            .Cells(r1, "C").Font.Color = RGB(27, 94, 32)
        End If

        .Range(.Cells(r1, "E"), .Cells(r2, "F")).Merge
        .Cells(r1, "E").Value = "New Risks This Week" & vbNewLine & newThisWeek
        .Cells(r1, "E").HorizontalAlignment = xlCenter
        .Cells(r1, "E").VerticalAlignment = xlCenter
        .Cells(r1, "E").Font.Bold = True
        .Cells(r1, "E").Font.Size = 14
        If newThisWeek > 0 Then
            .Range(.Cells(r1, "E"), .Cells(r2, "F")).Interior.Color = RGB(255, 205, 210)
            .Cells(r1, "E").Font.Color = RGB(183, 28, 28)
        Else
            .Range(.Cells(r1, "E"), .Cells(r2, "F")).Interior.Color = RGB(200, 230, 201)
            .Cells(r1, "E").Font.Color = RGB(27, 94, 32)
        End If

        .Range(.Cells(r1, "A"), .Cells(r2, "F")).Borders.Weight = xlThin
        .Rows(r1 & ":" & r2).RowHeight = 20
    End With
End Sub

'============================================================
' SAVE THIS WEEK'S BLOCK TO THE HIDDEN ARCHIVE
'============================================================
Sub SaveRiskArchive(wsRep As Worksheet, startRow As Long, endRow As Long)
    Dim wsArch As Worksheet
    Dim i As Long, r As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("RiskArchive").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsArch = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsArch.Name = "RiskArchive"
    wsArch.Visible = xlSheetHidden

    wsArch.Range("A1:B1").Value = Array("Company Code", "Customer")
    r = 2
    For i = startRow To endRow
        If wsRep.Cells(i, "B").Value <> "" Then
            wsArch.Cells(r, "A").Value = wsRep.Cells(i, "A").Value
            wsArch.Cells(r, "B").Value = wsRep.Cells(i, "B").Value
            r = r + 1
        End If
    Next i
End Sub

'============================================================
' ADD A HIGH-END SHAPE BUTTON NEXT TO SUB-TOTAL ROWS (COLUMN K)
'============================================================
Sub AddSleekEmailButton(ws As Worksheet, rowNum As Long, cc As String)
    Dim shp As Shape
    Dim btnRange As Range
    Dim btnName As String
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double

    btnName = "btnCC_" & cc & "_" & rowNum

    On Error Resume Next
    ws.Shapes(btnName).Delete
    For Each shp In ws.Shapes
        If shp.Type = msoShapeRoundedRectangle Then
            If shp.TopLeftCell.Row = rowNum And shp.TopLeftCell.Column = 11 Then
                shp.Delete
            End If
        End If
    Next shp
    On Error GoTo 0

    Set btnRange = ws.Range("K" & rowNum)

    btnWidth = 85
    btnHeight = 14
    btnLeft = btnRange.Left + (btnRange.Width - btnWidth) / 2
    btnTop = btnRange.Top + (btnRange.RowHeight - btnHeight) / 2

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With shp
        .Name = btnName
        .OnAction = "SendSingleEmailForCompanyCode"
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(21, 55, 87)
        .Line.Weight = 1

        With .TextFrame2
            .TextRange.Characters.Text = "Send Email " & cc
            .TextRange.Characters.Font.Name = "Calibri"
            .TextRange.Characters.Font.Size = 9
            .TextRange.Characters.Font.Bold = True
            .TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
    End With
End Sub

'============================================================
' TRIGGERED BY SELECTIVE CC EMAIL BUTTONS
'============================================================
Sub SendSingleEmailForCompanyCode()
    Dim wsRep As Worksheet, wsResp As Worksheet
    Dim btnRow As Long, i As Long
    Dim targetCC As String
    Dim OutApp As Object, OutMail As Object

    Set wsRep = ActiveSheet

    On Error Resume Next
    btnRow = wsRep.Shapes(Application.Caller).TopLeftCell.Row
    On Error GoTo 0

    If btnRow = 0 Then
        MsgBox "Please click a stylish email button directly on a generated sheet.", vbExclamation
        Exit Sub
    End If

    targetCC = Trim(wsRep.Cells(btnRow, "A").Value)
    If targetCC = "" Then
        MsgBox "Could not find a valid Company Code on this row.", vbCritical
        Exit Sub
    End If

    Dim customersList As Object
    Set customersList = CreateObject("System.Collections.ArrayList")

    Dim custNo As String, custName As String
    Dim d60 As String, d90 As String, dRisk As String

    i = btnRow + 1
    Do While Trim(wsRep.Cells(i, "A").Value) = targetCC
        custNo = Trim(wsRep.Cells(i, "B").Value)
        If custNo <> "" And Trim(wsRep.Cells(i, "C").Value) <> "" Then
            custName = Trim(wsRep.Cells(i, "C").Value)
            d60 = Format(wsRep.Cells(i, "D").Value, "#,##0.00")
            d90 = Format(wsRep.Cells(i, "E").Value, "#,##0.00")
            dRisk = Format(wsRep.Cells(i, "F").Value, "#,##0.00")

            customersList.Add Array(custNo, custName, d60, d90, dRisk)
        End If
        i = i + 1
        If i > wsRep.Rows.Count Then Exit Do
    Loop

    If customersList.Count = 0 Then
        MsgBox "No at-risk customers found for " & targetCC & " to email.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsResp = ThisWorkbook.Sheets("Responsible")
    On Error GoTo 0
    If wsResp Is Nothing Then
        MsgBox "Sheet 'Responsible' not found. Cannot search for emails.", vbCritical
        Exit Sub
    End If

    Dim respLast As Long, rIdx As Long
    Dim mailTo As String, mailCc As String
    mailTo = ""
    mailCc = ""

    respLast = wsResp.Cells(wsResp.Rows.Count, "A").End(xlUp).Row
    For rIdx = 2 To respLast
        If UCase(Trim(wsResp.Cells(rIdx, "A").Value)) = UCase(targetCC) Then
            mailTo = Trim(wsResp.Cells(rIdx, "B").Value)
            mailCc = Trim(wsResp.Cells(rIdx, "C").Value)
            Exit For
        End If
    Next rIdx

    If mailTo = "" And mailCc = "" Then
        MsgBox "No email configuration found for '" & targetCC & "' in the Responsible sheet.", vbExclamation
        Exit Sub
    End If

    Dim htmlTable As String, htmlRow As String, item As Variant
    htmlTable = "<table style='border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt; min-width: 800px; border: 1px solid #D9D9D9;'>" & _
                "<tr style='background-color: #44546A; color: white; font-weight: bold; text-align: left;'>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px;'>Company Code</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px;'>Customer</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px;'>Customer Name</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px; text-align: right;'>Overdue 60-89</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px; text-align: right;'>Overdue 90-179</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px; text-align: right; background-color: #E0ECFF; color: #0D47A1;'>Overdue in Risk</th>" & _
                "<th style='border: 1px solid #D9D9D9; padding: 8px; text-align: center; width: 250px; background-color: #F2F4F7; color: #333333;'>Status / Comments (Please fill up)</th>" & _
                "</tr>"

    For Each item In customersList
        htmlRow = "<tr>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px; font-weight: bold;'>" & targetCC & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px;'>" & item(0) & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px;'>" & item(1) & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px; text-align: right;'>" & item(2) & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px; text-align: right;'>" & item(3) & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px; text-align: right; background-color: #F5F9FF; font-weight: bold; color: #1F4E79;'>" & item(4) & "</td>" & _
                  "<td style='border: 1px solid #D9D9D9; padding: 6px; background-color: #FAFAFA; text-align: center; color: #888888; font-style: italic;'>&nbsp;</td>" & _
                  "</tr>"
        htmlTable = htmlTable & htmlRow
    Next item
    htmlTable = htmlTable & "</table>"

    Dim tempWB As Workbook, tempWS As Worksheet
    Dim tempFilePath As String, excelRow As Long

    Set tempWB = Workbooks.Add(xlWBATWorksheet)
    Set tempWS = tempWB.Sheets(1)
    tempWS.Name = targetCC & "_Risk_Report"

    tempWS.Range("A1:G1").Value = Array("Company Code", "Customer", "Customer Name", "Overdue 60-89", "Overdue 90-179", "Overdue in Risk", "Status / Comments (Please fill up)")
    tempWS.Range("A1:G1").Font.Bold = True
    tempWS.Range("A1:G1").Interior.Color = RGB(68, 84, 106)
    tempWS.Range("A1:G1").Font.Color = RGB(255, 255, 255)

    excelRow = 2
    For Each item In customersList
        tempWS.Cells(excelRow, 1).Value = targetCC
        tempWS.Cells(excelRow, 2).Value = item(0)
        tempWS.Cells(excelRow, 3).Value = item(1)
        tempWS.Cells(excelRow, 4).Value = CDbl(item(2))
        tempWS.Cells(excelRow, 5).Value = CDbl(item(3))
        tempWS.Cells(excelRow, 6).Value = CDbl(item(4))
        tempWS.Cells(excelRow, 7).Value = ""
        excelRow = excelRow + 1
    Next item

    tempWS.Range("D2:F" & (excelRow - 1)).NumberFormat = "#,##0.00"
    tempWS.Range("F2:F" & (excelRow - 1)).Interior.Color = RGB(224, 236, 255)
    tempWS.Range("G2:G" & (excelRow - 1)).Interior.Color = RGB(250, 250, 250)

    With tempWS.Range("A1:G" & (excelRow - 1)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With

    tempWS.Columns("A:G").AutoFit
    tempWS.Columns("G").ColumnWidth = 35

    tempFilePath = Environ("TEMP") & "\" & targetCC & "_Risk_Report_" & Format(Now, "yyyymmdd") & ".xlsx"
    On Error Resume Next
    Kill tempFilePath
    On Error GoTo 0
    tempWB.SaveAs tempFilePath
    tempWB.Close SaveChanges:=False

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = mailTo
        If mailCc <> "" Then .CC = mailCc
        .Subject = targetCC & " - Customers in risk (Overdue 60-179)"
        .Attachments.Add tempFilePath
        .HTMLBody = "<font face='Calibri' size='3' color='#1F4E79'><p>Hi Team,</p>" & _
                    "<p>Could you please check the following report and let me know what is the status of the following customers?</p>" & _
                    "<p>Shall we worry about these accounts slipping further into the <b>90+ days bucket</b> and assuming a larger credit risk under our factoring rules?</p><br>" & _
                    htmlTable & "<br>" & _
                    "<p>Please fill out your feedback directly in the attached Excel sheet or in the light-grey column above and hit reply.</p>" & _
                    "<p>Thank you for your prompt cooperation!</p>" & _
                    "<p>Best regards,<br>" & _
                    "<b>Desislav Viner Valev</b><br>" & _
                    "Professional Accounts Receivable<br>" & _
                    "Villeroy & Boch SSC office</p></font>"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

'============================================================
' EXTRACT A CLEAN DATE STRING FROM A "WEEK ..." BANNER TEXT
'============================================================
Function ExtractDateFromBanner(bannerText As String) As String
    Dim pos As Long, rawDate As String
    pos = InStr(bannerText, "generated:")
    If pos > 0 Then
        rawDate = Trim(Mid(bannerText, pos + Len("generated:")))
        rawDate = Trim(Left(rawDate, 10))
        rawDate = Replace(rawDate, "/", ".")
    Else
        rawDate = Format(Date, "dd.mm.yyyy")
    End If
    ExtractDateFromBanner = rawDate
End Function

'============================================================
' EXPORT A SELECTED FACTORING BLOCK OR CC TAB
'============================================================
Sub ExportGeneratedReport()
    Dim wsRep As Worksheet
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim itemList As Collection, itemType As Collection, itemRef As Collection
    Set itemList = New Collection
    Set itemType = New Collection
    Set itemRef = New Collection

    On Error Resume Next
    Set wsRep = ThisWorkbook.Sheets("Factoring Report")
    On Error GoTo 0

    Dim counter As Long
    counter = 0

    Dim bannerRows As Collection
    Set bannerRows = New Collection

    If Not wsRep Is Nothing Then
        lastRow = wsRep.Cells(wsRep.Rows.Count, "A").End(xlUp).Row
        For i = 1 To lastRow
            If InStr(1, wsRep.Cells(i, "A").Value, "WEEK ") > 0 Then
                bannerRows.Add i
            End If
        Next i

        Dim bIdx As Long, bStart As Long, bEnd As Long
        For bIdx = 1 To bannerRows.Count
            counter = counter + 1
            bStart = bannerRows(bIdx)
            If bIdx < bannerRows.Count Then
                bEnd = bannerRows(bIdx + 1) - 3
            Else
                bEnd = lastRow
            End If
            itemList.Add wsRep.Cells(bStart, "A").Value
            itemType.Add "FACTORING"
            itemRef.Add bStart & "|" & bEnd
        Next bIdx
    End If

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) And ws.Name <> "Factoring Report" And ws.Name <> "RiskArchive" _
           And ws.Name <> "TempBW_Working" And Left(ws.Name, 1) <> "_" And ws.Visible = xlSheetVisible Then
            counter = counter + 1
            itemList.Add ws.Name
            itemType.Add "CCTAB"
            itemRef.Add ws.Name
        End If
    Next ws

    If counter = 0 Then
        MsgBox "No generated reports found to export. Please run Option 1 or Option 2 first.", vbExclamation
        Exit Sub
    End If

    Dim frm As frmReportMenu
    Set frm = New frmReportMenu
    frm.lblTitle.Caption = "Select a report to export:"

    frm.lstOptions.Clear
    Dim idx As Long
    For idx = 1 To counter
        frm.lstOptions.AddItem itemList(idx)
    Next idx

    frm.Show

    If frm.UserCancelled Then
        Unload frm
        Exit Sub
    End If

    Dim selIdx As Long
    selIdx = frm.SelectedIndex + 1
    Unload frm

    Dim selType As String, selRef As String
    selType = itemType(selIdx)
    selRef = itemRef(selIdx)

    Dim newWB As Workbook
    Dim exportFileName As String
    Dim savePath As Variant

    Application.ScreenUpdating = False

    If selType = "FACTORING" Then
        Dim parts() As String
        parts = Split(selRef, "|")
        Dim rStart As Long, rEnd As Long
        rStart = CLng(parts(0))
        rEnd = CLng(parts(1))

        wsRep.Range("A" & rStart & ":K" & rEnd).Copy
        Set newWB = Workbooks.Add(xlWBATWorksheet)
        newWB.Sheets(1).Range("A1").PasteSpecial xlPasteColumnWidths
        newWB.Sheets(1).Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        newWB.Sheets(1).Name = "Factoring Report"

        Dim datePart As String
        datePart = ExtractDateFromBanner(wsRep.Cells(rStart, "A").Value)
        exportFileName = "R AGING_FACTORING_" & datePart
    Else
        Dim wsSel As Worksheet
        Set wsSel = ThisWorkbook.Sheets(selRef)
        wsSel.Cells.Copy
        Set newWB = Workbooks.Add(xlWBATWorksheet)
        newWB.Sheets(1).Range("A1").PasteSpecial xlPasteColumnWidths
        newWB.Sheets(1).Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        newWB.Sheets(1).Name = Left(selRef, 31)

        exportFileName = selRef
    End If

    Application.ScreenUpdating = True

    exportFileName = Replace(exportFileName, "/", ".")
    exportFileName = Replace(exportFileName, "\", ".")
    exportFileName = Replace(exportFileName, ":", ".")

    savePath = Application.GetSaveAsFilename(InitialFileName:=exportFileName, FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If savePath = False Then
        newWB.Close SaveChanges:=False
        Exit Sub
    End If

    newWB.SaveAs savePath
    MsgBox "Report exported successfully to:" & vbNewLine & savePath, vbInformation
End Sub

'============================================================
' EXPORT ALL GENERATED REPORTS TO ONE FOLDER
'============================================================
Function CountAvailableReports() As Long
    Dim wsRep As Worksheet, ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim total As Long
    total = 0

    On Error Resume Next
    Set wsRep = ThisWorkbook.Sheets("Factoring Report")
    On Error GoTo 0

    If Not wsRep Is Nothing Then
        lastRow = wsRep.Cells(wsRep.Rows.Count, "A").End(xlUp).Row
        For i = 1 To lastRow
            If InStr(1, wsRep.Cells(i, "A").Value, "WEEK ") > 0 Then
                total = total + 1
            End If
        Next i
    End If

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) And ws.Name <> "Factoring Report" And ws.Name <> "RiskArchive" _
           And ws.Name <> "TempBW_Working" And ws.Name <> "Home" And Left(ws.Name, 1) <> "_" And ws.Visible = xlSheetVisible Then
            total = total + 1
        End If
    Next ws

    CountAvailableReports = total
End Function

Sub ExportAllGeneratedReports()
    Dim wsRep As Worksheet, ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim folderPath As String
    Dim exportedCount As Long, failedCount As Long
    Dim summaryMsg As String

    If CountAvailableReports() = 0 Then
        MsgBox "No available reports to export." & vbNewLine & vbNewLine & _
               "Please generate a Factoring Report, Company Code Risk Tabs, or a Full Aging Report first.", _
               vbExclamation, "Nothing to Export"
        Exit Sub
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select a folder to export ALL generated reports"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    exportedCount = 0
    failedCount = 0
    summaryMsg = ""

    On Error Resume Next
    Set wsRep = ThisWorkbook.Sheets("Factoring Report")
    On Error GoTo 0

    If Not wsRep Is Nothing Then
        lastRow = wsRep.Cells(wsRep.Rows.Count, "A").End(xlUp).Row
        Dim bannerRows As Collection
        Set bannerRows = New Collection
        For i = 1 To lastRow
            If InStr(1, wsRep.Cells(i, "A").Value, "WEEK ") > 0 Then
                bannerRows.Add i
            End If
        Next i

        Dim bIdx As Long, bStart As Long, bEnd As Long
        For bIdx = 1 To bannerRows.Count
            bStart = bannerRows(bIdx)
            If bIdx < bannerRows.Count Then
                bEnd = bannerRows(bIdx + 1) - 3
            Else
                bEnd = lastRow
            End If

            Dim datePart As String
            datePart = ExtractDateFromBanner(wsRep.Cells(bStart, "A").Value)

            Dim fName As String
            fName = "R AGING_FACTORING_" & datePart & ".xlsx"
            fName = Replace(fName, "/", ".")
            fName = Replace(fName, "\", ".")
            fName = Replace(fName, ":", ".")

            Dim newWB As Workbook
            On Error Resume Next
            Err.Clear
            wsRep.Range("A" & bStart & ":K" & bEnd).Copy
            Set newWB = Workbooks.Add(xlWBATWorksheet)
            newWB.Sheets(1).Range("A1").PasteSpecial xlPasteColumnWidths
            newWB.Sheets(1).Range("A1").PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            newWB.Sheets(1).Name = "Factoring Report"
            newWB.SaveAs folderPath & fName, FileFormat:=51
            newWB.Close SaveChanges:=False

            If Err.Number <> 0 Then
                failedCount = failedCount + 1
                summaryMsg = summaryMsg & "FAILED: " & fName & " -- Error " & Err.Number & ": " & Err.Description & vbNewLine
                Err.Clear
            Else
                exportedCount = exportedCount + 1
            End If
            On Error GoTo 0
        Next bIdx
    End If

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) And ws.Name <> "Factoring Report" And ws.Name <> "RiskArchive" _
           And ws.Name <> "TempBW_Working" And Left(ws.Name, 1) <> "_" And ws.Visible = xlSheetVisible Then

            Dim fName2 As String
            fName2 = ws.Name & ".xlsx"
            fName2 = Replace(fName2, "/", ".")
            fName2 = Replace(fName2, "\", ".")
            fName2 = Replace(fName2, ":", ".")

            Dim newWB2 As Workbook
            On Error Resume Next
            Err.Clear
            ws.Cells.Copy
            Set newWB2 = Workbooks.Add(xlWBATWorksheet)
            newWB2.Sheets(1).Range("A1").PasteSpecial xlPasteColumnWidths
            newWB2.Sheets(1).Range("A1").PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            newWB2.Sheets(1).Name = Left(ws.Name, 31)
            newWB2.SaveAs folderPath & fName2, FileFormat:=51
            newWB2.Close SaveChanges:=False

            If Err.Number <> 0 Then
                failedCount = failedCount + 1
                summaryMsg = summaryMsg & "FAILED: " & fName2 & " -- Error " & Err.Number & ": " & Err.Description & vbNewLine
                Err.Clear
            Else
                exportedCount = exportedCount + 1
            End If
            On Error GoTo 0
        End If
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox exportedCount & " report(s) exported successfully to:" & vbNewLine & folderPath & _
           IIf(failedCount > 0, vbNewLine & vbNewLine & failedCount & " failed:" & vbNewLine & summaryMsg, ""), _
           vbInformation, "Export All Reports Complete"

    Dim frmClean As frmCleanupChoice
    Set frmClean = New frmCleanupChoice
    frmClean.Show

    Dim cleanupResult As String
    cleanupResult = frmClean.CleanupChoice
    Unload frmClean

    Select Case cleanupResult
        Case "ALL"
            DeleteGeneratedSheets True
            MsgBox "All generated sheets have been deleted.", vbInformation
        Case "NOFACTORING"
            DeleteGeneratedSheets False
            MsgBox "Company Code tabs deleted. Factoring Report was kept.", vbInformation
        Case "NONE"
    End Select

    EnsureTabOrder
End Sub

'============================================================
' HELPER - DELETE GENERATED SHEETS (USED BY CLEANUP DIALOG)
'============================================================
Sub DeleteGeneratedSheets(includeFactoring As Boolean)
    Dim ws As Worksheet
    Dim sheetsToDelete As New Collection

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) And ws.Name <> "TempBW_Working" And Left(ws.Name, 1) <> "_" Then
            If includeFactoring Then
                sheetsToDelete.Add ws.Name
            ElseIf ws.Name <> "Factoring Report" And ws.Name <> "RiskArchive" Then
                sheetsToDelete.Add ws.Name
            End If
        End If
    Next ws

    Application.DisplayAlerts = False
    On Error Resume Next
    Dim sName As Variant
    For Each sName In sheetsToDelete
        ThisWorkbook.Sheets(CStr(sName)).Delete
    Next sName
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

'============================================================
' DELETE ALL SHEETS EXCEPT FACTORING REPORT + MASTER DATA + HOME
'============================================================
Sub DeleteAllSheetsExceptFactoring()
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("This will permanently delete:" & vbNewLine & _
                      "- All Company Code Risk Tabs (R AGING_...)" & vbNewLine & _
                      "- All Full Aging / Single Report Tabs (F AGING_...)" & vbNewLine & vbNewLine & _
                      "The 'Factoring Report' sheet and its history will be KEPT." & vbNewLine & _
                      "Master data sheets will stay untouched." & vbNewLine & vbNewLine & _
                      "Continue?", vbYesNo + vbExclamation, "Delete All Sheets (Except Factoring)")

    If confirm = vbNo Then Exit Sub

    Dim ws As Worksheet
    Dim sheetsToDelete As New Collection

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) _
           And ws.Name <> "Factoring Report" _
           And ws.Name <> "RiskArchive" _
           And ws.Name <> "Home" Then
            sheetsToDelete.Add ws.Name
        End If
    Next ws

    If sheetsToDelete.Count = 0 Then
        MsgBox "No extra sheets found to delete.", vbInformation
        Exit Sub
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    Dim sName As Variant
    For Each sName In sheetsToDelete
        ThisWorkbook.Sheets(CStr(sName)).Delete
    Next sName
    On Error GoTo 0
    Application.DisplayAlerts = True

    EnsureTabOrder

    MsgBox sheetsToDelete.Count & " sheet(s) deleted. Factoring Report and master data were kept.", _
           vbInformation, "Cleanup Complete"
End Sub

'============================================================
' RESET TRACKER - WIPE HISTORY AND INDIVIDUAL TABS
'============================================================
Sub ResetTracker()
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("This will permanently delete:" & vbNewLine & _
                      "- The 'Factoring Report' sheet (all weekly blocks & KPIs)" & vbNewLine & _
                      "- The 'RiskArchive' sheet (NEW/Existing history)" & vbNewLine & _
                      "- All custom Company Code tabs" & vbNewLine & vbNewLine & _
                      "Your master data sheets will stay untouched." & vbNewLine & vbNewLine & _
                      "Continue?", vbYesNo + vbExclamation, "Reset Tracker Control")

    If confirm = vbNo Then Exit Sub

    Dim ws As Worksheet
    Dim sheetsToDelete As New Collection

    For Each ws In ThisWorkbook.Sheets
        If Not IsProtectedSheet(ws.Name) Then
            sheetsToDelete.Add ws.Name
        End If
    Next ws

    Application.DisplayAlerts = False
    On Error Resume Next

    ThisWorkbook.Sheets("BW_Summary").Visible = xlSheetVisible

    Dim sName As Variant
    For Each sName In sheetsToDelete
        ThisWorkbook.Sheets(CStr(sName)).Delete
    Next sName
    On Error GoTo 0
    Application.DisplayAlerts = True

    MsgBox "Tracker completely reset! All historic reports, archives, and custom tabs have been wiped clean.", vbInformation, "Reset Successful"
End Sub

'============================================================
' ENSURE FIXED TAB ORDER - Home always first, BW_Summary second
'============================================================
Sub EnsureTabOrder()
    Dim wsHome As Worksheet, wsBW As Worksheet

    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets("Home")
    Set wsBW = ThisWorkbook.Sheets("BW_Summary")
    On Error GoTo 0

    If Not wsHome Is Nothing Then
        wsHome.Visible = xlSheetVisible
        wsHome.Move Before:=ThisWorkbook.Sheets(1)
    End If

    If Not wsBW Is Nothing Then
        If Not wsHome Is Nothing Then
            wsBW.Move After:=wsHome
        Else
            wsBW.Move Before:=ThisWorkbook.Sheets(1)
        End If
    End If

    On Error Resume Next
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    On Error GoTo 0
End Sub

'============================================================
' JUMP BACK TO HOME
'============================================================
Sub GoToHome()
    On Error Resume Next
    ThisWorkbook.Sheets("Home").Activate
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    On Error GoTo 0
End Sub


---

# 📄 FILE 3: `02_Module_Home.md`

```markdown
# Module_Home — Dashboard & Navigation

## Purpose
Builds the app-style Home control panel: banner with live clock, three 
sections (Reports / Maintenance / Admin Panel), a full icon system embedded 
via a hidden IconLibrary sheet, and all navigation/settings helpers.

## Technical Summary

- **Shape-Based UI**: All elements are `Excel.Shape` objects, not cells — 
  enables full app-like appearance with hidden gridlines/headings and 
  locked cell selection.
- **Icon System**: Icons live pre-pasted on a hidden `IconLibrary` sheet, 
  named `icon_bars`, `icon_building`, etc. `CopyLibraryPicture` uses 
  `Pictures.Paste` (not `Worksheet.Paste`, which proved unreliable) to 
  clone them onto Home at build time.
- **Protection**: Sheet is protected with `DrawingObjects:=True` — shapes 
  cannot be moved/deleted/resized, but `.OnAction` macros still fire on 
  click.
- **Self-Sizing Background**: Background rectangle is drawn *last*, sized 
  to the actual final `y` position after all content renders — eliminates 
  manual height guessing.

```vb
Option Explicit

'============================================================
' MAIN ENTRY - BUILD THE HOME / CONTROL PANEL SHEET
'============================================================
Sub BuildHomeSheet()
    Dim ws As Worksheet
    Dim leftMargin As Double, colWidth As Double, colGap As Double
    Dim col1Left As Double, col2Left As Double, gridRight As Double
    Dim dividerX As Double, rightLeft As Double, rightWidth As Double
    Dim cardLeft As Double, cardTop As Double, cardWidth As Double, cardHeight As Double

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Home").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = "Home"
    ws.Tab.Color = RGB(30, 58, 95)
    ws.Visible = xlSheetVisible

    Application.ScreenUpdating = False
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    leftMargin = 20
    colWidth = 220
    colGap = 16
    col1Left = leftMargin
    col2Left = col1Left + colWidth + colGap
    gridRight = col2Left + colWidth

    dividerX = gridRight + 25
    rightLeft = dividerX + 20
    rightWidth = 180

    cardLeft = leftMargin - 15
    cardTop = 15
    cardWidth = (rightLeft + rightWidth + 15) - cardLeft
    cardHeight = 460

    Dim cardShape As Shape
    Set cardShape = ws.Shapes.AddShape(msoShapeRectangle, cardLeft, cardTop, cardWidth, cardHeight)
    With cardShape
        .Fill.ForeColor.RGB = RGB(247, 248, 250)
        .Line.ForeColor.RGB = RGB(220, 223, 227)
        .Line.Weight = 1
        .Shadow.Type = msoShadow21
        .ZOrder msoSendToBack
    End With

    Dim bannerShape As Shape
    Set bannerShape = ws.Shapes.AddShape(msoShapeRectangle, leftMargin, 30, gridRight - leftMargin, 46)
    With bannerShape
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "VILLEROY & BOCH" & vbNewLine & "AR RISK & AGING CENTER"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 13
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim subtitleShape As Shape
    Set subtitleShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftMargin, 80, gridRight - leftMargin, 18)
    With subtitleShape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "Select an action below to get started"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 9
            .TextRange.Font.Italic = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(110, 110, 110)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim hDivider As Shape
    Set hDivider = ws.Shapes.AddLine(leftMargin + 20, 102, gridRight - 20, 102)
    hDivider.Line.ForeColor.RGB = RGB(220, 223, 227)
    hDivider.Line.Weight = 1

    Dim vDivider As Shape
    Set vDivider = ws.Shapes.AddLine(dividerX, 30, dividerX, cardTop + cardHeight - 20)
    vDivider.Line.ForeColor.RGB = RGB(220, 223, 227)
    vDivider.Line.Weight = 1

    Dim btnWidth As Double, btnHeight As Double, rowTop As Double, rowStep As Double
    btnWidth = colWidth
    btnHeight = 36
    rowStep = btnHeight + 2 + 20 + 14
    rowTop = 118

    AddHomeButton ws, "Factoring Risk Report", "Factoring report. Risk above 4,999 EUR.", "GenerateSummaryReport", col1Left, colWidth, btnWidth, btnHeight, rowTop, RGB(31, 78, 121)
    AddHomeButton ws, "Company Code Risk Report", "One tab per company code. Risk above 4,999 EUR.", "GenerateCompanyCodeTabs", col2Left, colWidth, btnWidth, btnHeight, rowTop, RGB(39, 92, 44)
    rowTop = rowTop + rowStep

    AddHomeButton ws, "Full Aging Report", "Generates each company code as a separate file, no tabs created.", "GenerateFullAgingReportFiles", col1Left, colWidth, btnWidth, btnHeight, rowTop, RGB(120, 80, 30)
    AddHomeButton ws, "Full Aging Single Report", "Pick one company code, save directly as a file.", "GenerateSingleReportToFile", col2Left, colWidth, btnWidth, btnHeight, rowTop, RGB(70, 130, 140)
    rowTop = rowTop + rowStep

    AddHomeButton ws, "Export All Reports", "Saves every generated report into a folder.", "ExportAllGeneratedReports", col1Left, colWidth, btnWidth, btnHeight, rowTop, RGB(90, 60, 150)
    AddHomeButton ws, "Export a Single Report", "Pick one report and save it as a standalone file.", "ExportGeneratedReport", col2Left, colWidth, btnWidth, btnHeight, rowTop, RGB(158, 118, 20)
    rowTop = rowTop + rowStep

    Dim footerShape As Shape
    Set footerShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftMargin, rowTop + 8, gridRight - leftMargin, 16)
    With footerShape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "Villeroy & Boch  |  Accounts Receivable  |  AR Factoring Risk Automation"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 8
            .TextRange.Font.Fill.ForeColor.RGB = RGB(150, 150, 150)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim settingsHeader As Shape
    Set settingsHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, rightLeft, 30, rightWidth, 18)
    With settingsHeader
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "SETTINGS"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(70, 75, 82)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim settingsSub As Shape
    Set settingsSub = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, rightLeft, 48, rightWidth, 16)
    With settingsSub
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "Update reference lists"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 8
            .TextRange.Font.Italic = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim sBtnWidth As Double, sBtnHeight As Double, sBlockGap As Double, sBtnTop As Double
    sBtnWidth = rightWidth - 20
    sBtnHeight = 25
    sBlockGap = 8
    sBtnTop = 70

    AddSmallButton ws, "Prompts for Workbook", "OpenSAPPrompts", rightLeft, rightWidth, sBtnWidth, sBtnHeight, sBtnTop, RGB(212, 137, 26)
    sBtnTop = sBtnTop + sBtnHeight + sBlockGap

    AddSmallButton ws, "Partners List", "JumpToPartnerLists", rightLeft, rightWidth, sBtnWidth, sBtnHeight, sBtnTop, RGB(13, 71, 161)
    sBtnTop = sBtnTop + sBtnHeight + sBlockGap

    AddSmallButton ws, "Hide Reference Sheets", "RehideReferenceSheets", rightLeft, rightWidth, sBtnWidth, sBtnHeight, sBtnTop, RGB(100, 100, 100)
    sBtnTop = sBtnTop + sBtnHeight + sBlockGap

    Dim adminDivider As Shape
    Set adminDivider = ws.Shapes.AddLine(rightLeft + 10, sBtnTop + 6, rightLeft + rightWidth - 10, sBtnTop + 6)
    adminDivider.Line.ForeColor.RGB = RGB(220, 223, 227)
    adminDivider.Line.Weight = 1

    Dim adminTop As Double
    adminTop = sBtnTop + 16

    Dim adminHeader As Shape
    Set adminHeader = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, rightLeft, adminTop, rightWidth, 18)
    With adminHeader
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = "ADMIN PANEL"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(70, 75, 82)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim aBtnTop As Double, aBtnHeight As Double, aBlockGap As Double
    aBtnTop = adminTop + 22
    aBtnHeight = 27
    aBlockGap = 16

    AddHomeButton ws, "Build Home Screen", "If the Home screen looks stuck or overlapped", "BuildHomeSheet", rightLeft, rightWidth, rightWidth - 20, aBtnHeight, aBtnTop, RGB(70, 75, 82)
    aBtnTop = aBtnTop + aBtnHeight + aBlockGap + 10

    AddHomeButton ws, "Delete All Sheets", "Removes temporary tabs, keeps Factoring", "DeleteAllSheetsExceptFactoring", rightLeft, rightWidth, rightWidth - 20, aBtnHeight, aBtnTop, RGB(120, 40, 40)
    aBtnTop = aBtnTop + aBtnHeight + aBlockGap + 10

    AddHomeButton ws, "Reset Tracker", "Clears all reports and history to start fresh", "ResetTracker", rightLeft, rightWidth, rightWidth - 20, aBtnHeight, aBtnTop, RGB(150, 40, 40)

    LockAllShapesFreeFloating ws

    ws.Columns("A:Z").ColumnWidth = 9
    ws.Rows.RowHeight = 15

    ws.EnableSelection = xlNoSelection
    ws.Protect Password:="VB_AR_2026", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

    Application.ScreenUpdating = True
    EnsureTabOrder

    MsgBox "Home screen rebuilt successfully!", vbInformation, "Setup Complete"
End Sub

'============================================================
' HELPER - ADD A CENTERED BUTTON + DESCRIPTIVE CAPTION BELOW IT
'============================================================
Sub AddHomeButton(ws As Worksheet, caption As String, helperText As String, macroName As String, leftMargin As Double, contentWidth As Double, btnWidth As Double, btnHeight As Double, topPos As Double, fillColor As Long)
    Dim shp As Shape
    Dim btnLeft As Double
    btnLeft = leftMargin + (contentWidth - btnWidth) / 2

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, topPos, btnWidth, btnHeight)
    With shp
        .Name = "btnHome_" & macroName
        .OnAction = macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow21
        With .TextFrame2
            .TextRange.Characters.Text = caption
            .TextRange.Characters.Font.Name = "Calibri"
            .TextRange.Characters.Font.Size = 10
            .TextRange.Characters.Font.Bold = True
            .TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim helperShape As Shape
    Set helperShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftMargin, topPos + btnHeight + 1, contentWidth, 20)
    With helperShape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = helperText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(130, 130, 130)
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .WordWrap = msoTrue
        End With
    End With
End Sub

'============================================================
' HELPER - ADD A COMPACT BUTTON FOR SETTINGS/ADMIN
'============================================================
Sub AddSmallButton(ws As Worksheet, caption As String, macroName As String, boxLeft As Double, boxWidth As Double, btnWidth As Double, btnHeight As Double, topPos As Double, fillColor As Long)
    Dim shp As Shape
    Dim btnLeft As Double
    btnLeft = boxLeft + (boxWidth - btnWidth) / 2

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, topPos, btnWidth, btnHeight)
    With shp
        .Name = "btnSettings_" & macroName
        .OnAction = macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow21
        With .TextFrame2
            .TextRange.Characters.Text = caption
            .TextRange.Characters.Font.Name = "Calibri"
            .TextRange.Characters.Font.Size = 8.5
            .TextRange.Characters.Font.Bold = True
            .TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

'============================================================
' OPEN SAP ANALYSIS PROMPTS FOR WORKBOOK
'============================================================
Sub OpenSAPPrompts()
    Dim lResult As Variant
    Dim errOccurred As Boolean
    errOccurred = False

    On Error Resume Next
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
    If Err.Number <> 0 Then errOccurred = True

    lResult = Application.Run("SAPExecuteCommand", "ShowPrompts", "ALL")
    If Err.Number <> 0 Then errOccurred = True
    On Error GoTo 0

    If errOccurred Then
        MsgBox "Could not communicate with the SAP Analysis Add-In." & vbNewLine & vbNewLine & _
               "Please check that:" & vbNewLine & _
               "1. The 'Analysis' ribbon tab is enabled in Excel." & vbNewLine & _
               "2. Your SAP network / VPN is connected.", _
               vbExclamation, "SAP Analysis"
    End If
End Sub

'============================================================
' ROBUST COPY - WITH RETRY LOGIC TO ELIMINATE CLIPBOARD RACE CONDITIONS
'============================================================
Function CopyLibraryPicture(ws As Worksheet, shapeName As String) As Shape
    Dim wsLib As Worksheet
    Dim srcShp As Shape
    Dim newPic As Object
    Dim attempt As Long
    Dim maxAttempts As Long
    maxAttempts = 5

    On Error Resume Next
    Set wsLib = ThisWorkbook.Sheets("IconLibrary")
    On Error GoTo 0
    If wsLib Is Nothing Then Exit Function

    On Error Resume Next
    Set srcShp = wsLib.Shapes(shapeName)
    On Error GoTo 0
    If srcShp Is Nothing Then Exit Function

    For attempt = 1 To maxAttempts
        Application.CutCopyMode = False
        DoEvents

        srcShp.Copy

        Dim waitStart As Double
        waitStart = Timer
        Do While Timer < waitStart + 0.05
            DoEvents
        Loop

        On Error Resume Next
        Set newPic = Nothing
        Set newPic = ws.Pictures.Paste(Link:=False)
        On Error GoTo 0

        If Not newPic Is Nothing Then
            If newPic.Width > 1 And newPic.Height > 1 Then
                Exit For
            Else
                On Error Resume Next
                newPic.Delete
                On Error GoTo 0
                Set newPic = Nothing
            End If
        End If

        DoEvents
    Next attempt

    Application.CutCopyMode = False

    If newPic Is Nothing Then Exit Function

    Set CopyLibraryPicture = ws.Shapes(newPic.Name)
End Function

'============================================================
' DRAW ICON - PULLS FROM IconLibrary VIA CopyLibraryPicture
'============================================================
Sub DrawIcon(ws As Worksheet, iconType As String, boxLeft As Double, boxTop As Double, boxSize As Double, clr As Long, macroName As String)
    Dim newShp As Shape
    Dim iconName As String
    iconName = "icon_" & LCase(iconType)

    Set newShp = CopyLibraryPicture(ws, iconName)
    If newShp Is Nothing Then Exit Sub

    Dim padding As Double, targetSize As Double
    padding = boxSize * 0.24
    targetSize = boxSize - (padding * 2)

    With newShp
        .LockAspectRatio = msoFalse
        .Width = targetSize
        .Height = targetSize
        .Left = boxLeft + (boxSize - targetSize) / 2
        .Top = boxTop + (boxSize - targetSize) / 2
        If macroName <> "" Then .OnAction = macroName
        .Locked = True
    End With
End Sub

'============================================================
' JUMP TO REFERENCE SHEETS
'============================================================
Sub JumpToPartnerLists()
    JumpToReferenceSheet "AtradiusPartners"
    JumpToReferenceSheet "FactoringPartners"
End Sub

Sub JumpToFactoringPartners()
    JumpToReferenceSheet "FactoringPartners"
End Sub

Sub JumpToAtradiusPartners()
    JumpToReferenceSheet "AtradiusPartners"
End Sub

Sub JumpToReferenceSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' was not found in this workbook.", vbExclamation
        Exit Sub
    End If

    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.Activate
End Sub

'============================================================
' RE-HIDE FACTORING/ATRADIUS PARTNERS
'============================================================
Sub RehideReferenceSheets()
    Dim sheetNames As Variant
    Dim nm As Variant

    sheetNames = Array("FactoringPartners", "AtradiusPartners")

    For Each nm In sheetNames
        On Error Resume Next
        ThisWorkbook.Sheets(CStr(nm)).Visible = xlSheetVeryHidden
        On Error GoTo 0
    Next nm

    On Error Resume Next
    ThisWorkbook.Sheets("Home").Activate
    On Error GoTo 0
End Sub

'============================================================
' LOCK ALL SHAPES AS FREE-FLOATING & UNTOUCHABLE
'============================================================
Sub LockAllShapesFreeFloating(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        On Error Resume Next
        shp.Placement = xlFreeFloating
        shp.Locked = True
        On Error GoTo 0
    Next shp
End Sub

'============================================================
' ONE-TIME SETUP - HIDE ICON LIBRARY
'============================================================
Sub HideIconLibrary()
    ThisWorkbook.Sheets("IconLibrary").Visible = xlSheetVeryHidden
End Sub

'============================================================
' ONE-TIME SETUP - AUTO-NAME ICONS LEFT TO RIGHT
' (Run once after pasting icon images onto IconLibrary sheet)
'============================================================
Sub NameIconLibraryShapesByPosition()
    Dim ws As Worksheet
    Dim shapesList() As Shape
    Dim i As Long, j As Long, n As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("IconLibrary")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet 'IconLibrary' not found.", vbCritical
        Exit Sub
    End If

    n = ws.Shapes.Count
    If n = 0 Then
        MsgBox "No shapes found on 'IconLibrary'.", vbCritical
        Exit Sub
    End If

    ReDim shapesList(1 To n)
    For i = 1 To n
        Set shapesList(i) = ws.Shapes(i)
    Next i

    Dim tempShp As Shape
    For i = 1 To n - 1
        For j = 1 To n - i
            If shapesList(j).Left > shapesList(j + 1).Left Then
                Set tempShp = shapesList(j)
                Set shapesList(j) = shapesList(j + 1)
                Set shapesList(j + 1) = tempShp
            End If
        Next j
    Next i

    Dim iconNames As Variant
    iconNames = Array("icon_building", "icon_bars", "icon_database", "icon_documentcheck", "icon_document", _
        "icon_monitor", "icon_refresh", "icon_gear", "icon_grid", "icon_trash", _
        "icon_export", "icon_people", "icon_wrench")

    If n <> (UBound(iconNames) + 1) Then
        MsgBox "Found " & n & " shapes but expected " & (UBound(iconNames) + 1) & ".", vbExclamation
        Exit Sub
    End If

    For i = 1 To n
        On Error Resume Next
        shapesList(i).Name = CStr(iconNames(i - 1))
        On Error GoTo 0
    Next i

    MsgBox "Icons named successfully.", vbInformation
End Sub
