```vb
Option Explicit

Sub BuildGuideSheet()
    Dim ws As Worksheet
    Const DESIGN_WIDTH As Double = 1180
    Const MARGIN As Double = 24
    Const COL_GAP As Double = 40

    Dim navy As Long, lightBg As Long, titleClr As Long, descClr As Long, borderClr As Long, accentClr As Long
    navy = RGB(34, 63, 97)
    lightBg = RGB(241, 245, 249)
    titleClr = RGB(20, 35, 55)
    descClr = RGB(100, 110, 125)
    borderClr = RGB(226, 232, 240)
    accentClr = RGB(34, 63, 97)

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Guide").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Guide"
    ws.Tab.Color = navy
    ws.Visible = xlSheetVisible

    Application.ScreenUpdating = False
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    Dim colWidth As Double
    colWidth = (DESIGN_WIDTH - MARGIN * 2 - COL_GAP) / 2

    Dim leftColX As Double, rightColX As Double
    leftColX = MARGIN
    rightColX = MARGIN + colWidth + COL_GAP

    Dim y As Double
    y = 24

    ' ================= BANNER (FULL WIDTH) =================
    Dim bannerShp As Excel.Shape
    Set bannerShp = ws.Shapes.AddShape(msoShapeRectangle, MARGIN, y, DESIGN_WIDTH - MARGIN * 2, 56)
    bannerShp.Fill.ForeColor.RGB = navy
    bannerShp.Line.Visible = msoFalse
    bannerShp.Locked = True

    Guide_AddText ws, MARGIN + 20, y + 9, 500, 22, "USER GUIDE", 16, True, RGB(255, 255, 255), "Calibri", msoAlignLeft
    Guide_AddText ws, MARGIN + 20, y + 31, 500, 16, "AR Risk and Aging Center - Quick reference for everyday use", 8.5, False, RGB(200, 215, 230), "Calibri", msoAlignLeft

    Dim backBtn As Excel.Shape
    Set backBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, DESIGN_WIDTH - MARGIN - 110, y + 13, 90, 30)
    backBtn.Adjustments.item(1) = 0.3
    backBtn.Fill.ForeColor.RGB = RGB(255, 255, 255)
    backBtn.Line.Visible = msoFalse
    backBtn.OnAction = "Guide_GoToHome"
    backBtn.Locked = True
    With backBtn.TextFrame2
        .TextRange.Text = ChrW(8249) & " Home"
        .TextRange.Font.Name = "Calibri"
        .TextRange.Font.Size = 9
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = navy
        .VerticalAnchor = msoAnchorMiddle
        .HorizontalAnchor = msoAnchorCenter
    End With

    y = y + 78

    Dim topOfColumns As Double
    topOfColumns = y

    ' ================= LEFT COLUMN HEADER =================
    Guide_AddText ws, leftColX, y, colWidth, 20, "HOW TO USE THE APPLICATION", 12.5, True, navy, "Calibri", msoAlignLeft
    Dim leftHeaderDiv As Excel.Shape
    Set leftHeaderDiv = ws.Shapes.AddLine(leftColX, y + 22, leftColX + colWidth, y + 22)
    leftHeaderDiv.Line.ForeColor.RGB = navy
    leftHeaderDiv.Line.Weight = 2
    leftHeaderDiv.Locked = True

    ' ================= RIGHT COLUMN HEADER =================
    Guide_AddText ws, rightColX, y, colWidth, 20, "SYSTEM REQUIREMENTS AND SETUP", 12.5, True, navy, "Calibri", msoAlignLeft
    Dim rightHeaderDiv As Excel.Shape
    Set rightHeaderDiv = ws.Shapes.AddLine(rightColX, y + 22, rightColX + colWidth, y + 22)
    rightHeaderDiv.Line.ForeColor.RGB = navy
    rightHeaderDiv.Line.Weight = 2
    rightHeaderDiv.Locked = True

    y = y + 34

    Dim leftY As Double, rightY As Double
    leftY = y
    rightY = y

    ' ============================================================
    ' LEFT COLUMN CONTENT - HOW TO USE
    ' ============================================================
    Guide_AddSectionHeader ws, "GETTING STARTED", "Steps before your first report.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr, accentClr
    leftY = leftY + 34
    Guide_AddStep ws, 1, "Set Key Date and Filters", "Open Maintenance, click Report Data and Filters, and confirm your SAP key date.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr
    Guide_AddStep ws, 2, "Update Partner Lists", "Click Partners List to review Factoring and Atradius customers, then save to auto-hide.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr
    Guide_AddStep ws, 3, "You Are Ready", "Once data and partner lists are current, any report button is ready to use.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr

    leftY = leftY + 16

    Guide_AddSectionHeader ws, "REPORTS OVERVIEW", "What each report button generates.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr, accentClr
    leftY = leftY + 34
    Guide_AddRow ws, "Factoring Risk Report", "Weekly risk summary for BNL, ISG, SFR, IUK, AMI, ITA, VNL. Threshold: 4,999 EUR.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Company Code Risk Report", "Same risk logic applied to every company code. Threshold: 4,999 EUR.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Full Aging Report", "Complete, unfiltered aging breakdown across all buckets, every company code.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Full Aging Single Report", "Same as above, scoped to one company code you select.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr

    leftY = leftY + 16

    Guide_AddSectionHeader ws, "SENDING EMAILS", "Notify regional colleagues from the Factoring Report.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr, accentClr
    leftY = leftY + 34
    Guide_AddStep ws, 1, "Find the Region", "Locate the grey subtotal row for the company code.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr
    Guide_AddStep ws, 2, "Click Email", "Click the blue Email button in Column K next to that row.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr
    Guide_AddStep ws, 3, "Review and Send", "Outlook opens with recipients, a table, and an attachment already prepared.", leftColX, leftColX + colWidth, leftY, navy, titleClr, descClr

    leftY = leftY + 16

    Guide_AddSectionHeader ws, "EXPORT AND MAINTENANCE", "Save reports and keep the workbook tidy.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr, accentClr
    leftY = leftY + 34
    Guide_AddRow ws, "Export a Single Report", "Pick one report and save it as a standalone Excel file.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Export All Reports", "Saves every generated report into a folder of your choice.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Delete All Sheets", "Removes generated tabs after exporting. Keeps Factoring Report history.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Reset Tracker", "Full reset. Clears reports, history, and tabs to start a new cycle.", leftColX, leftColX + colWidth, leftY, titleClr, descClr, borderClr

    ' ============================================================
    ' RIGHT COLUMN CONTENT - SETUP AND REQUIREMENTS
    ' ============================================================
    Guide_AddSectionHeader ws, "OVERVIEW", "A fully self-contained workbook.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 30
    Guide_AddPlainDesc ws, "No installation, no server, no network dependency. Each colleague can download and run their own local copy independently.", rightColX, rightColX + colWidth, rightY, descClr
    rightY = rightY + 32

    Guide_AddSectionHeader ws, "SOFTWARE REQUIREMENTS", "What must be installed.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 34
    Guide_AddRow ws, "Microsoft Excel 2016+", "Required. Runs the application.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "SAP Analysis for Office", "Required for SAP refresh via Report Data and Filters.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Microsoft Outlook", "Required for one-click regional emails.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "VPN or corporate network", "Needed only when refreshing SAP data.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr

    rightY = rightY + 16

    Guide_AddSectionHeader ws, "FIRST-TIME SETUP", "One minute to get started.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 34
    Guide_AddStep ws, 1, "Save Locally", "Save the workbook to your own local drive. No shared location needed.", rightColX, rightColX + colWidth, rightY, navy, titleClr, descClr
    Guide_AddStep ws, 2, "Open the File", "Open the workbook as usual.", rightColX, rightColX + colWidth, rightY, navy, titleClr, descClr
    Guide_AddStep ws, 3, "Enable Content", "Click Enable Content when prompted. This is a macro-enabled workbook.", rightColX, rightColX + colWidth, rightY, navy, titleClr, descClr

    rightY = rightY + 16

    Guide_AddSectionHeader ws, "WHAT YOU DO NOT NEED", "No extra setup required.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 30
    Guide_AddPlainDesc ws, "No shared network drive required. No external files or images to install, everything is embedded in the workbook. No IT installation or admin rights needed beyond standard Office access.", rightColX, rightColX + colWidth, rightY, descClr
    rightY = rightY + 44

    Guide_AddSectionHeader ws, "TROUBLESHOOTING", "If something does not work.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 34
    Guide_AddRow ws, "Dashboard looks broken", "Click Build Home Screen in the Admin Panel.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Buttons do not respond", "Confirm macros are enabled in Trust Center settings.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "SAP warning appears", "SAP Analysis add-in may not be installed. Contact IT.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr
    Guide_AddRow ws, "Email button fails", "Confirm Outlook is installed and set as default mail app.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr

    rightY = rightY + 16

    Guide_AddSectionHeader ws, "SUPPORT", "Who to contact.", rightColX, rightColX + colWidth, rightY, titleClr, descClr, borderClr, accentClr
    rightY = rightY + 30
    Guide_AddPlainDesc ws, "For any issues not resolved above, please contact the AR Factoring Risk Automation owner.", rightColX, rightColX + colWidth, rightY, descClr
    rightY = rightY + 30

    ' ================= VERTICAL DIVIDER BETWEEN COLUMNS =================
    Dim maxY As Double
    maxY = leftY
    If rightY > maxY Then maxY = rightY

    Dim vDivider As Excel.Shape
    Set vDivider = ws.Shapes.AddLine(leftColX + colWidth + COL_GAP / 2, topOfColumns, leftColX + colWidth + COL_GAP / 2, maxY)
    vDivider.Line.ForeColor.RGB = borderClr
    vDivider.Line.Weight = 1
    vDivider.Locked = True

    ' ================= FOOTER (FULL WIDTH) =================
    Dim footY As Double
    footY = maxY + 14

    Dim footDiv As Excel.Shape
    Set footDiv = ws.Shapes.AddLine(MARGIN, footY, DESIGN_WIDTH - MARGIN, footY)
    footDiv.Line.ForeColor.RGB = borderClr
    footDiv.Line.Weight = 1
    footDiv.Locked = True

    Guide_AddText ws, MARGIN, footY + 8, 460, 14, "Villeroy and Boch  |  Accounts Receivable  |  AR Risk and Aging Center v1.0", 7.5, False, RGB(60, 70, 85), "Calibri", msoAlignLeft
    Guide_AddText ws, MARGIN, footY + 22, 520, 14, "Internal use only for V&B AR operations. (C) " & Year(Date) & " Villeroy and Boch. All rights reserved.", 6.5, False, descClr, "Calibri", msoAlignLeft

    footY = footY + 50

    ' ================= BACKGROUND, SIZED TO FINAL CONTENT =================
    Dim bgShape As Excel.Shape
    Set bgShape = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, DESIGN_WIDTH, footY)
    bgShape.Fill.ForeColor.RGB = lightBg
    bgShape.Line.Visible = msoFalse
    bgShape.Locked = True
    bgShape.ZOrder msoSendToBack

    ' ================= CENTER + LOCK + PROTECT =================
    Dim offsetX As Double
    offsetX = 0
    On Error Resume Next
    offsetX = (ActiveWindow.usableWidth - DESIGN_WIDTH) / 2
    On Error GoTo 0
    If offsetX < 0 Then offsetX = 0
    If offsetX > 2 Then
        Dim shp As Excel.Shape
        For Each shp In ws.Shapes
            shp.Left = shp.Left + offsetX
        Next shp
    End If

    Dim shpLock As Excel.Shape
    For Each shpLock In ws.Shapes
        On Error Resume Next
        shpLock.Placement = xlFreeFloating
        shpLock.Locked = True
        On Error GoTo 0
    Next shpLock

    ws.Columns("A:Z").ColumnWidth = 9
    ws.Rows.rowHeight = 15

    ws.EnableSelection = xlNoSelection
    ws.Protect Password:="VB_AR_2026", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

    Application.ScreenUpdating = True

    MsgBox "Guide sheet rebuilt with two-column layout.", vbInformation, "Setup Complete"
End Sub

Sub Guide_AddSectionHeader(ws As Worksheet, labelText As String, descText As String, leftPos As Double, rightPos As Double, topPos As Double, titleClr As Long, descClr As Long, borderClr As Long, accentClr As Long)
    Dim accentBar As Excel.Shape
    Set accentBar = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, 4, 18)
    accentBar.Fill.ForeColor.RGB = accentClr
    accentBar.Line.Visible = msoFalse
    accentBar.Locked = True

    Guide_AddText ws, leftPos + 10, topPos - 1, rightPos - leftPos - 10, 18, labelText, 11, True, titleClr, "Calibri", msoAlignLeft

    Dim divLine As Excel.Shape
    Set divLine = ws.Shapes.AddLine(leftPos, topPos + 22, rightPos, topPos + 22)
    divLine.Line.ForeColor.RGB = borderClr
    divLine.Line.Weight = 1
    divLine.Locked = True
End Sub

Sub Guide_AddStep(ws As Worksheet, num As Long, stepTitle As String, stepDesc As String, leftPos As Double, rightPos As Double, ByRef topPos As Double, circleClr As Long, titleClr As Long, descClr As Long)
    Dim shpCircle As Excel.Shape
    Dim titleShp As Excel.Shape
    Dim descShp As Excel.Shape

    Set shpCircle = ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, 20, 20)
    With shpCircle
        .Fill.ForeColor.RGB = circleClr
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = CStr(num)
            .TextRange.Font.Size = 9
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
    End With

    Set titleShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 28, topPos, rightPos - leftPos - 28, 16)
    With titleShp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = stepTitle
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = titleClr
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With

    Set descShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos + 28, topPos + 16, rightPos - leftPos - 28, 28)
    With descShp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = stepDesc
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Fill.ForeColor.RGB = descClr
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .WordWrap = msoTrue
        End With
    End With

    topPos = topPos + 46
End Sub

Sub Guide_AddRow(ws As Worksheet, rowTitle As String, rowDesc As String, leftPos As Double, rightPos As Double, ByRef topPos As Double, titleClr As Long, descClr As Long, borderClr As Long)
    Dim titleShp As Excel.Shape
    Dim descShp As Excel.Shape
    Dim divLine As Excel.Shape

    Set titleShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, rightPos - leftPos, 15)
    With titleShp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = rowTitle
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 8.5
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = titleClr
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End With
    End With

    Set descShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos + 14, rightPos - leftPos, 26)
    With descShp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = rowDesc
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Fill.ForeColor.RGB = descClr
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .WordWrap = msoTrue
        End With
    End With

    Set divLine = ws.Shapes.AddLine(leftPos, topPos + 40, rightPos, topPos + 40)
    divLine.Line.ForeColor.RGB = borderClr
    divLine.Line.Weight = 1
    divLine.Locked = True

    topPos = topPos + 46
End Sub

Sub Guide_AddPlainDesc(ws As Worksheet, descText As String, leftPos As Double, rightPos As Double, topPos As Double, descClr As Long)
    Dim descShp As Excel.Shape
    Set descShp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, rightPos - leftPos, 34)
    With descShp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Locked = True
        With .TextFrame2
            .TextRange.Text = descText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 8
            .TextRange.Font.Fill.ForeColor.RGB = descClr
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .WordWrap = msoTrue
        End With
    End With
End Sub

Sub Guide_AddText(ws As Worksheet, l As Double, t As Double, w As Double, h As Double, txt As String, sz As Double, bold As Boolean, clr As Long, fontName As String, align As Long)
    Dim shp As Excel.Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Locked = True
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Name = fontName
        .TextRange.Font.Size = sz
        .TextRange.Font.bold = bold
        .TextRange.Font.Fill.ForeColor.RGB = clr
        .TextRange.ParagraphFormat.Alignment = align
    End With
End Sub
Sub Guide_GoToHome()
    On Error Resume Next
    ThisWorkbook.Sheets("Home").Activate
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    On Error GoTo 0
End Sub
