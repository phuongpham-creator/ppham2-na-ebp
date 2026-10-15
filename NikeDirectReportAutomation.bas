Attribute VB_Name = "NikeDirectReportAutomation"
Option Explicit

' ==========================================================
' Nike Direct WOS report automation
' Recommended host workbook: a macro-enabled master copy of
' the Nike Direct WOS file (.xlsm). The macro writes into the
' existing BOP Data / Summary tabs without changing formulas.
' ==========================================================

Private Const DEST_BOP_SHEET As String = "BOP Data"
Private Const DEST_SUMMARY_SHEET As String = "Summary"
Private Const DEST_SALES_SHEET As String = "Sales Data"
Private Const SOURCE_PIVOT_SHEET_KEYWORD As String = "Channel WOS"
Private Const SOURCE_PIVOT_NAME As String = "PivotTable1"

' OLAP pivot fields
Private Const PF_PROMO As String = "[Query].[PROMO_IND].[PROMO_IND]"
Private Const PF_SAMPLE As String = "[Query].[SAMPLE_IND].[SAMPLE_IND]"
Private Const PF_CLOSEOUT As String = "[Query].[CLOSEOUT_IND].[CLOSEOUT_IND]"
Private Const PF_CHANNEL_GROUP As String = "[Query].[CHANNEL_GROUP].[CHANNEL_GROUP]"
Private Const PF_AA As String = "[Query].[AA_IND].[AA_IND]"
Private Const PF_TEAM As String = "[Query].[TEAM_IND].[TEAM_IND]"
Private Const PF_MONTH_REL As String = "[Query].[MONTH_RELEVANCE_DT].[MONTH_RELEVANCE_DT]"
Private Const PF_OFFER_DATE As String = "[Query].[RGN_FRST_PRODT_OFR_DT].[RGN_FRST_PRODT_OFR_DT]"

' OLAP member names
Private Const MEMBER_N As String = "&[N]"
Private Const MEMBER_Y As String = "&[Y]"
Private Const MEMBER_ALL As String = "[All]"

Private Type tChannelMap
    Channel As String
    Country As String
End Type

Public Sub CreateNextMonthReportCopy()
    Dim reportMonth As Date
    Dim newFilePath As String
    Dim newWb As Workbook
    Dim irPath As String
    Dim currentStep As String

    On Error GoTo CleanFail

    currentStep = "Prompt for report month"
    reportMonth = PromptForReportMonth()
    If reportMonth = 0 Then Exit Sub

    newFilePath = ThisWorkbook.Path & Application.PathSeparator & _
                  "Nike Direct WOS " & Format$(reportMonth, "mmmm yyyy") & ".xlsm"

    If Dir$(newFilePath) <> vbNullString Then
        If MsgBox("The target file already exists:" & vbCrLf & newFilePath & vbCrLf & vbCrLf & _
                  "Do you want to overwrite it?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        Kill newFilePath
    End If

    currentStep = "Create report copy"
    ThisWorkbook.SaveCopyAs newFilePath
    currentStep = "Open report copy"
    Set newWb = Workbooks.Open(newFilePath)
    If newWb Is Nothing Then
        Err.Raise vbObjectError + 2000, , "The new report copy could not be opened."
    End If

    currentStep = "Select IR workbook"
    irPath = GetIrWorkbookPath()
    If Len(irPath) = 0 Then
        MsgBox "No IR workbook was selected. The new report copy was created but not refreshed.", vbExclamation
        Exit Sub
    End If

    currentStep = "Run refresh against copied workbook"
    RunNikeDirectRefreshForWorkbook newWb, reportMonth, irPath
    MsgBox "New report created and refreshed:" & vbCrLf & newFilePath, vbInformation
    Exit Sub

CleanFail:
    MsgBox "Create/report setup failed during step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical
End Sub

Public Sub RunNikeDirectRefresh()
    RunNikeDirectRefreshForWorkbook ThisWorkbook
End Sub

Public Sub RunNikeDirectRefreshInternal(Optional ByVal reportMonth As Date = 0, Optional ByVal irPath As String = "")
    RunNikeDirectRefreshForWorkbook ThisWorkbook, reportMonth, irPath
End Sub

Private Sub RunNikeDirectRefreshForWorkbook(ByVal hostWb As Workbook, Optional ByVal reportMonth As Date = 0, Optional ByVal irPath As String = "")
    Dim irWb As Workbook
    Dim srcWs As Worksheet
    Dim destBopWs As Worksheet
    Dim destSummaryWs As Worksheet
    Dim pt As PivotTable
    Dim monthRelCurrent As Date
    Dim monthRelLy As Date
    Dim cutoffDate As Date
    Dim currentStep As String

    On Error GoTo CleanFail

    If reportMonth = 0 Then
        currentStep = "Prompt for report month"
        reportMonth = PromptForReportMonth()
        If reportMonth = 0 Then Exit Sub
    End If

    If Len(irPath) = 0 Then
        currentStep = "Select IR workbook"
        irPath = GetIrWorkbookPath()
        If Len(irPath) = 0 Then Exit Sub
    End If

    monthRelCurrent = DateAdd("m", -1, reportMonth)
    monthRelLy = DateAdd("yyyy", -1, monthRelCurrent)
    cutoffDate = DateAdd("d", -91, FirstBusinessDay(reportMonth))

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Opening and refreshing IR workbook..."

    currentStep = "Open IR workbook"
    Set irWb = GetOrOpenWorkbook(irPath)
    currentStep = "Refresh IR workbook connections"
    RefreshWorkbookAndWait irWb

    currentStep = "Resolve source and destination sheets"
    Set srcWs = FindWorksheetContains(irWb, SOURCE_PIVOT_SHEET_KEYWORD)
    If srcWs Is Nothing Then
        Err.Raise vbObjectError + 2001, , "Source worksheet containing '" & SOURCE_PIVOT_SHEET_KEYWORD & "' was not found."
    End If

    Set destBopWs = hostWb.Worksheets(DEST_BOP_SHEET)
    If destBopWs Is Nothing Then
        Err.Raise vbObjectError + 2002, , "Destination sheet '" & DEST_BOP_SHEET & "' was not found."
    End If

    Set destSummaryWs = hostWb.Worksheets(DEST_SUMMARY_SHEET)
    If destSummaryWs Is Nothing Then
        Err.Raise vbObjectError + 2003, , "Destination sheet '" & DEST_SUMMARY_SHEET & "' was not found."
    End If

    On Error Resume Next
    Set pt = srcWs.PivotTables(SOURCE_PIVOT_NAME)
    On Error GoTo CleanFail
    If pt Is Nothing Then
        Err.Raise vbObjectError + 2004, , "Pivot table '" & SOURCE_PIVOT_NAME & "' was not found on " & srcWs.Name & "."
    End If

    ' Step 1: current month BOP block
    Application.StatusBar = "Loading current BOP data..."
    currentStep = "Apply current-month filters for BOP"
    ApplyBaseFilters pt, monthRelCurrent, "All"
    currentStep = "Clear offer-date filter for current BOP"
    ClearOfferDateFilter pt
    currentStep = "Refresh current BOP pivot view"
    RefreshPivotAndWait pt
    currentStep = "Write current BOP block"
    WriteBopBlock srcWs, destBopWs.Range("A1"), True

    ' Step 2: LY BOP block
    Application.StatusBar = "Loading last-year BOP data..."
    currentStep = "Apply LY filters for BOP"
    ApplyBaseFilters pt, monthRelLy, "All"
    currentStep = "Clear offer-date filter for LY BOP"
    ClearOfferDateFilter pt
    currentStep = "Refresh LY BOP pivot view"
    RefreshPivotAndWait pt
    currentStep = "Write LY BOP block"
    WriteBopBlock srcWs, destBopWs.Range("A10"), False

    ' Step 3: closeout totals on Summary using current year month relevance
    Application.StatusBar = "Loading closeout totals..."
    currentStep = "Apply closeout filters"
    ApplyBaseFilters pt, monthRelCurrent, "Y"
    currentStep = "Clear offer-date filter for closeout totals"
    ClearOfferDateFilter pt
    currentStep = "Refresh closeout total pivot view"
    RefreshPivotAndWait pt
    currentStep = "Write closeout totals"
    WriteCloseoutTotals srcWs, destSummaryWs

    ' Step 4.1: active WOH (all products, <= cutoff)
    Application.StatusBar = "Loading active WOH..."
    currentStep = "Apply all-product filters for active WOH"
    ApplyBaseFilters pt, monthRelCurrent, "All"
    currentStep = "Apply offer-date cutoff for active WOH"
    SetOfferDateFilterUpToCutoff pt, cutoffDate
    currentStep = "Refresh active WOH pivot view"
    RefreshPivotAndWait pt
    currentStep = "Write active WOH"
    WriteWohBlock srcWs, destSummaryWs, False

    ' Step 4.2: closeout WOH (closeout only, <= cutoff)
    Application.StatusBar = "Loading closeout WOH..."
    currentStep = "Apply closeout-only filters for closeout WOH"
    ApplyBaseFilters pt, monthRelCurrent, "Y"
    currentStep = "Apply offer-date cutoff for closeout WOH"
    SetOfferDateFilterUpToCutoff pt, cutoffDate
    currentStep = "Refresh closeout WOH pivot view"
    RefreshPivotAndWait pt
    currentStep = "Write closeout WOH"
    WriteWohBlock srcWs, destSummaryWs, True

    ' Optional audit trail
    currentStep = "Write automation log"
    WriteAuditFields hostWb, reportMonth, monthRelCurrent, monthRelLy, cutoffDate, irWb.FullName

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Nike Direct inventory refresh is complete." & vbCrLf & vbCrLf & _
           "Please update the Sales Data tab from the external system, then review the Summary tab.", _
           vbInformation
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Refresh failed during step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical
End Sub

Private Sub ApplyBaseFilters(ByVal pt As PivotTable, ByVal monthRel As Date, ByVal closeoutMode As String)
    SetSingleMemberFilter pt, PF_PROMO, BuildSimpleMember("PROMO_IND", "N")
    SetSingleMemberFilter pt, PF_SAMPLE, BuildSimpleMember("SAMPLE_IND", "N")
    SetSingleMemberFilter pt, PF_CHANNEL_GROUP, BuildAllMember("CHANNEL_GROUP")
    SetSingleMemberFilter pt, PF_AA, BuildAllMember("AA_IND")
    SetSingleMemberFilter pt, PF_TEAM, BuildAllMember("TEAM_IND")
    SetSingleMemberFilter pt, PF_MONTH_REL, BuildMonthMember(monthRel)

    If UCase$(closeoutMode) = "Y" Then
        SetSingleMemberFilter pt, PF_CLOSEOUT, BuildSimpleMember("CLOSEOUT_IND", "Y")
    Else
        SetSingleMemberFilter pt, PF_CLOSEOUT, BuildAllMember("CLOSEOUT_IND")
    End If
End Sub

Private Sub ClearOfferDateFilter(ByVal pt As PivotTable)
    Dim pf As PivotField
    Set pf = pt.PivotFields(PF_OFFER_DATE)
    On Error Resume Next
    pf.ClearAllFilters
    On Error GoTo 0
End Sub

Private Sub SetOfferDateFilterUpToCutoff(ByVal pt As PivotTable, ByVal cutoffDate As Date)
    Dim pf As PivotField
    Dim itm As PivotItem
    Dim members() As Variant
    Dim count As Long
    Dim dt As Date

    Set pf = pt.PivotFields(PF_OFFER_DATE)
    pf.ClearAllFilters

    count = -1
    For Each itm In pf.PivotItems
        If TryParsePivotDate(itm.Caption, dt) Then
            If dt <= cutoffDate Then
                count = count + 1
                ReDim Preserve members(0 To count)
                members(count) = itm.Name
            End If
        End If
    Next itm

    If count < 0 Then
        Err.Raise vbObjectError + 1000, , "No offer-date members were found on or before cutoff " & Format$(cutoffDate, "m/d/yyyy")
    End If

    pf.VisibleItemsList = members
End Sub

Private Sub RefreshWorkbookAndWait(ByVal wb As Workbook)
    wb.RefreshAll
    WaitForExcelRefresh wb
End Sub

Private Sub RefreshPivotAndWait(ByVal pt As PivotTable)
    On Error Resume Next
    pt.RefreshTable
    On Error GoTo 0
    WaitForExcelRefresh pt.Parent.Parent
End Sub

Private Sub WaitForExcelRefresh(ByVal wb As Workbook)
    Dim c As WorkbookConnection
    Dim stillRefreshing As Boolean

    Do
        stillRefreshing = False

        For Each c In wb.Connections
            On Error Resume Next
            stillRefreshing = stillRefreshing Or c.OLEDBConnection.Refreshing
            On Error GoTo 0
        Next c

        If Application.CalculationState <> xlDone Then stillRefreshing = True
        DoEvents
    Loop While stillRefreshing
End Sub

Private Sub SetSingleMemberFilter(ByVal pt As PivotTable, ByVal fieldName As String, ByVal memberUniqueName As String)
    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)
    pf.ClearAllFilters
    pf.VisibleItemsList = Array(memberUniqueName)
End Sub

Private Function BuildSimpleMember(ByVal hierarchyName As String, ByVal memberCaption As String) As String
    BuildSimpleMember = "[Query].[" & hierarchyName & "].&[" & memberCaption & "]"
End Function

Private Function BuildAllMember(ByVal hierarchyName As String) As String
    BuildAllMember = "[Query].[" & hierarchyName & "].[All]"
End Function

Private Function BuildMonthMember(ByVal d As Date) As String
    BuildMonthMember = "[Query].[MONTH_RELEVANCE_DT].&[" & Format$(d, "yyyy-mm-dd") & "T00:00:00]"
End Function

Private Function TryParsePivotDate(ByVal rawCaption As String, ByRef outDate As Date) As Boolean
    On Error GoTo Fail
    If Len(Trim$(rawCaption)) = 0 Then Exit Function
    outDate = CDate(rawCaption)
    TryParsePivotDate = True
    Exit Function
Fail:
    TryParsePivotDate = False
End Function

Private Sub WriteBopBlock(ByVal srcWs As Worksheet, ByVal destTopLeft As Range, ByVal includeHeader As Boolean)
    Dim rowsOut() As Variant
    Dim r As Long, startRow As Long
    Dim channels As Variant

    channels = Array( _
        Array("NIKE FACTORY STORES", "CANADA"), _
        Array("NIKE FACTORY STORES", "UNITED STATES OF AMERICA"), _
        Array("NIKE STORES OWNED", "UNITED STATES OF AMERICA"), _
        Array("NIKE.COM", "UNITED STATES OF AMERICA"), _
        Array("Grand Total", "") _
    )

    If includeHeader Then
        ReDim rowsOut(1 To 6, 1 To 8)
        rowsOut(1, 1) = "CHANNEL_DESC"
        rowsOut(1, 2) = "COUNTRY"
        rowsOut(1, 3) = "IN TRANSIT QTY"
        rowsOut(1, 4) = "DC QTY"
        rowsOut(1, 5) = "TOTAL STORE IN TRANSIT QTY"
        rowsOut(1, 6) = "STORE OH QTY"
        rowsOut(1, 7) = "TOTAL INVENTORY"
        rowsOut(1, 8) = "Accessible"
        startRow = 2
    Else
        ReDim rowsOut(1 To 5, 1 To 8)
        startRow = 1
    End If

    For r = LBound(channels) To UBound(channels)
        Dim rowIndex As Long
        rowIndex = startRow + r
        rowsOut(rowIndex, 1) = channels(r)(0)
        rowsOut(rowIndex, 2) = channels(r)(1)
        rowsOut(rowIndex, 3) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "IN TRANSIT QTY")
        rowsOut(rowIndex, 4) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "DC QTY")
        rowsOut(rowIndex, 5) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "TOTAL STORE IN TRANSIT QTY")
        rowsOut(rowIndex, 6) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "STORE OH QTY")
        rowsOut(rowIndex, 7) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "TOTAL INVENTORY")
        rowsOut(rowIndex, 8) = ReadDisplayedMetric(srcWs, channels(r)(0), channels(r)(1), "AVAILABLE QTY")
    Next r

    destTopLeft.Resize(UBound(rowsOut, 1), UBound(rowsOut, 2)).Value = rowsOut
End Sub

Private Sub WriteCloseoutTotals(ByVal srcWs As Worksheet, ByVal summaryWs As Worksheet)
    summaryWs.Range("G11").Value = ReadDisplayedMetric(srcWs, "NIKE.COM", "UNITED STATES OF AMERICA", "TOTAL INVENTORY")
    summaryWs.Range("I11").Value = ReadDisplayedMetric(srcWs, "NIKE STORES OWNED", "UNITED STATES OF AMERICA", "TOTAL INVENTORY")
    summaryWs.Range("K11").Value = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "UNITED STATES OF AMERICA", "TOTAL INVENTORY")
    summaryWs.Range("M11").Value = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "CANADA", "TOTAL INVENTORY")
End Sub

Private Sub WriteWohBlock(ByVal srcWs As Worksheet, ByVal summaryWs As Worksheet, ByVal closeoutOnly As Boolean)
    Dim qtyNddc As Double, qtyNso As Double, qtyNvsUs As Double, qtyNvsCa As Double
    Dim wohNddc As Double, wohNso As Double, wohNvsUs As Double, wohNvsCa As Double
    Dim totalWoh As Double
    Dim targetRow As Long

    If closeoutOnly Then
        targetRow = 21
    Else
        targetRow = 20
    End If

    qtyNddc = ReadDisplayedMetric(srcWs, "NIKE.COM", "UNITED STATES OF AMERICA", "AVAILABLE QTY")
    qtyNso = ReadDisplayedMetric(srcWs, "NIKE STORES OWNED", "UNITED STATES OF AMERICA", "AVAILABLE QTY")
    qtyNvsUs = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "UNITED STATES OF AMERICA", "AVAILABLE QTY")
    qtyNvsCa = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "CANADA", "AVAILABLE QTY")

    wohNddc = ReadDisplayedMetric(srcWs, "NIKE.COM", "UNITED STATES OF AMERICA", "AVAILABLE WOH")
    wohNso = ReadDisplayedMetric(srcWs, "NIKE STORES OWNED", "UNITED STATES OF AMERICA", "AVAILABLE WOH")
    wohNvsUs = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "UNITED STATES OF AMERICA", "AVAILABLE WOH")
    wohNvsCa = ReadDisplayedMetric(srcWs, "NIKE FACTORY STORES", "CANADA", "AVAILABLE WOH")

    summaryWs.Range("G" & targetRow).Value = wohNddc
    summaryWs.Range("I" & targetRow).Value = wohNso
    summaryWs.Range("K" & targetRow).Value = wohNvsUs
    summaryWs.Range("M" & targetRow).Value = wohNvsCa

    totalWoh = WeightedWoh(Array(qtyNddc, qtyNso, qtyNvsUs, qtyNvsCa), _
                           Array(wohNddc, wohNso, wohNvsUs, wohNvsCa))
    summaryWs.Range("E" & targetRow).Value = totalWoh
End Sub

Private Function WeightedWoh(ByVal qtyArr As Variant, ByVal wohArr As Variant) As Double
    Dim i As Long
    Dim qtySum As Double
    Dim denom As Double

    For i = LBound(qtyArr) To UBound(qtyArr)
        qtySum = qtySum + CDbl(qtyArr(i))
        If CDbl(wohArr(i)) > 0 Then
            denom = denom + (CDbl(qtyArr(i)) / CDbl(wohArr(i)))
        End If
    Next i

    If denom = 0 Then
        WeightedWoh = 0
    Else
        WeightedWoh = qtySum / denom
    End If
End Function

Private Function ReadDisplayedMetric(ByVal ws As Worksheet, ByVal channelName As String, ByVal countryName As String, ByVal metricHeader As String) As Double
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim r As Long
    Dim cellChannel As String
    Dim cellCountry As String

    headerRow = 7
    dataStartRow = 8
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    colIndex = FindHeaderColumn(ws, headerRow, metricHeader, 2, lastCol)

    If colIndex = 0 Then
        Err.Raise vbObjectError + 1001, , "Metric header not found: " & metricHeader
    End If

    For r = dataStartRow To ws.Rows.Count
        cellChannel = Trim$(CStr(ws.Cells(r, 2).Value))
        cellCountry = Trim$(CStr(ws.Cells(r, 3).Value))

        If Len(cellChannel) = 0 Then Exit For

        If UCase$(channelName) = "GRAND TOTAL" Then
            If UCase$(cellChannel) = "GRAND TOTAL" Then
                ReadDisplayedMetric = CellToDouble(ws.Cells(r, colIndex).Value)
                Exit Function
            End If
        ElseIf StrComp(cellChannel, channelName, vbTextCompare) = 0 And _
               StrComp(cellCountry, countryName, vbTextCompare) = 0 Then
            ReadDisplayedMetric = CellToDouble(ws.Cells(r, colIndex).Value)
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 1002, , "Could not find row for channel [" & channelName & _
               "] country [" & countryName & "] in source pivot."
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = firstCol To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value)), headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function CellToDouble(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then
        CellToDouble = 0
    ElseIf Trim$(CStr(v)) = "-" Then
        CellToDouble = 0
    ElseIf IsNumeric(v) Then
        CellToDouble = CDbl(v)
    Else
        CellToDouble = Val(Replace$(CStr(v), ",", ""))
    End If
End Function

Private Sub WriteAuditFields(ByVal hostWb As Workbook, ByVal reportMonth As Date, ByVal monthRelCurrent As Date, ByVal monthRelLy As Date, ByVal cutoffDate As Date, ByVal irPath As String)
    Dim logWs As Worksheet

    On Error Resume Next
    Set logWs = hostWb.Worksheets("Automation Log")
    On Error GoTo 0

    If logWs Is Nothing Then
        Set logWs = hostWb.Worksheets.Add(After:=hostWb.Worksheets(hostWb.Worksheets.Count))
        logWs.Name = "Automation Log"
        logWs.Visible = xlSheetVeryHidden
    End If

    logWs.Cells.ClearContents
    logWs.Range("A1").Value = "Automation Run"
    logWs.Range("B1").Value = Now
    logWs.Range("A2").Value = "Report Month"
    logWs.Range("B2").Value = reportMonth
    logWs.Range("A3").Value = "Current Month Relevance"
    logWs.Range("B3").Value = monthRelCurrent
    logWs.Range("A4").Value = "LY Month Relevance"
    logWs.Range("B4").Value = monthRelLy
    logWs.Range("A5").Value = "13-Week Cutoff"
    logWs.Range("B5").Value = cutoffDate
    logWs.Range("A6").Value = "IR Workbook"
    logWs.Range("B6").Value = irPath
End Sub

Private Function PromptForReportMonth() As Date
    Dim rawInput As String
    rawInput = InputBox("Enter the report month as the first day of month." & vbCrLf & _
                        "Example: 10/1/2026", "Nike Direct Report Month", _
                        Format$(DateSerial(Year(Date), Month(Date), 1), "m/d/yyyy"))
    If Len(Trim$(rawInput)) = 0 Then Exit Function
    PromptForReportMonth = CDate(rawInput)
End Function

Private Function FirstBusinessDay(ByVal d As Date) As Date
    Dim firstDay As Date
    firstDay = DateSerial(Year(d), Month(d), 1)

    Do While Weekday(firstDay, vbMonday) > 5
        firstDay = firstDay + 1
    Loop

    FirstBusinessDay = firstDay
End Function

Private Function GetIrWorkbookPath() As String
    Dim pickedFile As Variant
    pickedFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        Title:="Select Monthly Inventory IR workbook")

    If VarType(pickedFile) = vbBoolean Then Exit Function
    GetIrWorkbookPath = CStr(pickedFile)
End Function

Private Function GetOrOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook = Workbooks.Open(fullPath)
End Function

Private Function FindWorksheetContains(ByVal wb As Workbook, ByVal textFragment As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, textFragment, vbTextCompare) > 0 Then
            Set FindWorksheetContains = ws
            Exit Function
        End If
    Next ws

    Err.Raise vbObjectError + 1003, , "Worksheet containing [" & textFragment & "] was not found in " & wb.Name
End Function
