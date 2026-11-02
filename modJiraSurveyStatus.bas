Attribute VB_Name = "modJiraSurveyStatus"
Option Explicit

'===============================================================================
' JIRA SURVEY STATUS AUTOMATION - VERSION 2.1
' Developed by: Perry Dave V. Geron
'
' PURPOSE
'   Generate a status workbook; optionally refresh an existing report by Issue Key.
'   The report applies date-based survey milestones, RVQ overdue checks,
'   and separate historical pending-issue / issue-resolution indicators.
'
' BUSINESS RULES
'   Stage priority: Approved > RVQ Available > VQ Released > Initial VQ
'                   > Not yet started.
'   RVQ overdue: no RVQ after MORE than 14 Monday-Friday business days,
'                starting the day after VQ release (day 15 is overdue).
'   Pending Issue and Resolution do NOT override the calculated survey stage.
'   VQ Released and Approved issue indicators show N/A; remarks are retained.
'   No general analyst turnaround calculation is performed.
'
' IMPORTANT
'   CSV date values are parsed as DD/MM/YYYY, regardless of PC locale.
'   Dates are displayed as MMM DD, with full year retained in the Excel value.
'   Original JIRA exports are never saved or modified. Before refreshing an
'   existing saved report, a backup is created and validated.
'===============================================================================

Private Const OUTPUT_HEADER_ROW As Long = 7
Private Const OUTPUT_FIRST_DATA_ROW As Long = 8
Private Const OVERDUE_THRESHOLD As Long = 14
Private Const JSS_METADATA_SHEET As String = "_JSS_Metadata"
Private Const JSS_MATCH_HEADER As String = "Latest Export Match"
' JIRA ticket URL: show only Issue Key in the visible Excel cell.
Private Const JIRA_ISSUE_BASE_URL As String = "https://jira.ehr.com/browse/"

' Entry point for Excel Ribbon XML. Ribbon calls this procedure directly.
Public Sub JSS_RibbonRun(ByVal control As IRibbonControl)
    JSS_GenerateSurveyStatus
End Sub

' Ribbon command for incremental refresh. The workbook on screen is the
' existing status report, not the new JIRA export.
Public Sub JSS_UpdateRibbon(ByVal control As IRibbonControl)
    JSS_UpdateSurveyStatus
End Sub

' Public, first-time report generation entry point.
Public Sub JSS_GenerateSurveyStatus()
    Dim produced As Workbook, unusedFallback As Boolean
    If Application.Workbooks.Count = 0 Then
        MsgBox "Open a JIRA export before clicking Generate Survey Status.", vbExclamation
        Exit Sub
    End If
    Set produced = JSS_GenerateCore(Application.ActiveWorkbook, False, unusedFallback)
End Sub

' Shared report builder: both buttons use the exact same existing date and
' status calculations. Quiet mode lets Update build a temporary source report.
Private Function JSS_GenerateCore(ByVal sourceBook As Workbook, _
    ByVal quiet As Boolean, ByRef usedCsvFallback As Boolean) As Workbook
    Dim sourceSheet As Worksheet
    Dim resultBook As Workbook, resultSheet As Worksheet
    Dim sourceData As Variant, resultData As Variant
    Dim headers As Object, seenKeys As Object, outUsed As Object
    Dim mainNames As Variant, extraNames As Collection
    Dim sourceHeaderRow As Long, sourceCount As Long, sourceCols As Long
    Dim outCols As Long, i As Long, j As Long, srcRow As Long, outRow As Long
    Dim columnName As String, canonical As String, key As String
    Dim stage As String, issueIndicator As String, issueResolution As String
    Dim overdueFlag As String, validationText As String
    Dim businessDays As Variant, dateValues As Object, dateValid As Object, dateGiven As Object
    Dim processed As Long, approved As Long, rvqAvailable As Long, released As Long
    Dim initialVQ As Long, notStarted As Long, overdueCount As Long
    Dim firstSeen As Long, rawValue As Variant, srcIndex As Long
    Dim is1904 As Boolean, sourceFile As String
    Dim startTime As Double, elapsedSeconds As Double, oldCalc As XlCalculation
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim problem As String, resultReady As Boolean
    Dim csvFallback As Boolean, csvReadProblem As String

    ' Start measurement before reading any input. Include report generation.
    startTime = Timer
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    On Error GoTo Failed

    If Application.Workbooks.Count = 0 Then
        Err.Raise vbObjectError + 700, , "Open the JIRA export in Excel before generating a report."
    End If
    If sourceBook Is Nothing Then Err.Raise vbObjectError + 701, , "No active JIRA export was found."
    If sourceBook.IsAddin Then Err.Raise vbObjectError + 702, , "Open the JIRA export instead of selecting an add-in."
    Set sourceSheet = sourceBook.ActiveSheet
    sourceFile = sourceBook.FullName

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Survey status: reading the JIRA export..."

    ' Prefer raw CSV bytes so DD/MM/YYYY values are not reinterpreted by
    ' Excel's regional date settings. Excel may open a cloud-backed CSV via
    ' a virtual/HTTPS path that ADODB.Stream cannot reopen directly.
    If LCase$(Right$(sourceBook.Name, 4)) = ".csv" Then
        On Error Resume Next
        Err.Clear
        sourceData = JSS_ReadCsvAsArray(sourceFile)
        If Err.Number <> 0 Then
            csvFallback = True
            csvReadProblem = Err.Description
            Err.Clear
        End If
        On Error GoTo Failed

        If csvFallback Then
            ' The document is already open. Use its visible worksheet as a
            ' fallback instead of failing the report completely. CAUTION:
            ' Excel may already have interpreted ambiguous DD/MM dates.
            ' A visible warning is added to the resulting report below.
            sourceData = JSS_ReadSheetAsArray(sourceSheet)
            is1904 = sourceBook.Date1904
        Else
            is1904 = False
        End If
    Else
        sourceData = JSS_ReadSheetAsArray(sourceSheet)
        is1904 = sourceBook.Date1904
    End If

    sourceCount = UBound(sourceData, 1)
    sourceCols = UBound(sourceData, 2)
    sourceHeaderRow = JSS_FindHeaderRow(sourceData)
    If sourceHeaderRow = 0 Then
        Err.Raise vbObjectError + 703, , "JIRA headers were not recognized. Required headers include Summary and Issue key."
    End If

    Set headers = JSS_CreateHeaders(sourceData, sourceHeaderRow, sourceCols)
    JSS_ValidateRequiredColumns headers

    ' The first 27 columns are a fixed, flat, stakeholder-friendly layout.
    ' Remaining JIRA fields appear after Data Validation, with no duplicates.
    mainNames = Array("Summary", "Issue Key", "Company Name", "Abbreviation", _
        "SPM Code", "Submission Type", "Submission Category", _
        "Ticket Status", "HL Comment", "DM Comment", "Created", _
        "Initial VQ Complete Date", _
        "Initial VQ Review Complete Date", "VQ Released Date", _
        "Returned VQ Received Date", "Business Days Without RVQ", _
        "RVQ Overdue Flag", "Pending Issue Date", "Pending Issue Indicator", _
        "Pending Issue Remarks", "Issue Resolution Date", _
        "Issue Resolution Status", "Issue Resolution Remarks", _
        "Validation Review Complete Date", "VQ Changes Complete Date", _
        "VQ Approvals Complete Date", "Data Validation")
    Set extraNames = New Collection
    Set outUsed = CreateObject("Scripting.Dictionary")
    outUsed.CompareMode = vbTextCompare
    For i = LBound(mainNames) To UBound(mainNames)
        outUsed.Add JSS_NormalizeHeader(CStr(mainNames(i))), True
    Next i
    For j = 1 To sourceCols
        columnName = Trim$(CStr(sourceData(sourceHeaderRow, j)))
        If Len(columnName) > 0 Then
            canonical = JSS_NormalizeHeader(columnName)
            If Not outUsed.Exists(canonical) And Not JSS_IsManualHeader(canonical) Then
                extraNames.Add JSS_DisplayHeader(columnName)
                outUsed.Add canonical, True
            End If
        End If
    Next j
    outCols = UBound(mainNames) - LBound(mainNames) + 1 + extraNames.Count
    If outCols > 16384 Then Err.Raise vbObjectError + 704, , "Too many columns for an Excel worksheet."

    ' Count only populated ticket rows. Blank and repeated keys are validation
    ' errors; they must never be silently dropped from the final report.
    processed = 0
    For srcRow = sourceHeaderRow + 1 To sourceCount
        If JSS_IsNonEmptyRow(sourceData, srcRow, sourceCols) Then processed = processed + 1
    Next srcRow
    If processed < 1 Then Err.Raise vbObjectError + 705, , "No data records were found below the JIRA headers."
    If processed + OUTPUT_HEADER_ROW > 1048576 Then Err.Raise vbObjectError + 706, , "Excel row limit exceeded."

    ReDim resultData(1 To processed, 1 To outCols)
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbTextCompare
    outRow = 0
    Application.StatusBar = "Survey status: evaluating " & processed & " tickets..."

    For srcRow = sourceHeaderRow + 1 To sourceCount
        If JSS_IsNonEmptyRow(sourceData, srcRow, sourceCols) Then
            outRow = outRow + 1
            Set dateValues = CreateObject("Scripting.Dictionary")
            Set dateValid = CreateObject("Scripting.Dictionary")
            Set dateGiven = CreateObject("Scripting.Dictionary")
            validationText = ""
            JSS_CollectDates sourceData, srcRow, headers, is1904, _
                             dateValues, dateValid, dateGiven, validationText

            ' Every JIRA issue key identifies one physical ticket row. Repeated
            ' identifiers are reported as errors but records remain visible.
            key = Trim$(CStr(JSS_FieldValue(sourceData, srcRow, headers, "Issue Key")))
            If Len(key) = 0 Then
                JSS_AddProblem validationText, "Missing Issue Key"
            ElseIf seenKeys.Exists(key) Then
                JSS_AddProblem validationText, "Duplicate Issue Key"
            Else
                seenKeys.Add key, srcRow
            End If

            stage = JSS_GetSurveyStage(dateValid, dateGiven)
            JSS_CheckChronology dateValues, dateValid, dateGiven, validationText
            JSS_GetIssueIndicators stage, dateValues, dateValid, dateGiven, _
                                   issueIndicator, issueResolution
            JSS_GetRVQTracking stage, dateValues, dateValid, dateGiven, _
                               Date, businessDays, overdueFlag

            Select Case stage
                Case "Approved": approved = approved + 1
                Case "RVQ Available": rvqAvailable = rvqAvailable + 1
                Case "VQ Released": released = released + 1
                Case "Initial VQ": initialVQ = initialVQ + 1
                Case Else: notStarted = notStarted + 1
            End Select
            If overdueFlag = "No RVQ Overdue" Then overdueCount = overdueCount + 1
            If Len(validationText) = 0 Then validationText = "Correct"

            ' Copy each requested field into its fixed output position.
            For i = LBound(mainNames) To UBound(mainNames)
                columnName = CStr(mainNames(i))
                Select Case columnName
                    Case "Ticket Status": rawValue = stage
                    Case "HL Comment", "DM Comment": rawValue = vbNullString
                    Case "Business Days Without RVQ": rawValue = businessDays
                    Case "RVQ Overdue Flag": rawValue = overdueFlag
                    Case "Pending Issue Indicator": rawValue = issueIndicator
                    Case "Issue Resolution Status": rawValue = issueResolution
                    Case "Data Validation": rawValue = validationText
                    Case Else
                        rawValue = JSS_OutputSourceValue(sourceData, srcRow, headers, _
                                                columnName, dateValues, dateValid, dateGiven)
                End Select
                resultData(outRow, i + 1) = JSS_ExcelSafeValue(rawValue)
            Next i
            ' Retain the other JIRA fields in their source order.
            For j = 1 To extraNames.Count
                resultData(outRow, UBound(mainNames) + 1 + j) = JSS_ExcelSafeValue( _
                    JSS_OutputSourceValue(sourceData, srcRow, headers, _
                            CStr(extraNames(j)), dateValues, dateValid, dateGiven))
            Next j
        End If
    Next srcRow

    Application.StatusBar = "Survey status: writing the report..."
    Set resultBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set resultSheet = resultBook.Worksheets(1)
    resultSheet.Name = "Survey Status"
    resultSheet.Range(resultSheet.Cells(OUTPUT_FIRST_DATA_ROW, 1), _
                      resultSheet.Cells(OUTPUT_HEADER_ROW + processed, outCols)).Value2 = resultData

    ' One continuous header row. Summary is column A, so sort is simply A-Z.
    For i = LBound(mainNames) To UBound(mainNames)
        resultSheet.Cells(OUTPUT_HEADER_ROW, i + 1).Value2 = CStr(mainNames(i))
    Next i
    For j = 1 To extraNames.Count
        resultSheet.Cells(OUTPUT_HEADER_ROW, UBound(mainNames) + 1 + j).Value2 = CStr(extraNames(j))
    Next j

    ' HL Comment and DM Comment are dedicated user-managed fields at I and J.
    ' Do not append an extra remarks column or overwrite these on refresh.
    JSS_RecordManagedHeaders resultBook, resultSheet, outCols
    JSS_SortBySummary resultSheet, processed, outCols
    ' Add live ticket links only after sorting so hyperlink addresses stay
    ' associated with the correct Issue Key and its complete ticket row.
    JSS_AddIssueKeyLinks resultSheet, processed
    JSS_StyleReport resultSheet, processed, outCols
    JSS_WriteSummary resultSheet, processed, approved, rvqAvailable, released, _
                     initialVQ, notStarted, overdueCount

    ' Disclose when raw CSV reading was unavailable. A user must verify
    ' ambiguous DD/MM dates because Excel may have parsed them on import.
    If csvFallback Then
        With resultSheet.Range("A6:M6")
            .Merge
            .Value2 = "DATE CHECK REQUIRED: Raw CSV could not be read; verify Excel-imported dates before using this report."
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(185, 28, 28)
            .EntireRow.RowHeight = 19
        End With
    End If

    resultSheet.Activate
    resultSheet.Cells(OUTPUT_FIRST_DATA_ROW, 3).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = OUTPUT_HEADER_ROW
    ActiveWindow.SplitColumn = 2
    ActiveWindow.FreezePanes = True

    ' Finish timing after sorting and formatting, not just after calculations.
    elapsedSeconds = Timer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    resultSheet.Range("B4").Value2 = Round(elapsedSeconds, 2)
    resultSheet.Range("B4").NumberFormat = "0.00 ""seconds"""
    resultReady = True
    usedCsvFallback = csvFallback
    Set JSS_GenerateCore = resultBook
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    If Not quiet And csvFallback Then
        MsgBox "The report was created using the open Excel sheet because the raw CSV could not be reopened." & vbCrLf & _
               "Reason: " & csvReadProblem & vbCrLf & vbCrLf & _
               "IMPORTANT: Verify dates, especially day/month values, before sharing the report." & vbCrLf & _
               "For most reliable results, open a locally downloaded CSV file.", _
               vbExclamation, "JIRA Survey Status - Verify Dates"
    ElseIf Not quiet Then
        MsgBox "Survey Status Report created successfully." & vbCrLf & _
               "Tickets: " & processed & vbCrLf & _
               "No RVQ Overdue: " & overdueCount, vbInformation, "JIRA Survey Status"
    End If
    Exit Function

Failed:
    problem = Err.Description
    On Error Resume Next
    If Not resultReady Then
        If Not resultBook Is Nothing Then resultBook.Close SaveChanges:=False
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0
    If quiet Then
        Err.Raise vbObjectError + 720, "JSS_GenerateCore", problem
    Else
        MsgBox "Unable to generate the survey report:" & vbCrLf & problem, vbExclamation, "JIRA Survey Status"
    End If
End Function

' Find Summary and Issue Key even when an export contains preamble rows.
Private Function JSS_FindHeaderRow(ByRef data As Variant) As Long
    Dim r As Long, c As Long, hasKey As Boolean, hasSummary As Boolean
    For r = 1 To WorksheetFunction.Min(25, UBound(data, 1))
        hasKey = False: hasSummary = False
        For c = 1 To UBound(data, 2)
            If JSS_NormalizeHeader(CStr(data(r, c))) = "issue key" Then hasKey = True
            If JSS_NormalizeHeader(CStr(data(r, c))) = "summary" Then hasSummary = True
        Next c
        If hasKey And hasSummary Then JSS_FindHeaderRow = r: Exit Function
    Next r
End Function

' Standardize JIRA's Custom field (Name) wrapper and whitespace.
Private Function JSS_NormalizeHeader(ByVal name As String) As String
    Dim s As String
    s = LCase$(Trim$(Replace(Replace(name, vbCr, " "), vbLf, " ")))
    If Left$(s, 14) = "custom field (" And Right$(s, 1) = ")" Then
        s = Mid$(s, 15, Len(s) - 15)
    End If
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    JSS_NormalizeHeader = s
End Function


' Remove JIRA's display-only Custom field (...) wrapper in the OUTPUT.
' Do not change the raw source header or the normalized source lookup key.
Private Function JSS_DisplayHeader(ByVal headerName As String) As String
    Dim caption As String
    caption = Trim$(headerName)
    If LCase$(Left$(caption, 14)) = "custom field (" And Right$(caption, 1) = ")" Then
        caption = Mid$(caption, 15, Len(caption) - 15)
    End If
    JSS_DisplayHeader = caption
End Function

Private Function JSS_CreateHeaders(ByRef data As Variant, ByVal headerRow As Long, _
                                   ByVal countCols As Long) As Object
    Dim result As Object, c As Long, k As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For c = 1 To countCols
        k = JSS_NormalizeHeader(CStr(data(headerRow, c)))
        If Len(k) > 0 Then
            If result.Exists(k) Then
                Err.Raise vbObjectError + 707, , "Duplicate header: " & CStr(data(headerRow, c))
            End If
            result.Add k, c
        End If
    Next c
    Set JSS_CreateHeaders = result
End Function

Private Sub JSS_ValidateRequiredColumns(ByVal headers As Object)
    Dim required As Variant, x As Variant, missing As String
    required = Array("Summary", "Issue Key", "Initial VQ Complete Date", _
        "Initial VQ Review Complete Date", "VQ Released Date", _
        "Returned VQ Received Date", "Validation Review Complete Date", _
        "VQ Changes Complete Date", "VQ Approvals Complete Date", _
        "Pending Issue Date", "Issue Resolution Date")
    For Each x In required
        If Not headers.Exists(JSS_NormalizeHeader(CStr(x))) Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & CStr(x)
        End If
    Next x
    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 708, , "Missing required JIRA fields: " & missing
    End If
End Sub

Private Function JSS_FieldValue(ByRef data As Variant, ByVal r As Long, _
                                ByVal headers As Object, ByVal name As String) As Variant
    Dim canonical As String
    canonical = JSS_NormalizeHeader(name)
    If headers.Exists(canonical) Then
        JSS_FieldValue = data(r, CLng(headers(canonical)))
    Else
        JSS_FieldValue = vbNullString
    End If
End Function

Private Function JSS_IsNonEmptyRow(ByRef data As Variant, ByVal r As Long, _
                                   ByVal ncols As Long) As Boolean
    Dim c As Long
    For c = 1 To ncols
        If Not IsError(data(r, c)) Then
            If Len(Trim$(CStr(data(r, c)))) > 0 Then
                JSS_IsNonEmptyRow = True
                Exit Function
            End If
        End If
    Next c
End Function

Private Function JSS_IsDateField(ByVal name As String) As Boolean
    Select Case JSS_NormalizeHeader(name)
        Case "created", "date received", "initial vq complete date", _
             "initial vq review complete date", "vq released date", _
             "returned vq received date", "validation review complete date", _
             "vq changes complete date", "vq approvals complete date", _
             "pending issue date", "issue resolution date"
            JSS_IsDateField = True
    End Select
End Function

' Date parsing: source CSV is DD/MM/YYYY HH:MM. Use DateSerial with explicit
' components rather than locale-dependent CDate or Excel text coercion.
Private Function JSS_ParseDate(ByVal inputValue As Variant, ByVal is1904 As Boolean, _
                               ByRef isProvided As Boolean, ByRef isValid As Boolean) As Date
    Dim textDate As String, bits As Variant, firstSpace As Long
    Dim d As Long, m As Long, y As Long, parsed As Date
    isProvided = False: isValid = False
    If IsError(inputValue) Or IsNull(inputValue) Or IsEmpty(inputValue) Then Exit Function
    textDate = Trim$(CStr(inputValue))
    If Len(textDate) = 0 Then Exit Function
    isProvided = True
    On Error GoTo Invalid
    If IsNumeric(inputValue) And VarType(inputValue) <> vbString Then
        ' Value2 returns real Excel dates as numeric serial values.
        parsed = DateSerial(1899, 12, 30) + CDbl(inputValue)
        If is1904 Then parsed = parsed + 1462
    ElseIf VarType(inputValue) = vbDate Then
        parsed = CDate(inputValue)
    Else
        firstSpace = InStr(textDate, " ")
        If firstSpace > 0 Then textDate = Left$(textDate, firstSpace - 1)
        bits = Split(textDate, "/")
        If UBound(bits) <> 2 Then GoTo Invalid
        If Not IsNumeric(bits(0)) Or Not IsNumeric(bits(1)) Or Not IsNumeric(bits(2)) Then GoTo Invalid
        d = CLng(bits(0)): m = CLng(bits(1)): y = CLng(bits(2))
        If y < 1900 Or y > 9999 Then GoTo Invalid
        parsed = DateSerial(y, m, d)
        ' DateSerial otherwise normalizes impossible days (e.g. 31 February).
        If Year(parsed) <> y Or Month(parsed) <> m Or Day(parsed) <> d Then GoTo Invalid
    End If
    If parsed < DateSerial(1900, 1, 1) Or parsed > DateSerial(9999, 12, 31) Then GoTo Invalid
    JSS_ParseDate = DateValue(parsed)
    isValid = True
    Exit Function
Invalid:
    isValid = False
End Function

Private Sub JSS_CollectDates(ByRef data As Variant, ByVal r As Long, _
    ByVal headers As Object, ByVal is1904 As Boolean, _
    ByVal dateValues As Object, ByVal dateValid As Object, _
    ByVal dateGiven As Object, ByRef errors As String)
    Dim labels As Variant, x As Variant, name As String, value As Date
    Dim isProvided As Boolean, valid As Boolean
    labels = Array("Created", "Date Received", "Initial VQ Complete Date", _
        "Initial VQ Review Complete Date", "VQ Released Date", _
        "Returned VQ Received Date", "Validation Review Complete Date", _
        "VQ Changes Complete Date", "VQ Approvals Complete Date", _
        "Pending Issue Date", "Issue Resolution Date")
    For Each x In labels
        name = CStr(x)
        value = JSS_ParseDate(JSS_FieldValue(data, r, headers, name), is1904, isProvided, valid)
        dateValues.Add name, value
        dateGiven.Add name, isProvided
        dateValid.Add name, valid
        If isProvided And Not valid Then JSS_AddProblem errors, "Invalid date: " & name
        If valid And value > Date Then JSS_AddProblem errors, "Future date: " & name
    Next x
End Sub

Private Sub JSS_AddProblem(ByRef errors As String, ByVal problem As String)
    If Len(errors) > 0 Then errors = errors & "; "
    errors = errors & problem
End Sub

Private Function JSS_GetSurveyStage(ByVal dateValid As Object, _
                                    ByVal dateGiven As Object) As String
    If dateValid("Validation Review Complete Date") And _
       dateValid("VQ Changes Complete Date") And _
       dateValid("VQ Approvals Complete Date") Then
        JSS_GetSurveyStage = "Approved"
    ElseIf dateValid("Returned VQ Received Date") Then
        JSS_GetSurveyStage = "RVQ Available"
    ElseIf dateValid("Initial VQ Review Complete Date") And _
           dateValid("VQ Released Date") Then
        JSS_GetSurveyStage = "VQ Released"
    ElseIf dateValid("Initial VQ Complete Date") Then
        JSS_GetSurveyStage = "Initial VQ"
    Else
        JSS_GetSurveyStage = "Not yet started"
    End If
End Function

' Report data-quality contradictions separately; they do not redefine the
' basic processing stage. Normal blank future-stage dates are NOT errors.
Private Sub JSS_CheckChronology(ByVal dates As Object, ByVal valid As Object, _
                                ByVal given As Object, ByRef errors As String)
    If valid("Returned VQ Received Date") And valid("VQ Released Date") Then
        If dates("Returned VQ Received Date") < dates("VQ Released Date") Then
            JSS_AddProblem errors, "RVQ received before VQ release"
        End If
    End If
    If valid("Initial VQ Review Complete Date") And valid("Initial VQ Complete Date") Then
        If dates("Initial VQ Review Complete Date") < dates("Initial VQ Complete Date") Then
            JSS_AddProblem errors, "Initial VQ review before Initial VQ complete"
        End If
    End If
    If valid("VQ Released Date") And valid("Initial VQ Review Complete Date") Then
        If dates("VQ Released Date") < dates("Initial VQ Review Complete Date") Then
            JSS_AddProblem errors, "VQ release before initial review"
        End If
    End If
    If valid("Issue Resolution Date") Then
        If Not valid("Pending Issue Date") Then
            JSS_AddProblem errors, "Resolution date without a pending issue date"
        ElseIf dates("Issue Resolution Date") < dates("Pending Issue Date") Then
            JSS_AddProblem errors, "Issue resolved before pending issue date"
        End If
    End If
    If valid("Initial VQ Review Complete Date") And _
       Not given("Initial VQ Complete Date") Then
        JSS_AddProblem errors, "Missing Initial VQ Complete Date"
    End If
    If valid("VQ Released Date") And Not given("Initial VQ Review Complete Date") Then
        JSS_AddProblem errors, "Missing Initial VQ Review Complete Date"
    End If
    If valid("Returned VQ Received Date") And Not given("VQ Released Date") Then
        JSS_AddProblem errors, "Missing VQ Released Date"
    End If
    If valid("VQ Approvals Complete Date") And _
       Not valid("Validation Review Complete Date") Then
        JSS_AddProblem errors, "Approval date without Validation Review Complete Date"
    End If
    If valid("VQ Approvals Complete Date") And _
       Not valid("VQ Changes Complete Date") Then
        JSS_AddProblem errors, "Approval date without VQ Changes Complete Date"
    End If
End Sub

' Pending issue information is historical; it must not override stages.
' Without an Issue Resolution Date, the honest classification is Not Recorded.
Private Sub JSS_GetIssueIndicators(ByVal stage As String, _
    ByVal dates As Object, ByVal valid As Object, ByVal given As Object, _
    ByRef pendingIndicator As String, ByRef resolution As String)
    If stage = "VQ Released" Or stage = "Approved" Then
        pendingIndicator = "N/A"
        resolution = "N/A"
        Exit Sub
    End If
    If given("Pending Issue Date") Then
        pendingIndicator = "Yes"
        If Not valid("Pending Issue Date") Then
            resolution = "Check Dates"
        ElseIf given("Issue Resolution Date") Then
            If Not valid("Issue Resolution Date") Then
                resolution = "Check Dates"
            ElseIf dates("Issue Resolution Date") < dates("Pending Issue Date") Then
                resolution = "Check Dates"
            Else
                resolution = "Resolved"
            End If
        Else
            resolution = "Not Recorded"
        End If
    Else
        pendingIndicator = "No"
        If given("Issue Resolution Date") Then
            resolution = "Check Dates"
        Else
            resolution = "No Issue Recorded"
        End If
    End If
End Sub

Private Sub JSS_GetRVQTracking(ByVal stage As String, _
    ByVal dates As Object, ByVal valid As Object, ByVal given As Object, _
    ByVal reportDate As Date, ByRef businessDays As Variant, _
    ByRef overdueFlag As String)
    Dim n As Long
    businessDays = vbNullString: overdueFlag = vbNullString
    If stage <> "VQ Released" Then Exit Sub
    If Not valid("VQ Released Date") Or given("Returned VQ Received Date") Then Exit Sub
    If dates("VQ Released Date") > reportDate Then Exit Sub

    ' NetworkDays is inclusive. Start at VQ release + 1 to exclude release day.
    ' An empty/negative date interval is handled without counting any days.
    If reportDate <= dates("VQ Released Date") Then
        n = 0
    Else
        n = Application.WorksheetFunction.NetworkDays( _
                dates("VQ Released Date") + 1, reportDate)
        If n < 0 Then n = 0
    End If
    businessDays = n
    If n > OVERDUE_THRESHOLD Then
        overdueFlag = "No RVQ Overdue"
    Else
        overdueFlag = "Within 14 Business Days"
    End If
End Sub

Private Function JSS_OutputSourceValue(ByRef data As Variant, ByVal rowIndex As Long, _
    ByVal headers As Object, ByVal name As String, ByVal dates As Object, _
    ByVal valid As Object, ByVal given As Object) As Variant
    Dim value As Variant, dateKey As String, candidate As Variant
    value = JSS_FieldValue(data, rowIndex, headers, name)
    If JSS_IsDateField(name) Then
        ' Extra columns may contain JIRA's "Custom field (...)" prefix.
        ' Match to the canonical date keys used by the validation engine.
        dateKey = name
        For Each candidate In dates.Keys
            If JSS_NormalizeHeader(CStr(candidate)) = JSS_NormalizeHeader(name) Then
                dateKey = CStr(candidate)
                Exit For
            End If
        Next candidate
        If valid(dateKey) Then
            ' Excel numeric serial, which can be formatted and sorted as a date.
            JSS_OutputSourceValue = CDbl(dates(dateKey))
        ElseIf given(dateKey) Then
            ' Leave unparseable source text visible for investigation.
            JSS_OutputSourceValue = value
        Else
            JSS_OutputSourceValue = vbNullString
        End If
    ElseIf JSS_NormalizeHeader(name) = "pending issue remarks" Or _
           JSS_NormalizeHeader(name) = "issue resolution remarks" Then
        If Len(Trim$(CStr(value))) = 0 Then
            JSS_OutputSourceValue = "Not Provided"
        Else
            JSS_OutputSourceValue = value
        End If
    Else
        JSS_OutputSourceValue = value
    End If
End Function

' Treat a leading equals sign as text when copying a CSV value into Excel.
' Prevents an exported client remark from becoming an executable formula.
Private Function JSS_ExcelSafeValue(ByVal inputValue As Variant) As Variant
    If VarType(inputValue) = vbString Then
        If Left$(CStr(inputValue), 1) = "=" Then
            JSS_ExcelSafeValue = "'" & CStr(inputValue)
            Exit Function
        End If
    End If
    JSS_ExcelSafeValue = inputValue
End Function

' Convert each valid Issue Key into a clickable JIRA link in column B.
' Visible text: OTS3-552751
' Link target:  https://jira.ehr.com/browse/OTS3-552751
' This changes only the new results workbook, never the original export.
Private Sub JSS_AddIssueKeyLinks(ByVal ws As Worksheet, ByVal rowsCount As Long)
    Dim r As Long, ticketKey As String, linkCell As Range
    Dim pattern As Object

    ' Require Jira-style project key and numeric ticket ID. Malformed values
    ' remain plain text and are still available for Data Validation.
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[A-Za-z0-9_]+-[0-9]+$"
    pattern.IgnoreCase = True
    pattern.Global = False

    For r = OUTPUT_FIRST_DATA_ROW To OUTPUT_HEADER_ROW + rowsCount
        Set linkCell = ws.Cells(r, 2)  ' Column B: Issue Key.
        ticketKey = Trim$(CStr(linkCell.Value2))
        If Len(ticketKey) > 0 And pattern.Test(ticketKey) Then
            If linkCell.Hyperlinks.Count > 0 Then linkCell.Hyperlinks.Delete
            ws.Hyperlinks.Add Anchor:=linkCell, _
                Address:=JIRA_ISSUE_BASE_URL & ticketKey, _
                TextToDisplay:=ticketKey, _
                ScreenTip:="Open JIRA ticket " & ticketKey

            ' Use only font styling, no additional cell background fill.
            linkCell.Font.Color = RGB(25, 89, 176)
            linkCell.Font.Underline = xlUnderlineStyleSingle
        End If
    Next r
End Sub

Private Sub JSS_SortBySummary(ByVal ws As Worksheet, _
                               ByVal rowsCount As Long, ByVal colsCount As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(OUTPUT_FIRST_DATA_ROW, 1), _
                                     ws.Cells(OUTPUT_HEADER_ROW + rowsCount, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(OUTPUT_HEADER_ROW, 1), _
                           ws.Cells(OUTPUT_HEADER_ROW + rowsCount, colsCount))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub JSS_WriteSummary(ByVal ws As Worksheet, _
    ByVal total As Long, ByVal approved As Long, ByVal rvq As Long, _
    ByVal released As Long, ByVal initial As Long, ByVal notStarted As Long, _
    ByVal overdue As Long)
    With ws
        .Range("A1").Value2 = "JIRA SURVEY STATUS REPORT"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(24, 49, 83)
        .Range("A2").Value2 = "Tickets Processed"
        .Range("B2").Value2 = total
        .Range("D2").Value2 = "Approved"
        .Range("E2").Value2 = approved
        .Range("G2").Value2 = "RVQ Available"
        .Range("H2").Value2 = rvq
        .Range("J2").Value2 = "VQ Released"
        .Range("K2").Value2 = released
        .Range("A3").Value2 = "Initial VQ"
        .Range("B3").Value2 = initial
        .Range("D3").Value2 = "Not yet started"
        .Range("E3").Value2 = notStarted
        .Range("G3").Value2 = "No RVQ Overdue"
        .Range("H3").Value2 = overdue
        .Range("J3").Value2 = "Generated"
        .Range("K3").Value = Now
        .Range("K3").NumberFormat = "mmm dd, yyyy hh:mm AM/PM"
        .Range("A4").Value2 = "Runtime"
        .Range("B4").Value2 = "Calculating..."
        .Range("A2:K4").Font.Name = "Segoe UI"
        .Range("A2:K4").Font.Size = 10
        .Range("A2:A4,D2:D3,G2:G3,J2:J3").Font.Bold = True
        .Range("B2:B4,E2:E3,H2:H3,K2:K3").HorizontalAlignment = xlLeft
        .Range("H3").Font.Color = RGB(185, 28, 28)
        .Range("B2:B3,E2:E3,H2:H3,K2").NumberFormat = "0"
        .Range("A5").Value2 = "Developed by Perry Dave V. Geron"
        .Range("A5").Font.Color = RGB(100, 116, 139)
        .Range("A5").Font.Italic = True
        .Range("A5").Font.Size = 9
    End With
End Sub

Private Sub JSS_AddFontCondition(ByVal target As Range, _
                                 ByVal compareValue As String, ByVal colorValue As Long)
    Dim rule As FormatCondition
    Set rule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
                                           Formula1:="=" & Chr$(34) & compareValue & Chr$(34))
    rule.Font.Color = colorValue
    rule.Font.Bold = True
End Sub

Private Sub JSS_StyleReport(ByVal ws As Worksheet, _
                             ByVal rowsCount As Long, ByVal colsCount As Long)
    Dim lastRow As Long, i As Long, h As String
    Dim fontRed As Long, fontGreen As Long
    Dim headersForStyle As Object, statusCol As String, overdueCol As String
    Dim resolutionCol As String, validationCol As String, daysCol As String
    lastRow = OUTPUT_HEADER_ROW + rowsCount
    fontRed = RGB(185, 28, 28)
    fontGreen = RGB(22, 101, 52)
    With ws.Range(ws.Cells(OUTPUT_HEADER_ROW, 1), ws.Cells(OUTPUT_HEADER_ROW, colsCount))
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(24, 49, 83)
        .WrapText = True
        .RowHeight = 34
    End With
    With ws.Range(ws.Cells(OUTPUT_FIRST_DATA_ROW, 1), ws.Cells(lastRow, colsCount))
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .VerticalAlignment = xlTop
    End With
    For i = 1 To colsCount
        h = CStr(ws.Cells(OUTPUT_HEADER_ROW, i).Value2)
        If JSS_IsDateField(h) Then
            ws.Range(ws.Cells(OUTPUT_FIRST_DATA_ROW, i), _
                     ws.Cells(lastRow, i)).NumberFormat = "mmm dd"
            ws.Columns(i).ColumnWidth = 18
        ElseIf InStr(1, LCase$(h), "remarks") > 0 Or _
               h = "HL Comment" Or h = "DM Comment" Then
            ws.Columns(i).ColumnWidth = 43
            ws.Range(ws.Cells(OUTPUT_FIRST_DATA_ROW, i), _
                     ws.Cells(lastRow, i)).WrapText = True
        ElseIf h = "Summary" Or h = "Company Name" Then
            ws.Columns(i).ColumnWidth = 38
        ElseIf h = "Data Validation" Then
            ws.Columns(i).ColumnWidth = 46
        ElseIf h = "Ticket Status" Then
            ws.Columns(i).ColumnWidth = 25
        Else
            ws.Columns(i).ColumnWidth = 20
        End If
    Next i
    Set headersForStyle = JSS_ReportHeaders(ws)
    statusCol = JSS_ExcelColumn(CLng(headersForStyle("ticket status")))
    overdueCol = JSS_ExcelColumn(CLng(headersForStyle("rvq overdue flag")))
    resolutionCol = JSS_ExcelColumn(CLng(headersForStyle("issue resolution status")))
    validationCol = JSS_ExcelColumn(CLng(headersForStyle("data validation")))
    daysCol = JSS_ExcelColumn(CLng(headersForStyle("business days without rvq")))
    ws.Range(daysCol & OUTPUT_FIRST_DATA_ROW & ":" & daysCol & lastRow).NumberFormat = "0"
    ws.Range("A" & OUTPUT_HEADER_ROW & ":" & JSS_ExcelColumn(colsCount) & lastRow).AutoFilter
    ' Font colors only for status results; the underlying values are unchanged.
    JSS_AddFontCondition ws.Range(statusCol & OUTPUT_FIRST_DATA_ROW & ":" & statusCol & lastRow), "Approved", fontGreen
    JSS_AddFontCondition ws.Range(overdueCol & OUTPUT_FIRST_DATA_ROW & ":" & overdueCol & lastRow), "No RVQ Overdue", fontRed
    JSS_AddFontCondition ws.Range(resolutionCol & OUTPUT_FIRST_DATA_ROW & ":" & resolutionCol & lastRow), "Check Dates", fontRed
    JSS_AddFontCondition ws.Range(resolutionCol & OUTPUT_FIRST_DATA_ROW & ":" & resolutionCol & lastRow), "Resolved", fontGreen
    JSS_AddFontCondition ws.Range(validationCol & OUTPUT_FIRST_DATA_ROW & ":" & validationCol & lastRow), "Correct", fontGreen
    ws.Range(validationCol & OUTPUT_FIRST_DATA_ROW & ":" & validationCol & lastRow).FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=AND($" & validationCol & OUTPUT_FIRST_DATA_ROW & _
        "<>""Correct"",$" & validationCol & OUTPUT_FIRST_DATA_ROW & "<>"""")").Font.Color = fontRed
End Sub

Private Function JSS_ExcelColumn(ByVal columnNumber As Long) As String
    Dim n As Long
    n = columnNumber
    Do While n > 0
        JSS_ExcelColumn = Chr$(65 + ((n - 1) Mod 26)) & JSS_ExcelColumn
        n = (n - 1) \ 26
    Loop
End Function

' Read a normal Excel workbook or worksheet into one efficient 2-D array.
Private Function JSS_ReadSheetAsArray(ByVal ws As Worksheet) As Variant
    Dim lastCell As Range, r As Long, c As Long, block As Variant
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then Err.Raise vbObjectError + 710, , "The active sheet is empty."
    r = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    c = lastCell.Column
    If r < 2 Or c < 2 Then Err.Raise vbObjectError + 711, , "This worksheet does not contain a JIRA ticket table."
    JSS_ReadSheetAsArray = ws.Range(ws.Cells(1, 1), ws.Cells(r, c)).Value2
End Function

' Read original CSV as UTF-8 text so source DD/MM/YYYY dates remain intact.
' ADO LoadFromFile sometimes refuses an Excel-opened CSV because of file
' sharing restrictions. First read the original bytes in SHARED/READ-ONLY
' mode, then use ADO only to decode the bytes in memory (no file reopening).
' Fall back to the original ADO method if shared binary file reading fails.
Private Function JSS_ReadCsvAsArray(ByVal fileName As String) As Variant
    Dim fileNo As Integer, byteCount As Long
    Dim fileBytes() As Byte, stream As Object, raw As String
    Dim sharedReadProblem As String

    On Error GoTo TryADOFileRead
    fileNo = FreeFile
    Open fileName For Binary Access Read Shared As #fileNo
    byteCount = LOF(fileNo)
    If byteCount < 1 Then Err.Raise vbObjectError + 715, , "CSV file is empty."
    ReDim fileBytes(0 To byteCount - 1)
    Get #fileNo, 1, fileBytes
    Close #fileNo
    fileNo = 0

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1  ' Binary stream: the file has already been read.
    stream.Open
    stream.Write fileBytes
    stream.Position = 0
    stream.Type = 2  ' Convert the in-memory bytes to UTF-8 text.
    stream.Charset = "utf-8"
    raw = stream.ReadText(-1)
    stream.Close
    GoTo ParseContents

TryADOFileRead:
    sharedReadProblem = Err.Description
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    Err.Clear
    On Error GoTo ADOReadFailed

    ' Second method: try opening the local CSV through ADODB directly.
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile fileName
    raw = stream.ReadText(-1)
    stream.Close

ParseContents:
    ' UTF-8 BOM, when present, is not part of the first CSV header.
    If Len(raw) > 0 Then
        If AscW(Left$(raw, 1)) = &HFEFF Then raw = Mid$(raw, 2)
    End If
    JSS_ReadCsvAsArray = JSS_ParseCsv(raw)
    Exit Function

ADOReadFailed:
    Err.Raise vbObjectError + 716, , _
        "Cannot read raw CSV (possibly cloud-only or locked). " & _
        "Shared read: " & sharedReadProblem & "; ADO: " & Err.Description
End Function

' RFC-style CSV parser: honors commas, escaped double quotes and line breaks
' inside quoted JIRA remarks. Returns a 1-based 2-D Variant array.
Private Function JSS_ParseCsv(ByVal contents As String) As Variant
    Dim allRows As Collection, rowFields As Collection
    Dim i As Long, n As Long, maxCols As Long, rowNum As Long, colNum As Long
    Dim fieldText As String, ch As String, inQuotes As Boolean, rowArray As Variant
    Dim result() As Variant
    Set allRows = New Collection
    Set rowFields = New Collection
    fieldText = "": n = Len(contents)
    For i = 1 To n
        ch = Mid$(contents, i, 1)
        If ch = """" Then
            If inQuotes And i < n And Mid$(contents, i + 1, 1) = """" Then
                fieldText = fieldText & """": i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            rowFields.Add fieldText: fieldText = ""
        ElseIf (ch = vbLf Or ch = vbCr) And Not inQuotes Then
            rowFields.Add fieldText: fieldText = ""
            If rowFields.Count > maxCols Then maxCols = rowFields.Count
            allRows.Add JSS_CollectionToArray(rowFields)
            Set rowFields = New Collection
            If ch = vbCr And i < n Then
                If Mid$(contents, i + 1, 1) = vbLf Then i = i + 1
            End If
        Else
            fieldText = fieldText & ch
        End If
    Next i
    If inQuotes Then Err.Raise vbObjectError + 712, , "CSV contains an unclosed quoted field."
    If Len(fieldText) > 0 Or rowFields.Count > 0 Then
        rowFields.Add fieldText
        If rowFields.Count > maxCols Then maxCols = rowFields.Count
        allRows.Add JSS_CollectionToArray(rowFields)
    End If
    If allRows.Count < 2 Or maxCols < 2 Then Err.Raise vbObjectError + 713, , "CSV contains no recognizable JIRA ticket data."
    ReDim result(1 To allRows.Count, 1 To maxCols)
    For rowNum = 1 To allRows.Count
        rowArray = allRows(rowNum)
        For colNum = LBound(rowArray) To UBound(rowArray)
            result(rowNum, colNum) = rowArray(colNum)
        Next colNum
    Next rowNum
    JSS_ParseCsv = result
End Function

Private Function JSS_CollectionToArray(ByVal fields As Collection) As Variant
    Dim result() As Variant, i As Long
    ReDim result(1 To fields.Count)
    For i = 1 To fields.Count
        result(i) = fields(i)
    Next i
    JSS_CollectionToArray = result
End Function

'===============================================================================
' VERSION 2.1 - INCREMENTAL UPDATE (UPDATE / APPEND WITHOUT LOSING NOTES)
'===============================================================================

' Persist an explicit ownership list when a NEW report is generated. This makes
' it possible to distinguish fields owned by the add-in from extra columns
' maintained by stakeholders. A VeryHidden worksheet does not clutter output.
Private Sub JSS_RecordManagedHeaders(ByVal reportBook As Workbook, _
    ByVal reportSheet As Worksheet, ByVal columnCount As Long)
    Dim meta As Worksheet, c As Long, n As Long, item As String
    Set meta = reportBook.Worksheets.Add(After:=reportBook.Worksheets(reportBook.Worksheets.Count))
    meta.Name = JSS_METADATA_SHEET
    meta.Cells(1, 1).Value2 = "JSS Managed Column Headers - Version 2.1"
    n = 1
    For c = 1 To columnCount
        item = CStr(reportSheet.Cells(OUTPUT_HEADER_ROW, c).Value2)
        If Not JSS_IsManualHeader(JSS_NormalizeHeader(item)) Then
            n = n + 1
            meta.Cells(n, 1).Value2 = item
        End If
    Next c
    meta.Visible = xlSheetVeryHidden
End Sub

' Read managed-field registry. Older v1 reports do not have the registry;
' safely migrate fixed fields and extras that match the incoming JIRA output.
' Unknown stakeholder columns remain USER-MANAGED and are NEVER overwritten.
Private Function JSS_LoadManagedHeaders(ByVal reportBook As Workbook, _
    ByVal reportSheet As Worksheet, ByVal latestHeaders As Object) As Object
    Dim managed As Object, meta As Worksheet, c As Long, lastRow As Long
    Dim name As String, key As String, fixedNames As Variant, item As Variant
    Set managed = CreateObject("Scripting.Dictionary")
    managed.CompareMode = vbTextCompare
    On Error Resume Next
    Set meta = reportBook.Worksheets(JSS_METADATA_SHEET)
    On Error GoTo 0
    If Not meta Is Nothing Then
        lastRow = meta.Cells(meta.Rows.Count, 1).End(xlUp).Row
        For c = 2 To lastRow
            key = JSS_NormalizeHeader(CStr(meta.Cells(c, 1).Value2))
            If Len(key) > 0 Then managed(key) = True
        Next c
    Else
        ' v1 did not record ownership. Fixed fields are always add-in-managed.
        ' Never assume any arbitrary custom column is safe to overwrite.
        fixedNames = Array("Summary", "Issue Key", "Company Name", _
            "Abbreviation", "SPM Code", "Submission Type", _
            "Submission Category", "Ticket Status", "Survey Status", "Created", _
            "Initial VQ Complete Date", "Initial VQ Review Complete Date", _
            "VQ Released Date", "Returned VQ Received Date", _
            "Business Days Without RVQ", "RVQ Overdue Flag", _
            "Pending Issue Date", "Pending Issue Indicator", _
            "Pending Issue Remarks", "Issue Resolution Date", _
            "Issue Resolution Status", "Issue Resolution Remarks", _
            "Validation Review Complete Date", "VQ Changes Complete Date", _
            "VQ Approvals Complete Date", "Data Validation")
        For Each item In fixedNames
            managed(JSS_NormalizeHeader(CStr(item))) = True
        Next item
        ' Optional JIRA extras originally following column Y may be treated
        ' as source-owned when the latest export also contains this header.
        ' Known stakeholder-note headers are always excluded.
        For c = 26 To JSS_LastHeaderCol(reportSheet)
            name = CStr(reportSheet.Cells(OUTPUT_HEADER_ROW, c).Value2)
            key = JSS_NormalizeHeader(name)
            If latestHeaders.Exists(key) And Not JSS_IsManualHeader(key) Then
                managed(key) = True
            End If
        Next c
    End If
    ' Do not let a legacy registry claim a stakeholder-owned field.
    For Each item In Array("HL Comment", "DM Comment", "Stakeholder Remarks", "Stakeholder Notes", _
                           "Follow-up Notes", "Action Required", "Manual Remarks", _
                           "Manual Notes", "Internal Notes", "Action Owner")
        key = JSS_NormalizeHeader(CStr(item))
        If managed.Exists(key) Then managed.Remove key
    Next item
    managed(JSS_NormalizeHeader(JSS_MATCH_HEADER)) = True
    Set JSS_LoadManagedHeaders = managed
End Function

' Common note labels must remain user-maintained, including when the previous
' year's JIRA export happens to contain a similarly named field.
Private Function JSS_IsManualHeader(ByVal normalized As String) As Boolean
    Select Case normalized
        Case "hl comment", "dm comment", "stakeholder remarks", "stakeholder notes", "follow-up notes", _
             "follow up notes", "action required", "manual remarks", _
             "manual notes", "internal notes", "action owner"
            JSS_IsManualHeader = True
    End Select
End Function

Private Function JSS_LastHeaderCol(ByVal ws As Worksheet) As Long
    JSS_LastHeaderCol = ws.Cells(OUTPUT_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

' Read the 7th-row headings by canonical name. Duplicates are unsafe because
' the updater otherwise cannot identify which column belongs to each field.
Private Function JSS_ReportHeaders(ByVal ws As Worksheet) As Object
    Dim found As Object, c As Long, key As String
    Set found = CreateObject("Scripting.Dictionary")
    found.CompareMode = vbTextCompare
    For c = 1 To JSS_LastHeaderCol(ws)
        key = JSS_NormalizeHeader(CStr(ws.Cells(OUTPUT_HEADER_ROW, c).Value2))
        If Len(key) > 0 Then
            If found.Exists(key) Then
                Err.Raise vbObjectError + 740, , "Duplicate report header: " & _
                    CStr(ws.Cells(OUTPUT_HEADER_ROW, c).Value2)
            End If
            found.Add key, c
        End If
    Next c
    Set JSS_ReportHeaders = found
End Function

' Maps Issue Keys to physical rows. Do not allow blanks or duplicate keys in
' either workbook: ambiguous updates could put notes on the wrong ticket.
Private Function JSS_IndexTickets(ByVal ws As Worksheet, _
    ByVal keyCol As Long) As Object
    Dim index As Object, r As Long, lastRow As Long, key As String
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < OUTPUT_FIRST_DATA_ROW Then
        Err.Raise vbObjectError + 741, , "No ticket records were found."
    End If
    For r = OUTPUT_FIRST_DATA_ROW To lastRow
        key = Trim$(CStr(ws.Cells(r, keyCol).Value2))
        If Len(key) = 0 Then
            Err.Raise vbObjectError + 742, , "Blank Issue Key at row " & r & ". No changes were made."
        End If
        If index.Exists(key) Then
            Err.Raise vbObjectError + 743, , "Duplicate Issue Key: " & key & _
                ". Resolve duplicates before updating."
        End If
        index.Add key, r
    Next r
    Set JSS_IndexTickets = index
End Function

' Value comparison must handle blank cells and Excel serial-number dates.
Private Function JSS_ValuesEqual(ByVal oldValue As Variant, _
    ByVal newValue As Variant) As Boolean
    If IsError(oldValue) Or IsError(newValue) Then Exit Function
    If (IsEmpty(oldValue) Or IsNull(oldValue)) Then oldValue = vbNullString
    If (IsEmpty(newValue) Or IsNull(newValue)) Then newValue = vbNullString
    If IsNumeric(oldValue) And IsNumeric(newValue) And _
       VarType(oldValue) <> vbString And VarType(newValue) <> vbString Then
        JSS_ValuesEqual = (Abs(CDbl(oldValue) - CDbl(newValue)) < 0.000001)
    Else
        JSS_ValuesEqual = (StrComp(CStr(oldValue), CStr(newValue), vbBinaryCompare) = 0)
    End If
End Function

Private Sub JSS_RequireExistingReport(ByVal reportBook As Workbook, _
    ByVal reportSheet As Worksheet)
    Dim existingHeaders As Object, name As Variant
    If reportBook Is Nothing Then Err.Raise vbObjectError + 744, , "Open an existing Survey Status report."
    If reportBook.IsAddin Then Err.Raise vbObjectError + 745, , "Select your Survey Status workbook."
    If Len(reportBook.Path) = 0 Then
        Err.Raise vbObjectError + 746, , "Save the existing report as XLSX before updating."
    End If
    If reportBook.ReadOnly Then
        Err.Raise vbObjectError + 747, , "The report is read-only. Save an editable copy first."
    End If
    If InStr(1, CStr(reportSheet.Range("A1").Value2), _
             "JIRA SURVEY STATUS REPORT", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 748, , "The selected workbook is not a generated JIRA Survey Status Report."
    End If
    Set existingHeaders = JSS_ReportHeaders(reportSheet)
    For Each name In Array("Summary", "Issue Key", _
                           "VQ Released Date", "RVQ Overdue Flag", "Data Validation")
        If Not existingHeaders.Exists(JSS_NormalizeHeader(CStr(name))) Then
            Err.Raise vbObjectError + 749, , "Existing report is missing: " & CStr(name)
        End If
    Next name
    If Not existingHeaders.Exists("ticket status") And _
       Not existingHeaders.Exists("survey status") Then
        Err.Raise vbObjectError + 749, , "Existing report is missing Ticket Status / Survey Status."
    End If
End Sub

' Extend the metadata registry only when updates have been validated and a
' backup already exists. This affects metadata only, never stakeholder notes.
Private Sub JSS_SaveManagedHeaders(ByVal reportBook As Workbook, _
    ByVal managed As Object)
    Dim meta As Worksheet, field As Variant, i As Long
    On Error Resume Next
    Set meta = reportBook.Worksheets(JSS_METADATA_SHEET)
    On Error GoTo 0
    If meta Is Nothing Then
        Set meta = reportBook.Worksheets.Add(After:=reportBook.Worksheets(reportBook.Worksheets.Count))
        meta.Name = JSS_METADATA_SHEET
    End If
    meta.Cells.ClearContents
    meta.Cells(1, 1).Value2 = "JSS Managed Column Headers - Version 2.1"
    i = 1
    For Each field In managed.Keys
        i = i + 1
        meta.Cells(i, 1).Value2 = CStr(field)
    Next field
    meta.Visible = xlSheetVeryHidden
End Sub

Private Function JSS_BackupPath(ByVal reportBook As Workbook) As String
    Dim fullName As String, dotAt As Long, stem As String, suffix As String
    fullName = reportBook.FullName
    dotAt = InStrRev(fullName, ".")
    If dotAt = 0 Then Err.Raise vbObjectError + 750, , "Cannot identify the report file extension."
    stem = Left$(fullName, dotAt - 1)
    suffix = Mid$(fullName, dotAt)
    JSS_BackupPath = stem & "_backup_" & Format$(Now, "yyyymmdd_hhnnss") & suffix
    If Len(Dir$(JSS_BackupPath)) > 0 Then
        JSS_BackupPath = stem & "_backup_" & Format$(Now, "yyyymmdd_hhnnss") & "_2" & suffix
    End If
End Function

' Upgrade older generated reports to the latest layout without deleting notes.
' This runs ONLY after a verified backup has been created. Cut/Insert
' moves table cells and their notes without shifting the top report metrics.
' For v2.0 the rightmost Stakeholder Remarks becomes HL Comment and moves
' immediately after Ticket Status; DM Comment is the next editable column.
Private Sub JSS_UpgradeCommentColumns(ByVal ws As Worksheet)
    Dim headings As Object, heading As String, desired As String
    Dim sourceColumn As Long, targetColumn As Long, n As Long
    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastDataRow < OUTPUT_HEADER_ROW Then lastDataRow = OUTPUT_HEADER_ROW
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Set headings = JSS_ReportHeaders(ws)

    If headings.Exists("survey status") And Not headings.Exists("ticket status") Then
        ws.Cells(OUTPUT_HEADER_ROW, CLng(headings("survey status"))).Value2 = "Ticket Status"
    End If
    Set headings = JSS_ReportHeaders(ws)
    If Not headings.Exists("ticket status") Then
        Err.Raise vbObjectError + 755, , "Cannot locate the report's Ticket Status column."
    End If
    If CLng(headings("ticket status")) <> 8 Then
        Err.Raise vbObjectError + 756, , "The Ticket Status column is no longer in column H; " & _
                  "please restore the approved report layout before refreshing."
    End If

    ' Preserve existing Stakeholder Remarks values by renaming that header
    ' rather than creating an empty replacement or dropping its notes.
    If headings.Exists("stakeholder remarks") And Not headings.Exists("hl comment") Then
        ws.Cells(OUTPUT_HEADER_ROW, CLng(headings("stakeholder remarks"))).Value2 = "HL Comment"
    End If

    For n = 1 To 2
        If n = 1 Then
            desired = "HL Comment"
            targetColumn = 9
        Else
            desired = "DM Comment"
            targetColumn = 10
        End If
        Set headings = JSS_ReportHeaders(ws)
        heading = JSS_NormalizeHeader(desired)
        If headings.Exists(heading) Then
            sourceColumn = CLng(headings(heading))
            If sourceColumn <> targetColumn Then
                ' Do not move a column from left of its final location; that
                ' would indicate a manually rearranged unsupported layout.
                If sourceColumn < targetColumn Then
                    Err.Raise vbObjectError + 757, , "Cannot reposition " & desired & _
                              " safely. Please restore the report's original column order."
                End If
                ' Move ONLY the report table rows. Shifting entire worksheet
                ' columns would also shift the KPI cells in rows 1 to 6.
                ws.Range(ws.Cells(OUTPUT_HEADER_ROW, sourceColumn), _
                         ws.Cells(lastDataRow, sourceColumn)).Cut
                ws.Range(ws.Cells(OUTPUT_HEADER_ROW, targetColumn), _
                         ws.Cells(lastDataRow, targetColumn)).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        Else
            ws.Range(ws.Cells(OUTPUT_HEADER_ROW, targetColumn), _
                     ws.Cells(lastDataRow, targetColumn)).Insert Shift:=xlToRight
            ws.Cells(OUTPUT_HEADER_ROW, targetColumn).Value2 = desired
            With ws.Cells(OUTPUT_HEADER_ROW, targetColumn)
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(24, 49, 83)
                .WrapText = True
                .EntireRow.RowHeight = 34
            End With
        End If
        ws.Columns(targetColumn).ColumnWidth = 43
        ws.Columns(targetColumn).WrapText = True
    Next n
End Sub

' Main Update command. Existing report is never overwritten without first
' saving a complete backup; data is validated before any report modification.
Public Sub JSS_UpdateSurveyStatus()
    Dim reportBook As Workbook, reportSheet As Worksheet
    Dim exportBook As Workbook, latestBook As Workbook, latestSheet As Worksheet
    Dim fileChoice As Variant, fallback As Boolean, wasOpen As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As XlCalculation, startTime As Double, elapsed As Double
    Dim reportHeaders As Object, latestHeaders As Object, managed As Object
    Dim reportKeys As Object, latestKeys As Object
    Dim key As Variant, normalized As String, field As Variant
    Dim latestKeyCol As Long, reportKeyCol As Long
    Dim r As Long, c As Long, destination As Long, lastRow As Long, lastCol As Long
    Dim newCount As Long, updatedCount As Long, unchangedCount As Long
    Dim missingCount As Long, changed As Boolean, backupFile As String
    Dim lastReportKey As String, oldValue As Variant, newValue As Variant
    Dim currentCol As Long, latestCol As Long, currentIndex As Long
    Dim appliedChanges As Boolean, completed As Boolean, message As String
    Dim oldTotal As Long, reportPath As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    startTime = Timer
    On Error GoTo Failed

    Set reportBook = Application.ActiveWorkbook
    If reportBook Is Nothing Then Err.Raise vbObjectError + 751, , "Open your generated status report first."
    Set reportSheet = reportBook.ActiveSheet
    JSS_RequireExistingReport reportBook, reportSheet
    reportPath = reportBook.FullName

    ' GetOpenFilename returns False if the stakeholder cancels. Nothing changes.
    fileChoice = Application.GetOpenFilename( _
        FileFilter:="JIRA exports (*.csv;*.xlsx;*.xls),*.csv;*.xlsx;*.xls", _
        Title:="Select the latest JIRA export")
    If VarType(fileChoice) = vbBoolean Then Exit Sub
    If StrComp(CStr(fileChoice), reportPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 752, , "Select the latest JIRA export, not the existing status report."
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Survey status: reading latest JIRA export..."
    Set exportBook = Application.Workbooks.Open( _
        Filename:=CStr(fileChoice), UpdateLinks:=0, ReadOnly:=True, Local:=True)
    wasOpen = True
    ' One calculation engine drives BOTH Generate and Update. The temporary
    ' workbook's values already include all date/status/validation rules.
    Set latestBook = JSS_GenerateCore(exportBook, True, fallback)
    If fallback Then
        Err.Raise vbObjectError + 753, , "Raw CSV could not be read safely. No update was applied. " & _
            "Download the CSV to a local folder and retry."
    End If
    Set latestSheet = latestBook.Worksheets("Survey Status")
    Set reportHeaders = JSS_ReportHeaders(reportSheet)
    Set latestHeaders = JSS_ReportHeaders(latestSheet)
    reportKeyCol = CLng(reportHeaders("issue key"))
    latestKeyCol = CLng(latestHeaders("issue key"))
    Set reportKeys = JSS_IndexTickets(reportSheet, reportKeyCol)
    Set latestKeys = JSS_IndexTickets(latestSheet, latestKeyCol)
    Set managed = JSS_LoadManagedHeaders(reportBook, reportSheet, latestHeaders)
    oldTotal = reportKeys.Count

    ' Old v1 reports may contain stakeholder-created columns at the right.
    ' Registry-managed columns only are permitted to refresh their values.
    ' Missing JIRA fields remain intact, including older JIRA extra fields.
    ' Create a backup BEFORE editing a single cell in the live report.
    backupFile = JSS_BackupPath(reportBook)
    Application.StatusBar = "Survey status: backing up stakeholder notes..."
    reportBook.SaveCopyAs backupFile
    If Len(Dir$(backupFile)) = 0 Then
        Err.Raise vbObjectError + 754, , "Backup could not be verified; update cancelled."
    End If

    appliedChanges = True
    ' One-time upgrade from v1/v2.0: rename and reposition the existing
    ' stakeholder notes before matching columns from the latest export.
    JSS_UpgradeCommentColumns reportSheet
    Set reportHeaders = JSS_ReportHeaders(reportSheet)
    If managed.Exists("survey status") Then managed.Remove "survey status"
    managed("ticket status") = True
    For Each field In Array("hl comment", "dm comment", "stakeholder remarks")
        If managed.Exists(CStr(field)) Then managed.Remove CStr(field)
    Next field
    ' Add missing fields from the new JIRA export to the RIGHT. Never shift
    ' stakeholder-added columns or replace their existing headings.
    lastCol = JSS_LastHeaderCol(reportSheet)
    For c = 1 To JSS_LastHeaderCol(latestSheet)
        normalized = JSS_NormalizeHeader(CStr(latestSheet.Cells(OUTPUT_HEADER_ROW, c).Value2))
        If Not reportHeaders.Exists(normalized) Then
            If normalized <> JSS_NormalizeHeader(JSS_MATCH_HEADER) Then
                lastCol = lastCol + 1
                reportSheet.Cells(OUTPUT_HEADER_ROW, lastCol).Value2 = _
                    latestSheet.Cells(OUTPUT_HEADER_ROW, c).Value2
                reportHeaders(normalized) = lastCol
                managed(normalized) = True
                With reportSheet.Cells(OUTPUT_HEADER_ROW, lastCol)
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(24, 49, 83)
                    .WrapText = True
                End With
                reportSheet.Columns(lastCol).ColumnWidth = 20
                If JSS_IsDateField(normalized) Then
                    reportSheet.Columns(lastCol).NumberFormat = "mmm dd"
                End If
            End If
        End If
    Next c

    ' Both stakeholder-owned comment fields are guaranteed by the layout
    ' migration, and are never overwritten from the JIRA export.
    If Not reportHeaders.Exists("hl comment") Or Not reportHeaders.Exists("dm comment") Then
        Err.Raise vbObjectError + 758, , "The editable comment fields could not be prepared."
    End If

    If Not reportHeaders.Exists(JSS_NormalizeHeader(JSS_MATCH_HEADER)) Then
        lastCol = lastCol + 1
        reportSheet.Cells(OUTPUT_HEADER_ROW, lastCol).Value2 = JSS_MATCH_HEADER
        reportHeaders.Add JSS_NormalizeHeader(JSS_MATCH_HEADER), lastCol
        With reportSheet.Cells(OUTPUT_HEADER_ROW, lastCol)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(24, 49, 83)
            .WrapText = True
        End With
        reportSheet.Columns(lastCol).ColumnWidth = 25
    End If
    currentIndex = CLng(reportHeaders(JSS_NormalizeHeader(JSS_MATCH_HEADER)))

    ' Update rows IN PLACE by Issue Key; this leaves user-defined columns,
    ' legacy Excel notes and threaded comments attached to existing cells.
    lastRow = OUTPUT_HEADER_ROW + reportKeys.Count
    Application.StatusBar = "Survey status: updating tickets and preserving notes..."
    For r = OUTPUT_FIRST_DATA_ROW To OUTPUT_HEADER_ROW + latestKeys.Count
        lastReportKey = Trim$(CStr(latestSheet.Cells(r, latestKeyCol).Value2))
        changed = False
        If reportKeys.Exists(lastReportKey) Then
            destination = CLng(reportKeys(lastReportKey))
        Else
            lastRow = lastRow + 1
            destination = lastRow
            reportKeys.Add lastReportKey, destination
            newCount = newCount + 1
            With reportSheet.Range(reportSheet.Cells(destination, 1), _
                                   reportSheet.Cells(destination, lastCol))
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                .VerticalAlignment = xlTop
            End With
        End If

        For c = 1 To JSS_LastHeaderCol(latestSheet)
            normalized = JSS_NormalizeHeader(CStr(latestSheet.Cells(OUTPUT_HEADER_ROW, c).Value2))
            If reportHeaders.Exists(normalized) And managed.Exists(normalized) Then
                currentCol = CLng(reportHeaders(normalized))
                latestCol = c
                newValue = latestSheet.Cells(r, latestCol).Value2
                oldValue = reportSheet.Cells(destination, currentCol).Value2
                If Not JSS_ValuesEqual(oldValue, newValue) Then
                    reportSheet.Cells(destination, currentCol).Value2 = newValue
                    changed = True
                End If
            End If
        Next c
        reportSheet.Cells(destination, currentIndex).Value2 = "Included"
        If reportKeys(lastReportKey) <= OUTPUT_HEADER_ROW + oldTotal Then
            If changed Then
                updatedCount = updatedCount + 1
            Else
                unchangedCount = unchangedCount + 1
            End If
        End If
    Next r

    ' A filtered CSV may omit tickets temporarily. Keep them and their notes;
    ' do not infer that they were deleted or change their original stage.
    For r = OUTPUT_FIRST_DATA_ROW To OUTPUT_HEADER_ROW + oldTotal
        key = Trim$(CStr(reportSheet.Cells(r, reportKeyCol).Value2))
        If Not latestKeys.Exists(CStr(key)) Then
            missingCount = missingCount + 1
            reportSheet.Cells(r, currentIndex).Value2 = "Not in Latest Export"
        End If
    Next r

    ' Sort the FULL current report, including all manual columns, so each
    ' note stays on the same ticket. Reapply links after the sort finishes.
    If reportSheet.AutoFilterMode Then reportSheet.AutoFilterMode = False
    JSS_SortBySummary reportSheet, reportKeys.Count, lastCol
    ' Date formats from an older report typically stop at its old last row.
    ' Extend them to appended tickets without touching the original notes.
    If newCount > 0 Then
        For Each field In reportHeaders.Keys
            currentCol = CLng(reportHeaders(CStr(field)))
            If JSS_IsDateField(CStr(field)) Then
                reportSheet.Range(reportSheet.Cells(OUTPUT_FIRST_DATA_ROW, currentCol), _
                    reportSheet.Cells(OUTPUT_HEADER_ROW + reportKeys.Count, currentCol)).NumberFormat = "mmm dd"
            ElseIf CStr(field) = "business days without rvq" Then
                reportSheet.Range(reportSheet.Cells(OUTPUT_FIRST_DATA_ROW, currentCol), _
                    reportSheet.Cells(OUTPUT_HEADER_ROW + reportKeys.Count, currentCol)).NumberFormat = "0"
            End If
        Next field
    End If
    JSS_AddIssueKeyLinks reportSheet, reportKeys.Count
    JSS_ExtendSystemFormats reportSheet, reportKeys.Count
    reportSheet.Range(reportSheet.Cells(OUTPUT_HEADER_ROW, 1), _
        reportSheet.Cells(OUTPUT_HEADER_ROW + reportKeys.Count, lastCol)).AutoFilter

    JSS_UpdateSummaryCounters reportSheet, reportKeys.Count
    With reportSheet
        .Range("D4").Value2 = "New Tickets"
        .Range("E4").Value2 = newCount
        .Range("G4").Value2 = "Updated Tickets"
        .Range("H4").Value2 = updatedCount
        .Range("J4").Value2 = "Unchanged"
        .Range("K4").Value2 = unchangedCount
        .Range("D5").Value2 = "Not in Export"
        .Range("E5").Value2 = missingCount
        .Range("G5").Value2 = "Last Updated"
        .Range("H5").Value = Now
        .Range("H5").NumberFormat = "mmm dd, yyyy hh:mm AM/PM"
        .Range("D4:D5,G4:G5,J4").Font.Bold = True
        .Range("E4:E5,H4,K4").NumberFormat = "0"
        .Range("J3").Value2 = "Last Refreshed"
        .Range("K3").Value = Now
        .Range("K3").NumberFormat = "mmm dd, yyyy hh:mm AM/PM"
    End With
    JSS_SaveManagedHeaders reportBook, managed
    elapsed = Timer - startTime
    If elapsed < 0 Then elapsed = elapsed + 86400#
    reportSheet.Range("B4").Value2 = Round(elapsed, 2)
    reportSheet.Range("B4").NumberFormat = "0.00 ""seconds"""
    reportBook.Save
    completed = True

    latestBook.Close SaveChanges:=False
    Set latestBook = Nothing
    exportBook.Close SaveChanges:=False
    Set exportBook = Nothing
    reportBook.Activate
    reportSheet.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    MsgBox "Survey Status updated and saved successfully." & vbCrLf & _
           "New: " & newCount & " | Updated: " & updatedCount & _
           " | Unchanged: " & unchangedCount & vbCrLf & _
           "Not in latest export: " & missingCount & vbCrLf & _
           "Backup: " & backupFile, vbInformation, "JIRA Survey Status v2.1"
    Exit Sub

Failed:
    message = Err.Description
    On Error Resume Next
    If Not latestBook Is Nothing Then latestBook.Close SaveChanges:=False
    If Not exportBook Is Nothing Then exportBook.Close SaveChanges:=False
    If Not reportBook Is Nothing Then reportBook.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    If appliedChanges And Not completed Then
        MsgBox "The refresh did not complete:" & vbCrLf & message & vbCrLf & vbCrLf & _
               "Do NOT save the partially updated report. Close it without saving " & _
               "and reopen the original or this backup:" & vbCrLf & _
               backupFile, vbCritical, "JIRA Status - Update not saved"
    Else
        MsgBox "The report was not changed:" & vbCrLf & message, vbExclamation, _
               "JIRA Status - Update cancelled"
    End If
End Sub

' Recount the five mutually exclusive survey stages and the separate overdue
' flag across ALL kept rows, including tickets absent from a filtered export.
Private Sub JSS_UpdateSummaryCounters(ByVal ws As Worksheet, ByVal countRows As Long)
    Dim r As Long, stage As String, overdue As String
    Dim approved As Long, rvq As Long, released As Long
    Dim initial As Long, notStarted As Long, overdueCount As Long
    Dim reportHeaders As Object, stageCol As Long, overdueCol As Long
    Set reportHeaders = JSS_ReportHeaders(ws)
    stageCol = CLng(reportHeaders("ticket status"))
    overdueCol = CLng(reportHeaders("rvq overdue flag"))
    For r = OUTPUT_FIRST_DATA_ROW To OUTPUT_HEADER_ROW + countRows
        stage = CStr(ws.Cells(r, stageCol).Value2)
        overdue = CStr(ws.Cells(r, overdueCol).Value2)
        Select Case stage
            Case "Approved": approved = approved + 1
            Case "RVQ Available": rvq = rvq + 1
            Case "VQ Released": released = released + 1
            Case "Initial VQ": initial = initial + 1
            Case Else: notStarted = notStarted + 1
        End Select
        If overdue = "No RVQ Overdue" Then overdueCount = overdueCount + 1
    Next r
    ws.Range("B2").Value2 = countRows
    ws.Range("E2").Value2 = approved
    ws.Range("H2").Value2 = rvq
    ws.Range("K2").Value2 = released
    ws.Range("B3").Value2 = initial
    ws.Range("E3").Value2 = notStarted
    ws.Range("H3").Value2 = overdueCount
End Sub

' Extend existing font-only status rules to new rows without deleting any
' custom stakeholder conditional formatting on other columns.
Private Sub JSS_ExtendSystemFormats(ByVal ws As Worksheet, ByVal countRows As Long)
    Dim name As Variant, r As Long, fc As FormatCondition
    Dim lastRow As Long, columnLetter As String, headers As Object
    lastRow = OUTPUT_HEADER_ROW + countRows
    Set headers = JSS_ReportHeaders(ws)
    For Each name In Array("ticket status", "rvq overdue flag", _
                           "issue resolution status", "data validation")
        If Not headers.Exists(CStr(name)) Then GoTo NextSystemFormat
        columnLetter = JSS_ExcelColumn(CLng(headers(CStr(name))))
        For r = 1 To ws.Range(columnLetter & OUTPUT_FIRST_DATA_ROW).FormatConditions.Count
            Set fc = ws.Range(columnLetter & OUTPUT_FIRST_DATA_ROW).FormatConditions(r)
            fc.ModifyAppliesToRange ws.Range(columnLetter & OUTPUT_FIRST_DATA_ROW & ":" & _
                                            columnLetter & lastRow)
        Next r
NextSystemFormat:
    Next name
    For Each name In Array("pending issue remarks", "issue resolution remarks", _
                           "hl comment", "dm comment")
        If headers.Exists(CStr(name)) Then
            ws.Range(ws.Cells(OUTPUT_FIRST_DATA_ROW, CLng(headers(CStr(name)))), _
                     ws.Cells(lastRow, CLng(headers(CStr(name))))).WrapText = True
        End If
    Next name
End Sub
