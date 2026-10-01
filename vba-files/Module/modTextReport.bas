Attribute VB_Name = "modTextReport"
Option Explicit

Private Const REPORT_OUTPUT_FOLDER As String = "기간보고"
Private Const WEEKLY_UNASSIGNED_MODULE As String = "대분류 미지정"
Private Const WEEKLY_UNASSIGNED_PROGRAM As String = "소분류 미지정"
' Path token -> modification. Rebuilt for each export; includes ancestor rows.
Private mReportModifications As Object

Public Sub 기간보고TXT_생성()
    Dim files As Collection
    Dim reportStart As Date, reportEnd As Date
    Dim outputFolder As String
    Dim eventsWereEnabled As Boolean, eventsCaptured As Boolean
    Dim errDescription As String

    On Error GoTo Failed
    eventsWereEnabled = Application.EnableEvents
    eventsCaptured = True
    Application.EnableEvents = False
    If Len(ThisWorkbook.Path) = 0 Then _
        Err.Raise vbObjectError + 7502, "기간보고TXT_생성", "통합문서를 먼저 저장하세요."
    EnsureConfigSheet
    GetOutputReportPeriod reportStart, reportEnd
    outputFolder = ThisWorkbook.Path & Application.PathSeparator & REPORT_OUTPUT_FOLDER
    EnsureReportOutputDirectory outputFolder
    outputFolder = CreateReportRunDirectory(outputFolder, reportStart, reportEnd)
    Set files = ExportPeriodTextReport(ActiveSheet, outputFolder)
    Application.EnableEvents = eventsWereEnabled
    MsgBox "기간보고 TXT 생성 완료" & vbCrLf & _
        "출력 기간: " & Format$(reportStart, "yyyy-mm-dd") & " ~ " & _
        Format$(reportEnd, "yyyy-mm-dd") & vbCrLf & _
        "TXT 파일: " & files.Count & "개" & vbCrLf & vbCrLf & outputFolder, vbInformation
    Exit Sub
Failed:
    errDescription = Err.Description
    On Error Resume Next
    If eventsCaptured Then Application.EnableEvents = eventsWereEnabled
    On Error GoTo 0
    MsgBox "기간보고 TXT를 생성할 수 없습니다: " & errDescription, vbExclamation
End Sub

' Noninteractive entry point. Returns the complete paths, in report file order.
' Completed: actual finish inside period (planned finish fallback).
' In Progress: actual interval overlaps period; an unfinished interval stays open.
' Planned: planned interval overlaps period; undated tasks remain visible.
Public Function ExportPeriodTextReport(ByVal ws As Worksheet, _
                                       ByVal outputDirectory As String) As Collection
    Dim reportStart As Date, reportEnd As Date
    Dim currentRows As New Collection, plannedRows As New Collection
    Dim groups As Collection, group As Object, assignments As Object
    Dim groupCurrentRows As Collection, groupPlannedRows As Collection
    Dim items As Collection, dates As Collection, levels As Collection, plans As Collection
    Dim files As New Collection
    Dim lastRow As Long, r As Long, groupIndex As Long, categoryDepth As Long
    Dim statusText As String, taskText As String, moduleText As String
    Dim programText As String, ownerText As String, groupingText As String
    Dim hierarchyPath As Variant, rowData As Variant
    Dim pageMode As String, groupLabel As String, outputPath As String
    Dim taskSheetWasProtected As Boolean, errNumber As Long, errDescription As String
    Dim eventsWereEnabled As Boolean, eventsCaptured As Boolean

    On Error GoTo Failed
    eventsWereEnabled = Application.EnableEvents
    eventsCaptured = True
    Application.EnableEvents = False
    If ws Is Nothing Then _
        Err.Raise vbObjectError + 7501, "ExportPeriodTextReport", "업무 시트를 지정하세요."
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Then _
        Err.Raise vbObjectError + 7501, "ExportPeriodTextReport", "업무 시트에서 실행하세요."
    EnsureConfigSheet
    GetOutputReportPeriod reportStart, reportEnd
    EnsureReportOutputDirectory outputDirectory

    taskSheetWasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If taskSheetWasProtected Then UnprotectTaskSheet ws
    SetupDataHeaders ws
    lastRow = GetLastDataRow(ws)
    SynchronizeTaskHierarchyModules ws, lastRow, False
    UpdateWeeklyReportStatuses ws, lastRow
    If taskSheetWasProtected Then
        ApplyCalculatedColumnsProtection ws, lastRow
        taskSheetWasProtected = False
    End If

    Set mReportModifications = CreateObject("Scripting.Dictionary")
    mReportModifications.CompareMode = vbTextCompare
    categoryDepth = GetWeeklyReportVisibleCategoryCount()
    pageMode = GetWeeklyReportPageMode()
    If pageMode = WEEKLY_REPORT_PAGE_MODE_CUSTOM Then _
        LoadWeeklyReportCustomPageAssignments assignments

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            mReportModifications(BuildWeeklyHierarchyPathToken(ws, r)) = _
                CleanReportModificationText(CStr(ws.Cells(r, COL_MODIFICATION).Value2))
            If Not HasChildTask(ws, r, lastRow) Then
                statusText = Trim$(CStr(ws.Cells(r, COL_WEEKLY_REPORT).Value2))
                taskText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_TASK).Value2))
                moduleText = GetWeeklyReportClassificationPath(ws, r)
                programText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_MINOR_CATEGORY).Value2))
                ownerText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_OWNER).Value2))
                groupingText = GetWeeklyReportPageGroupingPath(ws, r)
                hierarchyPath = BuildWeeklyHierarchyPath(ws, r)
                rowData = Array(moduleText, taskText, 0, 0#, "", False, _
                    GetTaskLevel(ws, r), ownerText, r, hierarchyPath, programText, groupingText)
                Select Case UCase$(statusText)
                    Case UCase$(REPORT_STATUS_COMPLETED)
                        If IsCompletedInReportPeriod(ws, r, reportStart, reportEnd) Then
                            rowData(2) = 1
                            rowData(3) = GetWeeklySortDate(ws.Cells(r, COL_ACTUAL_END).Value, _
                                ws.Cells(r, COL_PLAN_END).Value)
                            rowData(4) = BuildCompletedEndDateText(ws, r)
                            AddSortedWeeklyRow currentRows, rowData
                        End If
                    Case UCase$(REPORT_STATUS_IN_PROGRESS)
                        If IsInProgressInReportPeriod(ws, r, reportStart, reportEnd) Then
                            rowData(2) = 2
                            rowData(3) = GetWeeklySortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                            rowData(4) = BuildInProgressEndDateText(ws.Cells(r, COL_PLAN_END).Value)
                            AddSortedWeeklyRow currentRows, rowData
                        End If
                    Case UCase$(REPORT_STATUS_PLANNED)
                        If ReportIntervalsOverlap(ws.Cells(r, COL_PLAN_START).Value, _
                            ws.Cells(r, COL_PLAN_END).Value, reportStart, reportEnd) Then
                            rowData(2) = 3
                            rowData(3) = GetWeeklySortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                            AddSortedWeeklyRow plannedRows, rowData
                        End If
                End Select
            End If
        End If
    Next r

    Set groups = BuildWeeklyPageModuleGroups(currentRows, plannedRows, pageMode, assignments)
    For groupIndex = 1 To groups.Count
        Set group = groups(groupIndex)
        Set groupCurrentRows = New Collection
        Set groupPlannedRows = New Collection
        CopyWeeklyRowsForModuleGroup currentRows, group, groupCurrentRows, _
            (pageMode = WEEKLY_REPORT_PAGE_MODE_MODULE)
        CopyWeeklyRowsForModuleGroup plannedRows, group, groupPlannedRows, _
            (pageMode = WEEKLY_REPORT_PAGE_MODE_MODULE)
        Set items = New Collection
        Set dates = New Collection
        Set levels = New Collection
        Set plans = New Collection
        BuildWeeklyGroupedCurrentItems groupCurrentRows, items, dates, levels, _
            GetWeeklyReportShowModuleOwnerFlag(), GetWeeklyReportShowProgramOwnerFlag(), _
            GetWeeklyReportShowTaskOwnerFlag(), 0, categoryDepth
        BuildWeeklyGroupedPlanItems groupPlannedRows, plans, _
            GetWeeklyReportShowModuleOwnerFlag(), GetWeeklyReportShowProgramOwnerFlag(), _
            GetWeeklyReportShowTaskOwnerFlag(), 0, categoryDepth
        groupLabel = GetReportGroupLabel(group)
        outputPath = outputDirectory & Application.PathSeparator & _
            "기간보고_" & Format$(reportStart, "yyyymmdd") & "_" & _
            Format$(reportEnd, "yyyymmdd") & "_" & Format$(groupIndex, "000") & "_" & _
            SafeReportFileName(groupLabel) & ".txt"
        outputPath = GetUnusedReportFilePath(outputPath)
        WriteReportUtf8 outputPath, BuildPeriodReportText(reportStart, reportEnd, _
            groupLabel, items, dates, levels, plans, categoryDepth)
        files.Add outputPath
    Next groupIndex
    Set mReportModifications = Nothing
    Set ExportPeriodTextReport = files
    Application.EnableEvents = eventsWereEnabled
    Exit Function
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If taskSheetWasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Set mReportModifications = Nothing
    If eventsCaptured Then Application.EnableEvents = eventsWereEnabled
    On Error GoTo 0
    Err.Raise errNumber, "ExportPeriodTextReport", errDescription
End Function

Private Function IsCompletedInReportPeriod(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                           ByVal periodStart As Date, ByVal periodEnd As Date) As Boolean
    Dim finishDate As Variant
    finishDate = ws.Cells(rowNum, COL_ACTUAL_END).Value
    If Not IsDate(finishDate) Then finishDate = ws.Cells(rowNum, COL_PLAN_END).Value
    If IsDate(finishDate) Then
        IsCompletedInReportPeriod = (DateValue(CDate(finishDate)) >= DateValue(periodStart) And _
            DateValue(CDate(finishDate)) <= DateValue(periodEnd))
    End If
End Function

Private Function IsInProgressInReportPeriod(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                            ByVal periodStart As Date, ByVal periodEnd As Date) As Boolean
    Dim startDate As Variant, finishDate As Variant
    startDate = ws.Cells(rowNum, COL_ACTUAL_START).Value
    If Not IsDate(startDate) Then startDate = ws.Cells(rowNum, COL_PLAN_START).Value
    finishDate = ws.Cells(rowNum, COL_ACTUAL_END).Value
    If Not IsDate(finishDate) Then finishDate = DateSerial(9999, 12, 31)
    If Not IsDate(startDate) Then startDate = DateSerial(1900, 1, 1)
    IsInProgressInReportPeriod = ReportIntervalsOverlap(startDate, finishDate, periodStart, periodEnd)
End Function

Private Function ReportIntervalsOverlap(ByVal startValue As Variant, ByVal endValue As Variant, _
                                        ByVal periodStart As Date, ByVal periodEnd As Date) As Boolean
    If Not IsDate(startValue) And Not IsDate(endValue) Then
        ReportIntervalsOverlap = True
        Exit Function
    End If
    If Not IsDate(startValue) Then startValue = endValue
    If Not IsDate(endValue) Then endValue = startValue
    ReportIntervalsOverlap = (DateValue(CDate(startValue)) <= DateValue(periodEnd) And _
        DateValue(CDate(endValue)) >= DateValue(periodStart))
End Function

Private Function BuildPeriodReportText(ByVal periodStart As Date, ByVal periodEnd As Date, _
                                       ByVal groupLabel As String, ByVal items As Collection, _
                                       ByVal dates As Collection, ByVal levels As Collection, _
                                       ByVal plans As Collection, ByVal categoryDepth As Long) As String
    Dim result As String, displayText As String, dateText As String, periodText As String
    Dim unformattedText As String, requestTextStart As Long
    Dim i As Long, levelValue As Long, taskLevel As Long, categoryLevel As Long
    periodText = Format$(periodStart, "yyyy.mm.dd") & "~" & Format$(periodEnd, "yyyy.mm.dd")
    result = "파일 분류: " & groupLabel & vbCrLf & vbCrLf & _
        "기간 : " & periodText & vbCrLf & _
        "업무 현황" & vbCrLf
    ResetWeeklyReportNumbering
    For i = 1 To items.Count
        levelValue = CLng(levels(i))
        displayText = CStr(items(i))
        If levelValue >= 100 Then
            ' The preceding request was already formatted with its actual bullet.
            ' Reuse that text column without consuming another numbering value.
            displayText = BuildReportModificationLines(displayText, requestTextStart, vbCrLf)
        ElseIf levelValue < 0 Then
            categoryLevel = -levelValue
            displayText = FormatWeeklyReportLine(GetWeeklyReportVisibleCategoryPosition(categoryLevel) - 1, _
                GetWeeklyReportCategoryBullet(categoryLevel), displayText, categoryLevel)
        Else
            taskLevel = levelValue - categoryDepth
            If taskLevel < 1 Then taskLevel = 1
            unformattedText = displayText
            displayText = FormatWeeklyReportLine(levelValue - 1, _
                GetWeeklyReportLevelBullet(taskLevel), displayText, taskLevel + 4)
            requestTextStart = GetReportRequestTextStart(displayText, unformattedText, taskLevel)
        End If
        dateText = CStr(dates(i))
        result = result & displayText
        If Len(dateText) > 0 Then result = result & vbTab & dateText
        result = result & vbCrLf
    Next i
    If items.Count > 0 And plans.Count > 0 Then result = result & vbCrLf
    For i = 1 To plans.Count
        result = result & Replace$(CStr(plans(i)), ChrW(11), vbCrLf) & vbCrLf
    Next i
    BuildPeriodReportText = result
End Function

Private Function GetReportModification(ByVal pathToken As String) As String
    If Not GetOutputShowModificationFlag() Then Exit Function
    If mReportModifications Is Nothing Then Exit Function
    If mReportModifications.Exists(pathToken) Then _
        GetReportModification = CStr(mReportModifications(pathToken))
End Function

' Zero-based column of request content, after indentation, bullet and optional level label.
Private Function GetReportRequestTextStart(ByVal formattedLine As String, _
                                          ByVal unformattedText As String, _
                                          ByVal taskLevel As Long) As Long
    GetReportRequestTextStart = Len(formattedLine) - Len(unformattedText)
    If GetWeeklyReportShowTaskNameFlag() And GetWeeklyReportShowTaskLevelFlag() Then _
        GetReportRequestTextStart = GetReportRequestTextStart + Len("Level " & CStr(taskLevel) & " - ")
End Function

' Keep logical line breaks in modifications; whitespace-only cells have no output.
Private Function CleanReportModificationText(ByVal modificationText As String) As String
    Dim lines As Variant, i As Long, firstLine As Long, lastLine As Long
    modificationText = Replace$(modificationText, vbCrLf, vbLf)
    modificationText = Replace$(modificationText, vbCr, vbLf)
    modificationText = Replace$(modificationText, ChrW(11), vbLf)
    If Len(modificationText) = 0 Then Exit Function
    lines = Split(modificationText, vbLf)
    For i = LBound(lines) To UBound(lines)
        lines(i) = Trim$(Replace$(CStr(lines(i)), vbTab, " "))
    Next i
    firstLine = LBound(lines)
    lastLine = UBound(lines)
    Do While firstLine <= lastLine
        If Len(CStr(lines(firstLine))) > 0 Then Exit Do
        firstLine = firstLine + 1
    Loop
    Do While lastLine >= firstLine
        If Len(CStr(lines(lastLine))) > 0 Then Exit Do
        lastLine = lastLine - 1
    Loop
    For i = firstLine To lastLine
        If i > firstLine Then CleanReportModificationText = CleanReportModificationText & vbLf
        CleanReportModificationText = CleanReportModificationText & CStr(lines(i))
    Next i
End Function

Private Function BuildReportModificationLines(ByVal modificationText As String, _
                                               ByVal requestTextStart As Long, _
                                               ByVal lineSeparator As String) As String
    Dim lines As Variant, i As Long
    modificationText = CleanReportModificationText(modificationText)
    If Len(modificationText) = 0 Then Exit Function
    lines = Split(modificationText, vbLf)
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then BuildReportModificationLines = BuildReportModificationLines & lineSeparator
        If Len(CStr(lines(i))) > 0 Then _
            BuildReportModificationLines = BuildReportModificationLines & _
                Space$(requestTextStart) & ChrW(&H2514) & " " & CStr(lines(i))
    Next i
End Function

Private Function GetReportGroupLabel(ByVal group As Object) As String
    Dim key As Variant
    For Each key In group.Keys
        If Len(GetReportGroupLabel) > 0 Then GetReportGroupLabel = GetReportGroupLabel & "; "
        GetReportGroupLabel = GetReportGroupLabel & CStr(key)
    Next key
    If Len(GetReportGroupLabel) = 0 Then GetReportGroupLabel = "전체"
End Function

Private Function SafeReportFileName(ByVal label As String) As String
    Dim invalidCharacter As Variant
    For Each invalidCharacter In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
        label = Replace$(label, CStr(invalidCharacter), "_")
    Next invalidCharacter
    label = Trim$(label)
    If Len(label) > 50 Then label = Left$(label, 50)
    SafeReportFileName = label
End Function

Private Sub EnsureReportOutputDirectory(ByVal directoryPath As String)
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Len(Trim$(directoryPath)) = 0 Then _
        Err.Raise vbObjectError + 7503, "EnsureReportOutputDirectory", "TXT 출력 폴더를 지정하세요."
    If Not fileSystem.FolderExists(directoryPath) Then fileSystem.CreateFolder directoryPath
End Sub

Private Function CreateReportRunDirectory(ByVal baseDirectory As String, _
                                          ByVal periodStart As Date, ByVal periodEnd As Date) As String
    Dim candidate As String, stem As String, suffix As Long, fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    stem = baseDirectory & Application.PathSeparator & _
        Format$(periodStart, "yyyymmdd") & "_" & Format$(periodEnd, "yyyymmdd") & "_" & Format$(Now, "yyyymmdd_hhnnss")
    candidate = stem
    Do While fileSystem.FolderExists(candidate)
        suffix = suffix + 1
        candidate = stem & "_" & CStr(suffix)
    Loop
    fileSystem.CreateFolder candidate
    CreateReportRunDirectory = candidate
End Function

Private Function GetUnusedReportFilePath(ByVal proposedPath As String) As String
    Dim fileSystem As Object, suffix As Long, stem As String, candidate As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    stem = Left$(proposedPath, Len(proposedPath) - 4)
    candidate = proposedPath
    Do While fileSystem.FileExists(candidate)
        suffix = suffix + 1
        candidate = stem & "_" & CStr(suffix) & ".txt"
    Loop
    GetUnusedReportFilePath = candidate
End Function

Private Sub WriteReportUtf8(ByVal outputPath As String, ByVal textValue As String)
    Dim stream As Object, errNumber As Long, errDescription As String
    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.SaveToFile outputPath, 2
    stream.Close
    Exit Sub
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errNumber, "WriteReportUtf8", errDescription
End Sub

Private Function GetWeeklySortDate(ByVal primaryDate As Variant, _
                                   ByVal fallbackDate As Variant) As Double
    If IsDate(primaryDate) Then
        GetWeeklySortDate = CDbl(CDate(primaryDate))
    ElseIf IsDate(fallbackDate) Then
        GetWeeklySortDate = CDbl(CDate(fallbackDate))
    Else
        GetWeeklySortDate = CDbl(DateSerial(9999, 12, 31))
    End If
End Function

Private Sub AddSortedWeeklyRow(ByVal rows As Collection, ByVal newRow As Variant)
    Dim i As Long
    Dim existingRow As Variant

    For i = 1 To rows.Count
        existingRow = rows(i)
        If CLng(newRow(2)) < CLng(existingRow(2)) Or _
           (CLng(newRow(2)) = CLng(existingRow(2)) And _
            CDbl(newRow(3)) < CDbl(existingRow(3))) Then
            rows.Add newRow, Before:=i
            Exit Sub
        End If
    Next i

    rows.Add newRow
End Sub

Private Function BuildWeeklyHierarchyPath(ByVal ws As Worksheet, _
                                          ByVal rowNum As Long) As Variant
    Dim reversePath As Collection
    Dim pathValues() As String
    Dim currentLevel As Long
    Dim candidateLevel As Long
    Dim r As Long
    Dim i As Long

    Set reversePath = New Collection
    reversePath.Add BuildWeeklyHierarchyPathToken(ws, rowNum)
    currentLevel = GetTaskLevel(ws, rowNum)

    For r = rowNum - 1 To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            candidateLevel = GetTaskLevel(ws, r)
            If candidateLevel < currentLevel Then
                reversePath.Add BuildWeeklyHierarchyPathToken(ws, r)
                currentLevel = candidateLevel
                If currentLevel = 1 Then Exit For
            End If
        End If
    Next r

    ReDim pathValues(0 To reversePath.Count - 1)
    For i = 1 To reversePath.Count
        pathValues(i - 1) = CStr(reversePath(reversePath.Count - i + 1))
    Next i

    BuildWeeklyHierarchyPath = pathValues
End Function

Private Function BuildWeeklyHierarchyPathToken(ByVal ws As Worksheet, _
                                               ByVal rowNum As Long) As String
    Dim requestText As String
    requestText = CleanWeeklyReportTaskText(CStr(ws.Cells(rowNum, COL_TASK).Value2))
    If Len(requestText) = 0 And _
       Len(CleanWeeklyReportTaskText(CStr(ws.Cells(rowNum, COL_MODIFICATION).Value2))) > 0 Then _
        requestText = "(요청 내용 없음)"
    BuildWeeklyHierarchyPathToken = CStr(rowNum) & vbTab & requestText
End Function

Private Sub BuildWeeklyGroupedCurrentItems(ByVal rows As Collection, _
                                           ByVal items As Collection, _
                                           ByVal dates As Collection, _
                                           ByVal levels As Collection, _
                                           ByVal showModuleOwnerNames As Boolean, _
                                           ByVal showProgramOwnerNames As Boolean, _
                                           ByVal showTaskOwnerNames As Boolean, _
                                           ByVal taskOwnerLevel As Long, _
                                           ByVal categoryDepth As Long)
    Dim moduleNames As Collection
    Dim moduleSeen As Object
    Dim rowItem As Variant
    Dim moduleName As Variant
    Dim programNames As Collection
    Dim programName As Variant
    Dim categoryParts As Variant
    Dim categoryIndex As Long
    Dim groupByProgram As Boolean
    Dim previousCategoryParts As Variant

    groupByProgram = GetWeeklyReportShowCategoryFlag(4)

    Set moduleNames = New Collection
    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If Not moduleSeen.Exists(GetWeeklyRowModuleName(rowItem)) Then
            moduleSeen.Add GetWeeklyRowModuleName(rowItem), True
            moduleNames.Add GetWeeklyRowModuleName(rowItem)
        End If
    Next rowItem

    For Each moduleName In moduleNames
        categoryParts = Split(CStr(moduleName), " > ")
        For categoryIndex = LBound(categoryParts) To UBound(categoryParts)
            If GetWeeklyReportShowCategoryFlag(categoryIndex + 1) And _
               ShouldAppendWeeklyCategory(categoryParts, previousCategoryParts, _
                                          categoryIndex) Then
                items.Add AppendWeeklyOwnerText( _
                              CStr(categoryParts(categoryIndex)), _
                              GetWeeklyModuleOwnerText(rows, CStr(moduleName)), _
                              GetWeeklyReportShowCategoryOwnerFlag(categoryIndex + 1))
                If GetWeeklyReportShowExpectedDateFlag(categoryIndex + 1) Then
                    dates.Add GetWeeklyCategoryExpectedDateText( _
                                  rows, categoryParts, categoryIndex)
                Else
                    dates.Add ""
                End If
                levels.Add -(categoryIndex + 1)
            End If
        Next categoryIndex

        If groupByProgram Then
            Set programNames = CollectWeeklyProgramNames(rows, CStr(moduleName))
            For Each programName In programNames
                items.Add AppendWeeklyOwnerText( _
                              CStr(programName), _
                              GetWeeklyProgramOwnerText( _
                                  rows, CStr(moduleName), CStr(programName)), _
                              GetWeeklyReportShowCategoryOwnerFlag(4))
                If GetWeeklyReportShowExpectedDateFlag(4) Then
                    dates.Add GetWeeklyProgramExpectedDateText( _
                                  rows, CStr(moduleName), CStr(programName))
                Else
                    dates.Add ""
                End If
                levels.Add -4
                AppendWeeklyCurrentRowsForGroup _
                    rows, items, dates, levels, CStr(moduleName), _
                    CStr(programName), True, categoryDepth, _
                    showTaskOwnerNames, taskOwnerLevel
            Next programName
        Else
            AppendWeeklyCurrentRowsForGroup _
                rows, items, dates, levels, CStr(moduleName), _
                "", False, categoryDepth, showTaskOwnerNames, taskOwnerLevel
        End If
        previousCategoryParts = categoryParts
    Next moduleName
End Sub

Private Function ShouldAppendWeeklyCategory(ByVal currentParts As Variant, _
                                            ByVal previousParts As Variant, _
                                            ByVal categoryIndex As Long) As Boolean
    Dim compareIndex As Long

    If GetWeeklyReportRepeatTreeFlag(categoryIndex + 1) Then
        ShouldAppendWeeklyCategory = True
        Exit Function
    End If
    If Not IsArray(previousParts) Then
        ShouldAppendWeeklyCategory = True
        Exit Function
    End If
    For compareIndex = LBound(currentParts) To categoryIndex
        If compareIndex > UBound(previousParts) Or _
           StrComp(CStr(currentParts(compareIndex)), _
                   CStr(previousParts(compareIndex)), vbTextCompare) <> 0 Then
            ShouldAppendWeeklyCategory = True
            Exit Function
        End If
    Next compareIndex
End Function

Private Sub AppendWeeklyCurrentRowsForGroup(ByVal rows As Collection, _
                                            ByVal items As Collection, _
                                            ByVal dates As Collection, _
                                            ByVal levels As Collection, _
                                            ByVal moduleName As String, _
                                            ByVal programName As String, _
                                            ByVal filterByProgram As Boolean, _
                                            ByVal levelOffset As Long, _
                                            ByVal showTaskOwnerNames As Boolean, _
                                            ByVal taskOwnerLevel As Long)
    Dim currentPath As Variant
    Dim previousPath As Variant
    Dim hierarchyOwners As Object
    Dim hierarchyDates As Object
    Dim rowItem As Variant
    Dim sourceRow As Long
    Dim maxSourceRow As Long

    Set hierarchyOwners = BuildWeeklyHierarchyOwnerMap( _
                              rows, moduleName, programName, filterByProgram)
    Set hierarchyDates = BuildWeeklyHierarchyDateMap( _
                             rows, moduleName, programName, filterByProgram)
    previousPath = Empty
    maxSourceRow = GetWeeklyMaxSourceRow( _
                       rows, moduleName, programName, filterByProgram)

    For sourceRow = DATA_START_ROW To maxSourceRow
        For Each rowItem In rows
            If CLng(rowItem(8)) = sourceRow And _
               WeeklyRowMatchesGroup( _
                   rowItem, moduleName, programName, filterByProgram) Then
                currentPath = rowItem(9)
                AppendWeeklyCurrentHierarchyPath _
                    items, dates, levels, currentPath, previousPath, _
                    hierarchyDates, hierarchyOwners, showTaskOwnerNames, _
                    taskOwnerLevel, levelOffset
                previousPath = currentPath
                Exit For
            End If
        Next rowItem
    Next sourceRow
End Sub

Private Function BuildWeeklyHierarchyOwnerMap(ByVal rows As Collection, _
                                              ByVal moduleName As String, _
                                              ByVal programName As String, _
                                              ByVal filterByProgram As Boolean) As Object
    Dim hierarchyOwners As Object
    Dim ownerSet As Object
    Dim rowItem As Variant
    Dim itemPath As Variant
    Dim pathToken As String
    Dim ownerText As String
    Dim depth As Long

    Set hierarchyOwners = CreateObject("Scripting.Dictionary")
    hierarchyOwners.CompareMode = vbTextCompare

    For Each rowItem In rows
        If WeeklyRowMatchesGroup( _
               rowItem, moduleName, programName, filterByProgram) Then
            itemPath = rowItem(9)
            ownerText = Trim$(CStr(rowItem(7)))

            If Len(ownerText) > 0 Then
                For depth = LBound(itemPath) To UBound(itemPath)
                    pathToken = CStr(itemPath(depth))
                    If hierarchyOwners.Exists(pathToken) Then
                        Set ownerSet = hierarchyOwners(pathToken)
                    Else
                        Set ownerSet = CreateObject("Scripting.Dictionary")
                        ownerSet.CompareMode = vbTextCompare
                        hierarchyOwners.Add pathToken, ownerSet
                    End If
                    AddDistinctOwnerNames ownerSet, ownerText
                Next depth
            End If
        End If
    Next rowItem

    Set BuildWeeklyHierarchyOwnerMap = hierarchyOwners
End Function

Private Function BuildWeeklyHierarchyDateMap(ByVal rows As Collection, _
                                             ByVal moduleName As String, _
                                             ByVal programName As String, _
                                             ByVal filterByProgram As Boolean) As Object
    Dim hierarchyDates As Object
    Dim rowItem As Variant
    Dim itemPath As Variant
    Dim pathToken As String
    Dim dateValue As Variant
    Dim depth As Long

    Set hierarchyDates = CreateObject("Scripting.Dictionary")
    hierarchyDates.CompareMode = vbTextCompare

    For Each rowItem In rows
        If WeeklyRowMatchesGroup( _
               rowItem, moduleName, programName, filterByProgram) And _
           Len(CStr(rowItem(4))) > 0 Then
            itemPath = rowItem(9)
            For depth = LBound(itemPath) To UBound(itemPath)
                pathToken = CStr(itemPath(depth))
                If Not hierarchyDates.Exists(pathToken) Then
                    hierarchyDates.Add pathToken, _
                        Array(CDbl(rowItem(3)), CStr(rowItem(4)))
                Else
                    dateValue = hierarchyDates(pathToken)
                    If CDbl(rowItem(3)) > CDbl(dateValue(0)) Then
                        hierarchyDates(pathToken) = _
                            Array(CDbl(rowItem(3)), CStr(rowItem(4)))
                    End If
                End If
            Next depth
        End If
    Next rowItem

    Set BuildWeeklyHierarchyDateMap = hierarchyDates
End Function

Private Function GetWeeklyMaxSourceRow(ByVal rows As Collection, _
                                       ByVal moduleName As String, _
                                       ByVal programName As String, _
                                       ByVal filterByProgram As Boolean) As Long
    Dim rowItem As Variant

    For Each rowItem In rows
        If WeeklyRowMatchesGroup( _
               rowItem, moduleName, programName, filterByProgram) Then
            If CLng(rowItem(8)) > GetWeeklyMaxSourceRow Then
                GetWeeklyMaxSourceRow = CLng(rowItem(8))
            End If
        End If
    Next rowItem
End Function

Private Sub AppendWeeklyCurrentHierarchyPath(ByVal items As Collection, _
                                             ByVal dates As Collection, _
                                             ByVal levels As Collection, _
                                             ByVal currentPath As Variant, _
                                             ByVal previousPath As Variant, _
                                             ByVal hierarchyDates As Object, _
                                             ByVal hierarchyOwners As Object, _
                                             ByVal showTaskOwnerNames As Boolean, _
                                             ByVal taskOwnerLevel As Long, _
                                             ByVal levelOffset As Long)
    Dim commonDepth As Long
    Dim depth As Long
    Dim pathToken As String
    Dim displayText As String
    Dim dateText As String
    Dim showTaskName As Boolean
    Dim showTaskLevel As Boolean
    Dim showOwnerNames As Boolean

    showTaskName = GetWeeklyReportShowTaskNameFlag()
    showTaskLevel = GetWeeklyReportShowTaskLevelFlag()
    If Not showTaskName And Not showTaskLevel Then Exit Sub
    showOwnerNames = (showTaskName And showTaskOwnerNames) Or _
                     (showTaskLevel And GetWeeklyReportShowTaskLevelOwnerFlag())

    commonDepth = GetCommonWeeklyHierarchyDepth(previousPath, currentPath)
    If GetWeeklyReportRepeatTreeFlag(5) Or _
       GetWeeklyReportRepeatTreeFlag(6) Then commonDepth = 0

    For depth = commonDepth To UBound(currentPath)
        pathToken = CStr(currentPath(depth))
        displayText = ""
        If showTaskLevel Then displayText = "Level " & CStr(depth + 1)
        If showTaskName Then
            If Len(displayText) > 0 Then displayText = displayText & " - "
            displayText = displayText & GetWeeklyHierarchyPathText(pathToken)
        End If
        If showOwnerNames And hierarchyOwners.Exists(pathToken) Then
            displayText = displayText & " (" & _
                          JoinOwnerNameSet(hierarchyOwners(pathToken), ", ") & ")"
        End If

        dateText = ""
        If GetWeeklyReportShowTaskExpectedDateFlag(depth + 1) And _
           hierarchyDates.Exists(pathToken) Then
            dateText = CStr(hierarchyDates(pathToken)(1))
        End If

        items.Add displayText
        dates.Add dateText
        levels.Add depth + 1 + levelOffset
        If showTaskName Then
            If Len(GetReportModification(pathToken)) > 0 Then
                items.Add GetReportModification(pathToken)
                dates.Add ""
                levels.Add 100 + depth + 1 + levelOffset
            End If
        End If
    Next depth
End Sub

Private Function GetWeeklyHierarchyPathText(ByVal pathToken As String) As String
    Dim separatorPosition As Long

    separatorPosition = InStr(1, pathToken, vbTab, vbBinaryCompare)
    If separatorPosition > 0 Then
        GetWeeklyHierarchyPathText = Mid$(pathToken, separatorPosition + 1)
    Else
        GetWeeklyHierarchyPathText = pathToken
    End If
End Function

Private Function GetCommonWeeklyHierarchyDepth(ByVal previousPath As Variant, _
                                               ByVal currentPath As Variant) As Long
    Dim maxDepth As Long
    Dim depth As Long

    If Not IsArray(previousPath) Then Exit Function
    If Not IsArray(currentPath) Then Exit Function

    maxDepth = UBound(previousPath)
    If UBound(currentPath) < maxDepth Then maxDepth = UBound(currentPath)

    For depth = 0 To maxDepth
        If StrComp(CStr(previousPath(depth)), _
                   CStr(currentPath(depth)), vbTextCompare) <> 0 Then Exit For
        GetCommonWeeklyHierarchyDepth = depth + 1
    Next depth
End Function

Private Function BuildWeeklyPageModuleGroups(ByVal currentRows As Collection, _
                                             ByVal plannedRows As Collection, _
                                             ByVal pageMode As String, _
                                             ByVal customPageAssignments As Object) As Collection
    Dim result As Collection
    Dim moduleNames As Collection
    Dim pageGroupsByNumber As Object
    Dim unassignedModules As Collection
    Dim moduleName As Variant
    Dim moduleGroup As Object
    Dim pageNumber As Long
    Dim pageKey As String

    Set result = New Collection

    If pageMode = WEEKLY_REPORT_PAGE_MODE_ALL Then
        Set moduleGroup = CreateObject("Scripting.Dictionary")
        moduleGroup.CompareMode = vbTextCompare
        result.Add moduleGroup
        Set BuildWeeklyPageModuleGroups = result
        Exit Function
    End If

    Set moduleNames = New Collection
    CollectWeeklyModuleNames currentRows, moduleNames, _
        (pageMode = WEEKLY_REPORT_PAGE_MODE_MODULE)
    CollectWeeklyModuleNames plannedRows, moduleNames, _
        (pageMode = WEEKLY_REPORT_PAGE_MODE_MODULE)

    If pageMode = WEEKLY_REPORT_PAGE_MODE_MODULE Then
        For Each moduleName In moduleNames
            Set moduleGroup = CreateObject("Scripting.Dictionary")
            moduleGroup.CompareMode = vbTextCompare
            moduleGroup.Add CStr(moduleName), True
            result.Add moduleGroup
        Next moduleName
    Else
        If customPageAssignments Is Nothing Then
            Err.Raise vbObjectError + 7530, "BuildWeeklyPageModuleGroups", _
                      "커스텀 파일 모드에서는 config_outfoot 시트에 파일 번호와 분류 항목을 한 건 이상 설정해야 합니다."
        End If
        If customPageAssignments.Count = 0 Then
            Err.Raise vbObjectError + 7530, "BuildWeeklyPageModuleGroups", _
                      "커스텀 파일 모드에서는 config_outfoot 시트에 파일 번호와 분류 항목을 한 건 이상 설정해야 합니다."
        End If

        Set pageGroupsByNumber = CreateObject("Scripting.Dictionary")
        Set unassignedModules = New Collection

        For Each moduleName In moduleNames
            If customPageAssignments.Exists(Trim$(CStr(moduleName))) Then
                pageNumber = CLng(customPageAssignments(Trim$(CStr(moduleName))))
                pageKey = CStr(pageNumber)
                If Not pageGroupsByNumber.Exists(pageKey) Then
                    Set moduleGroup = CreateObject("Scripting.Dictionary")
                    moduleGroup.CompareMode = vbTextCompare
                    pageGroupsByNumber.Add pageKey, moduleGroup
                End If
                Set moduleGroup = pageGroupsByNumber(pageKey)
                moduleGroup.Add CStr(moduleName), True
            Else
                unassignedModules.Add CStr(moduleName)
            End If
        Next moduleName

        For pageNumber = 1 To 1000
            pageKey = CStr(pageNumber)
            If pageGroupsByNumber.Exists(pageKey) Then
                result.Add pageGroupsByNumber(pageKey)
            End If
        Next pageNumber

        For Each moduleName In unassignedModules
            Set moduleGroup = CreateObject("Scripting.Dictionary")
            moduleGroup.CompareMode = vbTextCompare
            moduleGroup.Add CStr(moduleName), True
            result.Add moduleGroup
        Next moduleName
    End If

    If result.Count = 0 Then
        Set moduleGroup = CreateObject("Scripting.Dictionary")
        moduleGroup.CompareMode = vbTextCompare
        result.Add moduleGroup
    End If

    Set BuildWeeklyPageModuleGroups = result
End Function

Private Sub CollectWeeklyModuleNames(ByVal rows As Collection, _
                                     ByVal moduleNames As Collection, _
                                     ByVal usePageGrouping As Boolean)
    Dim moduleSeen As Object
    Dim existingName As Variant
    Dim rowItem As Variant
    Dim moduleName As String

    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare
    For Each existingName In moduleNames
        moduleSeen.Add CStr(existingName), True
    Next existingName

    For Each rowItem In rows
        If usePageGrouping Then
            moduleName = GetWeeklyRowPageGroupingName(rowItem)
        Else
            moduleName = GetWeeklyRowModuleName(rowItem)
        End If
        If Not moduleSeen.Exists(moduleName) Then
            moduleSeen.Add moduleName, True
            moduleNames.Add moduleName
        End If
    Next rowItem
End Sub

Private Sub CopyWeeklyRowsForModuleGroup(ByVal sourceRows As Collection, _
                                         ByVal moduleGroup As Object, _
                                         ByVal destinationRows As Collection, _
                                         ByVal usePageGrouping As Boolean)
    Dim rowItem As Variant
    Dim moduleName As String

    For Each rowItem In sourceRows
        If usePageGrouping Then
            moduleName = GetWeeklyRowPageGroupingName(rowItem)
        Else
            moduleName = GetWeeklyRowModuleName(rowItem)
        End If
        If moduleGroup.Count = 0 Or moduleGroup.Exists(moduleName) Then
            destinationRows.Add rowItem
        End If
    Next rowItem
End Sub

Private Function GetWeeklyRowModuleName(ByVal rowItem As Variant) As String
    ' Empty classification slots retain their place for the formatting settings.
    GetWeeklyRowModuleName = CStr(rowItem(0))
    If Len(Trim$(GetWeeklyRowModuleName)) = 0 Then
        GetWeeklyRowModuleName = WEEKLY_UNASSIGNED_MODULE
    End If
End Function

Private Function GetWeeklyRowProgramName(ByVal rowItem As Variant) As String
    GetWeeklyRowProgramName = Trim$(CStr(rowItem(10)))
    If Len(GetWeeklyRowProgramName) = 0 Then
        GetWeeklyRowProgramName = WEEKLY_UNASSIGNED_PROGRAM
    End If
End Function

Private Function WeeklyRowMatchesGroup(ByVal rowItem As Variant, _
                                       ByVal moduleName As String, _
                                       ByVal programName As String, _
                                       ByVal filterByProgram As Boolean) As Boolean
    If StrComp(GetWeeklyRowModuleName(rowItem), _
               moduleName, vbTextCompare) <> 0 Then Exit Function

    If filterByProgram Then
        If StrComp(GetWeeklyRowProgramName(rowItem), _
                   programName, vbTextCompare) <> 0 Then Exit Function
    End If

    WeeklyRowMatchesGroup = True
End Function

Private Function CollectWeeklyProgramNames(ByVal rows As Collection, _
                                           ByVal moduleName As String) As Collection
    Dim result As Collection
    Dim programSeen As Object
    Dim rowItem As Variant
    Dim programName As String
    Dim sourceRow As Long
    Dim maxSourceRow As Long

    Set result = New Collection
    Set programSeen = CreateObject("Scripting.Dictionary")
    programSeen.CompareMode = vbTextCompare
    maxSourceRow = GetWeeklyMaxSourceRow(rows, moduleName, "", False)

    For sourceRow = DATA_START_ROW To maxSourceRow
        For Each rowItem In rows
            If CLng(rowItem(8)) = sourceRow And _
               StrComp(GetWeeklyRowModuleName(rowItem), _
                       moduleName, vbTextCompare) = 0 Then
                programName = GetWeeklyRowProgramName(rowItem)
                If Not programSeen.Exists(programName) Then
                    programSeen.Add programName, True
                    result.Add programName
                End If
                Exit For
            End If
        Next rowItem
    Next sourceRow

    Set CollectWeeklyProgramNames = result
End Function

Private Function GetWeeklyRowPageGroupingName(ByVal rowItem As Variant) As String
    If UBound(rowItem) >= 11 Then _
        GetWeeklyRowPageGroupingName = Trim$(CStr(rowItem(11)))
    If Len(GetWeeklyRowPageGroupingName) = 0 Then _
        GetWeeklyRowPageGroupingName = GetWeeklyRowModuleName(rowItem)
End Function

Private Sub BuildWeeklyGroupedPlanItems(ByVal rows As Collection, _
                                        ByVal items As Collection, _
                                        ByVal showModuleOwnerNames As Boolean, _
                                        ByVal showProgramOwnerNames As Boolean, _
                                        ByVal showTaskOwnerNames As Boolean, _
                                        ByVal taskOwnerLevel As Long, _
                                        ByVal categoryDepth As Long)
    Dim moduleNames As Collection
    Dim moduleSeen As Object
    Dim rowItem As Variant
    Dim moduleName As Variant
    Dim programNames As Collection
    Dim programName As Variant
    Dim blockText As String
    Dim categoryParts As Variant
    Dim categoryIndex As Long
    Dim categoryText As String
    Dim groupByProgram As Boolean
    Dim previousCategoryParts As Variant

    ResetWeeklyReportNumbering
    groupByProgram = GetWeeklyReportShowCategoryFlag(4)

    Set moduleNames = New Collection
    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If Not moduleSeen.Exists(GetWeeklyRowModuleName(rowItem)) Then
            moduleSeen.Add GetWeeklyRowModuleName(rowItem), True
            moduleNames.Add GetWeeklyRowModuleName(rowItem)
        End If
    Next rowItem

    For Each moduleName In moduleNames
        blockText = ""
        categoryParts = Split(CStr(moduleName), " > ")
        For categoryIndex = LBound(categoryParts) To UBound(categoryParts)
            If GetWeeklyReportShowCategoryFlag(categoryIndex + 1) And _
               ShouldAppendWeeklyCategory(categoryParts, previousCategoryParts, _
                                          categoryIndex) Then
                categoryText = AppendWeeklyOwnerText( _
                                   CStr(categoryParts(categoryIndex)), _
                                   GetWeeklyModuleOwnerText(rows, CStr(moduleName)), _
                                   GetWeeklyReportShowCategoryOwnerFlag(categoryIndex + 1))
                categoryText = FormatWeeklyReportLine( _
                                   GetWeeklyReportVisibleCategoryPosition(categoryIndex + 1) - 1, _
                                   GetWeeklyReportCategoryBullet(categoryIndex + 1), categoryText, _
                                   categoryIndex + 1)
                If Len(blockText) > 0 Then blockText = blockText & ChrW(11)
                blockText = blockText & categoryText
            End If
        Next categoryIndex

        If groupByProgram Then
            Set programNames = CollectWeeklyProgramNames(rows, CStr(moduleName))
            For Each programName In programNames
                categoryText = FormatWeeklyReportLine( _
                               GetWeeklyReportVisibleCategoryPosition(4) - 1, _
                               GetWeeklyReportCategoryBullet(4), _
                               AppendWeeklyOwnerText( _
                                   CStr(programName), _
                                   GetWeeklyProgramOwnerText( _
                                       rows, CStr(moduleName), CStr(programName)), _
                                   GetWeeklyReportShowCategoryOwnerFlag(4)), 4)
                If Len(blockText) = 0 Then
                    blockText = categoryText
                Else
                    blockText = blockText & ChrW(11) & categoryText
                End If
                AppendWeeklyPlanRowsForGroup _
                    rows, blockText, CStr(moduleName), CStr(programName), _
                    True, categoryDepth, showTaskOwnerNames, taskOwnerLevel
            Next programName
        Else
            AppendWeeklyPlanRowsForGroup _
                rows, blockText, CStr(moduleName), "", False, categoryDepth, _
                showTaskOwnerNames, taskOwnerLevel
        End If

        If Len(blockText) > 0 Then items.Add blockText
        previousCategoryParts = categoryParts
    Next moduleName
End Sub

Private Sub AppendWeeklyPlanRowsForGroup(ByVal rows As Collection, _
                                         ByRef blockText As String, _
                                         ByVal moduleName As String, _
                                         ByVal programName As String, _
                                         ByVal filterByProgram As Boolean, _
                                         ByVal levelOffset As Long, _
                                         ByVal showTaskOwnerNames As Boolean, _
                                         ByVal taskOwnerLevel As Long)
    Dim currentPath As Variant
    Dim previousPath As Variant
    Dim hierarchyOwners As Object
    Dim rowItem As Variant
    Dim sourceRow As Long
    Dim maxSourceRow As Long

    Set hierarchyOwners = BuildWeeklyHierarchyOwnerMap( _
                              rows, moduleName, programName, filterByProgram)
    previousPath = Empty
    maxSourceRow = GetWeeklyMaxSourceRow( _
                       rows, moduleName, programName, filterByProgram)

    For sourceRow = DATA_START_ROW To maxSourceRow
        For Each rowItem In rows
            If CLng(rowItem(8)) = sourceRow And _
               WeeklyRowMatchesGroup( _
                   rowItem, moduleName, programName, filterByProgram) Then
                currentPath = rowItem(9)
                AppendWeeklyPlanHierarchyPath _
                    blockText, currentPath, previousPath, _
                    hierarchyOwners, showTaskOwnerNames, _
                    taskOwnerLevel, levelOffset
                previousPath = currentPath
                Exit For
            End If
        Next rowItem
    Next sourceRow
End Sub

Private Sub AppendWeeklyPlanHierarchyPath(ByRef blockText As String, _
                                          ByVal currentPath As Variant, _
                                          ByVal previousPath As Variant, _
                                          ByVal hierarchyOwners As Object, _
                                          ByVal showTaskOwnerNames As Boolean, _
                                          ByVal taskOwnerLevel As Long, _
                                          ByVal levelOffset As Long)
    Dim commonDepth As Long
    Dim depth As Long
    Dim pathToken As String
    Dim displayText As String
    Dim lineText As String
    Dim modificationText As String
    Dim showTaskName As Boolean
    Dim showTaskLevel As Boolean
    Dim showOwnerNames As Boolean

    showTaskName = GetWeeklyReportShowTaskNameFlag()
    showTaskLevel = GetWeeklyReportShowTaskLevelFlag()
    If Not showTaskName And Not showTaskLevel Then Exit Sub
    showOwnerNames = (showTaskName And showTaskOwnerNames) Or _
                     (showTaskLevel And GetWeeklyReportShowTaskLevelOwnerFlag())

    commonDepth = GetCommonWeeklyHierarchyDepth(previousPath, currentPath)
    If GetWeeklyReportRepeatTreeFlag(5) Or _
       GetWeeklyReportRepeatTreeFlag(6) Then commonDepth = 0

    For depth = commonDepth To UBound(currentPath)
        pathToken = CStr(currentPath(depth))
        displayText = ""
        If showTaskLevel Then displayText = "Level " & CStr(depth + 1)
        If showTaskName Then
            If Len(displayText) > 0 Then displayText = displayText & " - "
            displayText = displayText & GetWeeklyHierarchyPathText(pathToken)
        End If
        If showOwnerNames And hierarchyOwners.Exists(pathToken) Then
            displayText = displayText & " (" & _
                          JoinOwnerNameSet(hierarchyOwners(pathToken), ", ") & ")"
        End If

        lineText = FormatWeeklyReportLine(depth + levelOffset, _
                       GetWeeklyReportLevelBullet(depth + 1), displayText, depth + 5)
        If Len(blockText) = 0 Then
            blockText = lineText
        Else
            blockText = blockText & ChrW(11) & lineText
        End If
        If showTaskName Then
            modificationText = GetReportModification(pathToken)
            If Len(modificationText) > 0 Then
                blockText = blockText & ChrW(11) & BuildReportModificationLines( _
                    modificationText, GetReportRequestTextStart(lineText, displayText, depth + 1), ChrW(11))
            End If
        End If
    Next depth
End Sub

Private Function GetWeeklyCategoryExpectedDateText( _
                         ByVal rows As Collection, _
                         ByVal categoryParts As Variant, _
                         ByVal categoryIndex As Long) As String
    Dim rowItem As Variant
    Dim rowParts As Variant
    Dim compareIndex As Long
    Dim pathMatches As Boolean
    Dim dateSerial As Double
    Dim latestDateSerial As Double

    latestDateSerial = -1
    For Each rowItem In rows
        rowParts = Split(GetWeeklyRowModuleName(rowItem), " > ")
        pathMatches = (UBound(rowParts) >= categoryIndex)
        If pathMatches Then
            For compareIndex = LBound(categoryParts) To categoryIndex
                If StrComp(CStr(rowParts(compareIndex)), _
                           CStr(categoryParts(compareIndex)), _
                           vbTextCompare) <> 0 Then
                    pathMatches = False
                    Exit For
                End If
            Next compareIndex
        End If
        If pathMatches And Len(CStr(rowItem(4))) > 0 Then
            dateSerial = CDbl(rowItem(3))
            If dateSerial > latestDateSerial Then
                latestDateSerial = dateSerial
                GetWeeklyCategoryExpectedDateText = CStr(rowItem(4))
            End If
        End If
    Next rowItem
End Function

Private Function GetWeeklyProgramExpectedDateText( _
                         ByVal rows As Collection, _
                         ByVal moduleName As String, _
                         ByVal programName As String) As String
    Dim rowItem As Variant
    Dim dateSerial As Double
    Dim latestDateSerial As Double

    latestDateSerial = -1
    For Each rowItem In rows
        If WeeklyRowMatchesGroup( _
               rowItem, moduleName, programName, True) And _
           Len(CStr(rowItem(4))) > 0 Then
            dateSerial = CDbl(rowItem(3))
            If dateSerial > latestDateSerial Then
                latestDateSerial = dateSerial
                GetWeeklyProgramExpectedDateText = CStr(rowItem(4))
            End If
        End If
    Next rowItem
End Function

Private Function AppendWeeklyOwnerText(ByVal displayText As String, _
                                       ByVal ownerText As String, _
                                       ByVal showOwnerNames As Boolean) As String
    AppendWeeklyOwnerText = displayText
    ownerText = Trim$(ownerText)

    If showOwnerNames And Len(ownerText) > 0 Then
        AppendWeeklyOwnerText = displayText & " (" & ownerText & ")"
    End If
End Function

Private Function GetWeeklyModuleOwnerText(ByVal rows As Collection, _
                                          ByVal moduleName As String) As String
    Dim ownerSeen As Object
    Dim rowItem As Variant
    Dim ownerText As String

    Set ownerSeen = CreateObject("Scripting.Dictionary")
    ownerSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If StrComp(GetWeeklyRowModuleName(rowItem), _
                   moduleName, vbTextCompare) = 0 Then
            ownerText = Trim$(CStr(rowItem(7)))
            AddDistinctOwnerNames ownerSeen, ownerText
        End If
    Next rowItem

    GetWeeklyModuleOwnerText = JoinOwnerNameSet(ownerSeen, ", ")
End Function

Private Function GetWeeklyProgramOwnerText(ByVal rows As Collection, _
                                           ByVal moduleName As String, _
                                           ByVal programName As String) As String
    Dim ownerSeen As Object
    Dim rowItem As Variant
    Dim ownerText As String

    Set ownerSeen = CreateObject("Scripting.Dictionary")
    ownerSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If WeeklyRowMatchesGroup(rowItem, moduleName, programName, True) Then
            ownerText = Trim$(CStr(rowItem(7)))
            AddDistinctOwnerNames ownerSeen, ownerText
        End If
    Next rowItem

    GetWeeklyProgramOwnerText = JoinOwnerNameSet(ownerSeen, ", ")
End Function

Private Function BuildInProgressEndDateText(ByVal planEndValue As Variant) As String
    If IsDate(planEndValue) Then
        BuildInProgressEndDateText = "~" & FormatReportMonthDay(CDate(planEndValue))
    Else
        BuildInProgressEndDateText = "~미정"
    End If
End Function

Private Function BuildCompletedEndDateText(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim planEndDate As Variant
    Dim completedDate As Variant

    planEndDate = ws.Cells(rowNum, COL_PLAN_END).Value
    completedDate = ws.Cells(rowNum, COL_ACTUAL_END).Value

    If IsDate(completedDate) Then
        BuildCompletedEndDateText = FormatReportMonthDay(CDate(completedDate))
    ElseIf IsDate(planEndDate) Then
        BuildCompletedEndDateText = FormatReportMonthDay(CDate(planEndDate))
    Else
        BuildCompletedEndDateText = ""
    End If
End Function

Private Function FormatReportMonthDay(ByVal targetDate As Date) As String
    FormatReportMonthDay = Format$(targetDate, "m") & "/" & Format$(targetDate, "d")
End Function

Private Function CleanWeeklyReportTaskText(ByVal taskText As String) As String
    taskText = Replace$(taskText, vbCr, " ")
    taskText = Replace$(taskText, vbLf, " ")
    taskText = Replace$(taskText, vbTab, " ")

    Do While InStr(taskText, "  ") > 0
        taskText = Replace$(taskText, "  ", " ")
    Loop

    CleanWeeklyReportTaskText = Trim$(taskText)
End Function

Private Function FormatWeeklyReportLine(ByVal indentLevel As Long, _
                                        ByVal bulletText As String, _
                                        ByVal displayText As String, _
                                        ByVal bulletIndex As Long) As String
    If indentLevel < 0 Then indentLevel = 0
    FormatWeeklyReportLine = Space$(GetWeeklyReportIndentSpaces(bulletIndex, indentLevel * 4))
    If Len(bulletText) > 0 Then _
        FormatWeeklyReportLine = FormatWeeklyReportLine & bulletText & " "
    FormatWeeklyReportLine = FormatWeeklyReportLine & displayText
End Function

