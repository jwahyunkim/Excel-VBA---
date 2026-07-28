Attribute VB_Name = "modWeeklyPptReport"
Option Explicit

Private Const WEEKLY_PPT_TEMPLATE_SHEET_NAME As String = "WeeklyPptTemplate"
Private Const WEEKLY_PPT_TEMPLATE_MARKER As String = "REO_WEEKLY_PPT_TEMPLATE_BASE64_V1"
Private Const WEEKLY_PPT_OUTPUT_FOLDER As String = "주간보고"
Private Const WEEKLY_PPT_FILE_PREFIX As String = "Digital MFG팀_주간보고_김좌현_"
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24
Private Const WEEKLY_UNASSIGNED_MODULE As String = "모듈 미지정"

Public Sub 주간보고PPT_생성()
    Call GenerateWeeklyPptReport(True)
End Sub

Public Sub 주간보고PPT_버튼_생성()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim anchorCell As Range
    Dim buttonName As String
    Dim lastRow As Long
    Dim wasProtected As Boolean

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws.Name = CONFIG_SHEET_NAME Or _
       ws.Name = REPORT_HISTORY_SHEET_NAME Or _
       ws.Name = WEEKLY_PPT_TEMPLATE_SHEET_NAME Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    buttonName = "btnWeeklyPptReport"
    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If wasProtected Then UnprotectTaskSheet ws

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH

    Set anchorCell = ws.Range("B2")
    Set btn = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        anchorCell.Left + (8 * 80), _
        anchorCell.Top + anchorCell.Height - 22, _
        72, _
        22)

    With btn
        .Name = buttonName
        .OnAction = "주간보고PPT_생성"
        .Placement = xlFreeFloating
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(212, 208, 200)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(128, 128, 128)
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
        .Adjustments.Item(1) = 0.05

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 1
            .MarginBottom = 1

            With .TextRange
                .Characters.Text = "주간 PPT"
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "굴림"
                .Font.Size = 9
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With
    End With

    If wasProtected Then
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If

    MsgBox "버튼 생성 완료: 주간 PPT", vbInformation
    Exit Sub

EH:
    If wasProtected Then
        On Error Resume Next
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
        On Error GoTo 0
    End If
    MsgBox "버튼을 생성할 수 없습니다: " & Err.Description, vbExclamation
End Sub

Public Function GenerateWeeklyPptReport(ByVal showCompletionMessage As Boolean) As String
    Dim ws As Worksheet
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim currentItems As Collection
    Dim currentDates As Collection
    Dim currentLevels As Collection
    Dim plannedItems As Collection
    Dim currentRows As Collection
    Dim plannedRows As Collection
    Dim pageCurrentRows As Collection
    Dim pagePlannedRows As Collection
    Dim pageModuleGroups As Collection
    Dim pageModuleGroup As Object
    Dim customPageAssignments As Object
    Dim outputCurrentItemPages As Collection
    Dim outputCurrentDatePages As Collection
    Dim outputCurrentLevelPages As Collection
    Dim outputPlannedItemPages As Collection
    Dim reportFriday As Date
    Dim currentWeekStart As Date
    Dim currentWeekEnd As Date
    Dim nextWeekStart As Date
    Dim nextWeekEnd As Date
    Dim templatePath As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim lastRow As Long
    Dim r As Long
    Dim statusText As String
    Dim taskText As String
    Dim moduleText As String
    Dim ownerText As String
    Dim useLegacyLayout As Boolean
    Dim showOwnerNames As Boolean
    Dim pageMode As String
    Dim overflowMode As String
    Dim pageIndex As Long
    Dim modulePageIndex As Long
    Dim currentPageCapacity As Long
    Dim planPageCapacity As Long
    Dim pptApp As Object
    Dim presentation As Object
    Dim slide As Object
    Dim duplicatedSlides As Object
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws.Name = CONFIG_SHEET_NAME Or _
       ws.Name = REPORT_HISTORY_SHEET_NAME Or _
       ws.Name = WEEKLY_PPT_TEMPLATE_SHEET_NAME Then
        Err.Raise vbObjectError + 7501, "GenerateWeeklyPptReport", "업무 시트에서 실행하세요."
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 7502, "GenerateWeeklyPptReport", "통합문서를 먼저 저장하세요."
    End If

    reportFriday = GetWeeklyReportFriday(Date)
    currentWeekStart = reportFriday - 4
    currentWeekEnd = reportFriday
    nextWeekStart = reportFriday + 3
    nextWeekEnd = reportFriday + 7

    Set currentItems = New Collection
    Set currentDates = New Collection
    Set currentLevels = New Collection
    Set plannedItems = New Collection
    Set currentRows = New Collection
    Set plannedRows = New Collection

    lastRow = GetLastDataRow(ws)
    EnsureConfigSheet
    showOwnerNames = GetWeeklyReportShowOwnerFlag()
    pageMode = GetWeeklyReportPageMode()
    overflowMode = GetWeeklyReportOverflowMode()
    If pageMode = WEEKLY_REPORT_PAGE_MODE_CUSTOM Then
        LoadWeeklyReportCustomPageAssignments customPageAssignments
    End If
    LoadHolidaySettings holidayDict, workdayDict
    UpdateDevelopmentProgressStatuses ws, lastRow, holidayDict, workdayDict

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            statusText = Trim$(CStr(ws.Cells(r, COL_WEEKLY_REPORT).Value2))
            taskText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_TASK).Value2))
            moduleText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_MODULE).Value2))
            ownerText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_OWNER).Value2))
            useLegacyLayout = (Len(moduleText) = 0 Or Len(ownerText) = 0)

            If Len(taskText) > 0 Then
                Select Case UCase$(statusText)
                    Case UCase$(REPORT_STATUS_IN_PROGRESS)
                        AddSortedWeeklyRow currentRows, Array( _
                            moduleText, _
                            taskText, 2, GetWeeklySortDate(ws.Cells(r, COL_PLAN_END).Value, Empty), _
                            BuildInProgressEndDateText(ws.Cells(r, COL_PLAN_END).Value), _
                            useLegacyLayout, GetTaskLevel(ws, r), ownerText)

                    Case UCase$(REPORT_STATUS_COMPLETED)
                        If IsCompletedInReportWeek(ws, r, currentWeekStart, currentWeekEnd) Then
                            AddSortedWeeklyRow currentRows, Array( _
                                moduleText, _
                                taskText, 1, GetWeeklySortDate(ws.Cells(r, COL_ACTUAL_END).Value, ws.Cells(r, COL_PLAN_END).Value), _
                                BuildCompletedEndDateText(ws, r), _
                                useLegacyLayout, GetTaskLevel(ws, r), ownerText)
                        End If

                    Case UCase$(REPORT_STATUS_PLANNED)
                        If IsPlannedForNextWeek(ws, r, nextWeekStart, nextWeekEnd) Then
                            AddSortedWeeklyRow plannedRows, Array( _
                                moduleText, _
                                taskText, 3, GetWeeklySortDate(ws.Cells(r, COL_PLAN_END).Value, Empty), "", _
                                useLegacyLayout, GetTaskLevel(ws, r), ownerText)
                        End If
                End Select
            End If
        End If
    Next r

    Set pageModuleGroups = BuildWeeklyPageModuleGroups( _
                               currentRows, _
                               plannedRows, _
                               pageMode, _
                               customPageAssignments)

    outputFolder = ThisWorkbook.Path & Application.PathSeparator & WEEKLY_PPT_OUTPUT_FOLDER
    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then MkDir outputFolder

    templatePath = outputFolder & Application.PathSeparator & _
                   "~weekly_ppt_template_" & Format$(Now, "yyyymmdd_hhnnss") & ".pptx"
    ExtractEmbeddedWeeklyPptTemplate templatePath

    outputPath = outputFolder & Application.PathSeparator & _
                 WEEKLY_PPT_FILE_PREFIX & Format$(reportFriday, "yyyymmdd") & ".pptx"

    If Len(Dir$(outputPath)) > 0 Then
        outputPath = outputFolder & Application.PathSeparator & _
                     WEEKLY_PPT_FILE_PREFIX & Format$(reportFriday, "yyyymmdd") & _
                     "_" & Format$(Now, "hhnnss") & ".pptx"
    End If

    Set pptApp = CreateObject("PowerPoint.Application")
    Set presentation = pptApp.Presentations.Open(templatePath, False, False, False)

    Set outputCurrentItemPages = New Collection
    Set outputCurrentDatePages = New Collection
    Set outputCurrentLevelPages = New Collection
    Set outputPlannedItemPages = New Collection

    If overflowMode = WEEKLY_REPORT_OVERFLOW_MODE_NEW_SLIDE Then
        currentPageCapacity = GetWeeklyReportCurrentPageCapacity(presentation.Slides(1))
        planPageCapacity = GetWeeklyReportPlanPageCapacity(presentation.Slides(1))
    End If

    For modulePageIndex = 1 To pageModuleGroups.Count
        Set pageModuleGroup = pageModuleGroups(modulePageIndex)
        Set pageCurrentRows = New Collection
        Set pagePlannedRows = New Collection
        CopyWeeklyRowsForModuleGroup currentRows, pageModuleGroup, pageCurrentRows
        CopyWeeklyRowsForModuleGroup plannedRows, pageModuleGroup, pagePlannedRows

        Set currentItems = New Collection
        Set currentDates = New Collection
        Set currentLevels = New Collection
        Set plannedItems = New Collection
        BuildWeeklyGroupedCurrentItems _
            pageCurrentRows, currentItems, currentDates, currentLevels, showOwnerNames
        BuildWeeklyGroupedPlanItems pagePlannedRows, plannedItems, showOwnerNames

        AppendWeeklyOutputPages _
            currentItems, currentDates, currentLevels, plannedItems, _
            overflowMode, currentPageCapacity, planPageCapacity, _
            outputCurrentItemPages, outputCurrentDatePages, _
            outputCurrentLevelPages, outputPlannedItemPages
    Next modulePageIndex

    For pageIndex = 2 To outputCurrentItemPages.Count
        Set duplicatedSlides = presentation.Slides(1).Duplicate
        Set duplicatedSlides = Nothing
    Next pageIndex

    For pageIndex = 1 To outputCurrentItemPages.Count
        Set currentItems = outputCurrentItemPages(pageIndex)
        Set currentDates = outputCurrentDatePages(pageIndex)
        Set currentLevels = outputCurrentLevelPages(pageIndex)
        Set plannedItems = outputPlannedItemPages(pageIndex)

        Set slide = presentation.Slides(pageIndex)
        FillWeeklyReportPeriodText _
            slide, currentWeekStart, currentWeekEnd, nextWeekStart, nextWeekEnd
        FillWeeklyReportCurrentTable slide, currentItems, currentDates, currentLevels
        FillWeeklyReportPlanArea _
            slide, _
            plannedItems, _
            (overflowMode = WEEKLY_REPORT_OVERFLOW_MODE_EXPAND)
    Next pageIndex

    presentation.SaveAs outputPath, PPT_SAVE_AS_OPEN_XML_PRESENTATION
    presentation.Close
    Set presentation = Nothing
    DeleteTemporaryWeeklyPptTemplate templatePath

    If showCompletionMessage Then
        pptApp.Visible = True
        Set presentation = pptApp.Presentations.Open(outputPath, False, False, True)
    Else
        pptApp.Quit
        Set pptApp = Nothing
    End If

    GenerateWeeklyPptReport = outputPath

    If showCompletionMessage Then
        MsgBox "주간보고 PPT 생성 완료" & vbCrLf & _
               "보고 기준일: " & Format$(reportFriday, "yyyy-mm-dd") & vbCrLf & _
               "업무 현황: " & currentRows.Count & "개" & vbCrLf & _
               "개발 계획: " & plannedRows.Count & "개" & vbCrLf & _
               "PPT 페이지: " & outputCurrentItemPages.Count & "개" & vbCrLf & vbCrLf & _
               outputPath, vbInformation
    End If
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not presentation Is Nothing Then presentation.Close
    If Not pptApp Is Nothing Then pptApp.Quit
    DeleteTemporaryWeeklyPptTemplate templatePath
    On Error GoTo 0

    If showCompletionMessage Then
        MsgBox "주간보고 PPT를 생성할 수 없습니다: " & errDescription, vbExclamation
        GenerateWeeklyPptReport = ""
    Else
        Err.Raise errNumber, "GenerateWeeklyPptReport", errDescription
    End If
End Function

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

Private Sub BuildWeeklyGroupedCurrentItems(ByVal rows As Collection, _
                                           ByVal items As Collection, _
                                           ByVal dates As Collection, _
                                           ByVal levels As Collection, _
                                           ByVal showOwnerNames As Boolean)
    Dim moduleNames As Collection
    Dim moduleSeen As Object
    Dim rowItem As Variant
    Dim moduleName As Variant

    Set moduleNames = New Collection
    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If Not CBool(rowItem(5)) And Not moduleSeen.Exists(CStr(rowItem(0))) Then
            moduleSeen.Add CStr(rowItem(0)), True
            moduleNames.Add CStr(rowItem(0))
        End If
    Next rowItem

    For Each moduleName In moduleNames
        items.Add AppendWeeklyOwnerText( _
                      CStr(moduleName), _
                      GetWeeklyModuleOwnerText(rows, CStr(moduleName)), _
                      showOwnerNames)
        dates.Add ""
        levels.Add 1

        For Each rowItem In rows
            If Not CBool(rowItem(5)) And _
               StrComp(CStr(rowItem(0)), CStr(moduleName), vbTextCompare) = 0 Then
                items.Add AppendWeeklyOwnerText( _
                              CStr(rowItem(1)), _
                              CStr(rowItem(7)), _
                              showOwnerNames)
                dates.Add CStr(rowItem(4))
                levels.Add 2
            End If
        Next rowItem
    Next moduleName

    For Each rowItem In rows
        If CBool(rowItem(5)) Then
            items.Add AppendWeeklyOwnerText( _
                          CStr(rowItem(1)), _
                          CStr(rowItem(7)), _
                          showOwnerNames)
            dates.Add CStr(rowItem(4))
            levels.Add CLng(rowItem(6))
        End If
    Next rowItem
End Sub

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
    CollectWeeklyModuleNames currentRows, moduleNames
    CollectWeeklyModuleNames plannedRows, moduleNames

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
                      "커스텀 페이지 모드에서는 config 시트에 페이지 번호와 모듈을 한 건 이상 설정해야 합니다."
        End If
        If customPageAssignments.Count = 0 Then
            Err.Raise vbObjectError + 7530, "BuildWeeklyPageModuleGroups", _
                      "커스텀 페이지 모드에서는 config 시트에 페이지 번호와 모듈을 한 건 이상 설정해야 합니다."
        End If

        Set pageGroupsByNumber = CreateObject("Scripting.Dictionary")
        Set unassignedModules = New Collection

        For Each moduleName In moduleNames
            If customPageAssignments.Exists(CStr(moduleName)) Then
                pageNumber = CLng(customPageAssignments(CStr(moduleName)))
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
                                     ByVal moduleNames As Collection)
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
        moduleName = GetWeeklyRowModuleName(rowItem)
        If Not moduleSeen.Exists(moduleName) Then
            moduleSeen.Add moduleName, True
            moduleNames.Add moduleName
        End If
    Next rowItem
End Sub

Private Sub CopyWeeklyRowsForModuleGroup(ByVal sourceRows As Collection, _
                                         ByVal moduleGroup As Object, _
                                         ByVal destinationRows As Collection)
    Dim rowItem As Variant
    Dim moduleName As String

    For Each rowItem In sourceRows
        moduleName = GetWeeklyRowModuleName(rowItem)
        If moduleGroup.Count = 0 Or moduleGroup.Exists(moduleName) Then
            destinationRows.Add rowItem
        End If
    Next rowItem
End Sub

Private Function GetWeeklyRowModuleName(ByVal rowItem As Variant) As String
    GetWeeklyRowModuleName = Trim$(CStr(rowItem(0)))
    If Len(GetWeeklyRowModuleName) = 0 Then
        GetWeeklyRowModuleName = WEEKLY_UNASSIGNED_MODULE
    End If
End Function

Private Function GetWeeklyReportCurrentPageCapacity(ByVal slide As Object) As Long
    Dim tableShape As Object
    Dim taskSlotCount As Long
    Dim dateSlotCount As Long

    Set tableShape = FindFirstTableShape(slide)
    If tableShape Is Nothing Then
        Err.Raise vbObjectError + 7540, "GetWeeklyReportCurrentPageCapacity", _
                  "PPT에서 업무 현황 표를 찾을 수 없습니다."
    End If

    taskSlotCount = tableShape.Table.Cell(2, 2).Shape.TextFrame.TextRange.Paragraphs.Count
    dateSlotCount = tableShape.Table.Cell(2, 3).Shape.TextFrame.TextRange.Paragraphs.Count
    GetWeeklyReportCurrentPageCapacity = taskSlotCount
    If dateSlotCount < GetWeeklyReportCurrentPageCapacity Then
        GetWeeklyReportCurrentPageCapacity = dateSlotCount
    End If
    If GetWeeklyReportCurrentPageCapacity < 1 Then GetWeeklyReportCurrentPageCapacity = 1
End Function

Private Function GetWeeklyReportPlanPageCapacity(ByVal slide As Object) As Long
    Dim planShape As Object
    Dim planTextRange As Object
    Dim maintenanceParagraphIndex As Long
    Dim lineHeight As Double
    Dim blankSlotHeight As Double
    Dim fixedContentHeight As Double
    Dim availableContentHeight As Double

    Set planShape = FindTextShape(slide, "(개발 항목)")
    If planShape Is Nothing Then
        Err.Raise vbObjectError + 7541, "GetWeeklyReportPlanPageCapacity", _
                  "PPT에서 '(개발 항목)' 영역을 찾을 수 없습니다."
    End If

    Set planTextRange = planShape.TextFrame.TextRange
    maintenanceParagraphIndex = FindPowerPointParagraphIndex( _
                                    planTextRange, _
                                    "(유지보수 항목)")
    If maintenanceParagraphIndex < 3 Then
        Err.Raise vbObjectError + 7542, "GetWeeklyReportPlanPageCapacity", _
                  "원본 PPT 개발 항목 영역의 문단 구조가 예상과 다릅니다."
    End If

    lineHeight = planTextRange.Paragraphs(2).BoundHeight
    If lineHeight <= 0 Then lineHeight = planTextRange.Paragraphs(2).Font.Size * 1.2
    If lineHeight <= 0 Then lineHeight = 12

    blankSlotHeight = 0
    blankSlotHeight = blankSlotHeight + planTextRange.Paragraphs(2).BoundHeight
    blankSlotHeight = blankSlotHeight + planTextRange.Paragraphs(3).BoundHeight
    fixedContentHeight = planTextRange.BoundHeight - blankSlotHeight
    availableContentHeight = planShape.Height - _
                             planShape.TextFrame.MarginTop - _
                             planShape.TextFrame.MarginBottom - _
                             fixedContentHeight
    GetWeeklyReportPlanPageCapacity = CLng(Fix(availableContentHeight / lineHeight))
    If GetWeeklyReportPlanPageCapacity < 1 Then GetWeeklyReportPlanPageCapacity = 1
End Function

Private Sub AppendWeeklyOutputPages(ByVal currentItems As Collection, _
                                    ByVal currentDates As Collection, _
                                    ByVal currentLevels As Collection, _
                                    ByVal plannedItems As Collection, _
                                    ByVal overflowMode As String, _
                                    ByVal currentPageCapacity As Long, _
                                    ByVal planPageCapacity As Long, _
                                    ByVal outputCurrentItemPages As Collection, _
                                    ByVal outputCurrentDatePages As Collection, _
                                    ByVal outputCurrentLevelPages As Collection, _
                                    ByVal outputPlannedItemPages As Collection)
    Dim currentItemPages As Collection
    Dim currentDatePages As Collection
    Dim currentLevelPages As Collection
    Dim plannedItemPages As Collection
    Dim pageItems As Collection
    Dim pageDates As Collection
    Dim pageLevels As Collection
    Dim pagePlans As Collection
    Dim pageCount As Long
    Dim pageIndex As Long

    If overflowMode = WEEKLY_REPORT_OVERFLOW_MODE_EXPAND Then
        outputCurrentItemPages.Add currentItems
        outputCurrentDatePages.Add currentDates
        outputCurrentLevelPages.Add currentLevels
        outputPlannedItemPages.Add plannedItems
        Exit Sub
    End If

    SplitWeeklyCurrentItems _
        currentItems, currentDates, currentLevels, currentPageCapacity, _
        currentItemPages, currentDatePages, currentLevelPages
    Set plannedItemPages = SplitWeeklyPlannedItems(plannedItems, planPageCapacity)

    pageCount = currentItemPages.Count
    If plannedItemPages.Count > pageCount Then pageCount = plannedItemPages.Count

    For pageIndex = 1 To pageCount
        If pageIndex <= currentItemPages.Count Then
            Set pageItems = currentItemPages(pageIndex)
            Set pageDates = currentDatePages(pageIndex)
            Set pageLevels = currentLevelPages(pageIndex)
        Else
            Set pageItems = New Collection
            Set pageDates = New Collection
            Set pageLevels = New Collection
        End If

        If pageIndex <= plannedItemPages.Count Then
            Set pagePlans = plannedItemPages(pageIndex)
        Else
            Set pagePlans = New Collection
        End If

        outputCurrentItemPages.Add pageItems
        outputCurrentDatePages.Add pageDates
        outputCurrentLevelPages.Add pageLevels
        outputPlannedItemPages.Add pagePlans
    Next pageIndex
End Sub

Private Sub SplitWeeklyCurrentItems(ByVal items As Collection, _
                                    ByVal dates As Collection, _
                                    ByVal levels As Collection, _
                                    ByVal pageCapacity As Long, _
                                    ByRef itemPages As Collection, _
                                    ByRef datePages As Collection, _
                                    ByRef levelPages As Collection)
    Dim pageItems As Collection
    Dim pageDates As Collection
    Dim pageLevels As Collection
    Dim sourceIndex As Long
    Dim levelValue As Long

    Set itemPages = New Collection
    Set datePages = New Collection
    Set levelPages = New Collection
    If pageCapacity < 1 Then pageCapacity = 1

    If items.Count = 0 Then
        Set pageItems = New Collection
        Set pageDates = New Collection
        Set pageLevels = New Collection
        itemPages.Add pageItems
        datePages.Add pageDates
        levelPages.Add pageLevels
        Exit Sub
    End If

    sourceIndex = 1
    Do While sourceIndex <= items.Count
        Set pageItems = New Collection
        Set pageDates = New Collection
        Set pageLevels = New Collection

        Do While sourceIndex <= items.Count And pageItems.Count < pageCapacity
            levelValue = CLng(levels(sourceIndex))
            If levelValue = 1 And _
               pageItems.Count = pageCapacity - 1 And _
               sourceIndex < items.Count And _
               pageItems.Count > 0 Then
                Exit Do
            End If

            pageItems.Add CStr(items(sourceIndex))
            pageDates.Add CStr(dates(sourceIndex))
            pageLevels.Add levelValue
            sourceIndex = sourceIndex + 1
        Loop

        itemPages.Add pageItems
        datePages.Add pageDates
        levelPages.Add pageLevels
    Loop
End Sub

Private Function SplitWeeklyPlannedItems(ByVal items As Collection, _
                                         ByVal pageCapacity As Long) As Collection
    Dim result As Collection
    Dim pageItems As Collection
    Dim itemChunks As Collection
    Dim itemText As Variant
    Dim itemChunk As Variant
    Dim chunkLineCount As Long
    Dim usedLineCount As Long

    Set result = New Collection
    If pageCapacity < 1 Then pageCapacity = 1

    If items.Count = 0 Then
        Set pageItems = New Collection
        result.Add pageItems
        Set SplitWeeklyPlannedItems = result
        Exit Function
    End If

    Set pageItems = New Collection
    usedLineCount = 0

    For Each itemText In items
        Set itemChunks = SplitWeeklyPlanItemText(CStr(itemText), pageCapacity)
        For Each itemChunk In itemChunks
            chunkLineCount = CountWeeklyPlanItemLines(CStr(itemChunk))
            If pageItems.Count > 0 And usedLineCount + chunkLineCount > pageCapacity Then
                result.Add pageItems
                Set pageItems = New Collection
                usedLineCount = 0
            End If

            pageItems.Add CStr(itemChunk)
            usedLineCount = usedLineCount + chunkLineCount
        Next itemChunk
    Next itemText

    If pageItems.Count > 0 Then result.Add pageItems

    Set SplitWeeklyPlannedItems = result
End Function

Private Function SplitWeeklyPlanItemText(ByVal itemText As String, _
                                         ByVal pageCapacity As Long) As Collection
    Dim result As Collection
    Dim lines As Variant
    Dim firstLine As String
    Dim chunkText As String
    Dim lineIndex As Long
    Dim chunkLineCount As Long

    Set result = New Collection
    If pageCapacity < 1 Then pageCapacity = 1
    lines = Split(itemText, ChrW(11))
    firstLine = CStr(lines(LBound(lines)))
    lineIndex = LBound(lines)

    Do While lineIndex <= UBound(lines)
        chunkText = ""
        chunkLineCount = 0

        If lineIndex > LBound(lines) And pageCapacity > 1 Then
            chunkText = firstLine
            chunkLineCount = 1
        End If

        Do While lineIndex <= UBound(lines) And chunkLineCount < pageCapacity
            If Len(chunkText) > 0 Then chunkText = chunkText & ChrW(11)
            chunkText = chunkText & CStr(lines(lineIndex))
            chunkLineCount = chunkLineCount + 1
            lineIndex = lineIndex + 1
        Loop

        result.Add chunkText
    Loop

    Set SplitWeeklyPlanItemText = result
End Function

Private Function CountWeeklyPlanItemLines(ByVal itemText As String) As Long
    CountWeeklyPlanItemLines = UBound(Split(itemText, ChrW(11))) + 1
End Function

Private Sub BuildWeeklyGroupedPlanItems(ByVal rows As Collection, _
                                         ByVal items As Collection, _
                                         ByVal showOwnerNames As Boolean)
    Dim moduleNames As Collection
    Dim moduleSeen As Object
    Dim rowItem As Variant
    Dim moduleName As Variant
    Dim blockText As String

    Set moduleNames = New Collection
    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If Not CBool(rowItem(5)) And Not moduleSeen.Exists(CStr(rowItem(0))) Then
            moduleSeen.Add CStr(rowItem(0)), True
            moduleNames.Add CStr(rowItem(0))
        End If
    Next rowItem

    For Each moduleName In moduleNames
        blockText = AppendWeeklyOwnerText( _
                        CStr(moduleName), _
                        GetWeeklyModuleOwnerText(rows, CStr(moduleName)), _
                        showOwnerNames)
        For Each rowItem In rows
            If Not CBool(rowItem(5)) And _
               StrComp(CStr(rowItem(0)), CStr(moduleName), vbTextCompare) = 0 Then
                blockText = blockText & ChrW(11) & "    " & ChrW(&H2022) & " " & _
                            AppendWeeklyOwnerText( _
                                CStr(rowItem(1)), _
                                CStr(rowItem(7)), _
                                showOwnerNames)
            End If
        Next rowItem
        items.Add blockText
    Next moduleName

    For Each rowItem In rows
        If CBool(rowItem(5)) Then
            items.Add AppendWeeklyOwnerText( _
                          CStr(rowItem(1)), _
                          CStr(rowItem(7)), _
                          showOwnerNames)
        End If
    Next rowItem
End Sub

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
    Dim ownerNames As Collection
    Dim ownerSeen As Object
    Dim rowItem As Variant
    Dim ownerText As String
    Dim ownerName As Variant
    Dim result As String

    Set ownerNames = New Collection
    Set ownerSeen = CreateObject("Scripting.Dictionary")
    ownerSeen.CompareMode = vbTextCompare

    For Each rowItem In rows
        If Not CBool(rowItem(5)) And _
           StrComp(CStr(rowItem(0)), moduleName, vbTextCompare) = 0 Then
            ownerText = Trim$(CStr(rowItem(7)))
            If Len(ownerText) > 0 And Not ownerSeen.Exists(ownerText) Then
                ownerSeen.Add ownerText, True
                ownerNames.Add ownerText
            End If
        End If
    Next rowItem

    For Each ownerName In ownerNames
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(ownerName)
    Next ownerName

    GetWeeklyModuleOwnerText = result
End Function

Private Function GetWeeklyReportFriday(ByVal targetDate As Date) As Date
    Dim weekdayNumber As Long

    weekdayNumber = Weekday(targetDate, vbMonday)

    If weekdayNumber <= 5 Then
        GetWeeklyReportFriday = DateValue(targetDate) + (5 - weekdayNumber)
    Else
        GetWeeklyReportFriday = DateValue(targetDate) - (weekdayNumber - 5)
    End If
End Function

Private Function IsCompletedInReportWeek(ByVal ws As Worksheet, _
                                         ByVal rowNum As Long, _
                                         ByVal weekStart As Date, _
                                         ByVal weekEnd As Date) As Boolean
    Dim completedDate As Variant

    completedDate = ws.Cells(rowNum, COL_ACTUAL_END).Value
    If Not IsDate(completedDate) Then completedDate = ws.Cells(rowNum, COL_PLAN_END).Value
    If Not IsDate(completedDate) Then Exit Function

    IsCompletedInReportWeek = _
        (CLng(CDate(completedDate)) >= CLng(weekStart) And _
         CLng(CDate(completedDate)) <= CLng(weekEnd))
End Function

Private Function IsPlannedForNextWeek(ByVal ws As Worksheet, _
                                      ByVal rowNum As Long, _
                                      ByVal nextWeekStart As Date, _
                                      ByVal nextWeekEnd As Date) As Boolean
    Dim planStart As Variant
    Dim planEnd As Variant

    planStart = ws.Cells(rowNum, COL_PLAN_START).Value
    planEnd = ws.Cells(rowNum, COL_PLAN_END).Value

    If Not IsDate(planStart) And Not IsDate(planEnd) Then
        IsPlannedForNextWeek = True
        Exit Function
    End If

    If Not IsDate(planStart) Then planStart = planEnd
    If Not IsDate(planEnd) Then planEnd = planStart

    IsPlannedForNextWeek = _
        (CLng(CDate(planStart)) <= CLng(nextWeekEnd) And _
         CLng(CDate(planEnd)) >= CLng(nextWeekStart))
End Function

Private Function BuildInProgressEndDateText(ByVal planEndValue As Variant) As String
    If IsDate(planEndValue) Then
        BuildInProgressEndDateText = "~" & FormatPptMonthDay(CDate(planEndValue))
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
        BuildCompletedEndDateText = FormatPptMonthDay(CDate(completedDate))
    ElseIf IsDate(planEndDate) Then
        BuildCompletedEndDateText = FormatPptMonthDay(CDate(planEndDate))
    Else
        BuildCompletedEndDateText = ""
    End If
End Function

Private Function FormatPptMonthDay(ByVal targetDate As Date) As String
    ' Zero-width spaces prevent PowerPoint from auto-converting 7/21 to 7-21.
    FormatPptMonthDay = _
        Format$(targetDate, "m") & ChrW(&H200B) & "/" & _
        ChrW(&H200B) & Format$(targetDate, "d")
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

Private Sub FillWeeklyReportPeriodText(ByVal slide As Object, _
                                       ByVal currentWeekStart As Date, _
                                       ByVal currentWeekEnd As Date, _
                                       ByVal nextWeekStart As Date, _
                                       ByVal nextWeekEnd As Date)
    Dim previousWeekShape As Object
    Dim nextWeekShape As Object

    Set previousWeekShape = FindTextShape(slide, "전주 중요 추진 업무 현황")
    Set nextWeekShape = FindTextShape(slide, "금주 주요 계획")

    If previousWeekShape Is Nothing Then
        Err.Raise vbObjectError + 7510, "FillWeeklyReportPeriodText", _
                  "PPT에서 '전주 중요 추진 업무 현황' 영역을 찾을 수 없습니다."
    End If

    If nextWeekShape Is Nothing Then
        Err.Raise vbObjectError + 7511, "FillWeeklyReportPeriodText", _
                  "PPT에서 '금주 주요 계획' 영역을 찾을 수 없습니다."
    End If

    ReplaceTextAfterMarkerPreservingStyle _
        previousWeekShape.TextFrame.TextRange, _
        ") ", _
        Format$(currentWeekStart, "yyyy.mm.dd") & "~" & Format$(currentWeekEnd, "yyyy.mm.dd")

    ReplaceTextAfterMarkerPreservingStyle _
        nextWeekShape.TextFrame.TextRange, _
        ") ", _
        Format$(nextWeekStart, "yyyy.mm.dd") & "~" & Format$(nextWeekEnd, "yyyy.mm.dd")
End Sub

Private Sub FillWeeklyReportCurrentTable(ByVal slide As Object, _
                                         ByVal items As Collection, _
                                         ByVal dateItems As Collection, _
                                         ByVal levelItems As Collection)
    Dim tableShape As Object
    Dim table As Object
    Dim taskTextRange As Object
    Dim dateTextRange As Object
    Dim taskParagraph As Object
    Dim levelValue As Long
    Dim displayText As String
    Dim dateDisplayText As String
    Dim baseLineHeight As Double
    Dim taskLineCount As Long
    Dim lineIndex As Long
    Dim originalTableHeight As Double
    Dim i As Long

    Set tableShape = FindFirstTableShape(slide)
    If tableShape Is Nothing Then
        Err.Raise vbObjectError + 7520, "FillWeeklyReportCurrentTable", _
                  "PPT에서 업무 현황 표를 찾을 수 없습니다."
    End If

    If tableShape.Table.Rows.Count < 2 Or tableShape.Table.Columns.Count < 3 Then
        Err.Raise vbObjectError + 7521, "FillWeeklyReportCurrentTable", _
                  "PPT 업무 현황 표의 구조가 예상과 다릅니다."
    End If

    ' The template table already applies PowerPoint bullet formatting.
    Set table = tableShape.Table
    originalTableHeight = tableShape.Height
    Set taskTextRange = table.Cell(2, 2).Shape.TextFrame.TextRange
    Set dateTextRange = table.Cell(2, 3).Shape.TextFrame.TextRange

    ClearPowerPointTextRangeParagraphs taskTextRange
    ClearPowerPointTextRangeParagraphs dateTextRange
    EnsurePowerPointParagraphCount taskTextRange, items.Count
    EnsurePowerPointParagraphCount dateTextRange, items.Count

    For i = 1 To items.Count
        levelValue = CLng(levelItems(i))
        displayText = CStr(items(i))

        Set taskParagraph = taskTextRange.Paragraphs(i)
        If levelValue > 1 Then
            displayText = "    " & displayText
            taskParagraph.ParagraphFormat.Bullet.Visible = False
        Else
            taskParagraph.ParagraphFormat.Bullet.Visible = True
            taskParagraph.ParagraphFormat.Bullet.Type = 1
            taskParagraph.ParagraphFormat.Bullet.Character = &H2022
            taskParagraph.ParagraphFormat.Bullet.RelativeSize = 1
            taskParagraph.ParagraphFormat.Bullet.Font.Name = "Arial"
        End If

        SetPowerPointParagraphText taskParagraph, displayText
        dateDisplayText = CStr(dateItems(i))
        SetPowerPointParagraphText dateTextRange.Paragraphs(i), dateDisplayText

        baseLineHeight = dateTextRange.Paragraphs(i).BoundHeight
        If baseLineHeight > 0 Then
            taskLineCount = CLng((taskParagraph.BoundHeight / baseLineHeight) + 0.25)
        Else
            taskLineCount = 1
        End If

        For lineIndex = 2 To taskLineCount
            dateDisplayText = dateDisplayText & ChrW(11)
        Next lineIndex

        If taskLineCount > 1 Then
            SetPowerPointParagraphText dateTextRange.Paragraphs(i), dateDisplayText
        End If
    Next i

    TrimTrailingPowerPointParagraphs taskTextRange, items.Count
    TrimTrailingPowerPointParagraphs dateTextRange, items.Count

    ' PowerPoint expands tables automatically when text wraps.
    ' Restore the exact template height so the original layout never changes.
    tableShape.Height = originalTableHeight
End Sub

Private Sub EnsurePowerPointParagraphCount(ByVal textRange As Object, _
                                           ByVal requiredParagraphCount As Long)
    Do While textRange.Paragraphs.Count < requiredParagraphCount
        textRange.InsertAfter vbCr
    Loop
End Sub

Private Sub FillWeeklyReportPlanArea(ByVal slide As Object, _
                                     ByVal plannedItems As Collection, _
                                     ByVal allowHeightExpansion As Boolean)
    Dim planShape As Object
    Dim planTextRange As Object
    Dim i As Long
    Dim itemText As String
    Dim originalLeft As Single
    Dim originalTop As Single
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim originalLockAspectRatio As Long
    Dim requiredHeight As Single
    Dim templateFontName As String
    Dim templateFontSize As Single
    Dim templateFontBold As Long
    Dim templateFontItalic As Long
    Dim templateFontUnderline As Long
    Dim templateFontColor As Long
    Dim maintenanceParagraphIndex As Long

    Set planShape = FindTextShape(slide, "(개발 항목)")
    If planShape Is Nothing Then
        Err.Raise vbObjectError + 7530, "FillWeeklyReportPlanArea", _
                  "PPT에서 '(개발 항목)' 영역을 찾을 수 없습니다."
    End If

    Set planTextRange = planShape.TextFrame.TextRange
    If planTextRange.Paragraphs.Count < 2 Then
        Err.Raise vbObjectError + 7531, "FillWeeklyReportPlanArea", _
                  "원본 PPT 개발 항목 영역의 문단 구조가 예상과 다릅니다."
    End If

    With planTextRange.Paragraphs(2).Font
        templateFontName = .Name
        templateFontSize = .Size
        templateFontBold = .Bold
        templateFontItalic = .Italic
        templateFontUnderline = .Underline
        templateFontColor = .Color.RGB
    End With

    originalLeft = planShape.Left
    originalTop = planShape.Top
    originalWidth = planShape.Width
    originalHeight = planShape.Height
    originalLockAspectRatio = planShape.LockAspectRatio

    SetPowerPointParagraphText planTextRange.Paragraphs(2), ""
    SetPowerPointParagraphText planTextRange.Paragraphs(3), ""

    For i = 1 To plannedItems.Count
        itemText = ChrW(&H2022) & " " & CStr(plannedItems(i))

        If i <= 2 Then
            SetPowerPointParagraphText planTextRange.Paragraphs(i + 1), itemText
        Else
            maintenanceParagraphIndex = FindPowerPointParagraphIndex( _
                                            planTextRange, "(유지보수 항목)")
            If maintenanceParagraphIndex = 0 Then
                Err.Raise vbObjectError + 7532, "FillWeeklyReportPlanArea", _
                          "원본 PPT에서 '(유지보수 항목)' 영역을 찾을 수 없습니다."
            End If
            planTextRange.Paragraphs(maintenanceParagraphIndex).InsertBefore _
                itemText & vbCr
        End If

        ApplyWeeklyPlanParagraphFont planTextRange.Paragraphs(i + 1), _
                                     templateFontName, templateFontSize, _
                                     templateFontBold, templateFontItalic, _
                                     templateFontUnderline, templateFontColor
    Next i

    ' Preserve the template format and position. Only extend the bottom edge.
    planShape.TextFrame.AutoSize = 0
    requiredHeight = planTextRange.BoundHeight + _
                     planShape.TextFrame.MarginTop + _
                     planShape.TextFrame.MarginBottom

    planShape.LockAspectRatio = 0
    planShape.Left = originalLeft
    planShape.Top = originalTop
    planShape.Width = originalWidth
    If allowHeightExpansion And requiredHeight > originalHeight Then
        planShape.Height = requiredHeight
    Else
        planShape.Height = originalHeight
    End If
    planShape.LockAspectRatio = originalLockAspectRatio
End Sub

Private Function FindPowerPointParagraphIndex(ByVal textRange As Object, _
                                              ByVal searchText As String) As Long
    Dim i As Long

    For i = 1 To textRange.Paragraphs.Count
        If InStr(1, CStr(textRange.Paragraphs(i).Text), _
                   searchText, vbTextCompare) > 0 Then
            FindPowerPointParagraphIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub ApplyWeeklyPlanParagraphFont(ByVal paragraphRange As Object, _
                                         ByVal fontName As String, _
                                         ByVal fontSize As Single, _
                                         ByVal fontBold As Long, _
                                         ByVal fontItalic As Long, _
                                         ByVal fontUnderline As Long, _
                                         ByVal fontColor As Long)
    With paragraphRange.Font
        .Name = fontName
        .Size = fontSize
        .Bold = fontBold
        .Italic = fontItalic
        .Underline = fontUnderline
        .Color.RGB = fontColor
    End With
End Sub

Private Function JoinCollection(ByVal items As Collection, _
                                ByVal linePrefix As String, _
                                ByVal delimiter As String) As String
    Dim i As Long
    Dim result As String

    For i = 1 To items.Count
        If i > 1 Then result = result & delimiter
        result = result & linePrefix & CStr(items(i))
    Next i

    JoinCollection = result
End Function

Private Sub ClearPowerPointTextRangeParagraphs(ByVal textRange As Object)
    Dim i As Long

    For i = 1 To textRange.Paragraphs.Count
        SetPowerPointParagraphText textRange.Paragraphs(i), ""
    Next i
End Sub

Private Sub TrimTrailingPowerPointParagraphs(ByVal textRange As Object, _
                                             ByVal usedParagraphCount As Long)
    Dim keepCount As Long

    keepCount = usedParagraphCount
    If keepCount < 1 Then keepCount = 1

    Do While textRange.Paragraphs.Count > keepCount
        textRange.Paragraphs(textRange.Paragraphs.Count).Delete
    Loop
End Sub

Private Sub SetPowerPointParagraphText(ByVal paragraphRange As Object, _
                                       ByVal textValue As String)
    Dim contentLength As Long
    Dim paragraphText As String

    paragraphText = CStr(paragraphRange.Text)
    contentLength = Len(paragraphText)

    If contentLength > 0 Then
        If Right$(paragraphText, 1) = vbCr Then contentLength = contentLength - 1
    End If

    If contentLength > 0 Then paragraphRange.Characters(1, contentLength).Delete
    If Len(textValue) > 0 Then paragraphRange.InsertBefore textValue
End Sub

Private Sub ReplaceTextAfterMarkerPreservingStyle(ByVal textRange As Object, _
                                                  ByVal markerText As String, _
                                                  ByVal replacementText As String)
    Dim markerPosition As Long
    Dim valueStart As Long
    Dim oldLength As Long

    markerPosition = InStr(1, CStr(textRange.Text), markerText, vbTextCompare)
    If markerPosition = 0 Then
        Err.Raise vbObjectError + 7550, "ReplaceTextAfterMarkerPreservingStyle", _
                  "PPT 날짜 영역의 기준 문자를 찾을 수 없습니다."
    End If

    valueStart = markerPosition + Len(markerText)
    oldLength = Len(CStr(textRange.Text)) - valueStart + 1

    If oldLength > 0 Then
        textRange.Characters(valueStart, oldLength).Text = replacementText
    Else
        textRange.InsertAfter replacementText
    End If
End Sub

Private Function FindFirstTableShape(ByVal slide As Object) As Object
    Dim shape As Object

    For Each shape In slide.Shapes
        On Error Resume Next
        If shape.HasTable Then
            Set FindFirstTableShape = shape
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next shape
End Function

Private Function FindTextShape(ByVal slide As Object, ByVal searchText As String) As Object
    Dim shape As Object
    Dim textValue As String

    For Each shape In slide.Shapes
        textValue = ""

        On Error Resume Next
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then textValue = CStr(shape.TextFrame.TextRange.Text)
        End If
        On Error GoTo 0

        If InStr(1, textValue, searchText, vbTextCompare) > 0 Then
            Set FindTextShape = shape
            Exit Function
        End If
    Next shape
End Function

Private Sub ExtractEmbeddedWeeklyPptTemplate(ByVal targetPath As String)
    Dim templateWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim base64Text As String
    Dim xmlDocument As Object
    Dim base64Node As Object
    Dim stream As Object

    On Error Resume Next
    Set templateWs = ThisWorkbook.Worksheets(WEEKLY_PPT_TEMPLATE_SHEET_NAME)
    On Error GoTo 0

    If templateWs Is Nothing Then
        Err.Raise vbObjectError + 7540, "ExtractEmbeddedWeeklyPptTemplate", _
                  "엑셀 내부의 주간보고 PPT 템플릿 시트를 찾을 수 없습니다."
    End If

    If CStr(templateWs.Range("A1").Value2) <> WEEKLY_PPT_TEMPLATE_MARKER Then
        Err.Raise vbObjectError + 7541, "ExtractEmbeddedWeeklyPptTemplate", _
                  "엑셀 내부의 주간보고 PPT 템플릿 데이터가 올바르지 않습니다."
    End If

    lastRow = templateWs.Cells(templateWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        base64Text = base64Text & CStr(templateWs.Cells(r, 1).Value2)
    Next r

    If Len(base64Text) = 0 Then
        Err.Raise vbObjectError + 7542, "ExtractEmbeddedWeeklyPptTemplate", _
                  "엑셀 내부의 주간보고 PPT 템플릿 데이터가 비어 있습니다."
    End If

    Set xmlDocument = CreateObject("MSXML2.DOMDocument.6.0")
    Set base64Node = xmlDocument.createElement("base64")
    base64Node.DataType = "bin.base64"
    base64Node.Text = base64Text

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write base64Node.nodeTypedValue
    stream.SaveToFile targetPath, 2
    stream.Close
End Sub

Private Sub DeleteTemporaryWeeklyPptTemplate(ByVal templatePath As String)
    If Len(templatePath) = 0 Then Exit Sub
    If Len(Dir$(templatePath)) = 0 Then Exit Sub
    Kill templatePath
End Sub
