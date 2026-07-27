Attribute VB_Name = "modWeeklyPptReport"
Option Explicit

Private Const WEEKLY_PPT_TEMPLATE_SHEET_NAME As String = "WeeklyPptTemplate"
Private Const WEEKLY_PPT_TEMPLATE_MARKER As String = "REO_WEEKLY_PPT_TEMPLATE_BASE64_V1"
Private Const WEEKLY_PPT_OUTPUT_FOLDER As String = "주간보고"
Private Const WEEKLY_PPT_FILE_PREFIX As String = "Digital MFG팀_주간보고_김좌현_"
Private Const PPT_SAVE_AS_OPEN_XML_PRESENTATION As Long = 24

Public Sub 주간보고PPT_생성()
    Call GenerateWeeklyPptReport(True)
End Sub

Public Function WeeklyPptReportSelfTest() As String
    Dim ws As Worksheet
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim inProgressCount As Long
    Dim completedCount As Long
    Dim plannedCount As Long
    Dim statusText As String

    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws)
    LoadHolidaySettings holidayDict, workdayDict
    UpdateDevelopmentProgressStatuses ws, lastRow, holidayDict, workdayDict

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If Not HasChildTask(ws, r, lastRow) Then
                statusText = UCase$(Trim$(CStr(ws.Cells(r, COL_DEV_PROGRESS).Value2)))

                Select Case statusText
                    Case UCase$(REPORT_STATUS_IN_PROGRESS)
                        inProgressCount = inProgressCount + 1
                    Case UCase$(REPORT_STATUS_COMPLETED)
                        completedCount = completedCount + 1
                    Case UCase$(REPORT_STATUS_PLANNED)
                        plannedCount = plannedCount + 1
                End Select
            End If
        End If
    Next r

    WeeklyPptReportSelfTest = _
        "In Progress=" & CStr(inProgressCount) & _
        "; Completed=" & CStr(completedCount) & _
        "; Planned=" & CStr(plannedCount)
End Function

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
        anchorCell.Left + (6 * 80), _
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
    Dim pptApp As Object
    Dim presentation As Object
    Dim slide As Object
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

    lastRow = GetLastDataRow(ws)
    LoadHolidaySettings holidayDict, workdayDict
    UpdateDevelopmentProgressStatuses ws, lastRow, holidayDict, workdayDict

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            statusText = Trim$(CStr(ws.Cells(r, COL_DEV_PROGRESS).Value2))
            taskText = CleanWeeklyReportTaskText(CStr(ws.Cells(r, COL_TASK).Value2))

            If Len(taskText) > 0 Then
                Select Case UCase$(statusText)
                    Case UCase$(REPORT_STATUS_IN_PROGRESS)
                        currentItems.Add taskText
                        currentDates.Add BuildInProgressEndDateText(ws.Cells(r, COL_PLAN_END).Value)
                        currentLevels.Add GetTaskLevel(ws, r)

                    Case UCase$(REPORT_STATUS_COMPLETED)
                        If IsCompletedInReportWeek(ws, r, currentWeekStart, currentWeekEnd) Then
                            currentItems.Add taskText
                            currentDates.Add BuildCompletedEndDateText(ws, r)
                            currentLevels.Add GetTaskLevel(ws, r)
                        End If

                    Case UCase$(REPORT_STATUS_PLANNED)
                        If IsPlannedForNextWeek(ws, r, nextWeekStart, nextWeekEnd) Then
                            plannedItems.Add taskText
                        End If
                End Select
            End If
        End If
    Next r

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
    Set slide = presentation.Slides(1)

    FillWeeklyReportPeriodText slide, currentWeekStart, currentWeekEnd, nextWeekStart, nextWeekEnd
    FillWeeklyReportCurrentTable slide, currentItems, currentDates, currentLevels
    FillWeeklyReportPlanArea slide, plannedItems

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
               "업무 현황: " & currentItems.Count & "개" & vbCrLf & _
               "개발 계획: " & plannedItems.Count & "개" & vbCrLf & vbCrLf & _
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
    Dim taskSlotCount As Long
    Dim dateSlotCount As Long
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

    taskSlotCount = taskTextRange.Paragraphs.Count
    dateSlotCount = dateTextRange.Paragraphs.Count
    If taskSlotCount > dateSlotCount Then taskSlotCount = dateSlotCount

    If items.Count > taskSlotCount Then
        Err.Raise vbObjectError + 7522, "FillWeeklyReportCurrentTable", _
                  "원본 PPT 업무 현황 영역의 최대 항목 수(" & _
                  CStr(taskSlotCount) & "개)를 초과했습니다."
    End If

    ClearPowerPointTextRangeParagraphs taskTextRange
    ClearPowerPointTextRangeParagraphs dateTextRange

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

Private Sub FillWeeklyReportPlanArea(ByVal slide As Object, ByVal plannedItems As Collection)
    Dim planShape As Object
    Dim planTextRange As Object
    Dim i As Long

    Set planShape = FindTextShape(slide, "(개발 항목)")
    If planShape Is Nothing Then
        Err.Raise vbObjectError + 7530, "FillWeeklyReportPlanArea", _
                  "PPT에서 '(개발 항목)' 영역을 찾을 수 없습니다."
    End If

    Set planTextRange = planShape.TextFrame.TextRange
    If planTextRange.Paragraphs.Count < 3 Then
        Err.Raise vbObjectError + 7531, "FillWeeklyReportPlanArea", _
                  "원본 PPT 개발 항목 영역의 문단 구조가 예상과 다릅니다."
    End If

    SetPowerPointParagraphText planTextRange.Paragraphs(2), ""
    SetPowerPointParagraphText planTextRange.Paragraphs(3), ""

    For i = 1 To plannedItems.Count
        If i > 2 Then
            Err.Raise vbObjectError + 7532, "FillWeeklyReportPlanArea", _
                      "원본 PPT 개발 항목 영역에는 최대 2개 업무를 표시할 수 있습니다."
        End If
        SetPowerPointParagraphText planTextRange.Paragraphs(i + 1), _
                                   ChrW(&H2022) & " " & CStr(plannedItems(i))
    Next i
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
