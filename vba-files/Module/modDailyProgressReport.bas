Attribute VB_Name = "modDailyProgressReport"
Option Explicit

Private Const DAILY_TEMPLATE_SHEET_NAME As String = "_일별진행현황템플릿"
Private Const DAILY_TITLE As String = "개발 진행 현황 관리표"
Private Const DAILY_HEADER_ROW As Long = 4
Private Const DAILY_DATA_START_ROW As Long = 5
Private Const DAILY_FIRST_DATE_COLUMN As Long = 8
Private Const DAILY_TEMPLATE_LAST_ROW As Long = 78
Private Const DAILY_FILE_FORMAT_XLSX As Long = 51

Public Sub 일별진행현황_생성버튼_생성(Optional ByVal showCompletionMessage As Boolean = True)
    Dim ws As Worksheet
    Dim button As Shape
    Dim anchorCell As Range
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet
    If Not IsDailyReportTaskSheet(ws) Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    UnprotectTaskSheet ws
    Set anchorCell = ws.Range("B2")

    On Error Resume Next
    ws.Shapes("btnDailyProgressReport").Delete
    On Error GoTo EH

    Set button = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        anchorCell.Left + (7 * 80), _
        anchorCell.Top + anchorCell.Height - 22, _
        72, _
        22)

    With button
        .Name = "btnDailyProgressReport"
        .OnAction = "일별진행현황_생성"
        .Placement = xlFreeFloating
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(212, 208, 200)
        .Fill.Transparency = 0
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
                .Characters.Text = "일별 현황"
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "맑은 고딕"
                .Font.Size = 9
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With
    End With

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
    ApplyCalculatedColumnsProtection ws, lastRow

    If showCompletionMessage Then
        MsgBox "버튼 생성 완료: 일별 현황", vbInformation
    End If
    Exit Sub

EH:
    MsgBox "일별 현황 버튼을 생성하는 중 오류가 발생했습니다: " & _
           Err.Description, vbExclamation
End Sub

Public Sub 일별진행현황_생성(Optional ByVal outputPath As String = "", _
                            Optional ByVal showCompletionMessage As Boolean = True)
    Dim sourceWs As Worksheet
    Dim templateWs As Worksheet
    Dim outputBook As Workbook
    Dim groupRows As Object
    Dim groupValues As Object
    Dim groupOrder As Collection
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim firstReportDate As Date
    Dim lastReportDate As Date
    Dim monthCursor As Date
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim savePath As Variant
    Dim lastRow As Long
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim stateCaptured As Boolean
    Dim errorText As String

    On Error GoTo EH

    Set sourceWs = ActiveSheet
    If Not IsDailyReportTaskSheet(sourceWs) Then
        If showCompletionMessage Then
            MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Else
            Err.Raise vbObjectError + 2101, , "업무 시트에서 실행하세요."
        End If
        Exit Sub
    End If

    Set templateWs = GetDailyTemplateSheet()
    If templateWs Is Nothing Then
        If showCompletionMessage Then
            MsgBox "통합문서에 일별 진행 현황 템플릿이 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2102, , "통합문서에 일별 진행 현황 템플릿이 없습니다."
        End If
        Exit Sub
    End If

    lastRow = GetLastDataRow(sourceWs)
    If lastRow < DATA_START_ROW Then
        If showCompletionMessage Then
            MsgBox "출력할 업무 데이터가 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2103, , "출력할 업무 데이터가 없습니다."
        End If
        Exit Sub
    End If

    Set groupRows = CreateObject("Scripting.Dictionary")
    Set groupValues = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection

    BuildDailyReportGroups sourceWs, lastRow, groupRows, groupValues, groupOrder
    If groupOrder.Count = 0 Then
        If showCompletionMessage Then
            MsgBox "출력할 Level 1 업무가 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2104, , "출력할 Level 1 업무가 없습니다."
        End If
        Exit Sub
    End If

    If Not GetDailyReportDateRange(sourceWs, lastRow, firstReportDate, lastReportDate) Then
        If showCompletionMessage Then
            MsgBox "실제 시작일이 입력된 업무가 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2105, , "실제 시작일이 입력된 업무가 없습니다."
        End If
        Exit Sub
    End If

    If Len(Trim$(outputPath)) > 0 Then
        savePath = outputPath
    Else
        savePath = Application.GetSaveAsFilename( _
            InitialFileName:=Format$(Date, "yyyymmdd") & "_TGS_개발_일별_진행현황표.xlsx", _
            FileFilter:="Excel 통합 문서 (*.xlsx), *.xlsx", _
            Title:="일별 진행 현황 저장")
        If VarType(savePath) = vbBoolean Then Exit Sub
    End If

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    stateCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    EnsureConfigSheet
    LoadHolidaySettings holidayDict, workdayDict

    monthCursor = DateSerial(Year(firstReportDate), Month(firstReportDate), 1)
    Do While monthCursor <= DateSerial(Year(lastReportDate), Month(lastReportDate), 1)
        monthStart = monthCursor
        If firstReportDate > monthStart Then monthStart = firstReportDate

        monthEnd = DateSerial(Year(monthCursor), Month(monthCursor) + 1, 0)
        If lastReportDate < monthEnd Then monthEnd = lastReportDate

        If HasWorkingDate(monthStart, monthEnd, holidayDict, workdayDict) Then
            AddDailyReportMonthSheet templateWs, outputBook, sourceWs, _
                groupRows, groupValues, groupOrder, monthStart, monthEnd, _
                holidayDict, workdayDict
        End If

        monthCursor = DateAdd("m", 1, monthCursor)
    Loop

    If outputBook Is Nothing Then
        If showCompletionMessage Then
            MsgBox "출력 기간에 근무일이 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2106, , "출력 기간에 근무일이 없습니다."
        End If
        GoTo SafeExit
    End If

    outputBook.SaveAs Filename:=CStr(savePath), FileFormat:=DAILY_FILE_FORMAT_XLSX
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing

    RestoreDailyReportApplicationState previousScreenUpdating, previousEnableEvents
    stateCaptured = False

    If showCompletionMessage Then
        MsgBox "일별 진행 현황을 생성했습니다." & vbCrLf & CStr(savePath), vbInformation
    End If
    Exit Sub

SafeExit:
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    On Error GoTo 0
    If stateCaptured Then
        RestoreDailyReportApplicationState previousScreenUpdating, previousEnableEvents
    End If
    Exit Sub

EH:
    errorText = Err.Description
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    On Error GoTo 0
    If stateCaptured Then
        RestoreDailyReportApplicationState previousScreenUpdating, previousEnableEvents
    End If
    If showCompletionMessage Then
        MsgBox "일별 진행 현황을 생성하는 중 오류가 발생했습니다." & vbCrLf & _
               "원인: " & errorText, vbExclamation
    Else
        Err.Raise vbObjectError + 2100, "일별진행현황_생성", errorText
    End If
End Sub

Private Sub BuildDailyReportGroups(ByVal ws As Worksheet, _
                                   ByVal lastRow As Long, _
                                   ByVal groupRows As Object, _
                                   ByVal groupValues As Object, _
                                   ByVal groupOrder As Collection)
    Dim rowNum As Long
    Dim taskLevel As Long
    Dim currentKey As String
    Dim groupKey As String
    Dim ownerText As String
    Dim values As Variant
    Dim rows As Collection

    For rowNum = DATA_START_ROW To lastRow
        taskLevel = GetTaskLevel(ws, rowNum)

        If taskLevel = 1 Then
            ownerText = Trim$(CStr(ws.Cells(rowNum, COL_OWNER).Value2))
            groupKey = BuildDailyGroupKey( _
                CStr(ws.Cells(rowNum, COL_TYPE).Value2), _
                CStr(ws.Cells(rowNum, COL_MAJOR_CATEGORY).Value2), _
                CStr(ws.Cells(rowNum, COL_MIDDLE_CATEGORY).Value2), _
                CStr(ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2), _
                ownerText)
            currentKey = groupKey

            If Not groupRows.Exists(groupKey) Then
                Set rows = New Collection
                groupRows.Add groupKey, rows
                values = Array( _
                    Trim$(CStr(ws.Cells(rowNum, COL_TYPE).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MAJOR_CATEGORY).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MIDDLE_CATEGORY).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2)), _
                    ownerText)
                groupValues.Add groupKey, values
                groupOrder.Add groupKey
            End If
        ElseIf Len(currentKey) = 0 Then
            ownerText = Trim$(CStr(ws.Cells(rowNum, COL_OWNER).Value2))
            groupKey = BuildDailyGroupKey( _
                CStr(ws.Cells(rowNum, COL_TYPE).Value2), _
                CStr(ws.Cells(rowNum, COL_MAJOR_CATEGORY).Value2), _
                CStr(ws.Cells(rowNum, COL_MIDDLE_CATEGORY).Value2), _
                CStr(ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2), _
                ownerText)
            currentKey = groupKey

            If Not groupRows.Exists(groupKey) Then
                Set rows = New Collection
                groupRows.Add groupKey, rows
                values = Array( _
                    Trim$(CStr(ws.Cells(rowNum, COL_TYPE).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MAJOR_CATEGORY).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MIDDLE_CATEGORY).Value2)), _
                    Trim$(CStr(ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2)), _
                    ownerText)
                groupValues.Add groupKey, values
                groupOrder.Add groupKey
            End If
        End If

        If Len(currentKey) > 0 And Len(Trim$(CStr(ws.Cells(rowNum, COL_TASK).Value2))) > 0 Then
            Set rows = groupRows(currentKey)
            rows.Add rowNum
        End If
    Next rowNum
End Sub

Private Function BuildDailyGroupKey(ByVal typeText As String, _
                                    ByVal majorText As String, _
                                    ByVal middleText As String, _
                                    ByVal minorText As String, _
                                    ByVal ownerText As String) As String
    BuildDailyGroupKey = Trim$(typeText) & ChrW(30) & _
                         Trim$(majorText) & ChrW(30) & _
                         Trim$(middleText) & ChrW(30) & _
                         Trim$(minorText) & ChrW(30) & _
                         Trim$(ownerText)
End Function

Private Function GetDailyReportDateRange(ByVal ws As Worksheet, _
                                         ByVal lastRow As Long, _
                                         ByRef firstDate As Date, _
                                         ByRef lastDate As Date) As Boolean
    Dim rowNum As Long
    Dim actualStart As Variant
    Dim actualEnd As Date

    For rowNum = DATA_START_ROW To lastRow
        actualStart = ws.Cells(rowNum, COL_ACTUAL_START).Value
        If IsDate(actualStart) Then
            If Not GetDailyReportDateRange Then
                firstDate = DateValue(CDate(actualStart))
                lastDate = firstDate
                GetDailyReportDateRange = True
            ElseIf DateValue(CDate(actualStart)) < firstDate Then
                firstDate = DateValue(CDate(actualStart))
            End If

            actualEnd = GetDailyTaskEndDate(ws, rowNum, DateValue(CDate(actualStart)))
            If actualEnd > lastDate Then lastDate = actualEnd
        End If
    Next rowNum
End Function

Private Function GetDailyTaskEndDate(ByVal ws As Worksheet, _
                                     ByVal rowNum As Long, _
                                     ByVal actualStart As Date) As Date
    Dim actualEnd As Variant

    actualEnd = ws.Cells(rowNum, COL_ACTUAL_END).Value
    If IsDate(actualEnd) Then
        GetDailyTaskEndDate = DateValue(CDate(actualEnd))
    ElseIf actualStart > Date Then
        GetDailyTaskEndDate = actualStart
    Else
        GetDailyTaskEndDate = Date
    End If

    If GetDailyTaskEndDate < actualStart Then
        GetDailyTaskEndDate = actualStart
    End If
End Function

Private Function HasWorkingDate(ByVal startDate As Date, _
                                ByVal endDate As Date, _
                                ByVal holidayDict As Object, _
                                ByVal workdayDict As Object) As Boolean
    Dim targetDate As Date

    For targetDate = startDate To endDate
        If IsWorkingDay(targetDate, holidayDict, workdayDict) Then
            HasWorkingDate = True
            Exit Function
        End If
    Next targetDate
End Function

Private Sub AddDailyReportMonthSheet(ByVal templateWs As Worksheet, _
                                     ByRef outputBook As Workbook, _
                                     ByVal sourceWs As Worksheet, _
                                     ByVal groupRows As Object, _
                                     ByVal groupValues As Object, _
                                     ByVal groupOrder As Collection, _
                                     ByVal monthStart As Date, _
                                     ByVal monthEnd As Date, _
                                     ByVal holidayDict As Object, _
                                     ByVal workdayDict As Object)
    Dim outputWs As Worksheet
    Dim originalVisibility As XlSheetVisibility
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CopyFailed

    originalVisibility = templateWs.Visible
    templateWs.Visible = xlSheetVisible

    If outputBook Is Nothing Then
        templateWs.Copy
        Set outputBook = ActiveWorkbook
        Set outputWs = outputBook.Worksheets(1)
    Else
        templateWs.Copy After:=outputBook.Worksheets(outputBook.Worksheets.Count)
        Set outputWs = outputBook.Worksheets(outputBook.Worksheets.Count)
    End If

    templateWs.Visible = originalVisibility

    outputWs.Visible = xlSheetVisible
    outputWs.Name = GetDailyMonthSheetName(outputBook, monthStart)
    PopulateDailyMonthSheet outputWs, sourceWs, groupRows, groupValues, _
        groupOrder, monthStart, monthEnd, holidayDict, workdayDict
    Exit Sub

CopyFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    templateWs.Visible = originalVisibility
    On Error GoTo 0
    Err.Raise errorNumber, "AddDailyReportMonthSheet", errorDescription
End Sub

Private Sub PopulateDailyMonthSheet(ByVal outputWs As Worksheet, _
                                    ByVal sourceWs As Worksheet, _
                                    ByVal groupRows As Object, _
                                    ByVal groupValues As Object, _
                                    ByVal groupOrder As Collection, _
                                    ByVal monthStart As Date, _
                                    ByVal monthEnd As Date, _
                                    ByVal holidayDict As Object, _
                                    ByVal workdayDict As Object)
    Dim groupCount As Long
    Dim outputLastRow As Long
    Dim clearLastRow As Long
    Dim lastUsedColumn As Long
    Dim dateColumn As Long
    Dim targetRow As Long
    Dim groupIndex As Long
    Dim targetDate As Date
    Dim groupKey As String
    Dim values As Variant
    Dim rows As Collection
    Dim shapeIndex As Long

    groupCount = groupOrder.Count
    outputLastRow = DAILY_DATA_START_ROW + groupCount - 1
    clearLastRow = outputWs.UsedRange.Row + outputWs.UsedRange.Rows.Count - 1
    If clearLastRow < DAILY_TEMPLATE_LAST_ROW Then clearLastRow = DAILY_TEMPLATE_LAST_ROW
    If clearLastRow < outputLastRow Then clearLastRow = outputLastRow

    lastUsedColumn = outputWs.UsedRange.Column + outputWs.UsedRange.Columns.Count - 1
    If lastUsedColumn < DAILY_FIRST_DATE_COLUMN Then lastUsedColumn = DAILY_FIRST_DATE_COLUMN

    On Error Resume Next
    outputWs.Range(outputWs.Cells(DAILY_DATA_START_ROW, 2), _
                   outputWs.Cells(clearLastRow, lastUsedColumn)).UnMerge
    On Error GoTo 0

    outputWs.Range(outputWs.Cells(DAILY_DATA_START_ROW, 2), _
                   outputWs.Cells(clearLastRow, lastUsedColumn)).ClearContents

    For shapeIndex = outputWs.Shapes.Count To 1 Step -1
        outputWs.Shapes(shapeIndex).Delete
    Next shapeIndex

    PrepareDailyDataRowFormats outputWs, outputLastRow

    outputWs.Range("B2").Value = DAILY_TITLE
    outputWs.Range("B4").Value = "Type"
    outputWs.Range("C4").Value = "대분류"
    outputWs.Range("D4").Value = "중분류"
    outputWs.Range("E4").Value = "메뉴명"
    outputWs.Range("F4").Value = "개발자_정"
    outputWs.Range("G4").Value = "개발자_부"

    For groupIndex = 1 To groupCount
        targetRow = DAILY_DATA_START_ROW + groupIndex - 1
        groupKey = CStr(groupOrder(groupIndex))
        values = groupValues(groupKey)

        outputWs.Cells(targetRow, 2).Value = values(0)
        outputWs.Cells(targetRow, 3).Value = values(1)
        outputWs.Cells(targetRow, 4).Value = values(2)
        outputWs.Cells(targetRow, 5).Value = values(3)
        outputWs.Cells(targetRow, 6).Value = values(4)
        outputWs.Cells(targetRow, 7).ClearContents
    Next groupIndex

    dateColumn = DAILY_FIRST_DATE_COLUMN
    For targetDate = monthStart To monthEnd
        If IsWorkingDay(targetDate, holidayDict, workdayDict) Then
            PrepareDailyDateColumn outputWs, dateColumn, clearLastRow
            outputWs.Cells(DAILY_HEADER_ROW, dateColumn).Value = targetDate
            outputWs.Cells(DAILY_HEADER_ROW, dateColumn).NumberFormatLocal = "m/d"

            For groupIndex = 1 To groupCount
                targetRow = DAILY_DATA_START_ROW + groupIndex - 1
                groupKey = CStr(groupOrder(groupIndex))
                Set rows = groupRows(groupKey)
                values = groupValues(groupKey)
                outputWs.Cells(targetRow, dateColumn).Value = _
                    BuildDailyCellText(sourceWs, rows, targetDate, CStr(values(4)))
            Next groupIndex

            dateColumn = dateColumn + 1
        End If
    Next targetDate

    If dateColumn <= lastUsedColumn Then
        outputWs.Range(outputWs.Cells(DAILY_HEADER_ROW, dateColumn), _
                       outputWs.Cells(clearLastRow, lastUsedColumn)).ClearContents
    End If

    MergeRepeatedDailyCategories outputWs, outputLastRow
    FormatDailyOutputArea outputWs, outputLastRow, dateColumn - 1
End Sub

Private Sub PrepareDailyDataRowFormats(ByVal ws As Worksheet, ByVal outputLastRow As Long)
    Dim rowNum As Long

    If outputLastRow <= DAILY_TEMPLATE_LAST_ROW Then Exit Sub

    For rowNum = DAILY_TEMPLATE_LAST_ROW + 1 To outputLastRow
        ws.Range(ws.Cells(DAILY_TEMPLATE_LAST_ROW, 2), _
                 ws.Cells(DAILY_TEMPLATE_LAST_ROW, 8)).Copy
        ws.Cells(rowNum, 2).PasteSpecial Paste:=xlPasteFormats
        ws.Rows(rowNum).RowHeight = ws.Rows(DAILY_TEMPLATE_LAST_ROW).RowHeight
    Next rowNum
    Application.CutCopyMode = False
End Sub

Private Sub PrepareDailyDateColumn(ByVal ws As Worksheet, _
                                   ByVal dateColumn As Long, _
                                   ByVal lastRow As Long)
    If dateColumn <> DAILY_FIRST_DATE_COLUMN Then
        ws.Range(ws.Cells(DAILY_HEADER_ROW, DAILY_FIRST_DATE_COLUMN), _
                 ws.Cells(lastRow, DAILY_FIRST_DATE_COLUMN)).Copy
        ws.Cells(DAILY_HEADER_ROW, dateColumn).PasteSpecial Paste:=xlPasteFormats
        ws.Columns(dateColumn).ColumnWidth = ws.Columns(DAILY_FIRST_DATE_COLUMN).ColumnWidth
        Application.CutCopyMode = False
    End If
End Sub

Private Function BuildDailyCellText(ByVal ws As Worksheet, _
                                    ByVal rows As Collection, _
                                    ByVal targetDate As Date, _
                                    ByVal ownerText As String) As String
    Dim rowNum As Variant
    Dim taskLevel As Long
    Dim taskText As String
    Dim lineText As String
    Dim bodyText As String
    Dim level1Number As Long

    For Each rowNum In rows
        If IsDailyTaskOnDate(ws, CLng(rowNum), targetDate) Then
            taskText = Trim$(CStr(ws.Cells(CLng(rowNum), COL_TASK).Value2))
            taskLevel = GetTaskLevel(ws, CLng(rowNum))

            Select Case taskLevel
                Case 1
                    level1Number = level1Number + 1
                    lineText = CStr(level1Number) & ". " & taskText
                Case 2
                    lineText = " - " & taskText
                Case Else
                    lineText = "  ㄴ " & taskText
            End Select

            If Len(bodyText) > 0 Then bodyText = bodyText & vbLf
            bodyText = bodyText & lineText
        End If
    Next rowNum

    If Len(bodyText) > 0 Then
        If Len(ownerText) > 0 Then
            BuildDailyCellText = "(" & ownerText & ")" & vbLf & bodyText
        Else
            BuildDailyCellText = bodyText
        End If
    End If
End Function

Private Function IsDailyTaskOnDate(ByVal ws As Worksheet, _
                                   ByVal rowNum As Long, _
                                   ByVal targetDate As Date) As Boolean
    Dim actualStart As Variant
    Dim actualEnd As Date

    actualStart = ws.Cells(rowNum, COL_ACTUAL_START).Value
    If Not IsDate(actualStart) Then Exit Function

    actualEnd = GetDailyTaskEndDate(ws, rowNum, DateValue(CDate(actualStart)))
    IsDailyTaskOnDate = _
        (targetDate >= DateValue(CDate(actualStart)) And targetDate <= actualEnd)
End Function

Private Sub MergeRepeatedDailyCategories(ByVal ws As Worksheet, ByVal lastRow As Long)
    MergeDailyCategoryColumn ws, 4, lastRow, 4
    MergeDailyCategoryColumn ws, 3, lastRow, 3
    MergeDailyCategoryColumn ws, 2, lastRow, 2
End Sub

Private Sub MergeDailyCategoryColumn(ByVal ws As Worksheet, _
                                     ByVal categoryColumn As Long, _
                                     ByVal lastRow As Long, _
                                     ByVal compareThroughColumn As Long)
    Dim blockStart As Long
    Dim rowNum As Long

    If lastRow < DAILY_DATA_START_ROW Then Exit Sub

    blockStart = DAILY_DATA_START_ROW
    For rowNum = DAILY_DATA_START_ROW + 1 To lastRow + 1
        If rowNum > lastRow Or _
           Not DailyCategoryPathMatches(ws, blockStart, rowNum, compareThroughColumn) Then
            If rowNum - blockStart > 1 Then
                Application.DisplayAlerts = False
                ws.Range(ws.Cells(blockStart, categoryColumn), _
                         ws.Cells(rowNum - 1, categoryColumn)).Merge
                Application.DisplayAlerts = True
            End If
            blockStart = rowNum
        End If
    Next rowNum
End Sub

Private Function DailyCategoryPathMatches(ByVal ws As Worksheet, _
                                          ByVal firstRow As Long, _
                                          ByVal secondRow As Long, _
                                          ByVal compareThroughColumn As Long) As Boolean
    Dim columnNum As Long

    For columnNum = 2 To compareThroughColumn
        If StrComp(Trim$(CStr(ws.Cells(firstRow, columnNum).Value2)), _
                   Trim$(CStr(ws.Cells(secondRow, columnNum).Value2)), _
                   vbTextCompare) <> 0 Then Exit Function
    Next columnNum

    DailyCategoryPathMatches = True
End Function

Private Sub FormatDailyOutputArea(ByVal ws As Worksheet, _
                                  ByVal lastRow As Long, _
                                  ByVal lastColumn As Long)
    Dim dataRange As Range

    Set dataRange = ws.Range(ws.Cells(DAILY_HEADER_ROW, 2), ws.Cells(lastRow, lastColumn))
    With dataRange
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(DAILY_DATA_START_ROW, DAILY_FIRST_DATE_COLUMN), _
                  ws.Cells(lastRow, lastColumn))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    ws.Range(ws.Cells(DAILY_DATA_START_ROW, 2), _
             ws.Cells(lastRow, 7)).VerticalAlignment = xlCenter

    ws.Rows(DAILY_DATA_START_ROW & ":" & lastRow).AutoFit
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(DAILY_HEADER_ROW, 2), ws.Cells(lastRow, lastColumn)).AutoFilter

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, lastColumn)).Address
    End With
End Sub

Private Function GetDailyMonthSheetName(ByVal outputBook As Workbook, _
                                        ByVal monthDate As Date) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long

    baseName = "개발일별진행현황(" & Format$(monthDate, "mm") & "월)"
    candidate = baseName
    suffix = 2

    Do While DailyWorksheetExists(outputBook, candidate)
        candidate = Left$(baseName, 27) & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop

    GetDailyMonthSheetName = candidate
End Function

Private Function DailyWorksheetExists(ByVal targetBook As Workbook, _
                                      ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = targetBook.Worksheets(sheetName)
    DailyWorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetDailyTemplateSheet() As Worksheet
    On Error Resume Next
    Set GetDailyTemplateSheet = ThisWorkbook.Worksheets(DAILY_TEMPLATE_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function IsDailyReportTaskSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    IsDailyReportTaskSheet = _
        (StrComp(ws.Name, CONFIG_SHEET_NAME, vbTextCompare) <> 0 And _
         StrComp(ws.Name, "WeeklyPptTemplate", vbTextCompare) <> 0 And _
         StrComp(ws.Name, DAILY_TEMPLATE_SHEET_NAME, vbTextCompare) <> 0 And _
         StrComp(ws.Name, "_버튼생성", vbTextCompare) <> 0 And _
         StrComp(ws.Name, WEEKLY_REPORT_CONFIG_SHEET_NAME, vbTextCompare) <> 0)
End Function

Private Sub RestoreDailyReportApplicationState(ByVal screenUpdating As Boolean, _
                                               ByVal enableEvents As Boolean)
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
End Sub
