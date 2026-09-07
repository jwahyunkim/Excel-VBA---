Attribute VB_Name = "modDailyProgressReport"
Option Explicit

Private Const DAILY_TEMPLATE_SHEET_NAME As String = "_일별진행현황템플릿"
Private Const DAILY_TITLE As String = "개발 진행 현황 관리표"
Private Const DAILY_HEADER_ROW As Long = 4
Private Const DAILY_DATA_START_ROW As Long = 5
Private Const DAILY_FIRST_DATE_COLUMN As Long = 8
Private Const DAILY_TEMPLATE_LAST_ROW As Long = 78
Private Const DAILY_FILE_FORMAT_XLSX As Long = 51
Private Const DAILY_HISTORY_SHEET_NAME As String = "_일별진척률이력"
Private Const DAILY_HISTORY_HEADER_ROW As Long = 1
Private Const DAILY_HISTORY_DATA_START_ROW As Long = 2
Private Const DAILY_HISTORY_SOURCE_SHEET_COLUMN As Long = 11
Private Const DAILY_HISTORY_KEY_SEPARATOR As Long = 29

Private mDailyProgressHistory As Object
Private mDailyProgressCache As Object

Public Sub 일별진행현황_생성버튼_생성()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet
    If Not IsDailyReportTaskSheet(ws) Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    UnprotectTaskSheet ws
    CreateVersionButton ws, "btnDailyProgressReport", "일별 현황", _
                            "일별진행현황_생성", GetNextVersionButtonOrder(ws), 72

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
    ApplyCalculatedColumnsProtection ws, lastRow

    MsgBox "버튼 생성 완료: 일별 현황", vbInformation
    Exit Sub

EH:
    MsgBox "일별 현황 버튼을 생성하는 중 오류가 발생했습니다: " & _
           Err.Description, vbExclamation
End Sub

Public Sub 일별진척률_이력초기화버튼_생성(Optional ByVal showCompletionMessage As Boolean = True)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim buttonOrder As Long
    Dim existingButton As Shape

    On Error GoTo EH
    Set ws = GetDailyReportSourceSheet()
    If ws Is Nothing Then Err.Raise vbObjectError + 2130, , "업무 시트에서 실행하세요."

    buttonOrder = GetNextVersionButtonOrder(ws)
    On Error Resume Next
    Set existingButton = ws.Shapes("btnDailyProgressHistoryReset")
    On Error GoTo EH
    If Not existingButton Is Nothing Then
        buttonOrder = CLng((existingButton.Left - ws.Range("B2").Left) / 80) + 1
    End If

    UnprotectTaskSheet ws
    CreateVersionButton ws, "btnDailyProgressHistoryReset", "이력 초기화", _
                        "일별진척률_이력초기화", buttonOrder, 72
    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
    ApplyCalculatedColumnsProtection ws, lastRow
    If showCompletionMessage Then MsgBox "버튼 생성 완료: 이력 초기화", vbInformation
    Exit Sub
EH:
    If showCompletionMessage Then
        MsgBox "이력 초기화 버튼 생성 오류: " & Err.Description, vbExclamation
    Else
        Err.Raise Err.Number, "일별진척률_이력초기화버튼_생성", Err.Description
    End If
End Sub

Public Sub 일별진척률_이력초기화(Optional ByVal interactive As Boolean = True)
    Dim historyWs As Worksheet
    Dim lastRow As Long
    Dim previousEnableEvents As Boolean
    Dim errorText As String

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    If ThisWorkbook.ReadOnly Then
        Err.Raise vbObjectError + 2131, , "이력을 초기화하려면 통합문서를 읽기/쓰기 상태로 여세요."
    End If
    On Error Resume Next
    Set historyWs = ThisWorkbook.Worksheets(DAILY_HISTORY_SHEET_NAME)
    On Error GoTo EH
    If historyWs Is Nothing Then
        Set mDailyProgressHistory = Nothing
        Set mDailyProgressCache = Nothing
        If interactive Then MsgBox "초기화할 일별 진척률 이력이 없습니다.", vbInformation
        Exit Sub
    End If

    If interactive Then
        If MsgBox("이 통합문서의 모든 일별 진척률 이력을 초기화하고 저장하시겠습니까?" & vbCrLf & _
                  "업무 데이터와 간트 진행률은 유지됩니다.", _
                  vbQuestion Or vbYesNo Or vbDefaultButton2, "일별 진척률 이력 초기화") <> vbYes Then Exit Sub
    End If

    Application.EnableEvents = False
    lastRow = historyWs.UsedRange.Row + historyWs.UsedRange.Rows.Count - 1
    If lastRow >= DAILY_HISTORY_DATA_START_ROW Then
        historyWs.Range(historyWs.Cells(DAILY_HISTORY_DATA_START_ROW, 1), _
                        historyWs.Cells(lastRow, DAILY_HISTORY_SOURCE_SHEET_COLUMN)).ClearContents
    End If
    Set mDailyProgressHistory = Nothing
    Set mDailyProgressCache = Nothing
    ThisWorkbook.Save
    Application.EnableEvents = previousEnableEvents
    If interactive Then MsgBox "일별 진척률 이력을 초기화했습니다.", vbInformation
    Exit Sub
EH:
    errorText = Err.Description
    Application.EnableEvents = previousEnableEvents
    If interactive Then
        MsgBox "일별 진척률 이력 초기화 오류: " & errorText, vbExclamation
    Else
        Err.Raise vbObjectError + 2132, "일별진척률_이력초기화", errorText
    End If
End Sub

Public Sub 일별진행현황_생성(Optional ByVal outputPath As String = "", _
                            Optional ByVal showCompletionMessage As Boolean = True)
    Dim sourceWs As Worksheet
    Dim templateWs As Worksheet
    Dim outputBook As Workbook
    Dim groupRows As Object
    Dim groupOrder As Collection
    Dim activeGroupOrder As Collection
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
    Dim displayStartDate As Date
    Dim displayEndDate As Date
    Dim hasDisplayDateRange As Boolean
    Dim errorStage As String

    On Error GoTo EH

    errorStage = "원본 업무 시트 확인"
    Set sourceWs = GetDailyReportSourceSheet()
    If Not IsDailyReportTaskSheet(sourceWs) Then
        If showCompletionMessage Then
            MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Else
            Err.Raise vbObjectError + 2101, , "업무 시트에서 실행하세요."
        End If
        Exit Sub
    End If

    If ThisWorkbook.ReadOnly Then
        If showCompletionMessage Then
            MsgBox "일별 진척률 이력을 저장하려면 통합문서를 읽기/쓰기 상태로 여세요.", _
                   vbExclamation
        Else
            Err.Raise vbObjectError + 2107, , _
                      "통합문서가 읽기 전용이어서 일별 진척률 이력을 저장할 수 없습니다."
        End If
        Exit Sub
    End If

    errorStage = "일별 현황 템플릿 확인"
    Set templateWs = GetDailyTemplateSheet()
    If templateWs Is Nothing Then
        If showCompletionMessage Then
            MsgBox "통합문서에 일별 진행 현황 템플릿이 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2102, , "통합문서에 일별 진행 현황 템플릿이 없습니다."
        End If
        Exit Sub
    End If

    errorStage = "간트 차트 새로고침"
    sourceWs.Activate
    RefreshGanttSheet False
    sourceWs.Calculate

    errorStage = "업무 데이터 확인"
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
    Set groupOrder = New Collection

    errorStage = "업무 그룹 구성"
    BuildDailyReportGroups sourceWs, lastRow, groupRows, groupOrder
    If groupOrder.Count = 0 Then
        If showCompletionMessage Then
            MsgBox "출력할 Level 1 업무가 없습니다.", vbExclamation
        Else
            Err.Raise vbObjectError + 2104, , "출력할 Level 1 업무가 없습니다."
        End If
        Exit Sub
    End If

    errorStage = "보고 기간 계산"
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

    errorStage = "휴일 및 표시 기간 설정 로드"
    EnsureConfigSheet
    LoadHolidaySettings holidayDict, workdayDict
    hasDisplayDateRange = TryGetDisplayDateRange(displayStartDate, displayEndDate)
    errorStage = "진척률 이력 로드"
    LoadDailyProgressHistory

    If hasDisplayDateRange Then
        If displayStartDate > firstReportDate Then firstReportDate = displayStartDate
        If displayEndDate < lastReportDate Then lastReportDate = displayEndDate
    End If

    If firstReportDate > lastReportDate Then
        If showCompletionMessage Then MsgBox "config 시트의 표시 기간에 출력할 업무가 없습니다.", vbInformation
        GoTo SafeExit
    End If

    monthCursor = DateSerial(Year(firstReportDate), Month(firstReportDate), 1)
    Do While monthCursor <= DateSerial(Year(lastReportDate), Month(lastReportDate), 1)
        errorStage = Format$(monthCursor, "yyyy-mm") & " 보고서 시트 생성"
        monthStart = monthCursor
        If firstReportDate > monthStart Then monthStart = firstReportDate

        monthEnd = DateSerial(Year(monthCursor), Month(monthCursor) + 1, 0)
        If lastReportDate < monthEnd Then monthEnd = lastReportDate

        Set activeGroupOrder = GetDailyActiveGroupOrder( _
            sourceWs, groupRows, groupOrder, monthStart, monthEnd, _
            holidayDict, workdayDict)

        If HasWorkingDate(monthStart, monthEnd, holidayDict, workdayDict) And _
           activeGroupOrder.Count > 0 Then
            AddDailyReportMonthSheet templateWs, outputBook, sourceWs, _
                groupRows, activeGroupOrder, monthStart, monthEnd, _
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

    errorStage = "결과 파일 저장"
    outputBook.SaveAs Filename:=CStr(savePath), FileFormat:=DAILY_FILE_FORMAT_XLSX
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing

    errorStage = "진척률 이력 저장"
    SaveDailyProgressHistory sourceWs, groupRows, groupOrder, Date
    errorStage = "원본 통합문서 저장"
    ThisWorkbook.Save
    ThisWorkbook.Activate
    sourceWs.Activate

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
    If Len(errorStage) > 0 Then errorText = errorStage & ": " & errorText
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
                                   ByVal groupOrder As Collection)
    Dim rowNum As Long
    Dim taskLevel As Long
    Dim currentKey As String
    Dim groupKey As String
    Dim ownerText As String
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
                groupOrder.Add groupKey
            End If
        End If

        If Len(currentKey) > 0 And Len(Trim$(CStr(ws.Cells(rowNum, COL_TASK).Value2))) > 0 Then
            Set rows = groupRows(currentKey)
            rows.Add rowNum
        End If
    Next rowNum

    RemoveDailyEmptyGroups groupRows, groupOrder
End Sub

Private Sub RemoveDailyEmptyGroups(ByVal groupRows As Object, _
                                   ByVal groupOrder As Collection)
    Dim groupIndex As Long
    Dim groupKey As String
    Dim rows As Collection

    For groupIndex = groupOrder.Count To 1 Step -1
        groupKey = CStr(groupOrder(groupIndex))
        Set rows = groupRows(groupKey)
        If rows.Count = 0 Then
            groupRows.Remove groupKey
            groupOrder.Remove groupIndex
        End If
    Next groupIndex
End Sub

Private Sub SaveDailyProgressHistory(ByVal ws As Worksheet, _
                                     ByVal groupRows As Object, _
                                     ByVal groupOrder As Collection, _
                                     ByVal snapshotDate As Date)
    Dim historyWs As Worksheet
    Dim groupIndex As Long
    Dim groupKey As String
    Dim rows As Collection
    Dim sourceRow As Long
    Dim targetRow As Long

    Set historyWs = GetDailyHistorySheet()
    For groupIndex = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(groupIndex))
        Set rows = groupRows(groupKey)
        If rows.Count > 0 Then
            sourceRow = CLng(rows(1))
            targetRow = FindDailyHistoryRow(historyWs, ws.Name, groupKey, snapshotDate)
            If targetRow = 0 Then
                targetRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row + 1
                If targetRow < DAILY_HISTORY_DATA_START_ROW Then _
                    targetRow = DAILY_HISTORY_DATA_START_ROW
            End If

            historyWs.Cells(targetRow, 1).Value = DateValue(snapshotDate)
            historyWs.Cells(targetRow, 2).Value = groupKey
            historyWs.Cells(targetRow, 3).Value = GetTaskProgressValue(ws, sourceRow)
            historyWs.Cells(targetRow, 4).Value = CStr(ws.Cells(sourceRow, COL_STATUS).Value)
            historyWs.Cells(targetRow, 5).Value = ws.Cells(sourceRow, COL_TYPE).Value2
            historyWs.Cells(targetRow, 6).Value = ws.Cells(sourceRow, COL_MAJOR_CATEGORY).Value2
            historyWs.Cells(targetRow, 7).Value = ws.Cells(sourceRow, COL_MIDDLE_CATEGORY).Value2
            historyWs.Cells(targetRow, 8).Value = ws.Cells(sourceRow, COL_MINOR_CATEGORY).Value2
            historyWs.Cells(targetRow, 9).Value = ws.Cells(sourceRow, COL_OWNER).Value2
            historyWs.Cells(targetRow, 10).Value = Now
            historyWs.Cells(targetRow, DAILY_HISTORY_SOURCE_SHEET_COLUMN).Value = ws.Name
            RemoveDuplicateDailyHistoryRows historyWs, ws.Name, groupKey, _
                                            snapshotDate, targetRow
        End If
    Next groupIndex

    historyWs.Columns(1).NumberFormat = "yyyy-mm-dd"
    historyWs.Columns(3).NumberFormat = "0%"
    historyWs.Columns(10).NumberFormat = "yyyy-mm-dd hh:mm"
    Set mDailyProgressHistory = Nothing
    Set mDailyProgressCache = Nothing
End Sub

Private Function FindDailyHistoryRow(ByVal historyWs As Worksheet, _
                                     ByVal sourceSheetName As String, _
                                     ByVal groupKey As String, _
                                     ByVal snapshotDate As Date) As Long
    Dim rowNum As Long
    Dim lastRow As Long
    Dim savedSheetName As String

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row
    For rowNum = lastRow To DAILY_HISTORY_DATA_START_ROW Step -1
        If IsDate(historyWs.Cells(rowNum, 1).Value) Then
            savedSheetName = Trim$(CStr( _
                historyWs.Cells(rowNum, DAILY_HISTORY_SOURCE_SHEET_COLUMN).Value2))
            If CLng(DateValue(historyWs.Cells(rowNum, 1).Value)) = _
               CLng(DateValue(snapshotDate)) And _
               CStr(historyWs.Cells(rowNum, 2).Value2) = groupKey And _
               (Len(savedSheetName) = 0 Or _
                StrComp(savedSheetName, sourceSheetName, vbTextCompare) = 0) Then
                FindDailyHistoryRow = rowNum
                Exit Function
            End If
        End If
    Next rowNum
End Function

Private Sub RemoveDuplicateDailyHistoryRows(ByVal historyWs As Worksheet, _
                                            ByVal sourceSheetName As String, _
                                            ByVal groupKey As String, _
                                            ByVal snapshotDate As Date, _
                                            ByVal keepRow As Long)
    Dim rowNum As Long
    Dim lastRow As Long
    Dim savedSheetName As String

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row
    For rowNum = lastRow To DAILY_HISTORY_DATA_START_ROW Step -1
        If rowNum <> keepRow And IsDate(historyWs.Cells(rowNum, 1).Value) Then
            savedSheetName = Trim$(CStr( _
                historyWs.Cells(rowNum, DAILY_HISTORY_SOURCE_SHEET_COLUMN).Value2))
            If CLng(DateValue(historyWs.Cells(rowNum, 1).Value)) = _
               CLng(DateValue(snapshotDate)) And _
               CStr(historyWs.Cells(rowNum, 2).Value2) = groupKey And _
               (Len(savedSheetName) = 0 Or _
                StrComp(savedSheetName, sourceSheetName, vbTextCompare) = 0) Then
                historyWs.Rows(rowNum).Delete
                If rowNum < keepRow Then keepRow = keepRow - 1
            End If
        End If
    Next rowNum
End Sub

Private Function GetDailyHistorySheet() As Worksheet
    On Error Resume Next
    Set GetDailyHistorySheet = ThisWorkbook.Worksheets(DAILY_HISTORY_SHEET_NAME)
    On Error GoTo 0

    If GetDailyHistorySheet Is Nothing Then
        Set GetDailyHistorySheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetDailyHistorySheet.Name = DAILY_HISTORY_SHEET_NAME
    End If

    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 1).Value = "기록일"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 2).Value = "업무키"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 3).Value = "진척률"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 4).Value = "상태"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 5).Value = "Type"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 6).Value = "대분류"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 7).Value = "중분류"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 8).Value = "소분류"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 9).Value = "담당자"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, 10).Value = "저장시각"
    GetDailyHistorySheet.Cells(DAILY_HISTORY_HEADER_ROW, _
                               DAILY_HISTORY_SOURCE_SHEET_COLUMN).Value = "원본시트"
    GetDailyHistorySheet.Rows(DAILY_HISTORY_HEADER_ROW).Font.Bold = True
    GetDailyHistorySheet.Visible = xlSheetVeryHidden
End Function

Private Sub LoadDailyProgressHistory()
    Dim historyWs As Worksheet
    Dim groupHistory As Object
    Dim lastRow As Long
    Dim rowNum As Long
    Dim groupKey As String
    Dim sourceSheetName As String
    Dim historyGroupKey As String
    Dim dateKey As String
    Dim progressValue As Variant

    Set mDailyProgressHistory = CreateObject("Scripting.Dictionary")
    Set mDailyProgressCache = CreateObject("Scripting.Dictionary")
    Set historyWs = GetDailyHistorySheet()
    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row

    For rowNum = DAILY_HISTORY_DATA_START_ROW To lastRow
        groupKey = CStr(historyWs.Cells(rowNum, 2).Value2)
        progressValue = historyWs.Cells(rowNum, 3).Value
        If Len(groupKey) > 0 And IsDate(historyWs.Cells(rowNum, 1).Value) And _
           IsNumeric(progressValue) Then
            sourceSheetName = Trim$(CStr( _
                historyWs.Cells(rowNum, DAILY_HISTORY_SOURCE_SHEET_COLUMN).Value2))
            historyGroupKey = BuildDailyHistoryGroupKey(sourceSheetName, groupKey)
            If Not mDailyProgressHistory.Exists(historyGroupKey) Then
                Set groupHistory = CreateObject("Scripting.Dictionary")
                mDailyProgressHistory.Add historyGroupKey, groupHistory
            Else
                Set groupHistory = mDailyProgressHistory(historyGroupKey)
            End If

            dateKey = CStr(CLng(DateValue(historyWs.Cells(rowNum, 1).Value)))
            groupHistory(dateKey) = CDbl(progressValue)
        End If
    Next rowNum
End Sub

Private Function GetDailyProgressForDate(ByVal ws As Worksheet, _
                                         ByVal groupKey As String, _
                                         ByVal targetDate As Date, _
                                         ByVal rows As Collection) As Variant
    Dim currentProgress As Double
    Dim cacheKey As String
    Dim result As Variant

    If rows.Count = 0 Then Exit Function

    currentProgress = GetTaskProgressValue(ws, CLng(rows(1)))
    If CLng(DateValue(targetDate)) = CLng(Date) Then
        GetDailyProgressForDate = currentProgress
        Exit Function
    End If

    If targetDate > Date Then
        GetDailyProgressForDate = currentProgress
        Exit Function
    End If

    cacheKey = BuildDailyHistoryGroupKey(ws.Name, groupKey) & _
               ChrW(DAILY_HISTORY_KEY_SEPARATOR) & CStr(CLng(DateValue(targetDate)))
    If mDailyProgressCache.Exists(cacheKey) Then
        GetDailyProgressForDate = mDailyProgressCache(cacheKey)
        Exit Function
    End If

    result = FindLatestDailyProgress(ws.Name, groupKey, targetDate)
    If IsEmpty(result) Then result = currentProgress
    mDailyProgressCache.Add cacheKey, result
    GetDailyProgressForDate = result
End Function

Private Function FindLatestDailyProgress(ByVal sourceSheetName As String, _
                                         ByVal groupKey As String, _
                                         ByVal targetDate As Date) As Variant
    Dim result As Variant

    result = FindLatestDailyProgressInGroup( _
        BuildDailyHistoryGroupKey(sourceSheetName, groupKey), targetDate)
    If IsEmpty(result) Then
        result = FindLatestDailyProgressInGroup( _
            BuildDailyHistoryGroupKey("", groupKey), targetDate)
    End If
    FindLatestDailyProgress = result
End Function

Private Function FindLatestDailyProgressInGroup(ByVal historyGroupKey As String, _
                                                ByVal targetDate As Date) As Variant
    Dim groupHistory As Object
    Dim itemDateKey As Variant
    Dim itemDate As Long
    Dim latestDate As Long

    If mDailyProgressHistory Is Nothing Then LoadDailyProgressHistory
    If Not mDailyProgressHistory.Exists(historyGroupKey) Then Exit Function

    Set groupHistory = mDailyProgressHistory(historyGroupKey)
    For Each itemDateKey In groupHistory.Keys
        itemDate = CLng(itemDateKey)
        If itemDate <= CLng(DateValue(targetDate)) And itemDate > latestDate Then
            latestDate = itemDate
            FindLatestDailyProgressInGroup = groupHistory(itemDateKey)
        End If
    Next itemDateKey
End Function

Private Function BuildDailyHistoryGroupKey(ByVal sourceSheetName As String, _
                                           ByVal groupKey As String) As String
    BuildDailyHistoryGroupKey = LCase$(Trim$(sourceSheetName)) & _
                                ChrW(DAILY_HISTORY_KEY_SEPARATOR) & groupKey
End Function

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
                                     ByVal activeGroupOrder As Collection, _
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
    PopulateDailyMonthSheet outputWs, sourceWs, groupRows, activeGroupOrder, _
        monthStart, monthEnd, holidayDict, workdayDict
    Exit Sub

CopyFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    templateWs.Visible = originalVisibility
    On Error GoTo 0
    Err.Raise vbObjectError + 2111, "AddDailyReportMonthSheet", _
              "원본 오류 " & CStr(errorNumber) & ": " & errorDescription
End Sub

Private Sub PopulateDailyMonthSheet(ByVal outputWs As Worksheet, _
                                    ByVal sourceWs As Worksheet, _
                                    ByVal groupRows As Object, _
                                    ByVal activeGroupOrder As Collection, _
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
    Dim rows As Collection
    Dim shapeIndex As Long
    Dim groupSourceRow As Long
    Dim populateStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo PopulateFailed

    populateStage = "출력 영역 계산"
    groupCount = activeGroupOrder.Count
    outputLastRow = DAILY_DATA_START_ROW + groupCount - 1
    clearLastRow = outputWs.UsedRange.Row + outputWs.UsedRange.Rows.Count - 1
    If clearLastRow < DAILY_TEMPLATE_LAST_ROW Then clearLastRow = DAILY_TEMPLATE_LAST_ROW
    If clearLastRow < outputLastRow Then clearLastRow = outputLastRow

    lastUsedColumn = outputWs.UsedRange.Column + outputWs.UsedRange.Columns.Count - 1
    If lastUsedColumn < DAILY_FIRST_DATE_COLUMN Then lastUsedColumn = DAILY_FIRST_DATE_COLUMN

    populateStage = "템플릿 병합 해제"
    On Error Resume Next
    populateStage = "템플릿 데이터 초기화"
    outputWs.Range(outputWs.Cells(DAILY_DATA_START_ROW, 2), _
                   outputWs.Cells(clearLastRow, lastUsedColumn)).UnMerge
    On Error GoTo PopulateFailed

    outputWs.Range(outputWs.Cells(DAILY_DATA_START_ROW, 2), _
                   outputWs.Cells(clearLastRow, lastUsedColumn)).ClearContents

    populateStage = "템플릿 개체 제거"
    For shapeIndex = outputWs.Shapes.Count To 1 Step -1
        outputWs.Shapes(shapeIndex).Delete
    Next shapeIndex

    populateStage = "행 서식 준비"
    PrepareDailyDataRowFormats outputWs, outputLastRow

    populateStage = "제목 및 헤더 작성"
    outputWs.Range("B2").Value = DAILY_TITLE
    outputWs.Range("B4").Value = "Type"
    outputWs.Range("C4").Value = "대분류"
    outputWs.Range("D4").Value = "중분류"
    outputWs.Range("E4").Value = "메뉴명"
    outputWs.Range("F4").Value = "개발자_정"
    outputWs.Range("G4").Value = "개발자_부"

    populateStage = "업무 그룹 작성"
    For groupIndex = 1 To groupCount
        targetRow = DAILY_DATA_START_ROW + groupIndex - 1
        groupKey = CStr(activeGroupOrder(groupIndex))
        Set rows = groupRows(groupKey)
        groupSourceRow = CLng(rows(1))

        outputWs.Cells(targetRow, 2).Value = sourceWs.Cells(groupSourceRow, COL_TYPE).Value2
        outputWs.Cells(targetRow, 3).Value = sourceWs.Cells(groupSourceRow, COL_MAJOR_CATEGORY).Value2
        outputWs.Cells(targetRow, 4).Value = sourceWs.Cells(groupSourceRow, COL_MIDDLE_CATEGORY).Value2
        outputWs.Cells(targetRow, 5).Value = sourceWs.Cells(groupSourceRow, COL_MINOR_CATEGORY).Value2
        outputWs.Cells(targetRow, 6).Value = sourceWs.Cells(groupSourceRow, COL_OWNER).Value2
        outputWs.Cells(targetRow, 7).ClearContents
    Next groupIndex

    populateStage = "날짜별 업무 작성"
    dateColumn = DAILY_FIRST_DATE_COLUMN
    For targetDate = monthStart To monthEnd
        If IsWorkingDay(targetDate, holidayDict, workdayDict) Then
            PrepareDailyDateColumn outputWs, dateColumn, clearLastRow
            outputWs.Cells(DAILY_HEADER_ROW, dateColumn).Value = targetDate
            outputWs.Cells(DAILY_HEADER_ROW, dateColumn).NumberFormatLocal = "m/d"

            For groupIndex = 1 To groupCount
                targetRow = DAILY_DATA_START_ROW + groupIndex - 1
                groupKey = CStr(activeGroupOrder(groupIndex))
                Set rows = groupRows(groupKey)
                outputWs.Cells(targetRow, dateColumn).Value = _
                    BuildDailyCellText(sourceWs, rows, targetDate, _
                                       CStr(sourceWs.Cells(CLng(rows(1)), COL_OWNER).Value2), _
                                       GetDailyProgressForDate(sourceWs, groupKey, _
                                                               targetDate, rows))
            Next groupIndex

            dateColumn = dateColumn + 1
        End If
    Next targetDate

    populateStage = "미사용 날짜 열 정리"
    If dateColumn <= lastUsedColumn Then
        If outputWs.AutoFilterMode Then outputWs.AutoFilterMode = False
        outputWs.Range(outputWs.Cells(DAILY_HEADER_ROW, dateColumn), _
                       outputWs.Cells(clearLastRow, lastUsedColumn)).ClearContents
        outputWs.Range(outputWs.Cells(DAILY_HEADER_ROW, dateColumn), _
                       outputWs.Cells(clearLastRow, lastUsedColumn)).ClearFormats
        outputWs.Range(outputWs.Cells(1, dateColumn), _
                       outputWs.Cells(1, lastUsedColumn)).EntireColumn.ColumnWidth = _
            outputWs.StandardWidth
    End If

    populateStage = "분류 셀 병합"
    MergeRepeatedDailyCategories outputWs, outputLastRow
    populateStage = "출력 영역 서식 적용"
    FormatDailyOutputArea outputWs, outputLastRow, dateColumn - 1
    populateStage = "미사용 템플릿 영역 정리"
    ClearDailyUnusedTemplateArea outputWs, outputLastRow, clearLastRow, lastUsedColumn
    Exit Sub

PopulateFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise vbObjectError + 2110, "PopulateDailyMonthSheet", _
              populateStage & " (원본 오류 " & CStr(errorNumber) & "): " & _
              errorDescription
End Sub

Private Function GetDailyActiveGroupOrder(ByVal sourceWs As Worksheet, _
                                          ByVal groupRows As Object, _
                                          ByVal groupOrder As Collection, _
                                          ByVal monthStart As Date, _
                                          ByVal monthEnd As Date, _
                                          ByVal holidayDict As Object, _
                                          ByVal workdayDict As Object) As Collection
    Dim activeGroups As Collection
    Dim groupIndex As Long
    Dim groupKey As String
    Dim rows As Collection
    Dim targetDate As Date
    Dim rowNum As Variant
    Dim groupIsActive As Boolean

    Set activeGroups = New Collection

    For groupIndex = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(groupIndex))
        Set rows = groupRows(groupKey)
        groupIsActive = False

        For targetDate = monthStart To monthEnd
            If IsWorkingDay(targetDate, holidayDict, workdayDict) Then
                For Each rowNum In rows
                    If IsDailyTaskOnDate(sourceWs, CLng(rowNum), targetDate) Then
                        groupIsActive = True
                        Exit For
                    End If
                Next rowNum
            End If

            If groupIsActive Then Exit For
        Next targetDate

        If groupIsActive Then activeGroups.Add groupKey
    Next groupIndex

    Set GetDailyActiveGroupOrder = activeGroups
End Function

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

Private Sub ClearDailyUnusedTemplateArea(ByVal ws As Worksheet, _
                                         ByVal lastDataRow As Long, _
                                         ByVal clearLastRow As Long, _
                                         ByVal lastUsedColumn As Long)
    Dim clearRange As Range

    If clearLastRow <= lastDataRow Then Exit Sub

    Set clearRange = ws.Range(ws.Cells(lastDataRow + 1, 2), _
                              ws.Cells(clearLastRow, lastUsedColumn))
    On Error Resume Next
    clearRange.UnMerge
    On Error GoTo 0
    clearRange.ClearContents
    clearRange.ClearFormats
    ws.Rows((lastDataRow + 1) & ":" & clearLastRow).RowHeight = _
        ws.StandardHeight
End Sub

Private Function BuildDailyCellText(ByVal ws As Worksheet, _
                                    ByVal rows As Collection, _
                                    ByVal targetDate As Date, _
                                    ByVal ownerText As String, _
                                    ByVal progressValue As Variant) As String
    Dim rowNum As Variant
    Dim taskLevel As Long
    Dim taskText As String
    Dim lineText As String
    Dim bodyText As String
    Dim level1Number As Long
    Dim progressText As String

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
        If Not IsEmpty(progressValue) And IsNumeric(progressValue) Then
            progressText = "진척률 " & Format$(CDbl(progressValue), "0%")
        Else
            progressText = "진척률 미기록"
        End If

        If Len(ownerText) > 0 Then
            BuildDailyCellText = "(" & ownerText & ") " & progressText & _
                                 vbLf & bodyText
        Else
            BuildDailyCellText = progressText & vbLf & bodyText
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
         StrComp(ws.Name, DAILY_HISTORY_SHEET_NAME, vbTextCompare) <> 0 And _
         StrComp(ws.Name, "_버튼생성", vbTextCompare) <> 0 And _
         StrComp(ws.Name, WEEKLY_REPORT_CONFIG_SHEET_NAME, vbTextCompare) <> 0)
End Function

Private Function GetDailyReportSourceSheet() As Worksheet
    Dim activeWs As Worksheet

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    Set activeWs = ActiveSheet
    If Not activeWs.Parent Is ThisWorkbook Then Exit Function
    If Not IsDailyReportTaskSheet(activeWs) Then Exit Function
    If GetLastDataRow(activeWs) < DATA_START_ROW Then Exit Function

    Set GetDailyReportSourceSheet = activeWs
End Function

Private Sub RestoreDailyReportApplicationState(ByVal screenUpdating As Boolean, _
                                               ByVal enableEvents As Boolean)
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
End Sub
