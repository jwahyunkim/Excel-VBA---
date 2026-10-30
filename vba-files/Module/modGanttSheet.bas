Attribute VB_Name = "modGanttSheet"
Option Explicit
' 2026-10-30 김좌현 수정 : 새로고침 서식 적용 시 Excel 기본 자동필터 조건과 범위 유지
' 2026-10-29 김좌현 수정 : 폐기한 간트 전용 캐시 무효화 호출 제거
' 2026-10-27 김좌현 수정 : 입력 열·변경 그룹만 계층 계산하고 UI 전용 보호의 재설정 생략
' 2026-10-23 김좌현 수정 : WEB UPDATE 헤더로 변경하고 기존 웹등록 열의 중복 이전 방지
' 2026-10-23 김좌현 수정 : 하위 작업시간의 상위 합계와 입력 즉시 갱신 및 중복 집계 방지
' 2026-10-22 김좌현 수정 : 날짜 열 너비 일괄 적용과 덮어쓰는 열 맞춤 제거로 간트 새로고침 가속
' 2026-10-22 김좌현 수정 : 간트 서식 중복 제거, 실제 표시 영역만 초기화와 선택적 너비/계층 갱신
' 2026-10-22 김좌현 수정 : 웹 설정/업무 ID/작업시간 열 이전과 표시 및 보호 통합
' 2026-10-22 김좌현 수정 : 상태에는 괄호 안 지연 일수만 표시하고 지연 완료도 기존 완료 초록 유지
' 2026-10-22 김좌현 수정 : 오늘 날짜 헤더와 빈 배경을 민트색으로 구분
' 2026-10-21 김좌현 수정 : 원래 상태값을 유지하면서 지연 일수와 하위 지연 및 완료일 누락 표시
' 2026-10-20 김좌현 수정 : 완료 숨김에서도 동일한 계층/상태 재계산 사용
' 2026-10-20 김좌현 수정 : 간트 only/보고 only 필터 및 전용 판정 함수 제거
' 2026-10-20 김좌현 수정 : 완료 항목 숨김 처리와 전용 번호 변환 제거

Private Const TASK_VALIDATION_MIN_LAST_ROW As Long = 5000
Private Const TASK_VALIDATION_EXTRA_ROWS As Long = 1000

Public Sub SetupDataHeaders(ws As Worksheet)
    EnsureTaskSheetLayout ws

    ws.Cells(HEADER_ROW, COL_NO).Value = "No."
    ws.Cells(HEADER_ROW, COL_LEVEL).Value = "Level"
    ws.Cells(HEADER_ROW, COL_TYPE).Value = "타입"
    ws.Cells(HEADER_ROW, COL_MAJOR_CATEGORY).Value = "대분류"
    ws.Cells(HEADER_ROW, COL_MIDDLE_CATEGORY).Value = "중분류"
    ws.Cells(HEADER_ROW, COL_MINOR_CATEGORY).Value = "소분류"
    ws.Cells(HEADER_ROW, COL_REQUEST).Value = "요청 내용"
    ws.Cells(HEADER_ROW, COL_MODIFICATION).Value = "수정 내용"
    ws.Cells(HEADER_ROW, COL_OWNER).Value = "담당"
    ws.Cells(HEADER_ROW, COL_NOTE).Value = "비고"
    ws.Cells(HEADER_ROW, COL_PLAN_START).Value = "계획 시작일"
    ws.Cells(HEADER_ROW, COL_PLAN_END).Value = "계획 종료일"
    ws.Cells(HEADER_ROW, COL_ACTUAL_START).Value = "실제 시작일"
    ws.Cells(HEADER_ROW, COL_ACTUAL_END).Value = "실제 종료일"
    ws.Cells(HEADER_ROW, COL_PROGRESS).Value = "진행률"
    ws.Cells(HEADER_ROW, COL_NORMAL_PROGRESS).Value = "정상 진행률"
    ws.Cells(HEADER_ROW, COL_MANUAL_PROGRESS).Value = "진행률 수동"
    ws.Cells(HEADER_ROW, COL_MANUAL_STATUS).Value = "상태 수동"
    ws.Cells(HEADER_ROW, COL_WEEKLY_REPORT).Value = "보고 상태"
    ws.Cells(HEADER_ROW, COL_PLAN_DAYS).Value = "계획일수"
    ws.Cells(HEADER_ROW, COL_ACTUAL_DAYS).Value = "실소요일수"
    ws.Cells(HEADER_ROW, COL_STATUS).Value = "상태"
End Sub

Private Sub NormalizeReportStatusCell(ByVal targetCell As Range)
    Dim statusText As String

    statusText = UCase$(Trim$(CStr(targetCell.Value2)))

    Select Case statusText
        Case "Y", "TRUE"
            targetCell.Value = REPORT_STATUS_IN_PROGRESS
        Case "N", "FALSE"
            targetCell.ClearContents
        Case UCase$(REPORT_STATUS_PLANNED)
            targetCell.Value = REPORT_STATUS_PLANNED
        Case UCase$(REPORT_STATUS_IN_PROGRESS)
            targetCell.Value = REPORT_STATUS_IN_PROGRESS
        Case UCase$(REPORT_STATUS_COMPLETED)
            targetCell.Value = REPORT_STATUS_COMPLETED
    End Select
End Sub

Public Sub UpdateTaskNumbers(ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim seqNo As Long

    If lastRow < DATA_START_ROW Then Exit Sub

    seqNo = 1

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            ws.Cells(r, COL_NO).Value = seqNo
            ws.Cells(r, COL_NO).NumberFormat = "0"
            seqNo = seqNo + 1
        Else
            ws.Cells(r, COL_NO).ClearContents
        End If
    Next r
End Sub

' 2026-10-23 김좌현 수정 : 이전 표시명도 현재 웹 열 구조로 인식해 데이터와 개체 위치 유지
Public Function HasTaskWebHeader(ByVal ws As Worksheet) As Boolean
    Dim headerText As String
    If ws Is Nothing Then Exit Function
    headerText = CStr(ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2)
    HasTaskWebHeader = (headerText = "WEB UPDATE" Or headerText = "웹등록")
End Function

' 2026-10-22 김좌현 수정 : 기존 레이아웃 이전 후 웹 설정 열과 분 단위 작업시간을 한 번만 삽입
Public Sub EnsureTaskSheetLayout(ByVal ws As Worksheet)
    Dim previousEnableEvents As Boolean, wasProtected As Boolean
    Dim insertedWebColumns As Boolean, legacyRequestColumn As Long
    Dim errorNumber As Long, errorDescription As String
    If ws Is Nothing Then Exit Sub
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or _
       IsWebUtilitySheet(ws) Then Exit Sub
    previousEnableEvents = Application.EnableEvents
    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    On Error GoTo Failed
    Application.EnableEvents = False
    If wasProtected Then ws.Unprotect

    If Not HasTaskWebHeader(ws) Then
        legacyRequestColumn = FindHeaderColumnByName(ws, "요청 내용")
        If legacyRequestColumn = 0 Then legacyRequestColumn = FindHeaderColumnByName(ws, "내용")
        If legacyRequestColumn > 0 Then MigrateTaskWebFoldSettings ws
        EnsureLegacyTaskSheetLayout ws
        ' Native insertion keeps values, formulas, names and cell-anchored objects together.
        ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_TASK_ID).Insert Shift:=xlToRight
        insertedWebColumns = True
    ElseIf CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) <> "요청 내용" Then
        Err.Raise vbObjectError + 7180, "EnsureTaskSheetLayout", _
                  "웹 등록 열 뒤의 요청 내용 위치를 확인하세요. 기존 데이터를 자동으로 덮어쓰지 않았습니다."
    End If

    If CStr(ws.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2) <> "작업시간(분)" Then
        If Len(CStr(ws.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2)) > 0 And _
           CStr(ws.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2) <> "진행률" Then _
            Err.Raise vbObjectError + 7181, "EnsureTaskSheetLayout", _
                      "실제 종료일 다음의 작업시간 열 위치를 확인하세요."
        ws.Columns(COL_WORK_MINUTES).Insert Shift:=xlToRight
        ws.Columns(COL_WORK_MINUTES).Hidden = True
    End If
    ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2 = "WEB UPDATE"
    ws.Cells(HEADER_ROW, COL_WEB_LEVEL).Value2 = "등록기준"
    ws.Cells(HEADER_ROW, COL_WEB_REQUESTER).Value2 = "웹요청자"
    ws.Cells(HEADER_ROW, COL_WEB_ACTION).Value2 = "웹처리구분"
    ws.Cells(HEADER_ROW, COL_WEB_STATUS).Value2 = "웹진행상태"
    ws.Cells(HEADER_ROW, COL_WEB_MODULE).Value2 = "웹모듈/업무"
    ws.Cells(HEADER_ROW, COL_WEB_HANDLER).Value2 = "웹처리자"
    ws.Cells(HEADER_ROW, COL_WEB_PROGRAM_ID).Value2 = "프로그램ID"
    ws.Cells(HEADER_ROW, COL_WEB_RESULT).Value2 = "웹등록결과"
    ws.Cells(HEADER_ROW, COL_WEB_TASK_ID).Value2 = "웹업무ID"
    ws.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2 = "작업시간(분)"
    If insertedWebColumns Then
        ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_RESULT).Hidden = True
        ws.Columns(COL_WORK_MINUTES).Hidden = True
    End If
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
    ApplyWebTaskValidation ws
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = previousEnableEvents
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
    Err.Raise errorNumber, "EnsureTaskSheetLayout", errorDescription
End Sub
Private Sub EnsureLegacyTaskSheetLayout(ByVal ws As Worksheet)
    Dim i As Long
    Dim r As Long
    Dim lastMigrationRow As Long
    Dim manualStatusText As String
    Dim targetRange As Range
    Dim lo As ListObject
    Dim previousEnableEvents As Boolean
    Dim wasProtected As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If ws Is Nothing Then Exit Sub
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or _
       IsWebUtilitySheet(ws) Then Exit Sub

    previousEnableEvents = Application.EnableEvents
    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    On Error GoTo MigrationError
    Application.EnableEvents = False
    If wasProtected Then ws.Unprotect

    If Trim$(CStr(ws.Cells(HEADER_ROW, "B").Value)) <> "No." Then
        ws.Columns("B").Insert Shift:=xlToRight
    End If
    ws.Cells(HEADER_ROW, "B").Value = "No."

    If Trim$(CStr(ws.Cells(HEADER_ROW, "D").Value)) <> "타입" Then
        ws.Columns("D").Insert Shift:=xlToRight
    End If
    ws.Cells(HEADER_ROW, "D").Value = "타입"

    If Trim$(CStr(ws.Cells(HEADER_ROW, "E").Value)) <> "대분류" Then
        If Trim$(CStr(ws.Cells(HEADER_ROW, "E").Value)) = "모듈" Then
            ws.Cells(HEADER_ROW, "E").Value = "대분류"
        Else
            ws.Columns("E").Insert Shift:=xlToRight
        End If
    End If
    ws.Cells(HEADER_ROW, "E").Value = "대분류"

    If Trim$(CStr(ws.Cells(HEADER_ROW, "F").Value)) <> "중분류" Then
        ws.Columns("F").Insert Shift:=xlToRight
    End If
    ws.Cells(HEADER_ROW, "F").Value = "중분류"

    If Trim$(CStr(ws.Cells(HEADER_ROW, "G").Value)) <> "소분류" Then
        If Trim$(CStr(ws.Cells(HEADER_ROW, "G").Value)) = "프로그램" Then
            ws.Cells(HEADER_ROW, "G").Value = "소분류"
        Else
            ws.Columns("G").Insert Shift:=xlToRight
        End If
    End If
    ws.Cells(HEADER_ROW, "G").Value = "소분류"

    ' Insert the second content column once. Excel shifts values, formulas and
    ' cell-anchored embedded objects together; the old H content remains intact.
    If Trim$(CStr(ws.Cells(HEADER_ROW, "I").Value)) <> "수정 내용" Then
        ws.Columns("I").Insert Shift:=xlToRight
        ws.Columns("I").ColumnWidth = ws.Columns("H").ColumnWidth
        On Error Resume Next
        ws.Columns("I").Validation.Delete
        On Error GoTo MigrationError
    End If
    ws.Cells(HEADER_ROW, "H").Value = "요청 내용"
    ws.Cells(HEADER_ROW, "I").Value = "수정 내용"

    If Trim$(CStr(ws.Cells(HEADER_ROW, "J").Value)) <> "담당" Then
        ws.Columns("J").Insert Shift:=xlToRight
    End If
    ws.Cells(HEADER_ROW, "J").Value = "담당"

    ' Newly inserted classification columns can inherit validation from adjacent cells.
    On Error Resume Next
    ws.Range("D" & DATA_START_ROW & ":" & _
             "G" & ws.Rows.Count).Validation.Delete
    On Error GoTo MigrationError
    EnsureTaskTypeConfig
    ApplyTaskTypeValidation ws

    ' Remove the retired development-report column from existing workbooks.
    If Trim$(CStr(ws.Range("T" & HEADER_ROW).Value)) = "개발진행" Then
        ws.Columns("T").Delete Shift:=xlToLeft
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, "T").Value)) <> "보고 상태" And _
       Trim$(CStr(ws.Cells(HEADER_ROW, "T").Value)) <> "주간보고" Then
        ws.Columns("T").Insert Shift:=xlToRight
    End If
    ws.Cells(HEADER_ROW, "T").Value = "보고 상태"

    lastMigrationRow = Application.Max( _
        ws.Cells(ws.Rows.Count, "H").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "I").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "S").End(xlUp).Row)

    For r = DATA_START_ROW To lastMigrationRow
        manualStatusText = Trim$(CStr(ws.Cells(r, "S").Value))

        If manualStatusText = STATUS_WEEKLY_REPORT Then
            ws.Cells(r, "T").Value = REPORT_STATUS_IN_PROGRESS
            ws.Cells(r, "S").ClearContents
        ElseIf manualStatusText = "개발진행" Or manualStatusText = "개발 진행" Then
            ws.Cells(r, "S").ClearContents
        End If
        NormalizeReportStatusCell ws.Cells(r, "T")
    Next r

    Set targetRange = ws.Range("B" & HEADER_ROW & ":" & ws.Cells(HEADER_ROW, ws.Columns.Count).Address(False, False))

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo MigrationError

    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = ws.ListObjects(i)
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            lo.Unlist
        End If
    Next i

    If wasProtected Then
        lastMigrationRow = GetLastDataRow(ws)
        If lastMigrationRow < DATA_START_ROW Then lastMigrationRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastMigrationRow
    End If
    Application.EnableEvents = previousEnableEvents
    Exit Sub

MigrationError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If wasProtected Then
        ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFiltering:=True, _
                   AllowInsertingRows:=True, AllowDeletingRows:=True
    End If
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
    Err.Raise errorNumber, "EnsureTaskSheetLayout", errorDescription
End Sub

Public Sub ClearCalculatedArea(ws As Worksheet, ByVal lastRow As Long, Optional ByVal clearFormats As Boolean = True)
    Dim normalProgressCol As Long
    Dim planDaysCol As Long
    Dim actualDaysCol As Long
    Dim statusCol As Long
    
    If lastRow < DATA_START_ROW Then Exit Sub
    
    normalProgressCol = FindHeaderColumnByName(ws, "정상 진행률")
    planDaysCol = FindHeaderColumnByName(ws, "계획일수")
    actualDaysCol = FindHeaderColumnByName(ws, "실소요일수")
    statusCol = FindHeaderColumnByName(ws, "상태")
    
    If normalProgressCol > 0 And planDaysCol > 0 And actualDaysCol > 0 And statusCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(lastRow, normalProgressCol)).ClearContents
        If clearFormats Then ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(lastRow, normalProgressCol)).ClearFormats
        
        ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(lastRow, actualDaysCol)).ClearContents
        If clearFormats Then ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(lastRow, actualDaysCol)).ClearFormats
        
        ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(lastRow, statusCol)).ClearContents
        If clearFormats Then ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(lastRow, statusCol)).ClearFormats
    Else
        ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).ClearContents
        If clearFormats Then ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).ClearFormats
        
        ws.Range(COL_PLAN_DAYS & DATA_START_ROW & ":" & COL_ACTUAL_DAYS & lastRow).ClearContents
        If clearFormats Then ws.Range(COL_PLAN_DAYS & DATA_START_ROW & ":" & COL_ACTUAL_DAYS & lastRow).ClearFormats
        
        ws.Range(COL_STATUS & DATA_START_ROW & ":" & COL_STATUS & lastRow).ClearContents
        If clearFormats Then ws.Range(COL_STATUS & DATA_START_ROW & ":" & COL_STATUS & lastRow).ClearFormats
    End If
End Sub

Public Sub ClearGanttArea(ws As Worksheet)
    Dim startCol As Long, lastCol As Long, lastRow As Long, chartArea As Range
    ClearTaskBarVisuals ws
    startCol = ws.Range(COL_GANTT_START & "1").Column
    lastCol = ws.Cells(GANTT_HEADER_ROW_DATE, ws.Columns.Count).End(xlToLeft).Column
    lastRow = Application.Max(GetLastDataRow(ws), ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    If lastCol < startCol Then Exit Sub
    Set chartArea = ws.Range(ws.Cells(1, startCol), ws.Cells(lastRow, lastCol))
    chartArea.UnMerge
    chartArea.ClearFormats
    chartArea.ClearContents
End Sub

Public Sub DrawDateHeader(ws As Worksheet, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim startCol As Long
    Dim endCol As Long
    Dim curCol As Long
    Dim d As Date
    Dim weekStartCol As Long
    Dim currentWeekLabel As String
    Dim nextWeekLabel As String
    Dim monthStartCol As Long
    Dim currentMonthLabel As String
    Dim nextMonthLabel As String

    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = startCol + CLng(chartEndDate - chartStartDate)

    ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol)).ClearContents
    ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol)).ClearFormats

    ws.Range(ws.Cells(1, startCol), ws.Cells(1, endCol)).EntireColumn.ColumnWidth = 3

    curCol = startCol
    weekStartCol = startCol
    monthStartCol = startCol
    currentWeekLabel = ""
    currentMonthLabel = ""

    For d = chartStartDate To chartEndDate
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Value = GetWeekdayKorShort(d)
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).HorizontalAlignment = xlCenter
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).VerticalAlignment = xlCenter

        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Value = d
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).NumberFormat = "mm-dd"
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Orientation = 90
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).HorizontalAlignment = xlCenter
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).VerticalAlignment = xlCenter

        nextWeekLabel = DatePart("ww", d, vbMonday, vbFirstFourDays) & " 주차"
        nextMonthLabel = Month(d) & "월"

        If currentMonthLabel = "" Then
            currentMonthLabel = nextMonthLabel
            monthStartCol = curCol
        ElseIf currentMonthLabel <> nextMonthLabel Then
            With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, monthStartCol), ws.Cells(GANTT_HEADER_ROW_MONTH, curCol - 1))
                .Merge
                .Value = currentMonthLabel
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
            End With

            currentMonthLabel = nextMonthLabel
            monthStartCol = curCol
        End If

        If currentWeekLabel = "" Then
            currentWeekLabel = nextWeekLabel
            weekStartCol = curCol
        ElseIf currentWeekLabel <> nextWeekLabel Then
            With ws.Range(ws.Cells(GANTT_HEADER_ROW_WEEK, weekStartCol), ws.Cells(GANTT_HEADER_ROW_WEEK, curCol - 1))
                .Merge
                .Value = currentWeekLabel
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
            End With

            currentWeekLabel = nextWeekLabel
            weekStartCol = curCol
        End If

        If Not IsWorkingDay(d, holidayDict, workdayDict) Then
            ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Interior.Color = RGB(230, 230, 230)
            ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Interior.Color = RGB(230, 230, 230)
        End If

        If CLng(d) = CLng(Date) Then
            ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Value = "오늘"
            With ws.Range(ws.Cells(GANTT_HEADER_ROW_DAY, curCol), _
                          ws.Cells(GANTT_HEADER_ROW_DATE, curCol))
                .Interior.Color = RGB(45, 212, 191)
                .Font.Color = RGB(17, 94, 89)
                .Font.Bold = True
            End With
        End If

        curCol = curCol + 1
    Next d

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, monthStartCol), ws.Cells(GANTT_HEADER_ROW_MONTH, curCol - 1))
        .Merge
        .Value = currentMonthLabel
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_WEEK, weekStartCol), ws.Cells(GANTT_HEADER_ROW_WEEK, curCol - 1))
        .Merge
        .Value = currentWeekLabel
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Public Sub UpdateWeeklyReportStatuses(ByVal ws As Worksheet, _
                                      ByVal lastRow As Long)
    Dim r As Long
    Dim actualStart As Variant
    Dim actualEnd As Variant

    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            actualStart = ws.Cells(r, COL_ACTUAL_START).Value
            actualEnd = ws.Cells(r, COL_ACTUAL_END).Value

            If IsDate(actualEnd) And CLng(CDate(actualEnd)) <= CLng(Date) Then
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_COMPLETED
            ElseIf IsDate(actualStart) And CLng(CDate(actualStart)) <= CLng(Date) Then
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_IN_PROGRESS
            Else
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_PLANNED
            End If
        Else
            ws.Cells(r, COL_WEEKLY_REPORT).ClearContents
        End If
    Next r
End Sub
Public Sub DrawTaskBars(ws As Worksheet, ByVal lastRow As Long, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim r As Long
    Dim d As Date
    Dim colIdx As Long
    Dim startCol As Long
    Dim planS As Variant
    Dim planE As Variant
    Dim actS As Variant
    Dim actE As Variant
    Dim statusText As String
    Dim actualColor As Long
    Dim actualDrawEnd As Variant
    Dim markerDate As Variant
    Dim markerText As String

    ApplyHierarchySummaryValues ws, lastRow, holidayDict, workdayDict
    UpdateWeeklyReportStatuses ws, lastRow

    startCol = ws.Range(COL_GANTT_START & "1").Column

    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then GoTo ContinueNextRow
        If HasTaskError(ws, r) Then GoTo ContinueNextRow

        planS = ws.Cells(r, COL_PLAN_START).Value
        planE = ws.Cells(r, COL_PLAN_END).Value
        actS = ws.Cells(r, COL_ACTUAL_START).Value
        actE = ws.Cells(r, COL_ACTUAL_END).Value
        statusText = CStr(ws.Cells(r, COL_STATUS).Value)
        markerDate = FindLevelMarkerDate(ws, r, holidayDict, workdayDict)
        markerText = GetTaskLevelMarker(ws, r)

        actualColor = GetTaskStatusFillColor(statusText)


        If IsDate(planS) And IsDate(planE) Then
            For d = CDate(planS) To CDate(planE)
                If d >= chartStartDate And d <= chartEndDate Then
                    If IsWorkingDay(d, holidayDict, workdayDict) Then
                        colIdx = startCol + CLng(d - chartStartDate)
                        With ws.Cells(r, colIdx)
                            .Interior.Color = RGB(180, 180, 180)
                            If Not IsEmpty(markerDate) Then
                                If CLng(d) = CLng(markerDate) Then
                                    .Value = markerText
                                    .Font.Color = RGB(0, 0, 0)
                                    .Font.Bold = True
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                End If
                            End If
                        End With
                    End If
                End If
            Next d
        End If

        actualDrawEnd = Empty
        If IsDate(actS) Then
            If IsDate(actE) Then
                actualDrawEnd = CDate(actE)
            ElseIf CDate(actS) <= Date Then
                actualDrawEnd = Date
            End If
        End If

        If IsDate(actS) And IsDate(actualDrawEnd) Then
            For d = CDate(actS) To CDate(actualDrawEnd)
                If d >= chartStartDate And d <= chartEndDate Then
                    If IsWorkingDay(d, holidayDict, workdayDict) Then
                        colIdx = startCol + CLng(d - chartStartDate)
                        With ws.Cells(r, colIdx)
                            .Interior.Color = actualColor
                            .Value = markerText
                            .Font.Color = RGB(0, 0, 0)
                            .Font.Bold = True
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                    End If
                End If
            Next d
        End If

ContinueNextRow:
    Next r
End Sub

Public Sub ApplyHierarchySummaryValues(ws As Worksheet, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim r As Long
    
    For r = lastRow To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            If HasChildTask(ws, r, lastRow) Then
                SetParentSummaryValues ws, r, lastRow, holidayDict, workdayDict
            Else
                NormalizeLeafProgressByActualEnd ws, r
            End If
            
            SetNormalProgressValue ws, r, holidayDict, workdayDict
            SetDurationValues ws, r, holidayDict, workdayDict
            SetStatusValue ws, r, lastRow, holidayDict, workdayDict
        End If
    Next r
End Sub

Public Sub UpdateTaskWorkMinutes(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                 Optional ByVal firstRow As Long = DATA_START_ROW)
    Dim r As Long, eventsEnabled As Boolean, wasProtected As Boolean, restoreProtection As Boolean
    Dim errorNumber As Long, errorText As String
    If ws Is Nothing Or lastRow < DATA_START_ROW Then Exit Sub
    firstRow = Application.Max(DATA_START_ROW, firstRow)
    If firstRow > lastRow Then Exit Sub
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or _
       IsWebUtilitySheet(ws) Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) <> "요청 내용" Or _
       CStr(ws.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2) <> "작업시간(분)" Then Exit Sub
    eventsEnabled = Application.EnableEvents
    wasProtected = ws.ProtectContents
    If wasProtected Then restoreProtection = Not ws.ProtectionMode
    On Error GoTo Failed
    Application.EnableEvents = False
    If restoreProtection Then UnprotectTaskSheet ws
    For r = lastRow To firstRow Step -1
        If HasTaskContent(ws, r) Then
            If HasChildTask(ws, r, lastRow) Then SetParentWorkMinutes ws, r, lastRow
        End If
    Next r
    If restoreProtection Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If restoreProtection Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "UpdateTaskWorkMinutes", errorText
End Sub

Private Sub SetParentWorkMinutes(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long)
    Dim child As Long, subtreeEnd As Long, raw As Variant, total As Double, minutes As Double
    Dim hasMinutes As Boolean, invalidMinutes As Boolean
    subtreeEnd = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    child = rowNum + 1
    Do While child <= subtreeEnd
        If HasTaskContent(ws, child) Then
            raw = ws.Cells(child, COL_WORK_MINUTES).Value2
            If IsError(raw) Then
                invalidMinutes = True
            ElseIf Len(Trim$(CStr(raw))) > 0 Then
                If VarType(raw) = vbBoolean Or Not IsNumeric(raw) Then
                    invalidMinutes = True
                Else
                    minutes = CDbl(raw)
                    If minutes < 0 Or minutes > 2147483647# Or minutes <> Fix(minutes) Then
                        invalidMinutes = True
                    Else
                        total = total + minutes
                        hasMinutes = True
                    End If
                End If
            End If
            ' Only direct children contribute; each nested subtotal is included once.
            child = GetTaskSubtreeEndRow(ws, child, lastRow) + 1
        Else
            child = child + 1
        End If
    Loop
    If invalidMinutes Then
        ws.Cells(rowNum, COL_WORK_MINUTES).Value = CVErr(xlErrValue)
    ElseIf hasMinutes Then
        ws.Cells(rowNum, COL_WORK_MINUTES).Value2 = total
    Else
        ws.Cells(rowNum, COL_WORK_MINUTES).ClearContents
    End If
End Sub

Private Sub NormalizeLeafProgressByActualEnd(ws As Worksheet, ByVal rowNum As Long)
    If IsDate(ws.Cells(rowNum, COL_ACTUAL_END).Value) Then
        ws.Cells(rowNum, COL_PROGRESS).Value = 1
        ws.Cells(rowNum, COL_PROGRESS).NumberFormat = "0%"
    End If
End Sub

Private Sub SetParentSummaryValues(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim currentLevel As Long
    Dim childRow As Long
    Dim childEndRow As Long
    
    Dim minPlanStart As Date
    Dim maxPlanEnd As Date
    Dim minActStart As Date
    Dim maxActEnd As Date
    
    Dim hasPlanStart As Boolean
    Dim hasPlanEnd As Boolean
    Dim hasActStart As Boolean
    Dim hasActEnd As Boolean
    Dim allChildActualDone As Boolean
    
    Dim vPlanS As Variant
    Dim vPlanE As Variant
    Dim vActS As Variant
    Dim vActE As Variant
    
    Dim childProgress As Double
    Dim childWeight As Double
    Dim weightedProgressSum As Double
    Dim totalWeight As Double

    SetParentWorkMinutes ws, rowNum, lastRow
    
    currentLevel = GetTaskLevel(ws, rowNum)
    allChildActualDone = True
    
    childRow = rowNum + 1
    
    Do While childRow <= lastRow
        If HasTaskContent(ws, childRow) Then
            If GetTaskLevel(ws, childRow) <= currentLevel Then Exit Do
            
            vPlanS = ws.Cells(childRow, COL_PLAN_START).Value
            vPlanE = ws.Cells(childRow, COL_PLAN_END).Value
            vActS = ws.Cells(childRow, COL_ACTUAL_START).Value
            vActE = ws.Cells(childRow, COL_ACTUAL_END).Value
            
            If IsDate(vPlanS) Then
                If Not hasPlanStart Then
                    minPlanStart = CDate(vPlanS)
                    hasPlanStart = True
                ElseIf CDate(vPlanS) < minPlanStart Then
                    minPlanStart = CDate(vPlanS)
                End If
            End If
            
            If IsDate(vPlanE) Then
                If Not hasPlanEnd Then
                    maxPlanEnd = CDate(vPlanE)
                    hasPlanEnd = True
                ElseIf CDate(vPlanE) > maxPlanEnd Then
                    maxPlanEnd = CDate(vPlanE)
                End If
            End If
            
            If IsDate(vActS) Then
                If Not hasActStart Then
                    minActStart = CDate(vActS)
                    hasActStart = True
                ElseIf CDate(vActS) < minActStart Then
                    minActStart = CDate(vActS)
                End If
            End If
            
            If IsDate(vActE) Then
                If Not hasActEnd Then
                    maxActEnd = CDate(vActE)
                    hasActEnd = True
                ElseIf CDate(vActE) > maxActEnd Then
                    maxActEnd = CDate(vActE)
                End If
            Else
                allChildActualDone = False
            End If
            
            childProgress = GetTaskProgressValue(ws, childRow)
            childWeight = GetProgressWeight(ws, childRow, holidayDict, workdayDict)
            
            weightedProgressSum = weightedProgressSum + (childProgress * childWeight)
            totalWeight = totalWeight + childWeight
            
            childEndRow = GetTaskSubtreeEndRow(ws, childRow, lastRow)
            childRow = childEndRow + 1
        Else
            childRow = childRow + 1
        End If
    Loop
    
    If hasPlanStart Then
        ws.Cells(rowNum, COL_PLAN_START).Value = minPlanStart
    Else
        ws.Cells(rowNum, COL_PLAN_START).ClearContents
    End If
    
    If hasPlanEnd Then
        ws.Cells(rowNum, COL_PLAN_END).Value = maxPlanEnd
    Else
        ws.Cells(rowNum, COL_PLAN_END).ClearContents
    End If
    
    If hasActStart Then
        ws.Cells(rowNum, COL_ACTUAL_START).Value = minActStart
    Else
        ws.Cells(rowNum, COL_ACTUAL_START).ClearContents
    End If
    
    If allChildActualDone And hasActEnd Then
        ws.Cells(rowNum, COL_ACTUAL_END).Value = maxActEnd
    Else
        ws.Cells(rowNum, COL_ACTUAL_END).ClearContents
    End If
    
    If Not IsManualProgressEnabled(ws, rowNum) Then
        If totalWeight > 0 Then
            ws.Cells(rowNum, COL_PROGRESS).Value = weightedProgressSum / totalWeight
        Else
            ws.Cells(rowNum, COL_PROGRESS).ClearContents
        End If
    End If
End Sub

Private Sub SetNormalProgressValue(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim planS As Variant
    Dim planE As Variant
    Dim actE As Variant
    Dim requiredProgress As Double
    
    actE = ws.Cells(rowNum, COL_ACTUAL_END).Value
    
    If IsDate(actE) Then
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).Value = 1
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).NumberFormat = "0%"
        Exit Sub
    End If
    
    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value
    
    If IsDate(planS) And IsDate(planE) Then
        requiredProgress = GetRequiredNormalProgress(CDate(planS), CDate(planE), holidayDict, workdayDict)
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).Value = requiredProgress
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).NumberFormat = "0%"
    Else
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).ClearContents
    End If
End Sub

Private Function IsManualProgressEnabled(ws As Worksheet, ByVal rowNum As Long) As Boolean
    IsManualProgressEnabled = (UCase$(Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_PROGRESS).Value))) = "Y")
End Function

Private Function GetProgressWeight(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object) As Double
    Dim v As Variant
    Dim calcDays As Variant
    
    v = ws.Cells(rowNum, COL_PLAN_DAYS).Value
    
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            GetProgressWeight = CDbl(v)
            Exit Function
        End If
    End If
    
    calcDays = CountWorkingDaysInclusive(ws.Cells(rowNum, COL_PLAN_START).Value, ws.Cells(rowNum, COL_PLAN_END).Value, holidayDict, workdayDict)
    
    If IsNumeric(calcDays) Then
        If CDbl(calcDays) > 0 Then
            GetProgressWeight = CDbl(calcDays)
            Exit Function
        End If
    End If
    
    GetProgressWeight = 1
End Function

Private Sub SetDurationValues(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim planDays As Variant
    Dim actualDays As Variant
    Dim actualStart As Variant
    Dim actualEnd As Variant
    Dim compareEnd As Date

    planDays = CountWorkingDaysInclusive(ws.Cells(rowNum, COL_PLAN_START).Value, ws.Cells(rowNum, COL_PLAN_END).Value, holidayDict, workdayDict)

    actualStart = ws.Cells(rowNum, COL_ACTUAL_START).Value
    actualEnd = ws.Cells(rowNum, COL_ACTUAL_END).Value
    
    If IsDate(actualStart) Then
        If IsDate(actualEnd) Then
            actualDays = CountWorkingDaysInclusive(actualStart, actualEnd, holidayDict, workdayDict)
        Else
            If CDate(actualStart) <= Date Then
                compareEnd = Date
                actualDays = CountWorkingDaysInclusive(actualStart, compareEnd, holidayDict, workdayDict)
            Else
                actualDays = ""
            End If
        End If
    Else
        actualDays = ""
    End If

    ws.Cells(rowNum, COL_PLAN_DAYS).Value = planDays
    ws.Cells(rowNum, COL_ACTUAL_DAYS).Value = actualDays

    ws.Cells(rowNum, COL_PLAN_DAYS).NumberFormat = "0"
    ws.Cells(rowNum, COL_ACTUAL_DAYS).NumberFormat = "0"
End Sub

Private Sub SetStatusValue(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    If HasChildTask(ws, rowNum, lastRow) Then
        ws.Cells(rowNum, COL_STATUS).Value = BuildParentStatusFromChildren(ws, rowNum, lastRow, holidayDict, workdayDict)
    Else
        ws.Cells(rowNum, COL_STATUS).Value = GetTaskStatus(ws, rowNum, holidayDict, workdayDict)
    End If
    ApplyTaskDelayStatusDisplay ws, rowNum, lastRow, holidayDict, workdayDict
End Sub

Private Sub ApplyTaskDelayStatusDisplay(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                        ByVal lastRow As Long, ByVal holidayDict As Object, _
                                        ByVal workdayDict As Object)
    Dim days As Variant
    ws.Cells(rowNum, COL_STATUS).NumberFormat = "General"
    If HasTaskError(ws, rowNum) Then Exit Sub
    days = GetTaskDelayDays(ws, rowNum, holidayDict, workdayDict)
    If IsNumeric(days) Then
        If CLng(days) > 0 Then
            ws.Cells(rowNum, COL_STATUS).NumberFormat = "@" & Chr$(34) & _
                                                       "(" & CLng(days) & ")" & Chr$(34)
        End If
    End If
End Sub

Private Function BuildParentStatusFromChildren(ws As Worksheet, ByVal rowNum As Long, _
                                                ByVal lastRow As Long, ByVal holidayDict As Object, _
                                                ByVal workdayDict As Object) As String
    Dim currentLevel As Long
    Dim childRow As Long
    Dim childEndRow As Long
    Dim childStatus As String
    
    Dim hasAnyChild As Boolean
    Dim hasError As Boolean
    Dim hasDelay As Boolean
    Dim hasCaution As Boolean
    Dim hasToday As Boolean
    Dim hasHold As Boolean
    Dim allDone As Boolean
    
    currentLevel = GetTaskLevel(ws, rowNum)
    allDone = True
    childRow = rowNum + 1
    
    Do While childRow <= lastRow
        If HasTaskContent(ws, childRow) Then
            If GetTaskLevel(ws, childRow) <= currentLevel Then Exit Do
            
            childStatus = CStr(ws.Cells(childRow, COL_STATUS).Value)
            hasAnyChild = True
            
            If IsDoneStatusText(childStatus) Then
                If HasHoldIncludedStatusText(childStatus) Then hasHold = True
            Else
                Select Case childStatus
                    Case STATUS_HOLD
                        hasHold = True
                        allDone = False
                    Case STATUS_ERROR
                        hasError = True
                        allDone = False
                    Case STATUS_DELAY
                        hasDelay = True
                        allDone = False
                    Case STATUS_CAUTION
                        hasCaution = True
                        allDone = False
                    Case STATUS_TODAY
                        hasToday = True
                        allDone = False
                    Case Else
                        allDone = False
                End Select
            End If
            
            childEndRow = GetTaskSubtreeEndRow(ws, childRow, lastRow)
            childRow = childEndRow + 1
        Else
            childRow = childRow + 1
        End If
    Loop
    
    If HasDelayedDescendantTask(ws, rowNum, lastRow, holidayDict, workdayDict) Then hasDelay = True
    If Not hasAnyChild Then
        BuildParentStatusFromChildren = ""
    ElseIf hasError Then
        BuildParentStatusFromChildren = STATUS_ERROR
    ElseIf allDone Then
        If hasHold Then
            BuildParentStatusFromChildren = STATUS_DONE & STATUS_DONE_WITH_HOLD_SUFFIX
        Else
            BuildParentStatusFromChildren = STATUS_DONE
        End If
    ElseIf hasDelay Then
        BuildParentStatusFromChildren = STATUS_DELAY
    ElseIf hasCaution Then
        BuildParentStatusFromChildren = STATUS_CAUTION
    ElseIf hasHold Then
        BuildParentStatusFromChildren = STATUS_HOLD
    ElseIf hasToday Then
        BuildParentStatusFromChildren = STATUS_TODAY
    Else
        BuildParentStatusFromChildren = STATUS_NORMAL
    End If
End Function

Private Function GetTaskStatusFillColor(ByVal statusText As String) As Long
    If IsDoneStatusText(statusText) Then
        GetTaskStatusFillColor = RGB(102, 255, 51)
        Exit Function
    End If

    Select Case statusText
        Case STATUS_NORMAL
            GetTaskStatusFillColor = RGB(91, 155, 213)
        Case STATUS_CAUTION
            GetTaskStatusFillColor = RGB(255, 192, 0)
        Case STATUS_TODAY
            GetTaskStatusFillColor = RGB(45, 160, 155)
        Case STATUS_HOLD
            GetTaskStatusFillColor = RGB(153, 102, 204)
        Case STATUS_DELAY
            GetTaskStatusFillColor = RGB(255, 99, 99)
        Case STATUS_ERROR
            GetTaskStatusFillColor = RGB(237, 125, 49)
        Case Else
            GetTaskStatusFillColor = RGB(91, 155, 213)
    End Select
End Function

Private Function IsTaskPlannedOnDate(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim planS As Variant
    Dim planE As Variant

    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value

    If Not IsDate(planS) Or Not IsDate(planE) Then Exit Function
    If Not IsWorkingDay(targetDate, holidayDict, workdayDict) Then Exit Function

    IsTaskPlannedOnDate = (CLng(targetDate) >= CLng(CDate(planS)) And _
                           CLng(targetDate) <= CLng(CDate(planE)))
End Function

Private Function IsTaskActualOnDate(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim actualS As Variant
    Dim actualE As Variant
    Dim actualDrawEnd As Date

    actualS = ws.Cells(rowNum, COL_ACTUAL_START).Value
    actualE = ws.Cells(rowNum, COL_ACTUAL_END).Value

    If Not IsDate(actualS) Then Exit Function
    If Not IsWorkingDay(targetDate, holidayDict, workdayDict) Then Exit Function

    If IsDate(actualE) Then
        actualDrawEnd = CDate(actualE)
    ElseIf CDate(actualS) <= Date Then
        actualDrawEnd = Date
    Else
        Exit Function
    End If

    IsTaskActualOnDate = (CLng(targetDate) >= CLng(CDate(actualS)) And _
                          CLng(targetDate) <= CLng(actualDrawEnd))
End Function

Private Function IsDoneStatusText(ByVal statusText As String) As Boolean
    IsDoneStatusText = (statusText = STATUS_DONE Or statusText = STATUS_DONE & STATUS_DONE_WITH_HOLD_SUFFIX)
End Function

Private Function HasHoldIncludedStatusText(ByVal statusText As String) As Boolean
    HasHoldIncludedStatusText = (InStr(1, statusText, STATUS_DONE_WITH_HOLD_SUFFIX, vbTextCompare) > 0)
End Function

Public Sub FormatBaseArea(ws As Worksheet, ByVal lastRow As Long, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object, _
                          Optional ByVal fitLayout As Boolean = True, Optional ByVal rebuildHierarchy As Boolean = True)
    Dim startCol As Long
    Dim endCol As Long
    Dim r As Long
    Dim c As Long
    Dim d As Date
    Dim prevDate As Date
    Dim filterLastRow As Long
    Dim statusText As String
    Dim dataLastCol As String
    Dim levelValue As Long
    Dim baseTaskText As String
    
    Dim detailColumnHidden As Variant

    detailColumnHidden = CaptureTaskDetailColumnHidden(ws)
    dataLastCol = COL_STATUS
    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = startCol + CLng(chartEndDate - chartStartDate)
    filterLastRow = IIf(lastRow < HEADER_ROW + 1, HEADER_ROW + 1, lastRow)

    ' Keep Excel's native filter object: recreating it discards every active criterion.
    ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).Borders.LineStyle = xlNone
    ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).Interior.Pattern = xlNone
    ws.Range(ws.Cells(DATA_START_ROW, startCol), ws.Cells(filterLastRow, endCol)).Borders.LineStyle = xlNone

    With ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & HEADER_ROW)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With

    With ws.Range(COL_PLAN_DAYS & HEADER_ROW & ":" & COL_STATUS & HEADER_ROW)
        .Interior.Color = RGB(217, 217, 217)
    End With

    If Not ws.AutoFilterMode Then
        ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).AutoFilter
    ElseIf Not ws.FilterMode Then
        ' Extend only our default, unfiltered range; retain user-selected filter ranges.
        With ws.AutoFilter.Range
            If .Row = HEADER_ROW And .Column = ws.Range(COL_NO & "1").Column And _
               .Columns.Count = ws.Range(COL_NO & "1:" & dataLastCol & "1").Columns.Count And _
               .Rows.Count <> filterLastRow - HEADER_ROW + 1 Then
                ws.AutoFilterMode = False
                ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).AutoFilter
            End If
        End With
    End If

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 210)
    End With

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            With ws.Range(COL_NO & r & ":" & dataLastCol & r)
                .Borders.LineStyle = xlContinuous
                .VerticalAlignment = xlCenter
            End With
            
            With ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(210, 210, 210)
            End With
            
            levelValue = GetTaskLevel(ws, r)
            baseTaskText = RemoveTaskLevelPrefix(CStr(ws.Cells(r, COL_TASK).Value))
            ws.Cells(r, COL_TASK).Value = baseTaskText
            ws.Cells(r, COL_TASK).HorizontalAlignment = xlLeft
            ws.Cells(r, COL_MODIFICATION).HorizontalAlignment = xlLeft
            
            ws.Cells(r, COL_NO).HorizontalAlignment = xlCenter
            ws.Cells(r, COL_NO).VerticalAlignment = xlCenter
            ws.Cells(r, COL_LEVEL).HorizontalAlignment = xlCenter
            ws.Cells(r, COL_LEVEL).VerticalAlignment = xlCenter
            
            statusText = CStr(ws.Cells(r, COL_STATUS).Value)
            
            If IsDoneStatusText(statusText) Then
                ws.Cells(r, COL_STATUS).Interior.Color = RGB(198, 239, 206)
            Else
                Select Case statusText
                    Case STATUS_NORMAL
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(221, 235, 247)
                    Case STATUS_CAUTION
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(255, 235, 156)
                    Case STATUS_TODAY
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(229, 247, 245)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(158, 225, 219)
                    Case STATUS_HOLD
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(217, 210, 233)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(180, 167, 214)
                    Case STATUS_DELAY
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(255, 235, 235)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(255, 199, 206)
                    Case STATUS_ERROR
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(255, 242, 204)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(244, 176, 132)
                End Select
            End If
        Else
            ws.Range(COL_NO & r & ":" & dataLastCol & r).Borders.LineStyle = xlNone
            If endCol >= startCol Then
                ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.LineStyle = xlNone
            End If
        End If
    Next r

    ' The final renderer paints weekend/today backgrounds once.
    For c = startCol To endCol
        d = chartStartDate + (c - startCol)
        If c > startCol Then
            prevDate = chartStartDate + (c - startCol - 1)
            
            If Month(prevDate) <> Month(d) Or Year(prevDate) <> Year(d) Then
                With ws.Range(ws.Cells(DATA_START_ROW, c), ws.Cells(lastRow, c)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(160, 160, 160)
                End With
            End If
        End If
    Next c

    If fitLayout Then
    ws.Range(COL_REQUEST & ":" & COL_MODIFICATION).WrapText = False
    ws.Columns(COL_WORK_MINUTES).ColumnWidth = 13
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
    ws.Columns(COL_NOTE).ColumnWidth = 4.5
    ws.Columns(COL_PLAN_START).ColumnWidth = 11
    ws.Columns(COL_PLAN_END).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_START).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_END).ColumnWidth = 11
    ws.Columns(COL_PROGRESS).ColumnWidth = 9
    ws.Columns(COL_NORMAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_MANUAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_PLAN_DAYS).ColumnWidth = 9
    ws.Columns(COL_ACTUAL_DAYS).ColumnWidth = 10

    ' 새로고침 후 화면의 문자열이 열 너비 때문에 잘리지 않도록
    ' 분류, 업무명, 담당자 및 상태 열을 실제 내용에 맞춘다.
    ws.Range(COL_NO & HEADER_ROW & ":" & _
             COL_MODIFICATION & lastRow).Columns.AutoFit
    ws.Range(COL_OWNER & HEADER_ROW & ":" & _
             COL_OWNER & lastRow).Columns.AutoFit
    ws.Range(COL_MANUAL_STATUS & HEADER_ROW & ":" & _
             COL_WEEKLY_REPORT & lastRow).Columns.AutoFit
    ws.Range(COL_STATUS & HEADER_ROW & ":" & _
             COL_STATUS & lastRow).Columns.AutoFit
    ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(lastRow)).AutoFit
    End If

    ws.Rows(GANTT_HEADER_ROW_MONTH).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_WEEK).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_DAY).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_DATE).RowHeight = 70

    ws.Range(COL_PLAN_START & DATA_START_ROW & ":" & COL_ACTUAL_END & lastRow).NumberFormat = "yyyy-mm-dd"
    ws.Range(COL_PROGRESS & DATA_START_ROW & ":" & COL_PROGRESS & lastRow).NumberFormat = "0%"
    ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).NumberFormat = "0%"
    ApplyTaskHierarchyAppearance ws, lastRow, startCol, endCol, fitLayout, rebuildHierarchy, True
    RestoreTaskDetailColumnHidden ws, detailColumnHidden
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub HandleTaskHierarchyChange(ByVal ws As Worksheet, ByVal Target As Range, _
                                     Optional ByVal taskStructureChanged As Boolean = True)
    Dim levelChanges As Range, classificationChanges As Range, bodyChanges As Range
    Dim ownerChanges As Range, minuteChanges As Range, structureChanges As Range
    Dim classChanges As Range, ownerAffected As Range, minuteAffected As Range, changedArea As Range
    Dim lastRow As Long, r As Long, firstChanged As Long, lastChanged As Long
    Dim firstGroupRow As Long, lastGroupRow As Long, changedRows As Object
    Dim issueMessages As String, structureProtectionNeeded As Boolean
    Dim errorNumber As Long, errorText As String

    If ws Is Nothing Or Target Is Nothing Then Exit Sub
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or _
       IsWebUtilitySheet(ws) Then Exit Sub

    Set levelChanges = Intersect(Target, ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_LEVEL & ws.Rows.Count))
    Set classificationChanges = Intersect(Target, ws.Range(COL_TYPE & DATA_START_ROW & ":" & COL_MINOR_CATEGORY & ws.Rows.Count))
    Set bodyChanges = Intersect(Target, ws.Range(COL_REQUEST & DATA_START_ROW & ":" & COL_MODIFICATION & ws.Rows.Count))
    Set ownerChanges = Intersect(Target, ws.Range(COL_OWNER & DATA_START_ROW & ":" & COL_OWNER & ws.Rows.Count))
    Set minuteChanges = Intersect(Target, ws.Range(COL_WORK_MINUTES & DATA_START_ROW & ":" & COL_WORK_MINUTES & ws.Rows.Count))

    Set classChanges = levelChanges
    If Not classificationChanges Is Nothing Then
        If classChanges Is Nothing Then
            Set classChanges = classificationChanges
        Else
            Set classChanges = Union(classChanges, classificationChanges)
        End If
    End If
    Set structureChanges = levelChanges
    If taskStructureChanged And Not bodyChanges Is Nothing Then
        If structureChanges Is Nothing Then
            Set structureChanges = bodyChanges
        Else
            Set structureChanges = Union(structureChanges, bodyChanges)
        End If
    End If
    Set ownerAffected = ownerChanges
    Set minuteAffected = minuteChanges
    If Not structureChanges Is Nothing Then
        If ownerAffected Is Nothing Then Set ownerAffected = structureChanges Else Set ownerAffected = Union(ownerAffected, structureChanges)
        If minuteAffected Is Nothing Then Set minuteAffected = structureChanges Else Set minuteAffected = Union(minuteAffected, structureChanges)
    End If
    If classChanges Is Nothing And ownerAffected Is Nothing And minuteAffected Is Nothing Then Exit Sub

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then Exit Sub
    If Not structureChanges Is Nothing Then
        If ws.ProtectContents Then structureProtectionNeeded = ws.ProtectionMode
    End If
    If Not classChanges Is Nothing Then
        Set changedRows = CreateObject("Scripting.Dictionary")
        firstChanged = lastRow + 1
        lastChanged = DATA_START_ROW - 1
        For Each changedArea In classChanges.Areas
            firstChanged = Application.Min(firstChanged, changedArea.Row)
            lastChanged = Application.Max(lastChanged, Application.Min(lastRow, changedArea.Row + changedArea.Rows.Count - 1))
            For r = changedArea.Row To Application.Min(lastRow, changedArea.Row + changedArea.Rows.Count - 1)
                changedRows(CStr(r)) = True
            Next r
        Next changedArea
        For r = firstChanged To lastChanged
            If changedRows.Exists(CStr(r)) Then
                EnforceTaskModuleFromParent ws, r, issueMessages
                PropagateTaskModuleToDescendants ws, r, lastRow
            End If
        Next r
    End If

    If Not ownerAffected Is Nothing Then
        TaskChangeGroupBounds ws, ownerAffected, lastRow, Not structureChanges Is Nothing, firstGroupRow, lastGroupRow
        SynchronizeTaskHierarchyOwners ws, lastGroupRow, firstGroupRow
    End If
    If Not minuteAffected Is Nothing Then
        TaskChangeGroupBounds ws, minuteAffected, lastRow, Not structureChanges Is Nothing, firstGroupRow, lastGroupRow
        UpdateTaskWorkMinutes ws, lastGroupRow, firstGroupRow
    End If
    If structureProtectionNeeded Then
        On Error GoTo ProtectionFailed
        UnprotectTaskSheet ws
        ApplyCalculatedColumnsProtection ws, lastRow
        On Error GoTo 0
    End If

    If Len(issueMessages) > 0 Then
        MsgBox "분류 계층 정합성을 자동으로 수정했습니다." & vbCrLf & vbCrLf & _
               issueMessages & vbCrLf & _
               "하위 작업의 네 분류 값은 가장 가까운 부모 작업과 같아야 합니다.", _
               vbExclamation, _
               "분류 정합성 오류"
    End If
    Exit Sub
ProtectionFailed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    ApplyCalculatedColumnsProtection ws, lastRow
    On Error GoTo 0
    Err.Raise errorNumber, "HandleTaskHierarchyChange", errorText
End Sub

Private Sub TaskChangeGroupBounds(ByVal ws As Worksheet, ByVal changed As Range, _
                                  ByVal lastRow As Long, ByVal includePreviousRoot As Boolean, _
                                  ByRef firstGroupRow As Long, ByRef lastGroupRow As Long)
    Dim area As Range, firstChanged As Long, lastChanged As Long, searchRow As Long, r As Long
    firstChanged = lastRow
    lastChanged = DATA_START_ROW
    For Each area In changed.Areas
        firstChanged = Application.Min(firstChanged, area.Row)
        lastChanged = Application.Max(lastChanged, Application.Min(lastRow, area.Row + area.Rows.Count - 1))
    Next area
    searchRow = firstChanged
    ' A new/removed level1 row also changes the group immediately above it.
    If includePreviousRoot Then searchRow = searchRow - 1
    firstGroupRow = DATA_START_ROW
    For r = searchRow To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) = 1 Then
                firstGroupRow = r
                Exit For
            End If
        End If
    Next r
    lastGroupRow = lastRow
    For r = lastChanged + 1 To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) = 1 Then
                lastGroupRow = r - 1
                Exit For
            End If
        End If
    Next r
End Sub

Public Sub SynchronizeTaskHierarchyModules(ByVal ws As Worksheet, _
                                           ByVal lastRow As Long, _
                                           Optional ByVal showIssueMessage As Boolean = True)
    Dim r As Long
    Dim issueMessages As String

    If ws Is Nothing Then Exit Sub
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Or _
           Len(Trim$(CStr(ws.Cells(r, COL_LEVEL).Value2))) > 0 Then
            EnforceTaskModuleFromParent ws, r, issueMessages
        End If
    Next r

    SynchronizeTaskHierarchyOwners ws, lastRow

    If showIssueMessage And Len(issueMessages) > 0 Then
        MsgBox "기존 분류 계층 오류를 자동으로 수정했습니다." & vbCrLf & vbCrLf & _
               issueMessages, _
               vbExclamation, _
               "분류 정합성 오류"
    End If
End Sub

Public Sub SynchronizeTaskHierarchyOwners(ByVal ws As Worksheet, _
                                          ByVal lastRow As Long, _
                                          Optional ByVal firstRow As Long = DATA_START_ROW)
    Dim r As Long
    Dim ownerText As String

    If ws Is Nothing Then Exit Sub
    If lastRow < DATA_START_ROW Then Exit Sub
    firstRow = Application.Max(DATA_START_ROW, firstRow)
    If firstRow > lastRow Then Exit Sub

    For r = lastRow To firstRow Step -1
        If HasTaskContent(ws, r) And HasChildTask(ws, r, lastRow) Then
            ownerText = GetDescendantLeafOwnerText(ws, r, lastRow)
            If Len(ownerText) = 0 Then
                ws.Cells(r, COL_OWNER).ClearContents
            Else
                ws.Cells(r, COL_OWNER).Value = ownerText
            End If
        End If
    Next r
End Sub

Private Function GetDescendantLeafOwnerText(ByVal ws As Worksheet, _
                                            ByVal parentRow As Long, _
                                            ByVal lastRow As Long) As String
    Dim ownerSet As Object
    Dim parentLevel As Long
    Dim r As Long

    Set ownerSet = CreateObject("Scripting.Dictionary")
    ownerSet.CompareMode = vbTextCompare
    parentLevel = GetTaskLevel(ws, parentRow)

    For r = parentRow + 1 To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) <= parentLevel Then Exit For
            If Not HasChildTask(ws, r, lastRow) Then
                AddDistinctOwnerNames ownerSet, CStr(ws.Cells(r, COL_OWNER).Value2)
            End If
        End If
    Next r

    GetDescendantLeafOwnerText = JoinOwnerNameSet(ownerSet, ", ")
End Function

Public Sub AddDistinctOwnerNames(ByVal ownerSet As Object, _
                                 ByVal rawOwnerText As String)
    Dim ownerNames As Variant
    Dim ownerName As Variant
    Dim normalizedText As String

    normalizedText = Trim$(rawOwnerText)
    If Len(normalizedText) = 0 Then Exit Sub

    If Left$(normalizedText, 1) = "(" And Right$(normalizedText, 1) = ")" Then
        normalizedText = Trim$(Mid$(normalizedText, 2, Len(normalizedText) - 2))
    End If
    normalizedText = Replace$(normalizedText, "，", ",")
    ownerNames = Split(normalizedText, ",")

    For Each ownerName In ownerNames
        normalizedText = Trim$(CStr(ownerName))
        If Len(normalizedText) > 0 Then
            If Not ownerSet.Exists(normalizedText) Then ownerSet.Add normalizedText, True
        End If
    Next ownerName
End Sub

Public Function JoinOwnerNameSet(ByVal ownerSet As Object, _
                                 ByVal delimiter As String) As String
    Dim ownerName As Variant
    Dim result As String

    For Each ownerName In ownerSet.Keys
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(ownerName)
    Next ownerName

    JoinOwnerNameSet = result
End Function

Public Sub ShowTaskInputErrorReasons(ByVal ws As Worksheet, ByVal lastRow As Long)
    Const MAX_DISPLAY_ERROR_COUNT As Long = 20
    Dim r As Long
    Dim errorCount As Long
    Dim errorText As String
    Dim reasonText As String
    Dim taskText As String
    Dim taskNo As String

    If ws Is Nothing Then Exit Sub

    For r = DATA_START_ROW To lastRow
        reasonText = GetTaskErrorReason(ws, r)
        If Len(reasonText) > 0 Then
            errorCount = errorCount + 1
            If errorCount <= MAX_DISPLAY_ERROR_COUNT Then
                taskText = Trim$(CStr(ws.Cells(r, COL_TASK).Value2))
                taskNo = Trim$(CStr(ws.Cells(r, COL_NO).Value2))
                If Len(taskText) = 0 Then taskText = "요청 내용 미입력"
                If Len(taskNo) = 0 Then taskNo = "-"

                If Len(errorText) > 0 Then errorText = errorText & vbCrLf
                errorText = errorText & "- 행 " & r & " / No. " & taskNo & _
                            " [" & taskText & "]: " & reasonText
            End If
        End If
    Next r

    If errorCount > MAX_DISPLAY_ERROR_COUNT Then
        errorText = errorText & vbCrLf & "- 그 외 " & _
                    (errorCount - MAX_DISPLAY_ERROR_COUNT) & "건"
    End If

    If errorCount > 0 Then
        MsgBox "입력 오류 " & errorCount & "건이 있습니다." & vbCrLf & vbCrLf & _
               errorText, _
               vbExclamation, _
               "간트 입력 오류 상세"
    End If
End Sub

Public Function GetNearestParentTaskRow(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long) As Long
    Dim currentLevel As Long
    Dim candidateLevel As Long
    Dim r As Long

    currentLevel = GetTaskLevel(ws, rowNum)
    If currentLevel <= 1 Then Exit Function

    For r = rowNum - 1 To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            candidateLevel = GetTaskLevel(ws, r)
            If candidateLevel < currentLevel Then
                GetNearestParentTaskRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub EnforceTaskModuleFromParent(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long, _
                                        ByRef issueMessages As String)
    Dim taskLevel As Long
    Dim parentRow As Long
    Dim parentTask As String
    Dim childTask As String
    Dim parentValues As Variant
    Dim childValues As Variant

    taskLevel = GetTaskLevel(ws, rowNum)
    If taskLevel <= 1 Then Exit Sub

    parentRow = GetNearestParentTaskRow(ws, rowNum)
    childTask = Trim$(CStr(ws.Cells(rowNum, COL_TASK).Value2))

    If parentRow = 0 Then
        ws.Range(COL_TYPE & rowNum & ":" & COL_MINOR_CATEGORY & rowNum).ClearContents
        AppendModuleConsistencyIssue issueMessages, rowNum, childTask, _
            "Level " & taskLevel & "이지만 위쪽에 부모 작업이 없습니다. " & _
            "부모 작업을 먼저 만들거나 Level을 1로 변경하세요."
        Exit Sub
    End If

    parentTask = Trim$(CStr(ws.Cells(parentRow, COL_TASK).Value2))
    parentValues = ws.Range(COL_TYPE & parentRow & ":" & _
                            COL_MINOR_CATEGORY & parentRow).Value2
    childValues = ws.Range(COL_TYPE & rowNum & ":" & _
                           COL_MINOR_CATEGORY & rowNum).Value2

    If Not ClassificationValuesEqual(parentValues, childValues) Then
        ws.Range(COL_TYPE & rowNum & ":" & _
                 COL_MINOR_CATEGORY & rowNum).Value2 = parentValues
        AppendModuleConsistencyIssue issueMessages, rowNum, childTask, _
            "입력한 분류가 부모 작업 '" & parentTask & "'(행 " & parentRow & _
            ")과 달라 부모의 타입/대분류/중분류/소분류로 맞췄습니다."
    End If
End Sub

Private Function ClassificationValuesEqual(ByVal leftValues As Variant, _
                                             ByVal rightValues As Variant) As Boolean
    Dim i As Long

    For i = 1 To 4
        If StrComp(Trim$(CStr(leftValues(1, i))), _
                   Trim$(CStr(rightValues(1, i))), vbTextCompare) <> 0 Then Exit Function
    Next i
    ClassificationValuesEqual = True
End Function

Private Sub PropagateTaskModuleToDescendants(ByVal ws As Worksheet, _
                                             ByVal rowNum As Long, _
                                             ByVal lastRow As Long)
    Dim parentLevel As Long
    Dim subtreeEndRow As Long
    Dim classificationValues As Variant
    Dim r As Long

    parentLevel = GetTaskLevel(ws, rowNum)
    subtreeEndRow = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    If subtreeEndRow <= rowNum Then Exit Sub

    classificationValues = ws.Range(COL_TYPE & rowNum & ":" & _
                                    COL_MINOR_CATEGORY & rowNum).Value2
    For r = rowNum + 1 To subtreeEndRow
        If GetTaskLevel(ws, r) > parentLevel And _
           (HasTaskContent(ws, r) Or _
            Len(Trim$(CStr(ws.Cells(r, COL_LEVEL).Value2))) > 0) Then
            ws.Range(COL_TYPE & r & ":" & _
                     COL_MINOR_CATEGORY & r).Value2 = classificationValues
        End If
    Next r
End Sub

Private Sub AppendModuleConsistencyIssue(ByRef issueMessages As String, _
                                         ByVal rowNum As Long, _
                                         ByVal taskText As String, _
                                         ByVal reasonText As String)
    If Len(taskText) = 0 Then taskText = "요청 내용 미입력"
    If Len(issueMessages) > 0 Then issueMessages = issueMessages & vbCrLf
    issueMessages = issueMessages & "- 행 " & rowNum & " [" & taskText & "]: " & reasonText
End Sub

Public Sub ApplyTaskInputValidation(ws As Worksheet)
    Dim lastSheetRow As Long
    Dim rngDate As Range
    Dim rngClassification As Range
    Dim rngLevel As Range
    Dim rngProgress As Range
    Dim rngManualProgress As Range
    Dim rngWeeklyReport As Range
    
    lastSheetRow = ws.Rows.Count

    Set rngClassification = ws.Range(COL_TYPE & DATA_START_ROW & ":" & _
                                     COL_MINOR_CATEGORY & lastSheetRow)
    On Error Resume Next
    rngClassification.Validation.Delete
    On Error GoTo 0

    EnsureTaskTypeConfig
    ApplyTaskTypeValidation ws
    ApplyTaskTextLengthValidation ws
    
    Set rngDate = Union( _
        ws.Range(COL_PLAN_START & DATA_START_ROW & ":" & COL_PLAN_START & lastSheetRow), _
        ws.Range(COL_PLAN_END & DATA_START_ROW & ":" & COL_PLAN_END & lastSheetRow), _
        ws.Range(COL_ACTUAL_START & DATA_START_ROW & ":" & COL_ACTUAL_START & lastSheetRow), _
        ws.Range(COL_ACTUAL_END & DATA_START_ROW & ":" & COL_ACTUAL_END & lastSheetRow))
    
    On Error Resume Next
    rngDate.Validation.Delete
    On Error GoTo 0
    
    rngDate.NumberFormat = "yyyy-mm-dd"
    
    rngDate.Validation.Add Type:=xlValidateDate, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:="2000-01-01", _
                           Formula2:="2100-12-31"
    
    rngDate.Validation.IgnoreBlank = True
    rngDate.Validation.InCellDropdown = True
    rngDate.Validation.InputTitle = "날짜 입력"
    rngDate.Validation.InputMessage = "날짜 형식으로 입력하세요."
    rngDate.Validation.ErrorTitle = "입력 오류"
    rngDate.Validation.ErrorMessage = "올바른 날짜만 입력할 수 있습니다."
    
    Set rngLevel = ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_LEVEL & lastSheetRow)
    
    On Error Resume Next
    rngLevel.Validation.Delete
    On Error GoTo 0
    
    rngLevel.Validation.Add Type:=xlValidateWholeNumber, _
                            AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, _
                            Formula1:="1", _
                            Formula2:="3"
    
    rngLevel.Validation.IgnoreBlank = True
    rngLevel.Validation.InputTitle = "Level 입력"
    rngLevel.Validation.InputMessage = "1~3 사이 정수만 입력하세요."
    rngLevel.Validation.ErrorTitle = "입력 오류"
    rngLevel.Validation.ErrorMessage = "Level은 1~3 사이 정수만 입력할 수 있습니다."
    
    Set rngProgress = ws.Range(COL_PROGRESS & DATA_START_ROW & ":" & COL_PROGRESS & lastSheetRow)
    
    On Error Resume Next
    rngProgress.Validation.Delete
    On Error GoTo 0
    
    rngProgress.NumberFormat = "0%"
    
    rngProgress.Validation.Add Type:=xlValidateDecimal, _
                               AlertStyle:=xlValidAlertStop, _
                               Operator:=xlBetween, _
                               Formula1:="0", _
                               Formula2:="1"
    
    rngProgress.Validation.IgnoreBlank = True
    rngProgress.Validation.InCellDropdown = True
    rngProgress.Validation.InputTitle = "진행률 입력"
    rngProgress.Validation.InputMessage = "0% ~ 100% 사이로 입력하세요. 예: 50%"
    rngProgress.Validation.ErrorTitle = "입력 오류"
    rngProgress.Validation.ErrorMessage = "진행률은 0% ~ 100% 사이만 입력할 수 있습니다."

    Set rngManualProgress = ws.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & COL_MANUAL_PROGRESS & lastSheetRow)

    On Error Resume Next
    rngManualProgress.Validation.Delete
    On Error GoTo 0

    rngManualProgress.Validation.Add Type:=xlValidateList, _
                                     AlertStyle:=xlValidAlertStop, _
                                     Operator:=xlBetween, _
                                     Formula1:="Y,N"

    rngManualProgress.Validation.IgnoreBlank = True
    rngManualProgress.Validation.InCellDropdown = True
    rngManualProgress.Validation.InputTitle = "진행률 수동 입력"
    rngManualProgress.Validation.InputMessage = "수동 입력 허용은 Y, 자동 계산은 N으로 입력하세요."
    rngManualProgress.Validation.ErrorTitle = "입력 오류"
    rngManualProgress.Validation.ErrorMessage = "Y 또는 N만 입력할 수 있습니다."
    
    Set rngWeeklyReport = ws.Range(COL_WEEKLY_REPORT & DATA_START_ROW & ":" & COL_WEEKLY_REPORT & lastSheetRow)

    On Error Resume Next
    rngWeeklyReport.Validation.Delete
    On Error GoTo 0

    rngWeeklyReport.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=REPORT_STATUS_PLANNED & "," & REPORT_STATUS_IN_PROGRESS & "," & REPORT_STATUS_COMPLETED
    rngWeeklyReport.Validation.IgnoreBlank = True
    rngWeeklyReport.Validation.InCellDropdown = True
    rngWeeklyReport.Validation.InputTitle = "보고 상태"
    rngWeeklyReport.Validation.InputMessage = "Planned, In Progress, Completed를 선택하세요."
    rngWeeklyReport.Validation.ErrorTitle = "입력 오류"
    rngWeeklyReport.Validation.ErrorMessage = "Planned, In Progress, Completed만 입력할 수 있습니다."

    ApplyManualStatusValidation ws, lastSheetRow
End Sub

Public Sub ApplyTaskTextLengthValidation(ByVal ws As Worksheet)
    Dim lastValidationRow As Long
    If ws Is Nothing Then Exit Sub
    ' Length is checked by GetTaskErrorReason during refresh. Remove the old volatile input rules.
    lastValidationRow = Application.Max(TASK_VALIDATION_MIN_LAST_ROW, GetLastDataRow(ws) + TASK_VALIDATION_EXTRA_ROWS)
    ws.Range(COL_REQUEST & DATA_START_ROW & ":" & COL_MODIFICATION & lastValidationRow).Validation.Delete
End Sub

Private Sub ApplyManualStatusValidation(ws As Worksheet, ByVal lastSheetRow As Long)
    Dim rngManualStatus As Range

    Set rngManualStatus = ws.Range(COL_MANUAL_STATUS & DATA_START_ROW & ":" & COL_MANUAL_STATUS & lastSheetRow)

    On Error Resume Next
    rngManualStatus.Validation.Delete
    On Error GoTo 0

    rngManualStatus.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=STATUS_HOLD

    rngManualStatus.Validation.IgnoreBlank = True
    rngManualStatus.Validation.InCellDropdown = True
End Sub

Private Function GetWeekdayKorShort(ByVal targetDate As Date) As String
    Select Case Weekday(targetDate, vbSunday)
        Case 1: GetWeekdayKorShort = "일"
        Case 2: GetWeekdayKorShort = "월"
        Case 3: GetWeekdayKorShort = "화"
        Case 4: GetWeekdayKorShort = "수"
        Case 5: GetWeekdayKorShort = "목"
        Case 6: GetWeekdayKorShort = "금"
        Case 7: GetWeekdayKorShort = "토"
    End Select
End Function

Private Function BuildTaskDisplayText(ByVal baseTaskText As String, ByVal levelValue As Long) As String
    Select Case levelValue
        Case 2
            BuildTaskDisplayText = Space$(4) & baseTaskText
        Case 3
            BuildTaskDisplayText = Space$(8) & baseTaskText
        Case Else
            BuildTaskDisplayText = baseTaskText
    End Select
End Function

Private Function RemoveTaskLevelPrefix(ByVal taskText As String) As String
    If Left$(taskText, 8) = Space$(8) Then
        RemoveTaskLevelPrefix = Mid$(taskText, 9)
    ElseIf Left$(taskText, 4) = Space$(4) Then
        RemoveTaskLevelPrefix = Mid$(taskText, 5)
    Else
        RemoveTaskLevelPrefix = taskText
    End If
End Function

Private Function FindHeaderColumnByName(ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(HEADER_ROW, c).Value)) = headerText Then
            FindHeaderColumnByName = c
            Exit Function
        End If
    Next c
End Function

Public Sub ShowAllTaskRows(ws As Worksheet, ByVal lastRow As Long)
    If lastRow < DATA_START_ROW Then Exit Sub
    ws.Rows(DATA_START_ROW & ":" & lastRow).Hidden = False
End Sub

Public Sub ShowAllDateColumns(ws As Worksheet)
    Dim startCol As Long
    Dim endCol As Long

    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = ws.Cells(GANTT_HEADER_ROW_DATE, ws.Columns.Count).End(xlToLeft).Column

    If endCol < startCol Then Exit Sub
    ws.Range(ws.Cells(1, startCol), ws.Cells(1, endCol)).EntireColumn.Hidden = False
End Sub

Public Sub ApplyCalculatedColumnsProtection(ws As Worksheet, ByVal lastRow As Long)
    Dim normalProgressCol As Long
    Dim planDaysCol As Long
    Dim actualDaysCol As Long
    Dim statusCol As Long
    Dim r As Long, c As Long, lastHeaderCol As Long
    Dim hiddenColumns As Range, column As Range
    Dim previousScreenUpdating As Boolean, errorNumber As Long, errorText As String

    ' 2026-10-30 김좌현 수정 : 숨긴 날짜 열이 병합 헤더와 겹치면 Excel의 Locked 설정이 실패하므로 잠금 중에만 표시
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo ProtectionFailed
    Application.ScreenUpdating = False
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then ws.Unprotect
    lastHeaderCol = ws.Cells(GANTT_HEADER_ROW_DATE, ws.Columns.Count).End(xlToLeft).Column
    For c = ws.Range(COL_GANTT_START & "1").Column To lastHeaderCol
        If ws.Columns(c).Hidden Then
            Set column = ws.Columns(c)
            If hiddenColumns Is Nothing Then
                Set hiddenColumns = column
            Else
                Set hiddenColumns = Application.Union(hiddenColumns, column)
            End If
        End If
    Next c
    If Not hiddenColumns Is Nothing Then
        For Each column In hiddenColumns.Areas
            column.EntireColumn.Hidden = False
        Next column
    End If
    ws.Cells.Locked = False

    ws.Range(ws.Cells(DATA_START_ROW, ws.Range(COL_NO & "1").Column), ws.Cells(ws.Rows.Count, ws.Range(COL_NO & "1").Column)).Locked = True

    normalProgressCol = FindHeaderColumnByName(ws, "정상 진행률")
    planDaysCol = FindHeaderColumnByName(ws, "계획일수")
    actualDaysCol = FindHeaderColumnByName(ws, "실소요일수")
    statusCol = FindHeaderColumnByName(ws, "상태")

    If normalProgressCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(ws.Rows.Count, normalProgressCol)).Locked = True
    End If

    If planDaysCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(ws.Rows.Count, planDaysCol)).Locked = True
    End If

    If actualDaysCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, actualDaysCol), ws.Cells(ws.Rows.Count, actualDaysCol)).Locked = True
    End If

    If statusCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(ws.Rows.Count, statusCol)).Locked = True
    End If

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If HasChildTask(ws, r, lastRow) Then
                ws.Cells(r, COL_PROGRESS).Locked = Not IsManualProgressEnabled(ws, r)
            End If
        End If
    Next r

    ws.Rows("1:" & HEADER_ROW).Locked = True
    ws.Rows(DATA_START_ROW & ":" & ws.Rows.Count).Locked = False
    ' 2026-10-22 김좌현 수정 : 정렬/삽입과 함께 이동하는 웹 업무 ID를 사용자 편집에서 보호
    ws.Columns(COL_WEB_TASK_ID).Locked = True
    ApplyWebTaskProtection ws
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True

    If Not hiddenColumns Is Nothing Then
        For Each column In hiddenColumns.Areas
            column.EntireColumn.Hidden = True
        Next column
    End If
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFiltering:=True, _
               AllowInsertingRows:=True, AllowDeletingRows:=True
    ws.EnableOutlining = True
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
ProtectionFailed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not hiddenColumns Is Nothing Then
        For Each column In hiddenColumns.Areas
            column.EntireColumn.Hidden = True
        Next column
    End If
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise errorNumber, "ApplyCalculatedColumnsProtection", errorText
End Sub

Public Sub UnprotectTaskSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
End Sub



