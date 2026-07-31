Attribute VB_Name = "modHoliday"
Option Explicit

Public Sub EnsureConfigSheet()
    Dim ws As Worksheet
    Dim lastSheetRow As Long
    Dim rngType As Range
    Dim rngHideLevel As Range
    Dim rngExcludeNo As Range
    Dim rngExcludeDate As Range
    Dim rngDisplayStart As Range
    Dim rngDisplayEnd As Range
    Dim rngDisplayGanttOnly As Range
    Dim rngDisplayReportOnly As Range
    Dim rngTaskMaxLength As Range
    Dim rngReportLayout As Range
    Dim rngDevReportOwner As Range
    Dim rngReportBullets As Range
    Dim rngWeeklyReportOwner As Range
    Dim rngWeeklyReportPageMode As Range
    Dim rngWeeklyReportOverflowMode As Range
    Dim rngWeeklyReportBullets As Range
    Dim rngWeeklyCustomPageNumbers As Range
    Dim legacyTaskMaxLength As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(LEGACY_CONFIG_SHEET_NAME)
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Name = CONFIG_SHEET_NAME
        End If
    End If

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
    End If

    ws.Range(HOLIDAY_COL_DATE & HOLIDAY_HEADER_ROW).Value = "날짜"
    ws.Range(HOLIDAY_COL_TYPE & HOLIDAY_HEADER_ROW).Value = "구분"
    ws.Range(HOLIDAY_COL_DESC & HOLIDAY_HEADER_ROW).Value = "설명"

    ws.Range("E1").Value = "간트 - 휴일 입력 예시"
    ws.Range("E2").Value = "A열: 날짜"
    ws.Range("E3").Value = "B열: 휴일 또는 근무일"
    ws.Range("E4").Value = "C열: 설명(선택)"

    ws.Range(HIDE_SETTING_TITLE_CELL).Value = "간트 - 숨김 설정"
    ws.Range(HIDE_SETTING_LEVEL_LABEL_CELL).Value = "완료 숨김 레벨"
    ws.Range(DISPLAY_SETTING_TITLE_CELL).Value = "간트 - 표시 기간 설정"
    ws.Range(DISPLAY_SETTING_START_LABEL_CELL).Value = "표시 시작일"
    ws.Range(DISPLAY_SETTING_END_LABEL_CELL).Value = "표시 종료일"
    ws.Range(DISPLAY_SETTING_GANTT_ONLY_LABEL_CELL).Value = "간트 only"
    ws.Range(DISPLAY_SETTING_REPORT_ONLY_LABEL_CELL).Value = "보고 only"
    ws.Range(INPUT_SETTING_TITLE_CELL).Value = "간트 - 입력 제한 설정"
    legacyTaskMaxLength = ws.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value
    ws.Range(TASK_MAX_LENGTH_LEVEL1_LABEL_CELL).Value = "Level 1 내용 최대 글자 수"
    ws.Range(TASK_MAX_LENGTH_LEVEL2_LABEL_CELL).Value = "Level 2 내용 최대 글자 수"
    ws.Range(TASK_MAX_LENGTH_LEVEL3_LABEL_CELL).Value = "Level 3 내용 최대 글자 수"
    ws.Range(DEV_REPORT_SETTING_TITLE_CELL).Value = "개발보고 설정"
    ws.Range(DEV_REPORT_LAYOUT_LABEL_CELL).Value = "출력 형식"
    ws.Range(DEV_REPORT_OWNER_LABEL_CELL).Value = "담당자 이름 출력"
    ws.Range(DEV_REPORT_BULLET_TITLE_CELL).Value = "레벨별 글머리 기호"
    ws.Range(DEV_REPORT_BULLET_LEVEL1_LABEL_CELL).Value = "Level 1"
    ws.Range(DEV_REPORT_BULLET_LEVEL2_LABEL_CELL).Value = "Level 2"
    ws.Range(DEV_REPORT_BULLET_LEVEL3_LABEL_CELL).Value = "Level 3"
    ws.Range(WEEKLY_REPORT_SETTING_TITLE_CELL).Value = "주간보고 설정"
    ws.Range(WEEKLY_REPORT_OWNER_LABEL_CELL).Value = "담당자 이름 출력"
    ws.Range(WEEKLY_REPORT_PAGE_MODE_LABEL_CELL).Value = "페이지 출력 모드"
    ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_LABEL_CELL).Value = "내용 넘침 처리"
    ws.Range(WEEKLY_REPORT_BULLET_TITLE_CELL).Value = "주간보고 글머리 기호"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_LABEL_CELL).Value = "Level 1"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_LABEL_CELL).Value = "Level 2"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_LABEL_CELL).Value = "Level 3"
    ws.Range(WEEKLY_REPORT_BULLET_MODULE_LABEL_CELL).Value = "모듈"
    ws.Range(WEEKLY_REPORT_CUSTOM_PAGE_HEADER_CELL).Value = "커스텀 페이지 번호"
    ws.Range(WEEKLY_REPORT_CUSTOM_MODULE_HEADER_CELL).Value = "모듈명"
    ws.Range(HIDE_EXCLUDE_NO_HEADER_CELL).Value = "숨김 제외 No."
    ws.Range(HIDE_EXCLUDE_DATE_HEADER_CELL).Value = "숨김 제외 날짜"

    If Trim$(CStr(ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value)) = "" Then
        ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value = 3
    End If
    If Trim$(CStr(legacyTaskMaxLength)) = "" Then legacyTaskMaxLength = DEFAULT_TASK_MAX_LENGTH
    If Trim$(CStr(ws.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value)) = "" Then
        ws.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value = legacyTaskMaxLength
    End If
    If Trim$(CStr(ws.Range(TASK_MAX_LENGTH_LEVEL2_VALUE_CELL).Value)) = "" Then
        ws.Range(TASK_MAX_LENGTH_LEVEL2_VALUE_CELL).Value = legacyTaskMaxLength
    End If
    If Trim$(CStr(ws.Range(TASK_MAX_LENGTH_LEVEL3_VALUE_CELL).Value)) = "" Then
        ws.Range(TASK_MAX_LENGTH_LEVEL3_VALUE_CELL).Value = legacyTaskMaxLength
    End If
    If StrComp(Trim$(CStr(ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL).Value)), _
               "현재 형식", vbTextCompare) = 0 Then _
        ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL).Value = DEV_REPORT_LAYOUT_CURRENT
    If Trim$(CStr(ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL).Value)) = "" Then _
        ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL).Value = DEV_REPORT_LAYOUT_CURRENT
    If Trim$(CStr(ws.Range(DEV_REPORT_OWNER_VALUE_CELL).Value)) = "" Then _
        ws.Range(DEV_REPORT_OWNER_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(DEV_REPORT_BULLET_LEVEL1_VALUE_CELL).Value)) = "" Then _
        ws.Range(DEV_REPORT_BULLET_LEVEL1_VALUE_CELL).Value = ChrW(&H2022)
    If Trim$(CStr(ws.Range(DEV_REPORT_BULLET_LEVEL2_VALUE_CELL).Value)) = "" Then _
        ws.Range(DEV_REPORT_BULLET_LEVEL2_VALUE_CELL).Value = "-"
    If Trim$(CStr(ws.Range(DEV_REPORT_BULLET_LEVEL3_VALUE_CELL).Value)) = "" Then _
        ws.Range(DEV_REPORT_BULLET_LEVEL3_VALUE_CELL).Value = ChrW(&HB7)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value = WEEKLY_REPORT_PAGE_MODE_ALL
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value = WEEKLY_REPORT_OVERFLOW_MODE_EXPAND
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Value = ChrW(&H2022)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_VALUE_CELL).Value = "-"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).Value = ChrW(&HB7)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_MODULE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_MODULE_VALUE_CELL).Value = ChrW(&H2022)


    ws.Columns(HOLIDAY_COL_DATE).ColumnWidth = 14
    ws.Columns(HOLIDAY_COL_TYPE).ColumnWidth = 12
    ws.Columns(HOLIDAY_COL_DESC).ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 26
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 20
    ws.Columns("M").ColumnWidth = 16
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("P").ColumnWidth = 14
    ws.Columns("R").ColumnWidth = 24
    ws.Columns("S").ColumnWidth = 16

    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(242, 242, 242)
    ws.Range("A1:C1").Borders.LineStyle = xlContinuous

    ws.Range("E1").Font.Bold = True
    ws.Range("E1").Interior.Color = RGB(221, 235, 247)
    ws.Range("E1").Borders.LineStyle = xlContinuous

    ws.Range("F1:G1").Font.Bold = True
    ws.Range("F1:G1").Interior.Color = RGB(221, 235, 247)
    ws.Range("F1:G1").Borders.LineStyle = xlContinuous

    ws.Range("I1:J1").Font.Bold = True
    ws.Range("I1:J1").Interior.Color = RGB(221, 235, 247)
    ws.Range("I1:J1").Borders.LineStyle = xlContinuous

    ws.Range(INPUT_SETTING_TITLE_CELL & ":J7").Font.Bold = True
    ws.Range(INPUT_SETTING_TITLE_CELL & ":J7").Interior.Color = RGB(221, 235, 247)
    ws.Range(INPUT_SETTING_TITLE_CELL & ":J7").Borders.LineStyle = xlContinuous
    ws.Range("I8:J10").Borders.LineStyle = xlContinuous
    ws.Range("L1:M1,L4:M4").Font.Bold = True
    ws.Range("L1:M1,L4:M4").Interior.Color = RGB(226, 239, 218)
    ws.Range("L1:M7").Borders.LineStyle = xlContinuous
    ws.Range("O1:P1").Font.Bold = True
    ws.Range("O1:P1").Interior.Color = RGB(252, 228, 214)
    ws.Range("O1:P4").Borders.LineStyle = xlContinuous
    ws.Range("O5:P5").Font.Bold = True
    ws.Range("O5:P5").Interior.Color = RGB(252, 228, 214)
    ws.Range("O5:P" & CStr(WEEKLY_REPORT_CUSTOM_END_ROW)).Borders.LineStyle = xlContinuous
    ws.Range("R1:S1").Font.Bold = True
    ws.Range("R1:S1").Interior.Color = RGB(252, 228, 214)
    ws.Range("R1:S5").Borders.LineStyle = xlContinuous

    ws.Range("F2:G2").Borders.LineStyle = xlContinuous
    ws.Range("I2:J5").Borders.LineStyle = xlContinuous
    ws.Range("F5:G5").Font.Bold = True
    ws.Range("F5:G5").Interior.Color = RGB(242, 242, 242)
    ws.Range("F5:G5").Borders.LineStyle = xlContinuous
    ws.Range("F6:G200").Borders.LineStyle = xlContinuous

    ws.Columns(HOLIDAY_COL_DATE).NumberFormat = "yyyy-mm-dd"
    ws.Columns("G").NumberFormat = "General"
    ws.Range(DISPLAY_SETTING_START_VALUE_CELL).NumberFormat = "yyyy-mm-dd"
    ws.Range(DISPLAY_SETTING_END_VALUE_CELL).NumberFormat = "yyyy-mm-dd"
    ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL).NumberFormat = "General"
    ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).NumberFormat = "General"
    ws.Range(TASK_MAX_LENGTH_VALUE_RANGE).NumberFormat = "0"
    ws.Range("M2:M7").NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_OWNER_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL & ":" & _
             WEEKLY_REPORT_BULLET_MODULE_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_CUSTOM_PAGE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_START_ROW) & ":" & _
             WEEKLY_REPORT_CUSTOM_PAGE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_END_ROW)).NumberFormat = "0"

    lastSheetRow = ws.rows.Count
    Set rngType = ws.Range(HOLIDAY_COL_TYPE & HOLIDAY_DATA_START_ROW & ":" & HOLIDAY_COL_TYPE & lastSheetRow)
    Set rngHideLevel = ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL)
    Set rngExcludeNo = ws.Range(HIDE_EXCLUDE_NO_START_CELL & ":F" & lastSheetRow)
    Set rngExcludeDate = ws.Range(HIDE_EXCLUDE_DATE_START_CELL & ":G" & lastSheetRow)
    Set rngDisplayStart = ws.Range(DISPLAY_SETTING_START_VALUE_CELL)
    Set rngDisplayEnd = ws.Range(DISPLAY_SETTING_END_VALUE_CELL)
    Set rngDisplayGanttOnly = ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL)
    Set rngDisplayReportOnly = ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL)
    Set rngTaskMaxLength = ws.Range(TASK_MAX_LENGTH_VALUE_RANGE)
    Set rngReportLayout = ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL)
    Set rngDevReportOwner = ws.Range(DEV_REPORT_OWNER_VALUE_CELL)
    Set rngReportBullets = ws.Range(DEV_REPORT_BULLET_LEVEL1_VALUE_CELL & ":" & DEV_REPORT_BULLET_LEVEL3_VALUE_CELL)
    Set rngWeeklyReportOwner = ws.Range(WEEKLY_REPORT_OWNER_VALUE_CELL)
    Set rngWeeklyReportPageMode = ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL)
    Set rngWeeklyReportOverflowMode = ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL)
    Set rngWeeklyReportBullets = ws.Range( _
        WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL & ":" & _
        WEEKLY_REPORT_BULLET_MODULE_VALUE_CELL)
    Set rngWeeklyCustomPageNumbers = ws.Range( _
        WEEKLY_REPORT_CUSTOM_PAGE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_START_ROW) & ":" & _
        WEEKLY_REPORT_CUSTOM_PAGE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_END_ROW))

    On Error Resume Next
    rngType.Validation.Delete
    rngHideLevel.Validation.Delete
    rngExcludeNo.Validation.Delete
    rngExcludeDate.Validation.Delete
    rngDisplayStart.Validation.Delete
    rngDisplayEnd.Validation.Delete
    rngDisplayGanttOnly.Validation.Delete
    rngDisplayReportOnly.Validation.Delete
    rngTaskMaxLength.Validation.Delete
    rngReportLayout.Validation.Delete
    rngDevReportOwner.Validation.Delete
    rngReportBullets.Validation.Delete
    rngWeeklyReportOwner.Validation.Delete
    rngWeeklyReportPageMode.Validation.Delete
    rngWeeklyReportOverflowMode.Validation.Delete
    rngWeeklyReportBullets.Validation.Delete
    rngWeeklyCustomPageNumbers.Validation.Delete
    On Error GoTo 0

    rngType.Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:=HOLIDAY_TYPE_HOLIDAY & "," & HOLIDAY_TYPE_WORKDAY

    rngType.Validation.IgnoreBlank = True
    rngType.Validation.InCellDropdown = True
    rngType.Validation.InputTitle = "구분 선택"
    rngType.Validation.InputMessage = "휴일 또는 근무일만 선택할 수 있습니다."
    rngType.Validation.ErrorTitle = "입력 오류"
    rngType.Validation.ErrorMessage = "휴일 또는 근무일만 입력할 수 있습니다."

    rngHideLevel.Validation.Add Type:=xlValidateWholeNumber, _
                                AlertStyle:=xlValidAlertStop, _
                                Operator:=xlBetween, _
                                Formula1:="1", _
                                Formula2:="3"

    rngExcludeNo.Validation.Add Type:=xlValidateWholeNumber, _
                                AlertStyle:=xlValidAlertStop, _
                                Operator:=xlBetween, _
                                Formula1:="1", _
                                Formula2:="100000"
    rngExcludeNo.Validation.IgnoreBlank = True

    rngExcludeDate.NumberFormat = "yyyy-mm-dd"
    rngExcludeDate.Validation.Add Type:=xlValidateDate, _
                                  AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlBetween, _
                                  Formula1:="2000-01-01", _
                                  Formula2:="2100-12-31"
    rngExcludeDate.Validation.IgnoreBlank = True

    rngDisplayStart.Validation.Add Type:=xlValidateDate, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:="2000-01-01", _
                                   Formula2:="2100-12-31"
    rngDisplayStart.Validation.IgnoreBlank = True
    rngDisplayStart.Validation.InputTitle = "표시 시작일"
    rngDisplayStart.Validation.InputMessage = "간트에서 처음 보여줄 날짜를 입력하세요."
    rngDisplayStart.Validation.ErrorTitle = "입력 오류"
    rngDisplayStart.Validation.ErrorMessage = "올바른 날짜를 입력하세요."

    rngDisplayEnd.Validation.Add Type:=xlValidateDate, _
                                 AlertStyle:=xlValidAlertStop, _
                                 Operator:=xlBetween, _
                                 Formula1:="2000-01-01", _
                                 Formula2:="2100-12-31"
    rngDisplayEnd.Validation.IgnoreBlank = True
    rngDisplayEnd.Validation.InputTitle = "표시 종료일"
    rngDisplayEnd.Validation.InputMessage = "간트에서 마지막으로 보여줄 날짜를 입력하세요."
    rngDisplayEnd.Validation.ErrorTitle = "입력 오류"
    rngDisplayEnd.Validation.ErrorMessage = "올바른 날짜를 입력하세요."

    rngDisplayGanttOnly.Validation.Add Type:=xlValidateList, _
                                        AlertStyle:=xlValidAlertStop, _
                                        Operator:=xlBetween, _
                                        Formula1:="Y,N"
    rngDisplayGanttOnly.Validation.IgnoreBlank = True
    rngDisplayGanttOnly.Validation.InCellDropdown = True

    rngDisplayReportOnly.Validation.Add Type:=xlValidateList, _
                                         AlertStyle:=xlValidAlertStop, _
                                         Operator:=xlBetween, _
                                         Formula1:=STATUS_WEEKLY_REPORT & "," & STATUS_DEV_PROGRESS & "," & REPORT_FILTER_ALL & "," & REPORT_FILTER_EMPTY
    rngDisplayReportOnly.Validation.IgnoreBlank = True
    rngDisplayReportOnly.Validation.InCellDropdown = True

    rngTaskMaxLength.Validation.Add Type:=xlValidateWholeNumber, _
                                    AlertStyle:=xlValidAlertStop, _
                                    Operator:=xlBetween, _
                                    Formula1:=CStr(MIN_TASK_MAX_LENGTH), _
                                    Formula2:=CStr(MAX_TASK_MAX_LENGTH)
    rngTaskMaxLength.Validation.IgnoreBlank = False
    rngTaskMaxLength.Validation.InputTitle = "내용 글자 수 제한"
    rngTaskMaxLength.Validation.InputMessage = _
        MIN_TASK_MAX_LENGTH & "~" & MAX_TASK_MAX_LENGTH & " 사이 정수를 입력하세요."
    rngTaskMaxLength.Validation.ErrorTitle = "설정값 오류"
    rngTaskMaxLength.Validation.ErrorMessage = _
        "레벨별 내용 최대 글자 수는 " & MIN_TASK_MAX_LENGTH & "~" & _
        MAX_TASK_MAX_LENGTH & " 사이 정수여야 합니다."

    rngReportLayout.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=DEV_REPORT_LAYOUT_CURRENT & "," & DEV_REPORT_LAYOUT_BY_STATUS
    rngReportLayout.Validation.IgnoreBlank = False
    rngReportLayout.Validation.InCellDropdown = True
    rngReportLayout.Validation.InputTitle = "개발 보고 출력 형식"
    rngReportLayout.Validation.InputMessage = _
        "모듈별 통합 또는 상태별 구분을 선택하세요. 보고서는 항상 하나의 파일로 생성됩니다."
    rngReportLayout.Validation.ErrorTitle = "설정값 오류"
    rngReportLayout.Validation.ErrorMessage = _
        "모듈별 통합 또는 상태별 구분만 선택할 수 있습니다."

    rngDevReportOwner.Validation.Add Type:=xlValidateList, _
                                     AlertStyle:=xlValidAlertStop, _
                                     Operator:=xlBetween, _
                                     Formula1:="Y,N"
    rngDevReportOwner.Validation.IgnoreBlank = False
    rngDevReportOwner.Validation.InCellDropdown = True
    rngDevReportOwner.Validation.InputTitle = "개발보고 담당자 표시"
    rngDevReportOwner.Validation.InputMessage = _
        "모듈명과 각 단계 업무명 뒤에 담당자를 표시하려면 Y, 숨기려면 N을 선택하세요."
    rngDevReportOwner.Validation.ErrorTitle = "설정값 오류"
    rngDevReportOwner.Validation.ErrorMessage = "Y 또는 N만 선택할 수 있습니다."

    rngReportBullets.Validation.Add Type:=xlValidateTextLength, _
                                    AlertStyle:=xlValidAlertStop, _
                                    Operator:=xlBetween, _
                                    Formula1:="1", _
                                    Formula2:="5"
    rngReportBullets.Validation.IgnoreBlank = False
    rngReportBullets.Validation.InputTitle = "레벨별 글머리 기호"
    rngReportBullets.Validation.InputMessage = "각 Level에 사용할 글머리 기호를 1~5자로 입력하세요."
    rngReportBullets.Validation.ErrorTitle = "설정값 오류"
    rngReportBullets.Validation.ErrorMessage = "글머리 기호는 1~5자로 입력해야 합니다."

    rngWeeklyReportOwner.Validation.Add Type:=xlValidateList, _
                                        AlertStyle:=xlValidAlertStop, _
                                        Operator:=xlBetween, _
                                        Formula1:="Y,N"
    rngWeeklyReportOwner.Validation.IgnoreBlank = False
    rngWeeklyReportOwner.Validation.InCellDropdown = True
    rngWeeklyReportOwner.Validation.InputTitle = "주간보고 담당자 표시"
    rngWeeklyReportOwner.Validation.InputMessage = _
        "모듈명과 업무명 뒤에 담당자를 표시하려면 Y, 숨기려면 N을 선택하세요."
    rngWeeklyReportOwner.Validation.ErrorTitle = "설정값 오류"
    rngWeeklyReportOwner.Validation.ErrorMessage = "Y 또는 N만 선택할 수 있습니다."

    rngWeeklyReportPageMode.Validation.Add Type:=xlValidateList, _
                                           AlertStyle:=xlValidAlertStop, _
                                           Operator:=xlBetween, _
                                           Formula1:=WEEKLY_REPORT_PAGE_MODE_ALL & "," & _
                                                     WEEKLY_REPORT_PAGE_MODE_MODULE & "," & _
                                                     WEEKLY_REPORT_PAGE_MODE_CUSTOM
    rngWeeklyReportPageMode.Validation.IgnoreBlank = False
    rngWeeklyReportPageMode.Validation.InCellDropdown = True
    rngWeeklyReportPageMode.Validation.InputTitle = "주간보고 페이지 출력"
    rngWeeklyReportPageMode.Validation.InputMessage = _
        "전체 통합, 모듈별 페이지 또는 커스텀 페이지를 선택하세요."
    rngWeeklyReportPageMode.Validation.ErrorTitle = "설정값 오류"
    rngWeeklyReportPageMode.Validation.ErrorMessage = "목록에 있는 페이지 출력 모드만 선택할 수 있습니다."

    rngWeeklyReportOverflowMode.Validation.Add Type:=xlValidateList, _
                                               AlertStyle:=xlValidAlertStop, _
                                               Operator:=xlBetween, _
                                               Formula1:=WEEKLY_REPORT_OVERFLOW_MODE_EXPAND & "," & _
                                                         WEEKLY_REPORT_OVERFLOW_MODE_NEW_SLIDE
    rngWeeklyReportOverflowMode.Validation.IgnoreBlank = False
    rngWeeklyReportOverflowMode.Validation.InCellDropdown = True
    rngWeeklyReportOverflowMode.Validation.InputTitle = "주간보고 내용 넘침 처리"
    rngWeeklyReportOverflowMode.Validation.InputMessage = _
        "영역을 계속 늘리거나 템플릿 수용량에 맞춰 새 슬라이드로 나눌 수 있습니다."
    rngWeeklyReportOverflowMode.Validation.ErrorTitle = "설정값 오류"
    rngWeeklyReportOverflowMode.Validation.ErrorMessage = "목록에 있는 내용 넘침 처리 모드만 선택할 수 있습니다."

    rngWeeklyReportBullets.Validation.Add Type:=xlValidateTextLength, _
                                           AlertStyle:=xlValidAlertStop, _
                                           Operator:=xlBetween, _
                                           Formula1:="1", _
                                           Formula2:="5"
    rngWeeklyReportBullets.Validation.IgnoreBlank = False
    rngWeeklyReportBullets.Validation.InputTitle = "주간보고 글머리 기호"
    rngWeeklyReportBullets.Validation.InputMessage = _
        "각 Level에 사용할 글머리 기호를 1~5자로 입력하세요."
    rngWeeklyReportBullets.Validation.ErrorTitle = "설정값 오류"
    rngWeeklyReportBullets.Validation.ErrorMessage = "글머리 기호는 1~5자로 입력해야 합니다."

    rngWeeklyCustomPageNumbers.Validation.Add Type:=xlValidateWholeNumber, _
                                                AlertStyle:=xlValidAlertStop, _
                                                Operator:=xlBetween, _
                                                Formula1:="1", _
                                                Formula2:="1000"
    rngWeeklyCustomPageNumbers.Validation.IgnoreBlank = True
    rngWeeklyCustomPageNumbers.Validation.InputTitle = "커스텀 페이지 번호"
    rngWeeklyCustomPageNumbers.Validation.InputMessage = _
        "같은 페이지에 넣을 모듈에는 같은 페이지 번호를 입력하세요."
    rngWeeklyCustomPageNumbers.Validation.ErrorTitle = "설정값 오류"
    rngWeeklyCustomPageNumbers.Validation.ErrorMessage = "페이지 번호는 1~1000 사이 정수여야 합니다."

    RefreshWeeklyReportModuleDropdown
End Sub

Public Sub LoadHolidaySettings(ByRef holidayDict As Object, ByRef workdayDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim vDate As Variant
    Dim parsedDate As Date
    Dim vType As String
    Dim key As String

    Set holidayDict = CreateObject("Scripting.Dictionary")
    Set workdayDict = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.rows.Count, HOLIDAY_COL_DATE).End(xlUp).Row

    If lastRow < HOLIDAY_DATA_START_ROW Then Exit Sub

    For r = HOLIDAY_DATA_START_ROW To lastRow
        vDate = ws.Cells(r, HOLIDAY_COL_DATE).Value
        vType = Trim$(CStr(ws.Cells(r, HOLIDAY_COL_TYPE).Value))

        If TryParseHolidayDate(vDate, parsedDate) Then
            key = NormalizeDateKey(parsedDate)

            If vType = HOLIDAY_TYPE_HOLIDAY Then
                If Not holidayDict.Exists(key) Then holidayDict.Add key, True
            ElseIf vType = HOLIDAY_TYPE_WORKDAY Then
                If Not workdayDict.Exists(key) Then workdayDict.Add key, True
            End If
        End If
    Next r
End Sub


Public Function TryGetDisplayDateRange(ByRef displayStartDate As Date, ByRef displayEndDate As Date) As Boolean
    Dim ws As Worksheet
    Dim startValue As Variant
    Dim endValue As Variant

    TryGetDisplayDateRange = False

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)

    startValue = ws.Range(DISPLAY_SETTING_START_VALUE_CELL).Value
    endValue = ws.Range(DISPLAY_SETTING_END_VALUE_CELL).Value

    If Trim$(CStr(startValue)) = "" And Trim$(CStr(endValue)) = "" Then Exit Function

    If Trim$(CStr(startValue)) = "" Or Trim$(CStr(endValue)) = "" Then
        Err.Raise vbObjectError + 7101, "TryGetDisplayDateRange", "config 시트의 표시 시작일과 표시 종료일을 모두 입력해야 합니다."
    End If

    If Not IsDate(startValue) Or Not IsDate(endValue) Then
        Err.Raise vbObjectError + 7102, "TryGetDisplayDateRange", "config 시트의 표시 기간은 날짜만 입력할 수 있습니다."
    End If

    displayStartDate = CDate(startValue)
    displayEndDate = CDate(endValue)

    If CLng(displayStartDate) > CLng(displayEndDate) Then
        Err.Raise vbObjectError + 7103, "TryGetDisplayDateRange", "config 시트의 표시 시작일은 표시 종료일보다 늦을 수 없습니다."
    End If

    TryGetDisplayDateRange = True
End Function

Public Function GetDisplayGanttOnlyFlag() As Boolean
    Dim ws As Worksheet
    Dim v As String

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = UCase$(Trim$(CStr(ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL).Value)))

    GetDisplayGanttOnlyFlag = (v = "Y")
End Function

Public Function GetDisplayReportOnlyFlag() As String
    Dim ws As Worksheet
    Dim v As String

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = Trim$(CStr(ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).Value))

    If v = STATUS_WEEKLY_REPORT Or v = STATUS_DEV_PROGRESS Or v = REPORT_FILTER_ALL Or v = REPORT_FILTER_EMPTY Then
        GetDisplayReportOnlyFlag = v
    Else
        GetDisplayReportOnlyFlag = ""
    End If
End Function

Public Function GetTaskMaxLength(Optional ByVal taskLevel As Long = 1) As Long
    Dim ws As Worksheet
    Dim settingValue As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        GetTaskMaxLength = DEFAULT_TASK_MAX_LENGTH
        Exit Function
    End If

    If taskLevel < 1 Or taskLevel > 3 Then taskLevel = 1
    settingValue = ws.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Offset(taskLevel - 1, 0).Value
    If IsNumeric(settingValue) Then
        GetTaskMaxLength = CLng(settingValue)
    Else
        GetTaskMaxLength = DEFAULT_TASK_MAX_LENGTH
    End If

    If GetTaskMaxLength < MIN_TASK_MAX_LENGTH Or _
       GetTaskMaxLength > MAX_TASK_MAX_LENGTH Then
        GetTaskMaxLength = DEFAULT_TASK_MAX_LENGTH
    End If
End Function

Public Function GetDevReportSeparateStatusFlag() As Boolean
    Dim ws As Worksheet
    Dim settingValue As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = Trim$(CStr(ws.Range(DEV_REPORT_LAYOUT_VALUE_CELL).Value2))
    GetDevReportSeparateStatusFlag = _
        (StrComp(settingValue, DEV_REPORT_LAYOUT_BY_STATUS, vbTextCompare) = 0)
End Function

Public Function GetDevReportLevelBullet(ByVal taskLevel As Long) As String
    Dim ws As Worksheet
    Dim bulletText As String

    If taskLevel < 1 Or taskLevel > 3 Then taskLevel = 1

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If Not ws Is Nothing Then
        bulletText = Trim$(CStr(ws.Range(DEV_REPORT_BULLET_LEVEL1_VALUE_CELL).Offset(taskLevel - 1, 0).Value2))
    End If

    If Len(bulletText) = 0 Then
        Select Case taskLevel
            Case 1: bulletText = ChrW(&H2022)
            Case 2: bulletText = "-"
            Case Else: bulletText = ChrW(&HB7)
        End Select
    End If

    GetDevReportLevelBullet = bulletText
End Function

Public Function GetWeeklyReportLevelBullet(ByVal taskLevel As Long) As String
    Dim ws As Worksheet
    Dim bulletText As String

    If taskLevel < 1 Or taskLevel > 3 Then taskLevel = 1

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If Not ws Is Nothing Then
        bulletText = Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Offset(taskLevel - 1, 0).Value2))
    End If

    If Len(bulletText) = 0 Then
        Select Case taskLevel
            Case 1: bulletText = ChrW(&H2022)
            Case 2: bulletText = "-"
            Case Else: bulletText = ChrW(&HB7)
        End Select
    End If

    GetWeeklyReportLevelBullet = bulletText
End Function

Public Function GetWeeklyReportModuleBullet() As String
    Dim ws As Worksheet
    Dim bulletText As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If Not ws Is Nothing Then
        bulletText = Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_MODULE_VALUE_CELL).Value2))
    End If

    If Len(bulletText) = 0 Then bulletText = ChrW(&H2022)
    GetWeeklyReportModuleBullet = bulletText
End Function

Public Function GetDevReportShowOwnerFlag() As Boolean
    Dim ws As Worksheet
    Dim settingValue As String

    GetDevReportShowOwnerFlag = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = UCase$(Trim$(CStr(ws.Range(DEV_REPORT_OWNER_VALUE_CELL).Value2)))
    If settingValue = "N" Then GetDevReportShowOwnerFlag = False
End Function

Public Function GetWeeklyReportShowOwnerFlag() As Boolean
    Dim ws As Worksheet
    Dim settingValue As String

    GetWeeklyReportShowOwnerFlag = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = UCase$(Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_VALUE_CELL).Value2)))
    If settingValue = "N" Then GetWeeklyReportShowOwnerFlag = False
End Function

Public Function GetWeeklyReportPageMode() As String
    Dim ws As Worksheet
    Dim settingValue As String

    GetWeeklyReportPageMode = WEEKLY_REPORT_PAGE_MODE_ALL

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = Trim$(CStr(ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value2))
    Select Case settingValue
        Case WEEKLY_REPORT_PAGE_MODE_ALL, _
             WEEKLY_REPORT_PAGE_MODE_MODULE, _
             WEEKLY_REPORT_PAGE_MODE_CUSTOM
            GetWeeklyReportPageMode = settingValue
    End Select
End Function

Public Function GetWeeklyReportOverflowMode() As String
    Dim ws As Worksheet
    Dim settingValue As String

    GetWeeklyReportOverflowMode = WEEKLY_REPORT_OVERFLOW_MODE_EXPAND

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = Trim$(CStr(ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value2))
    Select Case settingValue
        Case WEEKLY_REPORT_OVERFLOW_MODE_EXPAND, _
             WEEKLY_REPORT_OVERFLOW_MODE_NEW_SLIDE
            GetWeeklyReportOverflowMode = settingValue
    End Select
End Function

Public Sub LoadWeeklyReportCustomPageAssignments(ByRef modulePageDict As Object)
    Dim ws As Worksheet
    Dim lastPageRow As Long
    Dim lastModuleRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim pageValue As Variant
    Dim pageNumber As Long
    Dim moduleName As String

    Set modulePageDict = CreateObject("Scripting.Dictionary")
    modulePageDict.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)

    lastPageRow = ws.Cells(ws.Rows.Count, WEEKLY_REPORT_CUSTOM_PAGE_COLUMN).End(xlUp).Row
    lastModuleRow = ws.Cells(ws.Rows.Count, WEEKLY_REPORT_CUSTOM_MODULE_COLUMN).End(xlUp).Row
    lastRow = lastPageRow
    If lastModuleRow > lastRow Then lastRow = lastModuleRow
    If lastRow < WEEKLY_REPORT_CUSTOM_START_ROW Then Exit Sub

    For r = WEEKLY_REPORT_CUSTOM_START_ROW To lastRow
        pageValue = ws.Cells(r, WEEKLY_REPORT_CUSTOM_PAGE_COLUMN).Value2
        moduleName = Trim$(CStr(ws.Cells(r, WEEKLY_REPORT_CUSTOM_MODULE_COLUMN).Value2))

        If Len(Trim$(CStr(pageValue))) > 0 Or Len(moduleName) > 0 Then
            If Len(Trim$(CStr(pageValue))) = 0 Or Len(moduleName) = 0 Then
                Err.Raise vbObjectError + 7120, "LoadWeeklyReportCustomPageAssignments", _
                          "config 시트의 커스텀 페이지 설정 " & CStr(r) & _
                          "행에 페이지 번호와 모듈명을 모두 입력하세요."
            End If

            If Not IsNumeric(pageValue) Or CDbl(pageValue) <> Fix(CDbl(pageValue)) Then
                Err.Raise vbObjectError + 7121, "LoadWeeklyReportCustomPageAssignments", _
                          "config 시트의 커스텀 페이지 번호는 정수여야 합니다: " & CStr(r) & "행"
            End If

            pageNumber = CLng(pageValue)
            If pageNumber < 1 Or pageNumber > 1000 Then
                Err.Raise vbObjectError + 7122, "LoadWeeklyReportCustomPageAssignments", _
                          "config 시트의 커스텀 페이지 번호는 1~1000 사이여야 합니다: " & CStr(r) & "행"
            End If

            If modulePageDict.Exists(moduleName) Then
                Err.Raise vbObjectError + 7123, "LoadWeeklyReportCustomPageAssignments", _
                          "config 시트의 커스텀 페이지 설정에 같은 모듈이 중복되었습니다: " & moduleName
            End If

            modulePageDict.Add moduleName, pageNumber
        End If
    Next r
End Sub

Public Sub RefreshWeeklyReportModuleDropdown()
    Dim configWs As Worksheet
    Dim taskWs As Worksheet
    Dim moduleSeen As Object
    Dim selectedModuleSeen As Object
    Dim moduleNames As Collection
    Dim moduleName As String
    Dim moduleNameItem As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim outputRow As Long
    Dim listLastRow As Long
    Dim rngCustomModules As Range

    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If configWs Is Nothing Then Exit Sub

    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare
    Set selectedModuleSeen = CreateObject("Scripting.Dictionary")
    selectedModuleSeen.CompareMode = vbTextCompare
    Set moduleNames = New Collection

    For r = WEEKLY_REPORT_CUSTOM_START_ROW To WEEKLY_REPORT_CUSTOM_END_ROW
        moduleName = Trim$(CStr(configWs.Cells(r, WEEKLY_REPORT_CUSTOM_MODULE_COLUMN).Value2))
        If Len(moduleName) > 0 And Not selectedModuleSeen.Exists(moduleName) Then
            selectedModuleSeen.Add moduleName, True
        End If
    Next r

    For Each taskWs In ThisWorkbook.Worksheets
        If taskWs.Name <> CONFIG_SHEET_NAME And _
           taskWs.Name <> REPORT_HISTORY_SHEET_NAME And _
           taskWs.Name <> "WeeklyPptTemplate" Then
            lastRow = GetLastDataRow(taskWs)
            For r = DATA_START_ROW To lastRow
                moduleName = Trim$(CStr(taskWs.Cells(r, COL_MODULE).Value2))
                If Len(moduleName) > 0 And _
                   Not selectedModuleSeen.Exists(moduleName) And _
                   Not moduleSeen.Exists(moduleName) Then
                    moduleSeen.Add moduleName, True
                    moduleNames.Add moduleName
                End If
            Next r
        End If
    Next taskWs

    configWs.Columns(WEEKLY_REPORT_MODULE_LIST_COLUMN).ClearContents
    configWs.Cells(1, WEEKLY_REPORT_MODULE_LIST_COLUMN).Value = "주간보고 미선택 모듈 목록"
    outputRow = WEEKLY_REPORT_MODULE_LIST_START_ROW
    For Each moduleNameItem In moduleNames
        configWs.Cells(outputRow, WEEKLY_REPORT_MODULE_LIST_COLUMN).Value = CStr(moduleNameItem)
        outputRow = outputRow + 1
    Next moduleNameItem

    listLastRow = outputRow - 1
    If listLastRow < WEEKLY_REPORT_MODULE_LIST_START_ROW Then
        listLastRow = WEEKLY_REPORT_MODULE_LIST_START_ROW
        configWs.Cells(listLastRow, WEEKLY_REPORT_MODULE_LIST_COLUMN).Value = ""
    End If
    configWs.Columns(WEEKLY_REPORT_MODULE_LIST_COLUMN).Hidden = True

    Set rngCustomModules = configWs.Range( _
        WEEKLY_REPORT_CUSTOM_MODULE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_START_ROW) & ":" & _
        WEEKLY_REPORT_CUSTOM_MODULE_COLUMN & CStr(WEEKLY_REPORT_CUSTOM_END_ROW))

    On Error Resume Next
    rngCustomModules.Validation.Delete
    On Error GoTo 0
    rngCustomModules.Validation.Add Type:=xlValidateList, _
                                    AlertStyle:=xlValidAlertStop, _
                                    Operator:=xlBetween, _
                                    Formula1:="=$" & WEEKLY_REPORT_MODULE_LIST_COLUMN & "$" & _
                                              CStr(WEEKLY_REPORT_MODULE_LIST_START_ROW) & ":$" & _
                                              WEEKLY_REPORT_MODULE_LIST_COLUMN & "$" & CStr(listLastRow)
    rngCustomModules.Validation.IgnoreBlank = True
    rngCustomModules.Validation.InCellDropdown = True
    rngCustomModules.Validation.ShowError = True
    rngCustomModules.Validation.InputTitle = "커스텀 모듈 선택"
    rngCustomModules.Validation.InputMessage = "아직 선택하지 않은 모듈만 표시됩니다."
    rngCustomModules.Validation.ErrorTitle = "모듈 선택 오류"
    rngCustomModules.Validation.ErrorMessage = "드롭다운에 있는 미선택 모듈만 선택할 수 있습니다."
End Sub

Public Function GetDuplicateWeeklyReportCustomModule(ByVal configWs As Worksheet) As String
    Dim moduleSeen As Object
    Dim moduleName As String
    Dim r As Long

    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For r = WEEKLY_REPORT_CUSTOM_START_ROW To WEEKLY_REPORT_CUSTOM_END_ROW
        moduleName = Trim$(CStr(configWs.Cells(r, WEEKLY_REPORT_CUSTOM_MODULE_COLUMN).Value2))
        If Len(moduleName) > 0 Then
            If moduleSeen.Exists(moduleName) Then
                GetDuplicateWeeklyReportCustomModule = moduleName
                Exit Function
            End If
            moduleSeen.Add moduleName, True
        End If
    Next r
End Function

Public Sub RefreshTaskTextLengthValidation()
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim lastRow As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> CONFIG_SHEET_NAME And _
           ws.Name <> REPORT_HISTORY_SHEET_NAME And _
           ws.Name <> "WeeklyPptTemplate" Then
            wasProtected = _
                (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
            If wasProtected Then UnprotectTaskSheet ws

            ApplyTaskTextLengthValidation ws

            If wasProtected Then
                lastRow = GetLastDataRow(ws)
                If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
                ApplyCalculatedColumnsProtection ws, lastRow
            End If
        End If
    Next ws
    Exit Sub

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    If wasProtected And Not ws Is Nothing Then
        On Error Resume Next
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
        On Error GoTo 0
    End If

    Err.Raise errNumber, "RefreshTaskTextLengthValidation", errDescription
End Sub

Public Function GetHideCompletedMaxLevel() As Long
    Dim ws As Worksheet
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value

    If IsNumeric(v) Then
        GetHideCompletedMaxLevel = CLng(v)
    Else
        GetHideCompletedMaxLevel = 3
    End If

    If GetHideCompletedMaxLevel < 1 Then GetHideCompletedMaxLevel = 1
    If GetHideCompletedMaxLevel > 3 Then GetHideCompletedMaxLevel = 3
End Function

Public Sub LoadExcludedRowNos(ByRef excludeNoDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    Dim key As String

    Set excludeNoDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.rows.Count, "F").End(xlUp).Row

    If lastRow < 6 Then Exit Sub

    For r = 6 To lastRow
        v = ws.Cells(r, "F").Value
        If IsNumeric(v) Then
            key = Trim$(CStr(CLng(v)))
            If Len(key) > 0 Then
                If Not excludeNoDict.Exists(key) Then excludeNoDict.Add key, True
            End If
        End If
    Next r
End Sub

Public Sub LoadExcludedDates(ByRef excludeDateDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    Dim parsedDate As Date
    Dim key As String

    Set excludeDateDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.rows.Count, "G").End(xlUp).Row

    If lastRow < 6 Then Exit Sub

    For r = 6 To lastRow
        v = ws.Cells(r, "G").Value
        If TryParseHolidayDate(v, parsedDate) Then
            key = NormalizeDateKey(parsedDate)
            If Not excludeDateDict.Exists(key) Then excludeDateDict.Add key, True
        End If
    Next r
End Sub

Private Function TryParseHolidayDate(ByVal vDate As Variant, ByRef parsedDate As Date) As Boolean
    Dim s As String
    Dim yyyyPart As Long
    Dim mmPart As Long
    Dim ddPart As Long
    Dim tmpDate As Date

    TryParseHolidayDate = False

    If IsDate(vDate) Then
        parsedDate = CDate(vDate)
        TryParseHolidayDate = True
        Exit Function
    End If

    s = Trim$(CStr(vDate))
    s = Replace$(s, "-", "")
    s = Replace$(s, ".", "")
    s = Replace$(s, "/", "")

    If Len(s) = 8 And IsNumeric(s) Then
        yyyyPart = CLng(Left$(s, 4))
        mmPart = CLng(Mid$(s, 5, 2))
        ddPart = CLng(Right$(s, 2))

        On Error GoTo InvalidDate
        tmpDate = DateSerial(yyyyPart, mmPart, ddPart)
        On Error GoTo 0

        If Year(tmpDate) <> yyyyPart Then Exit Function
        If Month(tmpDate) <> mmPart Then Exit Function
        If Day(tmpDate) <> ddPart Then Exit Function

        parsedDate = tmpDate
        TryParseHolidayDate = True
        Exit Function
    End If

    If IsDate(s) Then
        parsedDate = CDate(s)
        TryParseHolidayDate = True
    End If

    Exit Function

InvalidDate:
    On Error GoTo 0
End Function

Public Function NormalizeDateKey(ByVal targetDate As Date) As String
    NormalizeDateKey = Format$(CLng(targetDate), "0")
End Function
