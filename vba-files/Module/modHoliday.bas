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
    Dim rngWeeklyReportModuleOwner As Range
    Dim rngWeeklyReportProgramOwner As Range
    Dim rngWeeklyReportTaskOwner As Range
    Dim rngWeeklyReportDisplay As Range
    Dim rngWeeklyReportPageMode As Range
    Dim rngWeeklyReportOverflowMode As Range
    Dim rngWeeklyReportProgramGroup As Range
    Dim rngWeeklyReportBullets As Range
    Dim rngWeeklyCustomPageNumbers As Range
    Dim legacyTaskMaxLength As Variant
    Dim legacyWeeklyOwnerValue As String
    Dim legacyWeeklyPageMode As String
    Dim legacyWeeklyOverflowMode As String
    Dim legacyWeeklyProgramGroup As String
    Dim migrateWeeklyOwnerSettings As Boolean
    Dim migrateClassificationSettings As Boolean
    Dim legacyMajorOwner As String
    Dim legacyMinorOwner As String
    Dim legacyTaskOwner As String
    Dim legacyTaskOwnerLevel As String
    Dim legacyMajorBullet As String
    Dim legacyMinorBullet As String
    Dim legacyLevel1Bullet As String
    Dim legacyLevel2Bullet As String
    Dim legacyLevel3Bullet As String
    Dim legacyCategoryDepth As String
    Dim migrateDisplaySettings As Boolean
    Dim migrationRow As Long

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

    migrateWeeklyOwnerSettings = _
        (Trim$(CStr(ws.Range("P2").Value2)) <> "모듈" And _
         Trim$(CStr(ws.Range("P2").Value2)) <> "대분류" And _
         Trim$(CStr(ws.Range("P2").Value2)) <> "대분류 경로" And _
         Trim$(CStr(ws.Range("P2").Value2)) <> "타입")
    If migrateWeeklyOwnerSettings Then
        legacyWeeklyOwnerValue = UCase$(Trim$(CStr(ws.Range("P2").Value2)))
        legacyWeeklyPageMode = Trim$(CStr(ws.Range("P3").Value2))
        legacyWeeklyOverflowMode = Trim$(CStr(ws.Range("P4").Value2))
        legacyWeeklyProgramGroup = UCase$(Trim$(CStr(ws.Range("P5").Value2)))
        If legacyWeeklyOwnerValue <> "N" Then legacyWeeklyOwnerValue = "Y"

        If Trim$(CStr(ws.Range("U1").Value2)) = "" And _
           Trim$(CStr(ws.Range("R1").Value2)) = "주간보고 글머리 기호" Then
            ws.Range("U1:V6").Value2 = ws.Range("R1:S6").Value2
        End If
        ws.Range("R1:S6").Clear
    End If

    migrateClassificationSettings = _
        (Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_TITLE_CELL).Value2)) <> _
         "주간보고 글머리 기호")
    If migrateClassificationSettings Then
        If migrateWeeklyOwnerSettings Then
            legacyMajorOwner = legacyWeeklyOwnerValue
            legacyMinorOwner = legacyWeeklyOwnerValue
            legacyTaskOwner = legacyWeeklyOwnerValue
            legacyTaskOwnerLevel = WEEKLY_REPORT_OWNER_TASK_LEVEL_ALL
        Else
            legacyMajorOwner = UCase$(Trim$(CStr(ws.Range("P3").Value2)))
            legacyMinorOwner = UCase$(Trim$(CStr(ws.Range("Q3").Value2)))
            legacyTaskOwner = UCase$(Trim$(CStr(ws.Range("R3").Value2)))
            legacyTaskOwnerLevel = Trim$(CStr(ws.Range("S3").Value2))
        End If
        legacyLevel1Bullet = Trim$(CStr(ws.Range("V2").Value2))
        legacyLevel2Bullet = Trim$(CStr(ws.Range("V3").Value2))
        legacyLevel3Bullet = Trim$(CStr(ws.Range("V4").Value2))
        legacyMajorBullet = Trim$(CStr(ws.Range("V5").Value2))
        legacyMinorBullet = Trim$(CStr(ws.Range("V6").Value2))
        If legacyMajorOwner <> "N" Then legacyMajorOwner = "Y"
        If legacyMinorOwner <> "N" Then legacyMinorOwner = "Y"
        If legacyTaskOwner <> "N" Then legacyTaskOwner = "Y"
        ws.Range("U1:V6").Clear
    End If

    migrateDisplaySettings = _
        (Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_LABEL_CELL).Value2)) <> _
         "주간보고 항목 출력")
    If migrateDisplaySettings Then
        legacyCategoryDepth = Trim$(CStr(ws.Range("P6").Value2))
        If Not migrateClassificationSettings Then
            legacyMajorOwner = UCase$(Trim$(CStr(ws.Range("P3").Value2)))
            legacyMinorOwner = UCase$(Trim$(CStr(ws.Range("S3").Value2)))
            legacyTaskOwner = UCase$(Trim$(CStr(ws.Range("T3").Value2)))
            If legacyMajorOwner <> "N" Then legacyMajorOwner = "Y"
            If legacyMinorOwner <> "N" Then legacyMinorOwner = "Y"
            If legacyTaskOwner <> "N" Then legacyTaskOwner = "Y"
        End If
    End If

    If Trim$(CStr(ws.Range("O5").Value2)) = "커스텀 페이지 번호" Then
        For migrationRow = 200 To 6 Step -1
            ws.Range("O" & CStr(migrationRow + 1) & ":P" & _
                     CStr(migrationRow + 1)).Value2 = _
                ws.Range("O" & CStr(migrationRow) & ":P" & _
                         CStr(migrationRow)).Value2
        Next migrationRow
    End If

    If Trim$(CStr(ws.Range("O6").Value2)) = "커스텀 페이지 번호" Then
        For migrationRow = 201 To 7 Step -1
            ws.Range("O" & CStr(migrationRow + 1) & ":P" & _
                     CStr(migrationRow + 1)).Value2 = _
                ws.Range("O" & CStr(migrationRow) & ":P" & _
                         CStr(migrationRow)).Value2
        Next migrationRow
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
    If Trim$(CStr(ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).Value)) = "개발진행" Then
        ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).ClearContents
    End If
    ws.Range(INPUT_SETTING_TITLE_CELL).Value = "간트 - 입력 제한 설정"
    legacyTaskMaxLength = ws.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value
    ws.Range(TASK_MAX_LENGTH_LEVEL1_LABEL_CELL).Value = "Level 1 내용 최대 글자 수"
    ws.Range(TASK_MAX_LENGTH_LEVEL2_LABEL_CELL).Value = "Level 2 내용 최대 글자 수"
    ws.Range(TASK_MAX_LENGTH_LEVEL3_LABEL_CELL).Value = "Level 3 내용 최대 글자 수"
    ' Clear settings left by the retired development-report feature.
    ws.Range("L1:M7").Clear
    ws.Range(WEEKLY_REPORT_SETTING_TITLE_CELL).Value = "주간보고 설정"
    ws.Range(WEEKLY_REPORT_OWNER_LABEL_CELL).Value = "담당자 이름 출력"
    ws.Range(WEEKLY_REPORT_DISPLAY_LABEL_CELL).Value = "주간보고 항목 출력"
    ws.Range(WEEKLY_REPORT_OWNER_TYPE_LABEL_CELL).Value = "타입"
    ws.Range(WEEKLY_REPORT_OWNER_MAJOR_LABEL_CELL).Value = "대분류"
    ws.Range(WEEKLY_REPORT_OWNER_MIDDLE_LABEL_CELL).Value = "중분류"
    ws.Range(WEEKLY_REPORT_OWNER_MINOR_LABEL_CELL).Value = "소분류"
    ws.Range(WEEKLY_REPORT_OWNER_TASK_LABEL_CELL).Value = "업무명"
    ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_LABEL_CELL).Value = "업무 레벨"
    ws.Range(WEEKLY_REPORT_PAGE_MODE_LABEL_CELL).Value = "페이지 출력 모드"
    ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_LABEL_CELL).Value = "내용 넘침 처리"
    ws.Range(WEEKLY_REPORT_CATEGORY_DEPTH_LABEL_CELL).Value = "분류 출력 기준"
    ws.Range(WEEKLY_REPORT_CATEGORY_DEPTH_VALUE_CELL).Value = "개별 Y/N 설정"
    ws.Range(WEEKLY_REPORT_BULLET_TITLE_CELL).Value = "주간보고 글머리 기호"
    ws.Range(WEEKLY_REPORT_BULLET_TYPE_LABEL_CELL).Value = "타입"
    ws.Range(WEEKLY_REPORT_BULLET_MAJOR_LABEL_CELL).Value = "대분류"
    ws.Range(WEEKLY_REPORT_BULLET_MIDDLE_LABEL_CELL).Value = "중분류"
    ws.Range(WEEKLY_REPORT_BULLET_MINOR_LABEL_CELL).Value = "소분류"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_LABEL_CELL).Value = "Level 1"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_LABEL_CELL).Value = "Level 2"
    ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_LABEL_CELL).Value = "Level 3"
    ws.Range(WEEKLY_REPORT_CUSTOM_PAGE_HEADER_CELL).Value = "커스텀 페이지 번호"
    ws.Range(WEEKLY_REPORT_CUSTOM_MODULE_HEADER_CELL).Value = "분류 경로"
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
    If migrateWeeklyOwnerSettings Then
        ws.Range(WEEKLY_REPORT_OWNER_MODULE_VALUE_CELL).Value = legacyWeeklyOwnerValue
        ws.Range(WEEKLY_REPORT_OWNER_PROGRAM_VALUE_CELL).Value = legacyWeeklyOwnerValue
        ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL).Value = legacyWeeklyOwnerValue
        ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value = _
            WEEKLY_REPORT_OWNER_TASK_LEVEL_ALL
        If Len(legacyWeeklyPageMode) > 0 Then _
            ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value = legacyWeeklyPageMode
        If Len(legacyWeeklyOverflowMode) > 0 Then _
            ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value = legacyWeeklyOverflowMode
        If Len(legacyWeeklyProgramGroup) > 0 Then
            If legacyWeeklyProgramGroup = "Y" Then
                ws.Range(WEEKLY_REPORT_CATEGORY_DEPTH_VALUE_CELL).Value = _
                    WEEKLY_REPORT_CATEGORY_DEPTH_MINOR
            Else
                ws.Range(WEEKLY_REPORT_CATEGORY_DEPTH_VALUE_CELL).Value = _
                    WEEKLY_REPORT_CATEGORY_DEPTH_MAJOR
            End If
        End If
    End If
    If migrateClassificationSettings Then
        ws.Range(WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MAJOR_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MIDDLE_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MINOR_VALUE_CELL).Value = legacyMinorOwner
        ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL).Value = legacyTaskOwner
        ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value = "Y"
        ws.Range(WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL).Value = legacyMajorBullet
        ws.Range(WEEKLY_REPORT_BULLET_MAJOR_VALUE_CELL).Value = legacyMajorBullet
        ws.Range(WEEKLY_REPORT_BULLET_MIDDLE_VALUE_CELL).Value = legacyMinorBullet
        ws.Range(WEEKLY_REPORT_BULLET_MINOR_VALUE_CELL).Value = legacyMinorBullet
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Value = legacyLevel1Bullet
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_VALUE_CELL).Value = legacyLevel2Bullet
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).Value = legacyLevel3Bullet
    End If
    If migrateDisplaySettings Then
        ws.Range(WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MAJOR_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MIDDLE_VALUE_CELL).Value = legacyMajorOwner
        ws.Range(WEEKLY_REPORT_OWNER_MINOR_VALUE_CELL).Value = legacyMinorOwner
        ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL).Value = legacyTaskOwner
        ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value = "Y"
        ws.Range(WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL).Value = "Y"
        ws.Range(WEEKLY_REPORT_DISPLAY_MAJOR_VALUE_CELL).Value = _
            IIf(legacyCategoryDepth = WEEKLY_REPORT_CATEGORY_DEPTH_TYPE, "N", "Y")
        ws.Range(WEEKLY_REPORT_DISPLAY_MIDDLE_VALUE_CELL).Value = _
            IIf(legacyCategoryDepth = WEEKLY_REPORT_CATEGORY_DEPTH_MIDDLE Or _
                legacyCategoryDepth = WEEKLY_REPORT_CATEGORY_DEPTH_MINOR, "Y", "N")
        ws.Range(WEEKLY_REPORT_DISPLAY_MINOR_VALUE_CELL).Value = _
            IIf(legacyCategoryDepth = WEEKLY_REPORT_CATEGORY_DEPTH_MINOR Or _
                UCase$(legacyCategoryDepth) = "Y", "Y", "N")
        ws.Range(WEEKLY_REPORT_DISPLAY_TASK_VALUE_CELL).Value = "Y"
        ws.Range(WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL).Value = "N"
    End If
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_MODULE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_MODULE_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_MIDDLE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_MIDDLE_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_PROGRAM_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_PROGRAM_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value)) = "" Or _
       Trim$(CStr(ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value)) = _
       WEEKLY_REPORT_OWNER_TASK_LEVEL_ALL Then _
        ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_MAJOR_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_MAJOR_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_MIDDLE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_MIDDLE_VALUE_CELL).Value = "N"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_MINOR_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_MINOR_VALUE_CELL).Value = "N"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_TASK_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_TASK_VALUE_CELL).Value = "Y"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL).Value = "N"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value = WEEKLY_REPORT_PAGE_MODE_ALL
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value)) = _
       WEEKLY_REPORT_PAGE_MODE_LEGACY_MODULE Then _
        ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).Value = WEEKLY_REPORT_PAGE_MODE_MODULE
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).Value = WEEKLY_REPORT_OVERFLOW_MODE_EXPAND
    ws.Range(WEEKLY_REPORT_CATEGORY_DEPTH_VALUE_CELL).Value = "개별 Y/N 설정"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL).Value = ChrW(&H2022)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_MAJOR_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_MAJOR_VALUE_CELL).Value = "-"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_MIDDLE_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_MIDDLE_VALUE_CELL).Value = ChrW(&HB7)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_MINOR_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_MINOR_VALUE_CELL).Value = ChrW(&H25E6)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL1_VALUE_CELL).Value = ChrW(&H2022)
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL2_VALUE_CELL).Value = "-"
    If Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).Value)) = "" Then _
        ws.Range(WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).Value = ChrW(&HB7)


    ws.Columns(HOLIDAY_COL_DATE).ColumnWidth = 14
    ws.Columns(HOLIDAY_COL_TYPE).ColumnWidth = 12
    ws.Columns(HOLIDAY_COL_DESC).ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 26
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("P").ColumnWidth = 14
    ws.Columns("Q").ColumnWidth = 14
    ws.Columns("R").ColumnWidth = 14
    ws.Columns("S").ColumnWidth = 14
    ws.Columns("T").ColumnWidth = 14
    ws.Columns("U").ColumnWidth = 16
    ws.Columns("W").ColumnWidth = 18
    ws.Columns("X").ColumnWidth = 12

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
    ws.Range("O1:U1").Font.Bold = True
    ws.Range("O1:U1").Interior.Color = RGB(252, 228, 214)
    ws.Range("O1:U1").Borders.LineStyle = xlContinuous
    ws.Range("O2:U2").Font.Bold = True
    ws.Range("O2:U3").Borders.LineStyle = xlContinuous
    ws.Range("O4:P6").Borders.LineStyle = xlContinuous
    ws.Range("O7:P7").Font.Bold = True
    ws.Range("O7:P7").Interior.Color = RGB(252, 228, 214)
    ws.Range("O7:P" & CStr(WEEKLY_REPORT_CUSTOM_END_ROW)).Borders.LineStyle = xlContinuous
    ws.Range("W1:X1").Font.Bold = True
    ws.Range("W1:X1").Interior.Color = RGB(252, 228, 214)
    ws.Range("W1:X8").Borders.LineStyle = xlContinuous

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
    ws.Range(WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL & ":" & _
             WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL & ":" & _
             WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_PROGRAM_GROUP_VALUE_CELL).NumberFormat = "General"
    ws.Range(WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL & ":" & _
             WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL).NumberFormat = "General"
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
    Set rngWeeklyReportModuleOwner = ws.Range( _
        WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL & ":" & _
        WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL)
    Set rngWeeklyReportProgramOwner = ws.Range(WEEKLY_REPORT_OWNER_PROGRAM_VALUE_CELL)
    Set rngWeeklyReportTaskOwner = ws.Range(WEEKLY_REPORT_OWNER_TASK_VALUE_CELL)
    Set rngWeeklyReportDisplay = ws.Range( _
        WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL & ":" & _
        WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL)
    Set rngWeeklyReportPageMode = ws.Range(WEEKLY_REPORT_PAGE_MODE_VALUE_CELL)
    Set rngWeeklyReportOverflowMode = ws.Range(WEEKLY_REPORT_OVERFLOW_MODE_VALUE_CELL)
    Set rngWeeklyReportProgramGroup = ws.Range(WEEKLY_REPORT_PROGRAM_GROUP_VALUE_CELL)
    Set rngWeeklyReportBullets = ws.Range( _
        WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL & ":" & _
        WEEKLY_REPORT_BULLET_LEVEL3_VALUE_CELL)
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
    rngWeeklyReportModuleOwner.Validation.Delete
    rngWeeklyReportProgramOwner.Validation.Delete
    rngWeeklyReportTaskOwner.Validation.Delete
    rngWeeklyReportDisplay.Validation.Delete
    rngWeeklyReportPageMode.Validation.Delete
    rngWeeklyReportOverflowMode.Validation.Delete
    rngWeeklyReportProgramGroup.Validation.Delete
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
                                         Formula1:=STATUS_WEEKLY_REPORT & "," & REPORT_FILTER_ALL & "," & REPORT_FILTER_EMPTY
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

    With rngWeeklyReportModuleOwner.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Y,N"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = "주간보고 담당자 표시"
        .InputMessage = "해당 구분의 담당자를 표시하려면 Y, 숨기려면 N을 선택하세요."
        .ErrorTitle = "설정값 오류"
        .ErrorMessage = "Y 또는 N만 선택할 수 있습니다."
    End With

    With rngWeeklyReportDisplay.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Y,N"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = "주간보고 항목 출력"
        .InputMessage = "해당 항목을 주간보고에 출력하려면 Y, 숨기려면 N을 선택하세요."
        .ErrorTitle = "설정값 오류"
        .ErrorMessage = "Y 또는 N만 선택할 수 있습니다."
    End With

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
        "전체 통합, 분류 경로별 페이지 또는 커스텀 페이지를 선택하세요."
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
        "같은 페이지에 넣을 분류 경로에는 같은 페이지 번호를 입력하세요."
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

    If v = STATUS_WEEKLY_REPORT Or v = REPORT_FILTER_ALL Or v = REPORT_FILTER_EMPTY Then
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

Public Function GetWeeklyReportLevelBullet(ByVal taskLevel As Long) As String
    Dim ws As Worksheet
    Dim bulletText As String

    If taskLevel < 1 Then taskLevel = 1
    If taskLevel > 3 Then taskLevel = 3

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

Public Function GetWeeklyReportGroupByProgramFlag() As Boolean
    GetWeeklyReportGroupByProgramFlag = GetWeeklyReportShowCategoryFlag(4)
End Function

Public Function GetWeeklyReportCategoryDepth() As Long
    GetWeeklyReportCategoryDepth = GetWeeklyReportVisibleCategoryCount()
End Function

Public Function GetWeeklyReportClassificationPath(ByVal taskWs As Worksheet, _
                                                  ByVal rowNum As Long) As String
    Dim typeText As String
    Dim majorText As String
    Dim middleText As String

    If Trim$(CStr(taskWs.Range("D" & HEADER_ROW).Value2)) = "타입" Then
        typeText = Trim$(CStr(taskWs.Cells(rowNum, COL_TYPE).Value2))
        majorText = Trim$(CStr(taskWs.Cells(rowNum, COL_MAJOR_CATEGORY).Value2))
        middleText = Trim$(CStr(taskWs.Cells(rowNum, COL_MIDDLE_CATEGORY).Value2))
    Else
        majorText = Trim$(CStr(taskWs.Range("D" & rowNum).Value2))
    End If

    If Len(typeText) = 0 Then typeText = "타입 미지정"
    If Len(majorText) = 0 Then majorText = "대분류 미지정"
    If Len(middleText) = 0 Then middleText = "중분류 미지정"
    GetWeeklyReportClassificationPath = typeText & " > " & majorText & _
                                        " > " & middleText
End Function

Public Function GetWeeklyReportShowCategoryFlag( _
                    ByVal categoryLevel As Long) As Boolean
    Dim valueCell As String

    Select Case categoryLevel
        Case 1: valueCell = WEEKLY_REPORT_DISPLAY_TYPE_VALUE_CELL
        Case 2: valueCell = WEEKLY_REPORT_DISPLAY_MAJOR_VALUE_CELL
        Case 3: valueCell = WEEKLY_REPORT_DISPLAY_MIDDLE_VALUE_CELL
        Case Else: valueCell = WEEKLY_REPORT_DISPLAY_MINOR_VALUE_CELL
    End Select
    GetWeeklyReportShowCategoryFlag = GetWeeklyReportDisplayFlag(valueCell, True)
End Function

Public Function GetWeeklyReportShowTaskNameFlag() As Boolean
    GetWeeklyReportShowTaskNameFlag = GetWeeklyReportDisplayFlag( _
                                          WEEKLY_REPORT_DISPLAY_TASK_VALUE_CELL, True)
End Function

Public Function GetWeeklyReportShowTaskLevelFlag() As Boolean
    GetWeeklyReportShowTaskLevelFlag = GetWeeklyReportDisplayFlag( _
                                           WEEKLY_REPORT_DISPLAY_TASK_LEVEL_VALUE_CELL, False)
End Function

Public Function GetWeeklyReportShowTaskLevelOwnerFlag() As Boolean
    GetWeeklyReportShowTaskLevelOwnerFlag = GetWeeklyReportOwnerFlag( _
                                                WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL)
End Function

Public Function GetWeeklyReportVisibleCategoryCount() As Long
    Dim categoryLevel As Long

    For categoryLevel = 1 To 4
        If GetWeeklyReportShowCategoryFlag(categoryLevel) Then
            GetWeeklyReportVisibleCategoryCount = _
                GetWeeklyReportVisibleCategoryCount + 1
        End If
    Next categoryLevel
End Function

Public Function GetWeeklyReportVisibleCategoryPosition( _
                    ByVal categoryLevel As Long) As Long
    Dim currentLevel As Long

    For currentLevel = 1 To categoryLevel
        If GetWeeklyReportShowCategoryFlag(currentLevel) Then
            GetWeeklyReportVisibleCategoryPosition = _
                GetWeeklyReportVisibleCategoryPosition + 1
        End If
    Next currentLevel
End Function

Private Function GetWeeklyReportDisplayFlag(ByVal valueCell As String, _
                                            ByVal defaultValue As Boolean) As Boolean
    Dim ws As Worksheet
    Dim settingValue As String

    GetWeeklyReportDisplayFlag = defaultValue
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = UCase$(Trim$(CStr(ws.Range(valueCell).Value2)))
    If settingValue = "Y" Then GetWeeklyReportDisplayFlag = True
    If settingValue = "N" Then GetWeeklyReportDisplayFlag = False
End Function

Public Function GetWeeklyReportCategoryBullet(ByVal categoryLevel As Long) As String
    Dim ws As Worksheet
    Dim bulletText As String
    Dim valueCell As String

    Select Case categoryLevel
        Case 1: valueCell = WEEKLY_REPORT_BULLET_TYPE_VALUE_CELL
        Case 2: valueCell = WEEKLY_REPORT_BULLET_MAJOR_VALUE_CELL
        Case 3: valueCell = WEEKLY_REPORT_BULLET_MIDDLE_VALUE_CELL
        Case Else: valueCell = WEEKLY_REPORT_BULLET_MINOR_VALUE_CELL
    End Select

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If Not ws Is Nothing Then bulletText = Trim$(CStr(ws.Range(valueCell).Value2))
    On Error GoTo 0

    If Len(bulletText) = 0 Then
        Select Case categoryLevel
            Case 1: bulletText = ChrW(&H2022)
            Case 2: bulletText = "-"
            Case 3: bulletText = ChrW(&HB7)
            Case Else: bulletText = ChrW(&H25E6)
        End Select
    End If
    GetWeeklyReportCategoryBullet = bulletText
End Function

Public Function GetWeeklyReportShowCategoryOwnerFlag( _
                    ByVal categoryLevel As Long) As Boolean
    Dim valueCell As String

    Select Case categoryLevel
        Case 1: valueCell = WEEKLY_REPORT_OWNER_TYPE_VALUE_CELL
        Case 2: valueCell = WEEKLY_REPORT_OWNER_MAJOR_VALUE_CELL
        Case 3: valueCell = WEEKLY_REPORT_OWNER_MIDDLE_VALUE_CELL
        Case Else: valueCell = WEEKLY_REPORT_OWNER_MINOR_VALUE_CELL
    End Select
    GetWeeklyReportShowCategoryOwnerFlag = GetWeeklyReportOwnerFlag(valueCell)
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

Public Function GetWeeklyReportProgramBullet() As String
    Dim ws As Worksheet
    Dim bulletText As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If Not ws Is Nothing Then
        bulletText = Trim$(CStr(ws.Range(WEEKLY_REPORT_BULLET_PROGRAM_VALUE_CELL).Value2))
    End If

    If Len(bulletText) = 0 Then bulletText = "-"
    GetWeeklyReportProgramBullet = bulletText
End Function

Public Function GetWeeklyReportShowModuleOwnerFlag() As Boolean
    GetWeeklyReportShowModuleOwnerFlag = GetWeeklyReportOwnerFlag( _
                                             WEEKLY_REPORT_OWNER_MODULE_VALUE_CELL)
End Function

Public Function GetWeeklyReportShowProgramOwnerFlag() As Boolean
    GetWeeklyReportShowProgramOwnerFlag = GetWeeklyReportOwnerFlag( _
                                              WEEKLY_REPORT_OWNER_PROGRAM_VALUE_CELL)
End Function

Public Function GetWeeklyReportShowTaskOwnerFlag() As Boolean
    GetWeeklyReportShowTaskOwnerFlag = GetWeeklyReportOwnerFlag( _
                                           WEEKLY_REPORT_OWNER_TASK_VALUE_CELL)
End Function

Private Function GetWeeklyReportOwnerFlag(ByVal valueCell As String) As Boolean
    Dim ws As Worksheet
    Dim settingValue As String

    GetWeeklyReportOwnerFlag = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = UCase$(Trim$(CStr(ws.Range(valueCell).Value2)))
    If settingValue = "N" Then GetWeeklyReportOwnerFlag = False
End Function

Public Function GetWeeklyReportTaskOwnerLevel() As Long
    Dim ws As Worksheet
    Dim settingValue As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    settingValue = Trim$(CStr( _
                       ws.Range(WEEKLY_REPORT_OWNER_TASK_LEVEL_VALUE_CELL).Value2))
    Select Case settingValue
        Case WEEKLY_REPORT_OWNER_TASK_LEVEL1: GetWeeklyReportTaskOwnerLevel = 1
        Case WEEKLY_REPORT_OWNER_TASK_LEVEL2: GetWeeklyReportTaskOwnerLevel = 2
        Case WEEKLY_REPORT_OWNER_TASK_LEVEL3: GetWeeklyReportTaskOwnerLevel = 3
        Case Else: GetWeeklyReportTaskOwnerLevel = 0
    End Select
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
        Case WEEKLY_REPORT_PAGE_MODE_LEGACY_MODULE
            GetWeeklyReportPageMode = WEEKLY_REPORT_PAGE_MODE_MODULE
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
                          "행에 페이지 번호와 분류 경로를 모두 입력하세요."
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
                          "config 시트의 커스텀 페이지 설정에 같은 분류 경로가 중복되었습니다: " & moduleName
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
           taskWs.Name <> "WeeklyPptTemplate" And _
           (Trim$(CStr(taskWs.Range("H" & HEADER_ROW).Value2)) = "내용" Or _
            Trim$(CStr(taskWs.Range("F" & HEADER_ROW).Value2)) = "내용" Or _
            Trim$(CStr(taskWs.Range("E" & HEADER_ROW).Value2)) = "내용") Then
            lastRow = GetLastDataRow(taskWs)
            For r = DATA_START_ROW To lastRow
                moduleName = GetWeeklyReportClassificationPath(taskWs, r)
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
    configWs.Cells(1, WEEKLY_REPORT_MODULE_LIST_COLUMN).Value = "주간보고 미선택 분류 경로 목록"
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
    rngCustomModules.Validation.InputTitle = "커스텀 분류 경로 선택"
    rngCustomModules.Validation.InputMessage = "아직 선택하지 않은 분류 경로만 표시됩니다."
    rngCustomModules.Validation.ErrorTitle = "분류 경로 선택 오류"
    rngCustomModules.Validation.ErrorMessage = "드롭다운에 있는 미선택 분류 경로만 선택할 수 있습니다."
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
