[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath)

# 2026-10-23 김좌현 수정 : WEB UPDATE 헤더 이전 시 값·수식·개체 보존 및 중복 웹 실행 버튼 제거 검증
# 2026-10-23 김좌현 수정 : 다른 통합문서 행 삽입 시 이름·외부참조 유입 방지 및 대상 선택목록 유지 검증
# 2026-10-23 김좌현 수정 : 하위 작업시간 입력 즉시 상위 합계 갱신과 중복 합산·빈값·보호 상태 보존 검증
# 2026-10-22 김좌현 수정 : 모두/간트 새로고침 분리 및 공통 버튼 정렬 검증
# 2026-10-22 김좌현 수정 : 웹 설정/작업시간 이전의 값·수식·개체와 반복 실행 보존 검증
# 2026-10-20 김좌현 수정 : 새로고침 통합과 기존 버튼 전환 회귀 검증
# 2026-10-20 김좌현 수정 : 항목 숨김 제거 후 5개 버튼 구성 검증
# 2026-10-20 김좌현 수정 : 표시 기간 유지 및 제거된 only 설정의 행 필터 미적용 검증
# 2026-10-20 김좌현 수정 : config 타입 목록 이전과 재초기화 시 값 보존 검증
# 2026-10-20 김좌현 수정 : 휴일 표 이전 시 날짜 보존 및 휴일 읽기 검증
# 2026-10-20 김좌현 수정 : 완료 숨김을 포함한 6개 버튼 구성 검증
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('task-layout-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath $WorkbookPath -Destination $taskTemp
$taskExcel = $null
$taskBook = $null
$taskPhase = 'Excel startup'
try {
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskExcel.AutomationSecurity = 1
    $taskPhase = 'open temporary workbook'
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    foreach ($obsolete in @('modDailyProgressReport', 'modWeeklyPptReport')) {
        try { $component = $taskComponents.Item($obsolete) } catch { $component = $null }
        if ($null -ne $component) { $taskComponents.Remove($component) }
    }
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
        $taskPhase = 'import ' + $file.BaseName
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        try { $source = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
        catch { $source = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
        $source = $source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        try { $component = $taskComponents.Item($file.BaseName) }
        catch { $component = $taskComponents.Add(1); $component.Name = $file.BaseName }
        $code = $component.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    $taskPhase = 'import workbook class'
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls'))
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $code = $taskComponents.Item([string]$taskBook.CodeName).CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($source)
    $harness = $taskComponents.Add(1)
    $harness.Name = 'TaskLayoutRegression'
    $harness.CodeModule.AddFromString(@'
Option Explicit
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 910, "TaskLayoutRegression", label
End Sub
Public Function RunTaskLayoutRegression() As String
    Dim ws As Worksheet, r As Long, c As Long, lastRow As Long
    Dim requests As Variant, modifications As Variant, details As Variant, manual As Variant
    Dim name As Variant, headerBefore As String, wasLegacy As Boolean
    Dim startDate As Date, endDate As Date, reportWs As Worksheet, retiredWs As Worksheet
    Dim configuredStart As Variant, configuredEnd As Variant, configuredModification As Variant
    Dim reversedPeriodError As Long
    Dim buttonLeft As Double, buttonTop As Double, buttonCount As Long
    Dim legacyButton As Shape, setupWs As Worksheet, shape As Shape
    Dim displayConfig As Worksheet, oldDisplaySettings As Variant
    Dim typeConfig As Worksheet, originalTypes As Variant, typeIndex As Long
    Dim hiddenBeforeRefresh() As Boolean
    Dim oldRequestCol As Long, oldModificationCol As Long, oldOwnerCol As Long, oldManualCol As Long
    Dim headerCell As Range
    Dim holidayValues As Variant, holidayStart As Long, holidayEnd As Long, holidayRows As Long
    Dim holidays As Object, workdays As Object, holidayKey As String
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    UnprotectTaskSheet ws
    Set headerCell = ws.Rows(HEADER_ROW).Find(What:="요청 내용", LookAt:=xlWhole, MatchCase:=False)
    If headerCell Is Nothing Then Set headerCell = ws.Rows(HEADER_ROW).Find(What:="내용", LookAt:=xlWhole, MatchCase:=False)
    Check Not headerCell Is Nothing, "original request header missing"
    oldRequestCol = headerCell.Column
    lastRow = ws.Cells(ws.Rows.Count, oldRequestCol).End(xlUp).Row
    Set headerCell = ws.Rows(HEADER_ROW).Find(What:="수정 내용", LookAt:=xlWhole, MatchCase:=False)
    wasLegacy = (headerCell Is Nothing)
    If Not wasLegacy Then oldModificationCol = headerCell.Column
    Set headerCell = ws.Rows(HEADER_ROW).Find(What:="담당", LookAt:=xlWhole, MatchCase:=False)
    Check Not headerCell Is Nothing, "original owner header missing"
    oldOwnerCol = headerCell.Column
    Set headerCell = ws.Rows(HEADER_ROW).Find(What:="진행률 수동", LookAt:=xlWhole, MatchCase:=False)
    Check Not headerCell Is Nothing, "original manual progress header missing"
    oldManualCol = headerCell.Column
    requests = ws.Range(ws.Cells(5, oldRequestCol), ws.Cells(lastRow, oldRequestCol)).Value2
    If Not wasLegacy Then modifications = ws.Range(ws.Cells(5, oldModificationCol), ws.Cells(lastRow, oldModificationCol)).Value2
    details = ws.Range(ws.Cells(5, oldOwnerCol), ws.Cells(lastRow, oldOwnerCol + 5)).Value2
    manual = ws.Range(ws.Cells(5, oldManualCol), ws.Cells(lastRow, oldManualCol + 2)).Value2
    Set typeConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    holidayStart = HOLIDAY_DATA_START_ROW
    If CStr(typeConfig.Range("A1").Value2) = "날짜" Then holidayStart = 2
    holidayEnd = typeConfig.Cells(typeConfig.Rows.Count, "A").End(xlUp).Row
    If holidayEnd >= holidayStart Then
        holidayValues = typeConfig.Range("A" & holidayStart & ":C" & holidayEnd).Value2
        holidayRows = holidayEnd - holidayStart + 1
    End If
    If CStr(typeConfig.Range("L9").Value2) = "허용 타입 목록" Then
        originalTypes = typeConfig.Range("L10:L29").Value2
    Else
        originalTypes = typeConfig.Range(TASK_TYPE_CONFIG_RANGE).Value2
    End If
    SetupDataHeaders ws
    SetupDataHeaders ws
    For typeIndex = 1 To 20
        Check CStr(typeConfig.Range(TASK_TYPE_CONFIG_RANGE).Cells(typeIndex, 1).Value2) = _
              CStr(originalTypes(typeIndex, 1)), "type list changed during migration"
    Next typeIndex
    Check Len(CStr(typeConfig.Range("L9").Value2)) = 0, "old type title survived"
    Check Len(CStr(typeConfig.Range("M10").Value2)) = 0, "old type guide survived"
    Check CStr(ws.Range("R4").Value2) = "요청 내용", "request header"
    Check CStr(ws.Range("H4").Value2) = "WEB UPDATE", "WEB UPDATE header missing"
    Check CStr(ws.Range("Q4").Value2) = "웹업무ID", "persistent task ID missing"
    Check CStr(ws.Range("Z4").Value2) = "작업시간(분)", "minutes column missing"
    Check ws.Columns("Q").Hidden, "private ID exposed"
    Check CStr(ws.Range("S4").Value2) = "수정 내용", "modification header"
    Check CStr(ws.Range("T4").Value2) = "담당", "owner header shifted twice"
    Check CStr(ws.Range("U4").Value2) = "비고", "note header shifted twice"
    Check CStr(ws.Range("V4").Value2) = "계획 시작일", "date header shifted twice"
    For r = 1 To UBound(requests, 1)
        Check CStr(ws.Cells(r + 4, COL_REQUEST).Value2) = CStr(requests(r, 1)), "request lost row " & r
        If wasLegacy Then Check Len(CStr(ws.Cells(r + 4, COL_MODIFICATION).Value2)) = 0, "new modification not blank"
        If Not wasLegacy Then Check CStr(ws.Cells(r + 4, COL_MODIFICATION).Value2) = CStr(modifications(r, 1)), "modification lost row " & r
        For c = 1 To 6
            Check CStr(ws.Cells(r + 4, c + ws.Range(COL_OWNER & "1").Column - 1).Value2) = CStr(details(r, c)), "owner/note/date lost row " & r & " column " & c
        Next c
        For c = 1 To 3
            Check CStr(ws.Cells(r + 4, c + ws.Range(COL_MANUAL_PROGRESS & "1").Column - 1).Value2) = CStr(manual(r, c)), "manual/status lost row " & r
        Next c
    Next r
    For Each name In Array("WeeklyPptTemplate", "_일별진행현황템플릿", "_일별진척률이력", "config_일일현황")
        Set retiredWs = Nothing
        On Error Resume Next
        Set retiredWs = ThisWorkbook.Worksheets(CStr(name))
        On Error GoTo Failed
        If Not retiredWs Is Nothing Then
            retiredWs.Visible = xlSheetVisible
            retiredWs.Delete
        End If
    Next name
    EnsureConfigSheet
    EnsureConfigSheet
    Check CStr(typeConfig.Range("A1").Value2) = "휴일·근무일 설정", "holiday table title missing"
    Check CStr(typeConfig.Cells(HOLIDAY_HEADER_ROW, 1).Value2) = "날짜", "holiday date header missing"
    LoadHolidaySettings holidays, workdays
    For r = 1 To holidayRows
        For c = 1 To 3
            Check CStr(typeConfig.Cells(HOLIDAY_DATA_START_ROW + r - 1, c).Value2) = _
                  CStr(holidayValues(r, c)), "holiday value lost during move"
        Next c
        If IsDate(typeConfig.Cells(HOLIDAY_DATA_START_ROW + r - 1, 1).Value) Then
            holidayKey = NormalizeDateKey(CDate(typeConfig.Cells(HOLIDAY_DATA_START_ROW + r - 1, 1).Value))
            If CStr(holidayValues(r, 2)) = HOLIDAY_TYPE_HOLIDAY Then Check holidays.Exists(holidayKey), "holiday no longer loaded"
            If CStr(holidayValues(r, 2)) = HOLIDAY_TYPE_WORKDAY Then Check workdays.Exists(holidayKey), "workday no longer loaded"
        End If
    Next r
    Set reportWs = ThisWorkbook.Worksheets(OUTPUT_CONFIG_SHEET_NAME)
    configuredStart = reportWs.Range(OUTPUT_REPORT_START_CELL).Value
    configuredEnd = reportWs.Range(OUTPUT_REPORT_END_CELL).Value
    configuredModification = reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "N"
    EnsureWeeklyReportConfigSheet
    Check Not GetOutputShowModificationFlag(), "modification toggle reset during refresh"
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "Y"
    Check GetOutputShowModificationFlag(), "modification toggle did not enable"
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = configuredModification
    GetOutputReportPeriod startDate, endDate
    Check startDate <= endDate, "default output period"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = DateSerial(2026, 9, 1)
    reportWs.Range(OUTPUT_REPORT_END_CELL).Value = DateSerial(2026, 9, 30)
    EnsureWeeklyReportConfigSheet
    GetOutputReportPeriod startDate, endDate
    Check startDate = DateSerial(2026, 9, 1) And endDate = DateSerial(2026, 9, 30), "period reset during refresh"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = DateSerial(2026, 10, 1)
    On Error Resume Next
    GetOutputReportPeriod startDate, endDate
    reversedPeriodError = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check reversedPeriodError <> 0, "reversed output period accepted"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = configuredStart
    reportWs.Range(OUTPUT_REPORT_END_CELL).Value = configuredEnd
    CheckWebLayoutMigration
    CheckCopiedTaskRows
    CheckTaskWorkMinutes
    ApplyTaskInputValidation ws
    Check InStr(ws.Range("R5").Validation.Formula1, "LEN(R5)") > 0, "request length validation"
    Check InStr(ws.Range("S5").Validation.Formula1, "LEN(S5)") > 0, "modification length validation"
    ws.Activate
    UnprotectTaskSheet ws
    If ws.FilterMode Then ws.ShowAllData
    ws.Outline.ShowLevels RowLevels:=8
    ShowAllTaskRows ws, GetLastDataRow(ws)
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    ReDim hiddenBeforeRefresh(DATA_START_ROW To GetLastDataRow(ws))
    For r = DATA_START_ROW To GetLastDataRow(ws)
        hiddenBeforeRefresh(r) = ws.Rows(r).Hidden
    Next r
    Set displayConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    oldDisplaySettings = displayConfig.Range("J2:J5").Value2
    displayConfig.Range("J2").Value = DateSerial(2026, 10, 10)
    displayConfig.Range("J3").Value = DateSerial(2026, 10, 20)
    ' Old workbook values must no longer filter rows.
    displayConfig.Range("T4").Value2 = "Y"
    displayConfig.Range("J5").Value2 = "주간보고"
    RefreshGanttSheet False
    Check CDate(ws.Range("AI4").Value) = DateSerial(2026, 10, 10), "display start lost"
    Check CDate(ws.Range("AS4").Value) = DateSerial(2026, 10, 20), "display end lost"
    For r = DATA_START_ROW To GetLastDataRow(ws)
        If HasTaskContent(ws, r) Then Check Not ws.Rows(r).Hidden, "retired only filter hid row " & r & "; before=" & hiddenBeforeRefresh(r)
    Next r
    displayConfig.Range("J2:J5").Value2 = oldDisplaySettings
    UpgradeTaskReportButtons ws
    ' Legacy create-only buttons migrate into the shared action row.
    UnprotectTaskSheet ws
    On Error Resume Next
    ws.Shapes("btnGanttRefresh").Delete
    On Error GoTo Failed
    CreateVersionButton ws, "btnGanttCreate", "생성", "칸트차트_생성", 2, 72
    buttonLeft = ws.Shapes("btnGanttCreate").Left
    buttonTop = ws.Shapes("btnGanttCreate").Top
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    ' Compare the migration itself after Excel has normalized protected shape anchors.
    buttonLeft = ws.Shapes("btnGanttCreate").Left
    buttonTop = ws.Shapes("btnGanttCreate").Top
    UpgradeGanttRefreshButtons ws
    Check ws.ProtectContents, "button migration lost protection"
    ' Excel may quantize shape coordinates by a fraction of one point after protection.
    Check Abs(ws.Shapes("btnGanttRefresh").Top - ws.Shapes("btnGanttRefreshOnly").Top) < 0.5, "refresh buttons are not aligned"
    Check ws.Shapes("btnGanttRefreshOnly").Left > ws.Shapes("btnGanttRefresh").Left, "Gantt-only refresh missing after full refresh"
    Check ws.Shapes("btnGanttRefresh").OnAction = "칸트차트_모두새로고침", "legacy button action"
    Check ws.Shapes("btnGanttRefresh").TextFrame2.TextRange.Text = "모두 새로고침", "legacy button caption"
    Check ws.Shapes("btnGanttRefreshOnly").OnAction = "칸트차트_간트새로고침", "Gantt-only refresh action"
    Check Right$(ws.Shapes("btnWebPreview").OnAction, Len("!PreviewWebTasks")) = "!PreviewWebTasks", "web preview button missing"
    Set legacyButton = Nothing
    On Error Resume Next
    Set legacyButton = ws.Shapes("btnWebUpload")
    On Error GoTo Failed
    Check legacyButton Is Nothing, "duplicate main web upload button survived"
    Check Right$(ws.Shapes("btnWebConfig").OnAction, Len("!OpenWebConfig")) = "!OpenWebConfig", "web config button missing"
    ' When both exist, keep the existing refresh button; also update custom actions.
    UnprotectTaskSheet ws
    CreateVersionButton ws, "btnGanttCreate", "생성", "칸트차트_생성", 5, 72
    Set legacyButton = ws.Shapes.AddShape(msoShapeRectangle, 1, 1, 72, 22)
    legacyButton.Name = "customLegacyRefresh"
    legacyButton.OnAction = "'old.xlsm'!modGanttMain.칸트차트_새로고침"
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    buttonLeft = ws.Shapes("btnGanttRefresh").Left
    buttonTop = ws.Shapes("btnGanttRefresh").Top
    UpgradeGanttRefreshButtons ws
    ThisWorkbook.Saved = True
    UpgradeGanttRefreshButtons ws
    Check ThisWorkbook.Saved, "repeated migration dirtied workbook"
    ' Excel may quantize shape coordinates by a fraction of one point after protection.
    Check Abs(ws.Shapes("btnGanttRefresh").Left - buttonLeft) < 0.5, "existing refresh button replaced " & ws.Shapes("btnGanttRefresh").Left & " vs " & buttonLeft
    Check ws.Shapes("customLegacyRefresh").OnAction = "칸트차트_모두새로고침", "qualified legacy action"
    Set legacyButton = Nothing
    On Error Resume Next
    Set legacyButton = ws.Shapes("btnGanttCreate")
    On Error GoTo Failed
    Check legacyButton Is Nothing, "duplicate create button survived"
    UnprotectTaskSheet ws
    ws.Shapes("customLegacyRefresh").Delete
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    버튼_생성_선택
    Set setupWs = ActiveSheet
    Check setupWs.CheckBoxes.Count >= 7, "split refresh button options missing"
    Check setupWs.DropDowns.Count >= 7, "button order options missing"
    Check setupWs.CheckBoxes("chkButtonOption2").Caption = "모두 새로고침", "full refresh option caption"
    Check setupWs.CheckBoxes("chkButtonOption7").Caption = "간트 새로고침", "Gantt refresh option caption"
    선택한_버튼_생성 False
    Check ActiveSheet.Name = ws.Name, "button creation did not return to task sheet"
    buttonCount = 0
    For Each shape In ws.Shapes
        If shape.Name = "btnGanttRefresh" Then buttonCount = buttonCount + 1
    Next shape
    Check buttonCount = 1, "refresh button duplicated"
    Check ws.Shapes("btnGanttReset").OnAction = "칸트차트_초기화", "reset option shifted"
    Check ws.Shapes("btnGanttObjectInsert").OnAction = "칸트차트_개체삽입", "object option shifted"
    Check ws.Shapes(COMPLETED_TASK_BUTTON_NAME).OnAction = "칸트차트_완료숨김", "completed button action"
    Check Abs(ws.Shapes(COMPLETED_TASK_BUTTON_NAME).Width - ws.Shapes("btnGanttRefresh").Width) < 0.5, "completed button template width " & ws.Shapes(COMPLETED_TASK_BUTTON_NAME).Width & " vs " & ws.Shapes("btnGanttRefresh").Width
    Check ws.Shapes(COMPLETED_TASK_BUTTON_NAME).TextFrame2.TextRange.Font.Size = ws.Shapes("btnGanttRefresh").TextFrame2.TextRange.Font.Size, "completed button template font"
    Check ws.Shapes("btnPeriodTextReport").Left > ws.Shapes("btnGanttObjectInsert").Left, "button order shifted"
    버튼_생성_선택
    버튼_생성_취소
    Check ws.Shapes("btnPeriodTextReport").OnAction = "기간보고TXT_생성", "TXT button action"
    Check ws.Shapes("btnPeriodTextReport").TextFrame2.TextRange.Text = "TXT 출력", "TXT button caption"
    Check ws.Range("S5").Locked = False, "modification input locked"
    Check ws.ProtectContents, "task sheet protection disabled"
    RunTaskLayoutRegression = "PASS: data preservation, config type migration, holiday title/migration/loading, idempotent columns, period validation/persistence, input validation, cross-workbook row insertion without name/link imports, receiving dropdown lists and updates, live hierarchical work-minute sums without double counting, zero/blank and leaf preservation, split refresh, display period preserved without only filters, legacy button migration, seven button options, TXT button, protection"
    Exit Function
Failed:
    RunTaskLayoutRegression = "FAIL: " & Err.Description
End Function
Private Sub CheckWebLayoutMigration()
    Dim fixture As Worksheet, shape As Shape, beforeValue As String, headers As Variant
    Dim c As Long, beforeWidth As Long
    Set fixture = ThisWorkbook.Worksheets.Add
    headers = Array("No.", "Level", "타입", "대분류", "중분류", "소분류", "요청 내용", _
                    "수정 내용", "담당", "비고", "계획 시작일", "계획 종료일", "실제 시작일", _
                    "실제 종료일", "진행률", "정상 진행률", "진행률 수동", "상태 수동", _
                    "보고 상태", "계획일수", "실소요일수", "상태")
    For c = 0 To UBound(headers)
        fixture.Cells(4, c + 2).Value2 = headers(c)
    Next c
    fixture.Range("C5").Value2 = 1
    fixture.Range("H5").Value2 = "migration original"
    fixture.Range("I5").Formula = "=H5&""-formula"""
    fixture.Range("J5").Value2 = "owner"
    fixture.Range("K5").Value2 = "note"
    fixture.Range("N5").Value = DateSerial(2026, 10, 1)
    fixture.Range("O5").Value = DateSerial(2026, 10, 2)
    fixture.Range("P5").Value2 = 0.5
    fixture.Hyperlinks.Add Anchor:=fixture.Range("K5"), Address:="https://example.invalid/migration"
    Set shape = fixture.Shapes.AddShape(msoShapeRectangle, fixture.Range("K6").Left, fixture.Range("K6").Top, 12, 12)
    shape.Placement = xlMoveAndSize
    beforeValue = CStr(fixture.Range("I5").Value2)
    SetupDataHeaders fixture
    SetupDataHeaders fixture
    fixture.Range("H4").Value2 = "웹등록"
    fixture.Range("H5").Value2 = "Y"
    fixture.Range("Q5").Value2 = "preserved-task-id"
    SetupDataHeaders fixture
    SetupDataHeaders fixture
    Check CStr(fixture.Range("H4").Value2) = "WEB UPDATE", "legacy web header was not renamed"
    Check CStr(fixture.Range("H5").Value2) = "Y", "WEB UPDATE selection lost during rename"
    Check CStr(fixture.Range("Q5").Value2) = "preserved-task-id", "task ID moved during header rename"
    Check CStr(fixture.Range("R5").Value2) = "migration original", "fixture request moved incorrectly"
    Check CStr(fixture.Range("S5").Value2) = beforeValue, "formula result lost"
    Check CStr(fixture.Range("S5").Formula) = "=R5&""-formula""", "formula reference not shifted"
    Check CStr(fixture.Range("T5").Value2) = "owner", "fixture owner lost"
    Check fixture.Range("U5").Hyperlinks.Count = 1, "hyperlink anchor lost"
    Check shape.TopLeftCell.Column = fixture.Range("U6").Column, "anchored shape did not move"
    Check CDbl(fixture.Range("AA5").Value2) = 0.5, "progress moved incorrectly"
    Check Len(CStr(fixture.Range("Z5").Value2)) = 0, "new minutes inherited data"
    ApplyCalculatedColumnsProtection fixture, 5
    Check fixture.Range("Q5").Locked And fixture.Columns("Q").Hidden, "private ID not protected"
    Check Not fixture.Range("Z5").Locked, "minutes not editable"
    fixture.Delete
End Sub
Private Sub CheckTaskWorkMinutes()
    Dim fixture As Worksheet, guardSheet As Worksheet, guardBook As Workbook
    Dim r As Long, values(1 To 2, 1 To 1) As Variant, holidays As Object, workdays As Object
    Dim invalidValue As Variant
    Dim oldEvents As Boolean, oldAlerts As Boolean, errorNumber As Long, errorText As String
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set fixture = ThisWorkbook.Worksheets.Add
    fixture.Cells(HEADER_ROW, COL_NO).Value2 = "No."
    fixture.Cells(HEADER_ROW, COL_LEVEL).Value2 = "Level"
    fixture.Cells(HEADER_ROW, COL_REQUEST).Value2 = "요청 내용"
    fixture.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2 = "작업시간(분)"
    For r = 5 To 13
        If r <> 9 Then
            fixture.Cells(r, COL_LEVEL).Value2 = 1
            fixture.Cells(r, COL_REQUEST).Value2 = "work minutes fixture " & CStr(r)
        End If
    Next r
    fixture.Cells(6, COL_LEVEL).Value2 = 2
    fixture.Cells(7, COL_LEVEL).Value2 = 3
    fixture.Cells(8, COL_LEVEL).Value2 = 3
    fixture.Cells(10, COL_LEVEL).Value2 = 2
    fixture.Cells(12, COL_LEVEL).Value2 = 2
    fixture.Cells(5, COL_WORK_MINUTES).Value2 = 999
    fixture.Cells(6, COL_WORK_MINUTES).Value2 = 888
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 30
    fixture.Cells(8, COL_WORK_MINUTES).Value2 = 45
    fixture.Cells(10, COL_WORK_MINUTES).Value2 = 15
    fixture.Cells(12, COL_WORK_MINUTES).Value2 = 20
    fixture.Cells(13, COL_WORK_MINUTES).Formula = "=12+13"
    fixture.Cells(7, COL_PROGRESS).Value2 = 0.5
    fixture.Cells(7, COL_WEEKLY_REPORT).Value2 = "minutes must preserve report"
    PrepareWebTaskSheet fixture
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 75, "setup did not aggregate L3 work minutes"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "setup double counted nested work minutes"
    Check CDbl(fixture.Cells(11, COL_WORK_MINUTES).Value2) = 20, "independent root work minutes missing"
    Check Not Application.EnableEvents And Not fixture.ProtectContents, "setup changed work-minute event/protection state"
    ApplyCalculatedColumnsProtection fixture, 13
    fixture.Rows(8).Hidden = True
    Application.EnableEvents = True

    ' These edits use the production workbook event; no refresh is called.
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 60
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 105, "live child edit did not update parent minutes"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 120, "live child edit did not update all ancestors"
    Check fixture.ProtectContents And Application.EnableEvents, "live minutes edit lost protection or events"
    values(1, 1) = 30: values(2, 1) = 45
    fixture.Range(COL_WORK_MINUTES & "7:" & COL_WORK_MINUTES & "8").Value2 = values
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "multirow minute paste did not sum children"
    Union(fixture.Cells(7, COL_WORK_MINUTES), fixture.Cells(10, COL_WORK_MINUTES)).Value2 = 10
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 55, "multiarea minutes omitted hidden child"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 65, "multiarea minute edit left stale ancestor"
    Check fixture.Rows(8).Hidden, "minute aggregation expanded hidden child"
    Check CDbl(fixture.Cells(11, COL_WORK_MINUTES).Value2) = 20, "minute edit crossed next root boundary"
    Union(fixture.Cells(7, COL_WORK_MINUTES), fixture.Cells(10, COL_WORK_MINUTES)).ClearContents
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 45, "cleared child minutes were retained"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 45, "cleared multiarea minutes left stale root"
    fixture.Cells(8, COL_WORK_MINUTES).ClearContents
    Check IsEmpty(fixture.Cells(6, COL_WORK_MINUTES).Value2), "all-blank child minutes became zero"
    Check IsEmpty(fixture.Cells(5, COL_WORK_MINUTES).Value2), "all-blank branch minutes retained root total"
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 0
    Check Not IsEmpty(fixture.Cells(6, COL_WORK_MINUTES).Value2), "explicit zero minutes became blank"
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 0 And _
          CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 0, "zero minutes did not reach ancestors"
    fixture.Cells(7, COL_WORK_MINUTES).ClearContents
    Check IsEmpty(fixture.Cells(5, COL_WORK_MINUTES).Value2), "clearing last zero left a zero subtotal"
    fixture.Range(COL_WORK_MINUTES & "7:" & COL_WORK_MINUTES & "8").Value2 = values
    fixture.Cells(10, COL_WORK_MINUTES).Value2 = 15
    fixture.Cells(6, COL_WORK_MINUTES).Value2 = 999
    fixture.Cells(5, COL_WORK_MINUTES).Value2 = 999
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 75, "manual parent minutes replaced child total"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "manual root minutes replaced child total"
    fixture.Cells(10, COL_LEVEL).Value2 = 3
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 90, "level change did not move minutes into branch"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "level change double counted child minutes"
    fixture.Cells(10, COL_LEVEL).Value2 = 2
    Check CDbl(fixture.Cells(6, COL_WORK_MINUTES).Value2) = 75, "level change left old branch minutes"
    fixture.Cells(8, COL_REQUEST).ClearContents
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 45, "blank task row contributed stale minutes"
    fixture.Cells(8, COL_MODIFICATION).Value2 = "modification-only child"
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "modification-only task minutes were ignored"
    Check CStr(fixture.Cells(13, COL_WORK_MINUTES).Formula) = "=12+13", "aggregation replaced manual leaf formula"
    Check CDbl(fixture.Cells(13, COL_WORK_MINUTES).Value2) = 25, "aggregation changed standalone leaf minutes"
    Check CDbl(fixture.Cells(7, COL_PROGRESS).Value2) = 0.5, "minute edit changed leaf progress"
    Check CStr(fixture.Cells(7, COL_WEEKLY_REPORT).Value2) = "minutes must preserve report", "minute edit changed report state"
    Check fixture.ProtectContents And Application.EnableEvents, "multiple minute edits lost protection or events"
    For Each invalidValue In Array("invalid work minutes", -1, 1.5, True, 2147483648#, CVErr(xlErrDiv0))
        fixture.Cells(7, COL_WORK_MINUTES).Value2 = invalidValue
        Check CStr(fixture.Cells(7, COL_WORK_MINUTES).Value2) = CStr(invalidValue), "aggregation replaced invalid leaf minutes"
        Check CStr(fixture.Cells(6, COL_WORK_MINUTES).Value2) = CStr(CVErr(xlErrValue)), "invalid child produced a valid-looking subtotal"
        Check CStr(fixture.Cells(5, COL_WORK_MINUTES).Value2) = CStr(CVErr(xlErrValue)), "invalid child did not propagate to root"
        fixture.Cells(7, COL_WORK_MINUTES).Value2 = 30
        Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "correcting invalid child did not restore ancestor totals"
    Next invalidValue
    Check fixture.ProtectContents And Application.EnableEvents, "invalid minute recovery lost protection/events"

    Application.EnableEvents = False
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 25
    UpdateTaskWorkMinutes fixture, 13
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 85, "shared minute aggregation is stale"
    Check fixture.ProtectContents And Not Application.EnableEvents, "shared minute aggregation changed protection/events"
    Set holidays = CreateObject("Scripting.Dictionary")
    Set workdays = CreateObject("Scripting.Dictionary")
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 30
    ApplyHierarchySummaryValues fixture, 13, holidays, workdays
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 90, "hierarchy refresh omitted minutes"
    fixture.Cells(7, COL_WORK_MINUTES).Value2 = 5
    RefreshWebTaskPeriods fixture
    Check CDbl(fixture.Cells(5, COL_WORK_MINUTES).Value2) = 65, "web period refresh omitted minutes"
    Check CStr(fixture.Cells(13, COL_WORK_MINUTES).Formula) = "=12+13", "refresh replaced leaf minute formula"
    Check Not Application.EnableEvents And fixture.ProtectContents, "minute refresh changed protection/events"
    Application.EnableEvents = True
    fixture.Range(COL_WORK_MINUTES & DATA_START_ROW & ":" & COL_WORK_MINUTES & fixture.Rows.Count).ClearContents
    Check IsEmpty(fixture.Cells(5, COL_WORK_MINUTES).Value2) And _
          IsEmpty(fixture.Cells(6, COL_WORK_MINUTES).Value2) And _
          IsEmpty(fixture.Cells(11, COL_WORK_MINUTES).Value2), "clearing full minutes input range retained a subtotal"
    Check Application.CountA(fixture.Range(COL_WORK_MINUTES & "5:" & COL_WORK_MINUTES & "13")) = 0, "full minutes clear retained leaf values"
    Check Application.EnableEvents And fixture.ProtectContents, "full minutes clear lost protection/events"
    Check fixture.Rows(8).Hidden, "full minutes clear expanded hidden child"
    Application.EnableEvents = False

    ' A sheet with similarly placed data but no work-minute header is unrelated.
    Set guardSheet = ThisWorkbook.Worksheets.Add
    guardSheet.Cells(5, COL_LEVEL).Value2 = 1
    guardSheet.Cells(6, COL_LEVEL).Value2 = 2
    guardSheet.Cells(5, COL_REQUEST).Value2 = "unrelated root"
    guardSheet.Cells(6, COL_REQUEST).Value2 = "unrelated child"
    guardSheet.Cells(5, COL_OWNER).Value2 = "unrelated parent owner"
    guardSheet.Cells(6, COL_OWNER).Value2 = "unrelated child owner"
    guardSheet.Cells(5, COL_WORK_MINUTES).Value2 = 100
    Application.EnableEvents = True
    guardSheet.Cells(6, COL_WORK_MINUTES).Value2 = 25
    Check CDbl(guardSheet.Cells(5, COL_WORK_MINUTES).Value2) = 100, "minutes event modified unrelated sheet"
    Check CStr(guardSheet.Cells(5, COL_OWNER).Value2) = "unrelated parent owner", "minutes event modified unrelated owner"
    Application.EnableEvents = False
    guardSheet.Delete
    Set guardSheet = Nothing
    Set guardBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set guardSheet = guardBook.Worksheets(1)
    guardSheet.Name = WEB_PREVIEW_SHEET_NAME
    guardSheet.Cells(HEADER_ROW, COL_REQUEST).Value2 = "요청 내용"
    guardSheet.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2 = "작업시간(분)"
    guardSheet.Cells(5, COL_LEVEL).Value2 = 1
    guardSheet.Cells(6, COL_LEVEL).Value2 = 2
    guardSheet.Cells(5, COL_REQUEST).Value2 = "utility root"
    guardSheet.Cells(6, COL_REQUEST).Value2 = "utility child"
    guardSheet.Cells(5, COL_WORK_MINUTES).Value2 = 100
    guardSheet.Cells(6, COL_WORK_MINUTES).Value2 = 25
    HandleTaskHierarchyChange guardSheet, guardSheet.Cells(6, COL_WORK_MINUTES)
    Check CDbl(guardSheet.Cells(5, COL_WORK_MINUTES).Value2) = 100, "minutes handler modified utility sheet"
    GoTo CleanUp
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
CleanUp:
    On Error Resume Next
    Application.EnableEvents = False
    If Not guardBook Is Nothing Then guardBook.Close SaveChanges:=False
    If guardBook Is Nothing And Not guardSheet Is Nothing Then guardSheet.Delete
    If Not fixture Is Nothing Then fixture.Delete
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "CheckTaskWorkMinutes", errorText
End Sub
Private Sub CheckCopiedTaskRows()
    Dim fixture As Worksheet, blankBook As Workbook, targetBook As Workbook
    Dim blankSheet As Worksheet, targetSheet As Worksheet, listSheet As Worksheet
    Dim listNames As Variant, listColumns As Variant, refs(0 To 5) As String
    Dim i As Long, r As Long, copyIndex As Long, destination As Long, copiedRows As Long
    Dim expectedNames As Long, links As Variant, item As Name
    Dim oldAlerts As Boolean, errorNumber As Long, errorText As String
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    listNames = Array(TASK_TYPE_LIST_NAME, "WebRequesters", "WebActions", "WebStates", "WebModules", "WebHandlers")
    listColumns = Array(COL_TYPE, COL_WEB_REQUESTER, COL_WEB_ACTION, COL_WEB_STATUS, COL_WEB_MODULE, COL_WEB_HANDLER)
    Set fixture = ThisWorkbook.Worksheets.Add
    fixture.Cells(HEADER_ROW, COL_NO).Value2 = "No."
    fixture.Cells(HEADER_ROW, COL_REQUEST).Value2 = "요청 내용"
    For r = 5 To 6
        fixture.Cells(r, COL_REQUEST).Value2 = "source row " & CStr(r)
        For i = 0 To UBound(listNames)
            fixture.Cells(r, CStr(listColumns(i))).Value2 = "source choice " & CStr(i) & " row " & CStr(r)
        Next i
    Next r
    ApplyTaskTypeValidation fixture
    ApplyWebTaskValidation fixture
    ' Local fold/filter metadata must not be imported merely because it spans the copied rows.
    fixture.Names.Add Name:="'" & fixture.Name & "'!_GanttDetailFoldRange", RefersTo:="='" & fixture.Name & "'!$S:$AH"
    fixture.Names.Add Name:="'" & fixture.Name & "'!_GanttHideCompleted", RefersTo:="=1"
    fixture.Names.Add Name:="'" & fixture.Name & "'!_GanttDoneOwned_1", RefersTo:="='" & fixture.Name & "'!$5:$6"
    fixture.Range("B4:AH6").AutoFilter
    Set blankBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set blankSheet = blankBook.Worksheets(1)
    Set targetBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set targetSheet = targetBook.Worksheets(1)
    Set listSheet = targetBook.Worksheets.Add
    listSheet.Name = "ReceivingLists"
    For i = 0 To UBound(listNames)
        listSheet.Cells(1, i + 1).Value2 = "target first " & CStr(i)
        listSheet.Cells(2, i + 1).Value2 = "target second " & CStr(i)
        Set item = targetBook.Names.Add(Name:=CStr(listNames(i)), _
            RefersTo:="='" & listSheet.Name & "'!" & listSheet.Range(listSheet.Cells(1, i + 1), listSheet.Cells(2, i + 1)).Address)
        refs(i) = item.RefersTo
    Next i
    targetSheet.Names.Add Name:="'" & targetSheet.Name & "'!_GanttDetailFoldRange", RefersTo:="='" & targetSheet.Name & "'!$T:$AI"
    targetSheet.Names.Add Name:="'" & targetSheet.Name & "'!_GanttHideCompleted", RefersTo:="=0"
    expectedNames = targetBook.Names.Count
    targetSheet.Range("R10").Value2 = "existing destination row"
    For copyIndex = 1 To 3
        destination = copyIndex * 10
        copiedRows = IIf(copyIndex = 2, 2, 1)
        fixture.Rows("5:" & CStr(4 + copiedRows)).Copy
        blankSheet.Rows(CStr(destination) & ":" & CStr(destination + copiedRows - 1)).Insert Shift:=xlDown
        Check blankBook.Names.Count = 0, "copied rows imported names into blank workbook"
        links = blankBook.LinkSources(xlExcelLinks)
        Check IsEmpty(links), "copied rows created external links in blank workbook"
        fixture.Rows("5:" & CStr(4 + copiedRows)).Copy
        targetSheet.Rows(CStr(destination) & ":" & CStr(destination + copiedRows - 1)).Insert Shift:=xlDown
        Application.CutCopyMode = False
        Check targetBook.Names.Count = expectedNames, "copied rows changed receiving workbook name count"
        links = targetBook.LinkSources(xlExcelLinks)
        Check IsEmpty(links), "copied rows created external links in receiving workbook"
        For r = 0 To copiedRows - 1
            Check CStr(blankSheet.Cells(destination + r, COL_REQUEST).Value2) = "source row " & CStr(5 + r), "blank workbook lost copied task data"
            Check CStr(targetSheet.Cells(destination + r, COL_REQUEST).Value2) = "source row " & CStr(5 + r), "receiving workbook lost copied task data"
            For i = 0 To UBound(listNames)
                Check CStr(targetSheet.Cells(destination + r, CStr(listColumns(i))).Value2) = _
                    CStr(fixture.Cells(5 + r, CStr(listColumns(i))).Value2), "copied selection value changed"
                Check CStr(targetSheet.Cells(destination + r, CStr(listColumns(i))).Validation.Formula1) = _
                    CStr(fixture.Cells(5 + r, CStr(listColumns(i))).Validation.Formula1), "copied dropdown formula changed"
            Next i
        Next r
        For i = 0 To UBound(listNames)
            Check targetBook.Names(CStr(listNames(i))).RefersTo = refs(i), "receiving list reference changed"
            targetSheet.Cells(destination, CStr(listColumns(i))).Value2 = listSheet.Cells(1, i + 1).Value2
            Check targetSheet.Cells(destination, CStr(listColumns(i))).Validation.Value, "copied dropdown rejected receiving list value"
            targetSheet.Cells(destination, CStr(listColumns(i))).Value2 = "source choice " & CStr(i) & " row 5"
            Check Not targetSheet.Cells(destination, CStr(listColumns(i))).Validation.Value, "copied dropdown ignored receiving list"
        Next i
    Next copyIndex
    Check CStr(targetSheet.Range("R11").Value2) = "existing destination row", "row insertion overwrote existing task"
    Check targetSheet.Names("_GanttHideCompleted").RefersTo = "=0", "receiving completed-view flag changed"
    Check targetSheet.Names("_GanttDetailFoldRange").RefersToRange.Column = 20, "receiving folded columns changed"
    For i = 0 To UBound(listNames)
        listSheet.Cells(1, i + 1).Value2 = "updated receiving choice " & CStr(i)
        targetSheet.Cells(30, CStr(listColumns(i))).Value2 = listSheet.Cells(1, i + 1).Value2
        Check targetSheet.Cells(30, CStr(listColumns(i))).Validation.Value, "copied dropdown did not follow receiving list update"
        targetSheet.Cells(30, CStr(listColumns(i))).Value2 = "target first " & CStr(i)
        Check Not targetSheet.Cells(30, CStr(listColumns(i))).Validation.Value, "copied dropdown retained obsolete receiving value"
        Check CStr(fixture.Cells(5, CStr(listColumns(i))).Value2) = "source choice " & CStr(i) & " row 5", "copy changed source task"
    Next i
    GoTo CleanUp
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not blankBook Is Nothing Then blankBook.Close SaveChanges:=False
    If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    If Not fixture Is Nothing Then fixture.Delete
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "CheckCopiedTaskRows", errorText
End Sub
'@)
    $taskPhase = 'save and reload imported VBA'
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('TaskLayoutRegression')
    $taskPhase = 'run VBA layout regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunTaskLayoutRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    $taskComponents.Remove($harness)
    if ($SavePath) {
        $taskPhase = 'save updated copy'
        $taskBook.SaveAs([IO.Path]::GetFullPath($SavePath), 52)
        Write-Output ('UPDATED COPY: ' + [IO.Path]::GetFullPath($SavePath))
    }
}
catch {
    Write-Error ('Failed during ' + $taskPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
    throw
}
finally {
    if ($null -ne $taskBook) { try { $taskBook.Close($false) } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook) }
    if ($null -ne $taskExcel) { try { $taskExcel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp -Force }
}
