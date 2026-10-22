[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath)

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
    Check CStr(ws.Range("H4").Value2) = "웹등록", "web block missing"
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
    Check Right$(ws.Shapes("btnWebUpload").OnAction, Len("!UploadWebTasks")) = "!UploadWebTasks", "web upload button missing"
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
    RunTaskLayoutRegression = "PASS: data preservation, config type migration, holiday title/migration/loading, idempotent columns, period validation/persistence, input validation, split refresh, display period preserved without only filters, legacy button migration, seven button options, TXT button, protection"
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
