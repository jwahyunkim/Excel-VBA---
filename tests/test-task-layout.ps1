[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath)

# 2026-10-20 김좌현 수정 : 새로고침 통합과 기존 버튼 전환 회귀 검증
# 2026-10-20 김좌현 수정 : 항목 숨김 제거 후 5개 버튼 구성 검증
# 2026-10-20 김좌현 수정 : 표시 기간 유지 및 제거된 only 설정의 행 필터 미적용 검증
# 2026-10-20 김좌현 수정 : config 타입 목록 이전과 재초기화 시 값 보존 검증
# 2026-10-20 김좌현 수정 : 휴일 표 이전 시 날짜 보존 및 휴일 읽기 검증
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
    Dim holidayValues As Variant, holidayStart As Long, holidayEnd As Long, holidayRows As Long
    Dim holidays As Object, workdays As Object, holidayKey As String
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    UnprotectTaskSheet ws
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    headerBefore = CStr(ws.Range("I4").Value2)
    wasLegacy = (headerBefore <> "수정 내용")
    requests = ws.Range("H5:H" & lastRow).Value2
    If wasLegacy Then
        details = ws.Range("I5:N" & lastRow).Value2
        manual = ws.Range("Q5:S" & lastRow).Value2
    Else
        modifications = ws.Range("I5:I" & lastRow).Value2
        details = ws.Range("J5:O" & lastRow).Value2
        manual = ws.Range("R5:T" & lastRow).Value2
    End If
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
    Check CStr(ws.Range("H4").Value2) = "요청 내용", "request header"
    Check CStr(ws.Range("I4").Value2) = "수정 내용", "modification header"
    Check CStr(ws.Range("J4").Value2) = "담당", "owner header shifted twice"
    Check CStr(ws.Range("K4").Value2) = "비고", "note header shifted twice"
    Check CStr(ws.Range("L4").Value2) = "계획 시작일", "date header shifted twice"
    For r = 1 To UBound(requests, 1)
        Check CStr(ws.Cells(r + 4, "H").Value2) = CStr(requests(r, 1)), "request lost row " & r
        If wasLegacy Then Check Len(CStr(ws.Cells(r + 4, "I").Value2)) = 0, "new modification not blank"
        If Not wasLegacy Then Check CStr(ws.Cells(r + 4, "I").Value2) = CStr(modifications(r, 1)), "modification lost row " & r
        For c = 1 To 6
            Check CStr(ws.Cells(r + 4, c + 9).Value2) = CStr(details(r, c)), "owner/note/date lost row " & r & " column " & c
        Next c
        For c = 1 To 3
            Check CStr(ws.Cells(r + 4, c + 17).Value2) = CStr(manual(r, c)), "manual/status lost row " & r
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
    ApplyTaskInputValidation ws
    Check InStr(ws.Range("H5").Validation.Formula1, "LEN(H5)") > 0, "request length validation"
    Check InStr(ws.Range("I5").Validation.Formula1, "LEN(I5)") > 0, "modification length validation"
    ws.Activate
    Set displayConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    oldDisplaySettings = displayConfig.Range("J2:J5").Value2
    displayConfig.Range("J2").Value = DateSerial(2026, 10, 10)
    displayConfig.Range("J3").Value = DateSerial(2026, 10, 20)
    ' Old workbook values must no longer filter rows.
    displayConfig.Range("J4").Value2 = "Y"
    displayConfig.Range("J5").Value2 = "주간보고"
    RefreshGanttSheet False
    Check CDate(ws.Range("X4").Value) = DateSerial(2026, 10, 10), "display start lost"
    Check CDate(ws.Range("AH4").Value) = DateSerial(2026, 10, 20), "display end lost"
    For r = DATA_START_ROW To GetLastDataRow(ws)
        If HasTaskContent(ws, r) Then Check Not ws.Rows(r).Hidden, "retired only filter hid row " & r
    Next r
    displayConfig.Range("J2:J5").Value2 = oldDisplaySettings
    UpgradeTaskReportButtons ws
    ' A legacy create-only button must retain its position and remain executable.
    UnprotectTaskSheet ws
    On Error Resume Next
    ws.Shapes("btnGanttRefresh").Delete
    On Error GoTo Failed
    CreateVersionButton ws, "btnGanttCreate", "생성", "칸트차트_생성", 2, 72
    buttonLeft = ws.Shapes("btnGanttCreate").Left
    buttonTop = ws.Shapes("btnGanttCreate").Top
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    UpgradeGanttRefreshButtons ws
    Check ws.ProtectContents, "button migration lost protection"
    Check ws.Shapes("btnGanttRefresh").Left = buttonLeft, "legacy button moved horizontally"
    Check ws.Shapes("btnGanttRefresh").Top = buttonTop, "legacy button moved vertically"
    Check ws.Shapes("btnGanttRefresh").OnAction = "칸트차트_새로고침", "legacy button action"
    Check ws.Shapes("btnGanttRefresh").TextFrame2.TextRange.Text = "새로고침", "legacy button caption"
    ' When both exist, keep the existing refresh button; also update custom actions.
    UnprotectTaskSheet ws
    CreateVersionButton ws, "btnGanttCreate", "생성", "칸트차트_생성", 5, 72
    Set legacyButton = ws.Shapes.AddShape(msoShapeRectangle, 1, 1, 72, 22)
    legacyButton.Name = "customLegacyRefresh"
    legacyButton.OnAction = "'old.xlsm'!modGanttMain.칸트차트_생성"
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    UpgradeGanttRefreshButtons ws
    ThisWorkbook.Saved = True
    UpgradeGanttRefreshButtons ws
    Check ThisWorkbook.Saved, "repeated migration dirtied workbook"
    Check ws.Shapes("btnGanttRefresh").Left = buttonLeft, "existing refresh button replaced"
    Check ws.Shapes("customLegacyRefresh").OnAction = "칸트차트_새로고침", "qualified legacy action"
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
    Check setupWs.CheckBoxes.Count = 5, "button options were not consolidated"
    Check setupWs.DropDowns.Count = 5, "button order options were not consolidated"
    Check setupWs.CheckBoxes("chkButtonOption2").Caption = "간트차트 새로고침", "refresh option caption"
    선택한_버튼_생성 False
    Check ActiveSheet.Name = ws.Name, "button creation did not return to task sheet"
    buttonCount = 0
    For Each shape In ws.Shapes
        If shape.Name = "btnGanttRefresh" Then buttonCount = buttonCount + 1
    Next shape
    Check buttonCount = 1, "refresh button duplicated"
    Check ws.Shapes("btnGanttReset").OnAction = "칸트차트_초기화", "reset option shifted"
    Check ws.Shapes("btnGanttObjectInsert").OnAction = "칸트차트_개체삽입", "object option shifted"
    Check ws.Shapes("btnPeriodTextReport").Left > ws.Shapes("btnGanttObjectInsert").Left, "button order shifted"
    버튼_생성_선택
    버튼_생성_취소
    Check ws.Shapes("btnPeriodTextReport").OnAction = "기간보고TXT_생성", "TXT button action"
    Check ws.Shapes("btnPeriodTextReport").TextFrame2.TextRange.Text = "TXT 출력", "TXT button caption"
    Check ws.Range("I5").Locked = False, "modification input locked"
    Check ws.ProtectContents, "task sheet protection disabled"
    RunTaskLayoutRegression = "PASS: data preservation, config type migration, holiday title/migration/loading, idempotent columns, period validation/persistence, input validation, Gantt refresh, display period preserved without only filters, legacy button migration, five button options, TXT button, protection"
    Exit Function
Failed:
    RunTaskLayoutRegression = "FAIL: " & Err.Description
End Function
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
