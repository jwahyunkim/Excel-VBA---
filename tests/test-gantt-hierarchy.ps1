[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath, [switch]$ViewOnly, [switch]$SelectedColumnsOnly)

# 2026-10-23 김좌현 수정 : 모든 열 접기 단계의 헤더 고정과 펼친 단계의 가로 고정 해제 검증
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
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('gantt-hierarchy-' + [guid]::NewGuid().ToString('N') + '.xlsm')
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
    $harness.Name = 'GanttHierarchyRegression'
    $harness.CodeModule.AddFromString(@'
Option Explicit
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 920, "GanttHierarchyRegression", label
End Sub
Private Sub CheckStage(ByVal ws As Worksheet, ByVal stage As Long)
    Dim c As Long, header As Long
    ws.Outline.ShowLevels ColumnLevels:=stage
    EnsureTaskRequestFreeze ws
    Check ActiveWindow.FreezePanes And ActiveWindow.SplitRow = HEADER_ROW, "fixed header missing at stage " & stage
    Check ActiveWindow.Panes(1).ScrollRow = 1, "frozen headers start below row 1 at stage " & stage
    If stage > 1 Then Check ActiveWindow.SplitColumn = 0, "expanded stage has frozen columns " & stage
    For header = 1 To HEADER_ROW
        Check Not ws.Rows(header).Hidden, "hidden header at stage " & stage
    Next header
    For c = ws.Range(COL_WEB_ENABLED & "1").Column To ws.Range(COL_WEB_RESULT & "1").Column
        Check ws.Columns(c).OutlineLevel = 3, "web column depth " & c
        Check CBool(ws.Columns(c).Hidden) = (stage < 3), "web visibility " & stage & "/" & c
    Next c
    Check ws.Columns(COL_WORK_MINUTES).OutlineLevel = 3, "minutes depth"
    Check CBool(ws.Columns(COL_WORK_MINUTES).Hidden) = (stage < 3), "minutes visibility " & stage
    Check Not ws.Columns(COL_REQUEST).Hidden, "request hidden at stage " & stage
    Check ws.Columns(COL_WEB_TASK_ID).Hidden And ws.Columns(COL_WEB_TASK_ID).OutlineLevel = 1, "private ID exposed/grouped"
    Check ws.Rows(6).OutlineLevel = 2 And ws.Rows(7).OutlineLevel = 3, "column stage changed row groups"
End Sub
Public Function RunGanttHierarchyRegression() As String
    Dim ws As Worksheet, original As Worksheet, config As Worksheet, levels As Variant
    Dim headers As Variant, c As Long, r As Long, stage As Long, hidden As Variant
    Dim fixtureValue As String, selectedColumns As String
    On Error GoTo Failed
    Application.EnableEvents = False
    EnsureConfigSheet
    Set config = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    config.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    config.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "G"
    On Error Resume Next
    ThisWorkbook.Names("_GanttWebFoldV6").Delete
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "WebOutlineRegression"
    headers = Array("No.", "Level", "타입", "대분류", "중분류", "소분류", "요청 내용", _
                    "수정 내용", "담당", "비고", "계획 시작일", "계획 종료일", "실제 시작일", _
                    "실제 종료일", "진행률", "정상 진행률", "진행률 수동", "상태 수동", _
                    "보고 상태", "계획일수", "실소요일수", "상태")
    For c = 0 To UBound(headers)
        ws.Cells(4, c + 2).Value2 = headers(c)
    Next c
    levels = Array(1, 2, 3, 3, 2, 1, 1, 3)
    For r = 5 To 12
        ws.Cells(r, "C").Value2 = levels(r - 5)
        ws.Cells(r, "H").Value2 = "Hierarchy " & r
        ws.Cells(r, "I").Value2 = "Modification " & r
    Next r
    SetupDataHeaders ws
    Check CStr(config.Range(TASK_FOLD_START_VALUE_CELL).Value2) = COL_MODIFICATION, "old fold start not migrated"
    Check CStr(config.Range(TASK_FOLD_END_VALUE_CELL).Value2) = COL_STATUS, "old fold end not migrated"
    Check CStr(config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2) = "G", "old G freeze changed"
    SetupDataHeaders ws
    Check CStr(config.Range(TASK_FOLD_START_VALUE_CELL).Value2) = COL_MODIFICATION, "fold migrated twice"
    ws.Cells(5, COL_WORK_MINUTES).Value2 = 30
    ws.Cells(5, COL_WEB_TASK_ID).Value2 = "stable-id-test"
    fixtureValue = CStr(ws.Cells(7, COL_MODIFICATION).Value2)
    ws.Activate
    RebuildTaskHierarchyOutline ws, 12
    EnsureTaskDetailColumnOutline ws
    ApplyCalculatedColumnsProtection ws, 12
    Check ws.Rows(5).OutlineLevel = 1 And ws.Rows(10).OutlineLevel = 1, "root boundary"
    Check ws.Rows(12).OutlineLevel = 2, "level-1-to-3 boundary"
    For stage = 1 To 3
        CheckStage ws, stage
        For c = ws.Range(COL_MODIFICATION & "1").Column To ws.Range(COL_STATUS & "1").Column
            If c <> ws.Range(COL_WORK_MINUTES & "1").Column Then
                Check ws.Columns(c).OutlineLevel = 2, "ordinary detail depth " & c
                Check CBool(ws.Columns(c).Hidden) = (stage = 1), "ordinary visibility " & stage & "/" & c
            End If
        Next c
        If stage = 1 Then
            Check ActiveWindow.FreezePanes And ActiveWindow.SplitRow = HEADER_ROW, "compact freeze missing"
            Check ActiveWindow.Panes(1).ScrollColumn = 7, "G freeze not retained"
        Else
            Check ActiveWindow.FreezePanes And ActiveWindow.SplitRow = HEADER_ROW And ActiveWindow.SplitColumn = 0, "expanded header-only freeze missing"
        End If
    Next stage
    CheckStage ws, 2
    hidden = CaptureTaskDetailColumnHidden(ws)
    UnprotectTaskSheet ws
    ws.Columns("H:AH").Hidden = False
    RestoreTaskDetailColumnHidden ws, hidden
    Check ws.Columns("H").Hidden And ws.Columns(COL_WORK_MINUTES).Hidden, "refresh lost stage 2"
    Check Not ws.Columns(COL_MODIFICATION).Hidden, "refresh collapsed ordinary details"
    ApplyCalculatedColumnsProtection ws, 12
    ThisWorkbook.Saved = True
    EnsureTaskDetailColumnOutline ws
    Check ThisWorkbook.Saved, "unchanged outline dirtied workbook"
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "G,H,R,Z"
    UnprotectTaskSheet ws
    EnsureTaskDetailColumnOutline ws
    ApplyCalculatedColumnsProtection ws, 12
    CheckStage ws, 1
    Check ws.Columns("G").OutlineLevel = 1 And ws.Columns("R").OutlineLevel = 1, "selected columns grouped"
    Check ws.Columns("Z").OutlineLevel = 3 And ws.Columns("Z").Hidden, "selected minutes bypassed stage 3"
    CheckStage ws, 2
    Check Not ws.Columns("T").Hidden, "custom stage 2 did not expand ordinary columns"
    CheckStage ws, 3
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "H,Z"
    UnprotectTaskSheet ws
    EnsureTaskDetailColumnOutline ws
    ApplyCalculatedColumnsProtection ws, 12
    CheckStage ws, 1
    CheckStage ws, 3
    selectedColumns = ""
    For c = 1 To ws.Range(COL_STATUS & "1").Column
        If Not IsTaskWebDetailColumn(ws, c) And c <> ws.Range(COL_WEB_TASK_ID & "1").Column Then
            If Len(selectedColumns) > 0 Then selectedColumns = selectedColumns & ","
            selectedColumns = selectedColumns & Replace(ws.Cells(1, c).Address(False, False), "1", "")
        End If
    Next c
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = selectedColumns
    UnprotectTaskSheet ws
    EnsureTaskDetailColumnOutline ws
    ApplyCalculatedColumnsProtection ws, 12
    CheckStage ws, 1
    CheckStage ws, 3
    Check ActiveWindow.FreezePanes And ActiveWindow.SplitRow = HEADER_ROW And ActiveWindow.SplitColumn = 0, "all ordinary fields selected lost header-only freeze"
    config.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "L1"
    EnsureTaskDetailColumnOutline ws
    Check ws.Columns("Z").OutlineLevel = 3, "invalid config changed valid groups"
    config.Range(TASK_FOLD_START_VALUE_CELL).Value2 = COL_MODIFICATION
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "G"
    UnprotectTaskSheet ws
    EnsureTaskDetailColumnOutline ws
    ApplyCalculatedColumnsProtection ws, 12
    CheckStage ws, 2
    Check CStr(ws.Cells(7, COL_MODIFICATION).Value2) = fixtureValue, "outline changed content"
    Check CDbl(ws.Cells(5, COL_WORK_MINUTES).Value2) = 30, "outline changed minutes"
    Check CStr(ws.Cells(5, COL_WEB_TASK_ID).Value2) = "stable-id-test", "outline changed ID"
    Check ws.Cells(5, COL_WEB_TASK_ID).Locked And ws.ProtectContents, "private ID protection missing"
    For Each original In ThisWorkbook.Worksheets
        If original.Name <> ws.Name And CStr(original.Cells(HEADER_ROW, COL_NO).Value2) = "No." Then _
            EnsureTaskSheetLayout original
    Next original
    ws.Activate
    RunGanttHierarchyRegression = "PASS: one-time fold migration/G freeze, fixed headers in all three native stages, selected fields, row boundaries, refresh state and preserved data"
    Exit Function
Failed:
    RunGanttHierarchyRegression = "FAIL: " & Err.Description
End Function
Public Function ReopenGanttHierarchyRegression() As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("WebOutlineRegression")
    ws.Activate
    InitializeTaskHierarchyOutlines
    CheckStage ws, 2
    Check ws.ProtectContents And ws.EnableOutlining, "reopen protection missing"
    Check ws.Cells(5, COL_WEB_TASK_ID).Locked, "reopen ID unlocked"
    ReopenGanttHierarchyRegression = "PASS: saved stage 2 fixed headers and private ID/row groups restored on reopen"
    Exit Function
Failed:
    ReopenGanttHierarchyRegression = "FAIL: " & Err.Description
End Function
Public Function RunOutlineSelectionRegression() As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("WebOutlineRegression")
    ws.Activate
    Application.EnableEvents = True
    ws.Outline.ShowLevels ColumnLevels:=3
    StartTaskOutlineViewMonitor
    Check ActiveWindow.SplitColumn = 0, "stage 3 should not freeze a detail column"
    ws.Outline.ShowLevels ColumnLevels:=1
    ws.Cells(7, COL_REQUEST).Select
    Check ActiveWindow.SplitColumn = 0, "selection event changed panes before returning"
    RunOutlineSelectionRegression = "PASS: stage 1 selection returned without changing panes inside the event"
Finished:
    Exit Function
Failed:
    RunOutlineSelectionRegression = "FAIL: " & Err.Source & " / " & Err.Description
    Application.EnableEvents = False
    Resume Finished
End Function
Public Function VerifyScheduledOutlineSelection() As String
    On Error GoTo Failed
    Check ActiveWindow.FreezePanes And ActiveWindow.SplitRow = HEADER_ROW, "scheduled stage 1 lost frozen headers"
    Check ActiveWindow.SplitColumn > 0, "scheduled stage 1 did not freeze the selected column"
    Check ActiveWindow.Panes(1).ScrollColumn = 7, "scheduled stage 1 lost G viewport"
    VerifyScheduledOutlineSelection = "PASS: deferred stage 1 pane refresh completed"
    Application.EnableEvents = False
    Exit Function
Failed:
    VerifyScheduledOutlineSelection = "FAIL: " & Err.Source & " / " & Err.Description
    Application.EnableEvents = False
End Function
'@)
    $taskPhase = 'save and reload imported VBA'
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('GanttHierarchyRegression')
    $taskPhase = 'three-stage VBA regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunGanttHierarchyRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('GanttHierarchyRegression')
    $taskPhase = 'reopen regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!ReopenGanttHierarchyRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    $taskPhase = 'stage 1 selection event regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunOutlineSelectionRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    Start-Sleep -Seconds 3
    $taskPhase = 'scheduled stage 1 pane regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!VerifyScheduledOutlineSelection")
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
