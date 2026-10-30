[CmdletBinding()]
param([string]$WorkbookPath)

# 2026-10-29 김좌현 수정 : 통합 새로고침 호출에 맞춰 검증 단계 추적 갱신
# 2026-10-22 김좌현 수정 : 실제 Excel 임시 사본에서 웹 모드/간트 전용 경로와 재열기 검증
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('web-view-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath $WorkbookPath -Destination $taskTemp
$taskExcel = $null
$taskBook = $null
function Read-TaskVba([string]$taskPath) {
    $taskBytes = [IO.File]::ReadAllBytes($taskPath)
    try { $taskText = [Text.UTF8Encoding]::new($false, $true).GetString($taskBytes) }
    catch { $taskText = [Text.Encoding]::GetEncoding(949).GetString($taskBytes) }
    return $taskText.TrimStart([char]0xFEFF)
}
try {
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskExcel.AutomationSecurity = 1
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    foreach ($taskFile in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
        $taskSource = Read-TaskVba $taskFile.FullName
        $taskSource = $taskSource -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        if ($taskFile.BaseName -eq 'modHoliday') {
            $taskSource = $taskSource.Replace('Public Sub EnsureConfigSheet()', "Public Sub EnsureConfigSheet()`r`n    WebViewRegression.ConfigurationCalls = WebViewRegression.ConfigurationCalls + 1")
        }
        if ($taskFile.BaseName -eq 'modWebUpload') {
            $taskSource = $taskSource.Replace('Public Sub PrepareWebTaskSheet(ByVal ws As Worksheet)', "Public Sub PrepareWebTaskSheet(ByVal ws As Worksheet)`r`n    WebViewRegression.WebSetupCalls = WebViewRegression.WebSetupCalls + 1")
        }
        if ($taskFile.BaseName -eq 'modGanttBars') {
            $taskSource = $taskSource.Replace('    ClearTaskBarVisuals ws', "    WebViewRegression.RenderCalls = WebViewRegression.RenderCalls + 1`r`n    ClearTaskBarVisuals ws")
        }
        if ($taskFile.BaseName -eq 'modGanttMain') {
            foreach ($taskCall in @('EnsureTaskSheetLayout ws', 'EnsureConfigSheet', 'SetupDataHeaders ws', 'UpgradeGanttRefreshButtons ws', 'PrepareWebTaskSheet ws', 'LoadHolidaySettings holidayDict, workdayDict', 'ApplyHierarchySummaryValues ws, lastRow, holidayDict, workdayDict', 'ClearGanttArea ws', 'FormatBaseArea ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict', 'ApplyCalculatedColumnsProtection ws, lastRow', 'RefreshWebModeView ws, True')) {
                $taskSource = $taskSource.Replace("    $taskCall`r`n", "    WebViewRegression.RefreshPhase = `"$taskCall`"`r`n    $taskCall`r`n")
            }
        }
        try { $taskComponent = $taskBook.VBProject.VBComponents.Item($taskFile.BaseName) }
        catch { $taskComponent = $taskBook.VBProject.VBComponents.Add(1); $taskComponent.Name = $taskFile.BaseName }
        $taskCode = $taskComponent.CodeModule
        if ($taskCode.CountOfLines -gt 0) { $taskCode.DeleteLines(1, $taskCode.CountOfLines) }
        $taskCode.AddFromString($taskSource)
    }
    # 2026-10-29 김좌현 수정 : 더 이상 생성하지 않는 구형 CommandBars 이벤트 클래스 제거
    try { $taskComponent = $taskBook.VBProject.VBComponents.Item('CGanttOutlineView') } catch { $taskComponent = $null }
    if ($taskComponent) { $taskBook.VBProject.VBComponents.Remove($taskComponent) }
    $taskSource = Read-TaskVba (Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls')
    $taskSource = $taskSource.Substring($taskSource.IndexOf('Option Explicit'))
    $taskCode = $taskBook.VBProject.VBComponents.Item([string]$taskBook.CodeName).CodeModule
    $taskCode.DeleteLines(1, $taskCode.CountOfLines)
    $taskCode.AddFromString($taskSource)
    $taskComponent = $taskBook.VBProject.VBComponents.Add(1)
    $taskComponent.Name = 'WebViewRegression'
    $taskSource = Read-TaskVba (Join-Path $PSScriptRoot 'WebViewRegression.bas')
    $taskSource = $taskSource -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    $taskComponent.CodeModule.AddFromString($taskSource)
    $taskResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunWebViewRegression")
    Write-Output $taskResult
    if (-not $taskResult.StartsWith('PASS:')) { throw $taskResult }
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskExcel.EnableEvents = $true
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunSavedWebViewRegression")
    Write-Output $taskResult
    if (-not $taskResult.StartsWith('PASS:')) { throw $taskResult }
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunPartialSavedWebViewRegression")
    Write-Output $taskResult
    if (-not $taskResult.StartsWith('PASS:')) { throw $taskResult }
}
finally {
    if ($taskExcel) { try { $taskExcel.EnableEvents = $false } catch {} }
    if ($taskBook) { try { $taskBook.Close($false) } catch {} }
    if ($taskExcel) { try { $taskExcel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
    if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp -Force }
}
