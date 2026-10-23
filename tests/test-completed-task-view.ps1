[CmdletBinding()]
param([string]$WorkbookPath, [switch]$VisibilityOnly)

# 2026-10-23 김좌현 수정 : 1·2·3단계 헤더 고정·선택 보존·저장 후 시트 전환 회귀 실행
# 2026-10-20 김좌현 수정 : 완료 숨김 기능의 실제 Excel 회귀 검사
# 2026-10-20 김좌현 수정 : 표시 창의 실제 틀 고정/모니터 회귀 및 임시 소스 오류 주입으로 새로고침 복구 검사
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $config = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $config.excel_file
}
$testPath = Join-Path $env:TEMP ('completed-task-view-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath $WorkbookPath -Destination $testPath
$excel = $null
$book = $null
function Read-VbaSource([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    try { $sourceText = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
    catch { $sourceText = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
    return $sourceText.TrimStart([char]0xFEFF)
}
try {
    $excel = New-Object -ComObject Excel.Application
    # Hidden or minimized Excel windows do not reproduce the visible-row split calculation.
    $excel.Visible = $true
    $excel.WindowState = -4143
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
        $source = (Read-VbaSource $file.FullName) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        if ($file.BaseName -eq 'modGanttHierarchy') {
            $freezePattern = '(?m)^(?<indent>[ \t]*)\.FreezePanes = False[ \t]*\r?$'
            if (-not [regex]::IsMatch($source, $freezePattern)) { throw 'Cannot instrument frozen-pane recreation.' }
            $source = [regex]::Replace($source, $freezePattern, {
                param($match)
                $match.Groups['indent'].Value + 'CompletedTaskViewRegression.RecordFreezeReset' + "`r`n" + $match.Value
            })
        }
        if ($file.BaseName -eq 'modGanttMain') {
            $failurePattern = '(?m)^(?<indent>[ \t]*)ClearGanttArea ws[ \t]*\r?$'
            $failureMatch = [regex]::Match($source, $failurePattern)
            if (-not $failureMatch.Success) { throw 'Cannot inject refresh failure after ShowAllTaskRows.' }
            $injectedSource = $failureMatch.Groups['indent'].Value + 'If ws.Name = "CompletedViewRefreshFail" Then Err.Raise vbObjectError + 7611, "CompletedTaskViewRegression", "Injected after ShowAllTaskRows"' + "`r`n"
            $source = $source.Insert($failureMatch.Index, $injectedSource)
        }
        try { $component = $book.VBProject.VBComponents.Item($file.BaseName) }
        catch { $component = $book.VBProject.VBComponents.Add(1); $component.Name = $file.BaseName }
        $code = $component.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    # Import the event class as well, so tests use the current monitor implementation.
    $source = Read-VbaSource (Join-Path $taskRoot 'vba-files/Class/CGanttOutlineView.cls')
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $source = $source.Replace('Private Sub mCommandBars_OnUpdate()', "Private Sub mCommandBars_OnUpdate()`r`n    CompletedTaskViewRegression.RecordNativeUpdate")
    try { $component = $book.VBProject.VBComponents.Item('CGanttOutlineView') }
    catch { $component = $book.VBProject.VBComponents.Add(2); $component.Name = 'CGanttOutlineView' }
    $code = $component.CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($source)
    $source = Read-VbaSource (Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls')
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $code = $book.VBProject.VBComponents.Item([string]$book.CodeName).CodeModule
    $code.DeleteLines(1, $code.CountOfLines)
    $code.AddFromString($source)
    try { $harness = $book.VBProject.VBComponents.Item('CompletedTaskViewRegression') }
    catch { $harness = $book.VBProject.VBComponents.Add(1); $harness.Name = 'CompletedTaskViewRegression' }
    $source = Read-VbaSource (Join-Path $PSScriptRoot 'CompletedTaskViewRegression.bas')
    $source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    if ($harness.CodeModule.CountOfLines -gt 0) { $harness.CodeModule.DeleteLines(1, $harness.CodeModule.CountOfLines) }
    $harness.CodeModule.AddFromString($source)
    $book.Save()
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    if ($VisibilityOnly) {
        $result = [string]$excel.Run("'$($book.Name)'!RunCompletedTaskVisibilityRegression")
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw $result }
        return
    }
    $result = [string]$excel.Run("'$($book.Name)'!RunCompletedTaskRegression")
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
    $taskFailures = @()
    foreach ($procedure in @('RunCompletedTaskVisibilityRegression', 'RunCompletedTaskViewOrderRegression', 'RunRefreshFailureRegression')) {
        $result = [string]$excel.Run("'$($book.Name)'!$procedure")
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { $taskFailures += $result }
    }
    if ($taskFailures.Count -gt 0) { throw ($taskFailures -join "`n") }
    $book.Save()
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $excel.EnableEvents = $true
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    $result = [string]$excel.Run("'$($book.Name)'!RunSavedCompletedTaskRegression")
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
}
finally {
    if ($excel) { try { $excel.EnableEvents = $false } catch {} }
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if (Test-Path -LiteralPath $testPath) { Remove-Item -LiteralPath $testPath -Force }
}
