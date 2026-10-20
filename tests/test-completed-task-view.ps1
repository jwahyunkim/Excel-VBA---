[CmdletBinding()]
param([string]$WorkbookPath)

# 2026-10-20 김좌현 수정 : 완료 숨김 기능의 실제 Excel 회귀 검사
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
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        try { $source = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
        catch { $source = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
        $source = $source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        try { $component = $book.VBProject.VBComponents.Item($file.BaseName) }
        catch { $component = $book.VBProject.VBComponents.Add(1); $component.Name = $file.BaseName }
        $code = $component.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls'))
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $code = $book.VBProject.VBComponents.Item([string]$book.CodeName).CodeModule
    $code.DeleteLines(1, $code.CountOfLines)
    $code.AddFromString($source)
    $harness = $book.VBProject.VBComponents.Add(1)
    $harness.Name = 'CompletedTaskViewRegression'
    $source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'CompletedTaskViewRegression.bas'))
    $source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    $harness.CodeModule.AddFromString($source)
    $book.Save()
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    $result = [string]$excel.Run("'$($book.Name)'!RunCompletedTaskRegression")
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
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
