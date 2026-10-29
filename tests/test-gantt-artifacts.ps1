[CmdletBinding()]
param([string]$SourcePath, [switch]$Baseline)
# 2026-10-29 김좌현 수정 : 숨긴 행의 첨부 개체와 실제 데이터 가져오기 경로 회귀 검증
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$config = Get-Content (Join-Path $root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$testPath = Join-Path ([IO.Path]::GetTempPath()) ('gantt-artifacts-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath (Join-Path $root $config.excel_file) -Destination $testPath
function Read-Vba([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    try { $source = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
    catch { $source = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
    return $source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute .*\r?\n', ''
}
$excel = $null; $book = $null; $sourceBook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    foreach ($file in Get-ChildItem (Join-Path $root 'vba-files/Module') -Filter '*.bas') {
        $source = Read-Vba $file.FullName
        if ($file.BaseName -eq 'modVersionMigration') {
            $source += "`r`n" + (Read-Vba (Join-Path $PSScriptRoot 'GanttArtifactsRegression.bas'))
        }
        $code = $book.VBProject.VBComponents.Item($file.BaseName).CodeModule
        if ($code.CountOfLines) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    if ($SourcePath) {
        $excel.AutomationSecurity = 3
        $sourceBook = $excel.Workbooks.Open($SourcePath, 0, $true)
        $excel.AutomationSecurity = 1
        $result = [string]$excel.Run("'$($book.Name)'!RunActualArtifactRegression", $sourceBook.Name, [bool]$Baseline)
    } else {
        $result = [string]$excel.Run("'$($book.Name)'!RunGanttArtifactsRegression")
    }
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
    $book.Save()
    $book.Close($false)
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    $result = [string]$excel.Run("'$($book.Name)'!RunSavedArtifactRegression")
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
} finally {
    if ($sourceBook) { $sourceBook.Close($false) }
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    if (Test-Path -LiteralPath $testPath) { Remove-Item -LiteralPath $testPath -Force }
}
