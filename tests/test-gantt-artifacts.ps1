[CmdletBinding()]
param([string]$SourcePath, [switch]$Baseline, [switch]$ActualRefresh, [switch]$InspectCompile)
# 2026-10-29 김좌현 수정 : 삭제한 캐시 콜백이 남지 않도록 최신 통합문서 이벤트 주입
# 2026-10-29 김좌현 수정 : 숨긴 행의 첨부 개체와 실제 데이터 가져오기 경로 회귀 검증
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$config = Get-Content (Join-Path $root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$testPath = Join-Path ([IO.Path]::GetTempPath()) ('gantt-artifacts-' + [guid]::NewGuid().ToString('N') + '.xlsm')
if ($ActualRefresh) { Copy-Item -LiteralPath $SourcePath -Destination $testPath }
else { Copy-Item -LiteralPath (Join-Path $root $config.excel_file) -Destination $testPath }
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
    Write-Output 'Excel fixture opened; loading VBA sources.'
    foreach ($file in Get-ChildItem (Join-Path $root 'vba-files/Module') -Filter '*.bas') {
        $source = Read-Vba $file.FullName
        if ($file.BaseName -eq 'modGanttMain') {
            # Exercise the real menu action without displaying modal UI.
            $source = $source.Replace('Private mPptDeleteTarget As Range', "Private mPptDeleteTarget As Range`r`nPublic PptTestDecision As VbMsgBoxResult`r`nPublic PptTestExecuteMenu As Boolean`r`nPublic PptTestPopupCount As Long`r`nPublic PptTestFailCopy As Boolean`r`nPublic PptTestLastMessage As String")
            $source = ([regex]::Replace($source, '\bMsgBox\b', 'PptTestMessageBox')).Replace('    menu.ShowPopup', '    PptTestShowPopup menu')
        }
        if ($file.BaseName -eq 'modVersionMigration') {
            $source = $source.Replace('    SetMigratedPptCellIcon targetSheet, targetRow', "    If PptTestFailCopy Then Err.Raise vbObjectError + 7953, , ""Injected copy failure after binding""`r`n    SetMigratedPptCellIcon targetSheet, targetRow")
            if ($Baseline) {
                $source = $source.Replace("CopyFailed:`r`n", "CopyFailed:`r`n    Err.Raise vbObjectError + 7951, , Err.Description`r`n")
            }
            $regression = Read-Vba (Join-Path $PSScriptRoot 'GanttArtifactsRegression.bas')
            $regression = [regex]::Replace($regression, '(?m)^(\s*)(stepName|stage) = (.+)$', '$1$2 = $3' + "`r`n" + '$1PptTestTrace $2')
            $source += "`r`n" + $regression
        }
        $code = $book.VBProject.VBComponents.Item($file.BaseName).CodeModule
        if ($code.CountOfLines) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    $source = Read-Vba (Join-Path $root 'vba-files/Class/현재_통합_문서.cls')
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $source += "`r`nPublic Sub PptTestRightClick(ByVal ws As Worksheet, ByVal target As Range, ByRef cancel As Boolean)`r`n    Workbook_SheetBeforeRightClick ws, target, cancel`r`nEnd Sub`r`n"
    $code = $book.VBProject.VBComponents.Item($book.CodeName).CodeModule
    if ($code.CountOfLines) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($source)
    if ($InspectCompile) {
        $excel.Visible = $true
        $excel.VBE.MainWindow.Visible = $true
    }
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) { throw 'VBA compile did not complete.' }
    Write-Output 'PASS: VBA project compilation.'
    Write-Output 'VBA loaded; running attachment regression.'
    if ($ActualRefresh) {
        $result = [string]$excel.Run("'$($book.Name)'!RunActualRowRefreshRegression")
    } elseif ($SourcePath) {
        $excel.AutomationSecurity = 3
        $sourceBook = $excel.Workbooks.Open($SourcePath, 0, $true)
        $excel.AutomationSecurity = 1
        $result = [string]$excel.Run("'$($book.Name)'!RunActualArtifactRegression", $sourceBook.Name, [bool]$Baseline)
    } else {
        $result = [string]$excel.Run("'$($book.Name)'!RunGanttArtifactsRegression")
    }
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
    if (-not $ActualRefresh -and -not $SourcePath) {
        Write-Output 'Running PPT deletion and replacement regression.'
        $result = [string]$excel.Run("'$($book.Name)'!RunGanttPptDeleteRegression")
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw $result }
    }
    $book.Save()
    $book.Close($false)
    $book = $excel.Workbooks.Open($testPath, 0, $false)
    $result = [string]$excel.Run("'$($book.Name)'!RunSavedArtifactRegression")
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw $result }
} finally {
    $tracePath = $testPath + '.trace'
    if (Test-Path -LiteralPath $tracePath) {
        Get-Content -LiteralPath $tracePath -Encoding UTF8 | Write-Output
        Remove-Item -LiteralPath $tracePath -Force
    }
    if ($sourceBook) { try { $sourceBook.Close($false) } catch { Write-Warning $_ } }
    if ($book) { try { $book.Close($false) } catch { Write-Warning $_ } }
    if ($excel) {
        try { $excel.Quit() } catch { Write-Warning $_ }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if (Test-Path -LiteralPath $testPath) { try { Remove-Item -LiteralPath $testPath -Force } catch { Write-Warning $_ } }
}
