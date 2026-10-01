[CmdletBinding()]
param(
    [string]$WorkbookPath,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$reportRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $reportConfig = Get-Content -LiteralPath (Join-Path $reportRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $reportRoot $reportConfig.excel_file
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $reportRoot 'artifacts/text-report-verification'
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$reportTemp = Join-Path ([IO.Path]::GetTempPath()) ('text-report-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($reportTemp)
$reportCopy = Join-Path $reportTemp 'text-report-test.xlsm'
Copy-Item -LiteralPath $WorkbookPath -Destination $reportCopy
$reportExcel = $null
$reportBook = $null
$reportPhase = 'Excel startup'

function Read-ReportVbaSource([string]$SourcePath) {
    $sourceBytes = [IO.File]::ReadAllBytes($SourcePath)
    try { $sourceText = [Text.UTF8Encoding]::new($false, $true).GetString($sourceBytes) }
    catch { $sourceText = [Text.Encoding]::GetEncoding(949).GetString($sourceBytes) }
    $sourceText = $sourceText.TrimStart([char]0xFEFF)
    $sourceText = $sourceText -replace '(?m)^Attribute .*\r?\n', ''
    $sourceText = $sourceText -replace '(?s)^VERSION 1\.0 CLASS\r?\nBEGIN\r?\n.*?\r?\nEND\r?\n', ''
    return $sourceText
}

try {
    $reportExcel = New-Object -ComObject Excel.Application
    $reportExcel.Visible = $false
    $reportExcel.DisplayAlerts = $false
    $reportExcel.EnableEvents = $false
    $reportExcel.AskToUpdateLinks = $false
    $reportExcel.AutomationSecurity = 1
    $reportPhase = 'open temporary workbook'
    $reportBook = $reportExcel.Workbooks.Open($reportCopy, 0, $false)
    foreach ($obsoleteName in @('modWeeklyPptReport', 'modDailyProgressReport')) {
        $obsoleteComponent = $null
        try { $obsoleteComponent = $reportBook.VBProject.VBComponents.Item($obsoleteName) } catch {}
        if ($null -ne $obsoleteComponent) {
            $reportBook.VBProject.VBComponents.Remove($obsoleteComponent)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obsoleteComponent)
        }
    }
    foreach ($reportFile in Get-ChildItem -LiteralPath (Join-Path $reportRoot 'vba-files/Module') -Filter '*.bas') {
        if ($reportFile.BaseName -in @('modWeeklyPptReport', 'modDailyProgressReport')) { continue }
        $reportPhase = 'import ' + $reportFile.BaseName
        $reportSource = Read-ReportVbaSource $reportFile.FullName
        if ($reportFile.BaseName -eq 'modTextReport') {
            # Exercise the production private builders in the temporary project only.
            $reportSource += "`r`n" + (Read-ReportVbaSource (Join-Path $PSScriptRoot 'TextReportRegression.bas'))
        }
        $reportComponent = $null
        try { $reportComponent = $reportBook.VBProject.VBComponents.Item($reportFile.BaseName) } catch {}
        if ($null -eq $reportComponent) {
            $reportComponent = $reportBook.VBProject.VBComponents.Add(1)
            $reportComponent.Name = $reportFile.BaseName
        }
        $reportCode = $reportComponent.CodeModule
        if ($reportCode.CountOfLines -gt 0) { $reportCode.DeleteLines(1, $reportCode.CountOfLines) }
        $reportCode.AddFromString($reportSource)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportCode)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportComponent)
    }
    $reportPhase = 'import workbook event class'
    $reportSource = Read-ReportVbaSource (Join-Path $reportRoot 'vba-files/Class/현재_통합_문서.cls')
    $reportComponent = $reportBook.VBProject.VBComponents.Item($reportBook.CodeName)
    $reportCode = $reportComponent.CodeModule
    if ($reportCode.CountOfLines -gt 0) { $reportCode.DeleteLines(1, $reportCode.CountOfLines) }
    $reportCode.AddFromString($reportSource)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportCode)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportComponent)
    $reportPhase = 'run VBA regression'
    $reportResult = [string]$reportExcel.Run("'text-report-test.xlsm'!RunTextReportRegression", $OutputDirectory)
    if (-not $reportResult.StartsWith('PASS:')) { throw $reportResult }

    $reportPhase = 'verify Korean UTF-8 files'
    $reportFiles = @(Get-ChildItem -LiteralPath $OutputDirectory -Filter '*.txt' -Recurse | Where-Object { $_.Name -ne 'result.txt' })
    if ($reportFiles.Count -lt 7) { throw 'Too few TXT samples produced.' }
    foreach ($reportFile in $reportFiles) {
        $reportBytes = [IO.File]::ReadAllBytes($reportFile.FullName)
        if ($reportBytes.Length -lt 3 -or $reportBytes[0] -ne 0xEF -or $reportBytes[1] -ne 0xBB -or $reportBytes[2] -ne 0xBF) {
            throw ('Missing UTF-8 BOM: ' + $reportFile.FullName)
        }
        $reportText = [Text.UTF8Encoding]::new($false, $true).GetString($reportBytes)
        if (-not $reportText.Contains('업무 현황') -or $reportText.Contains([char]0xFFFD)) {
            throw ('UTF-8/Korean verification failed: ' + $reportFile.FullName)
        }
    }
    $reportResult += '; ' + $reportFiles.Count + ' valid UTF-8 TXT files'
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'result.txt'), $reportResult, [Text.UTF8Encoding]::new($false))
    Write-Output $reportResult
}
catch {
    Write-Error ('Regression failed during ' + $reportPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
    throw
}
finally {
    if ($null -ne $reportBook) {
        try { $reportBook.Close($false) } catch { Write-Warning ('Test workbook cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportBook)
    }
    if ($null -ne $reportExcel) {
        try { $reportExcel.Quit() } catch { Write-Warning ('Test Excel cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $reportCopy) { Remove-Item -LiteralPath $reportCopy -Force }
    if (Test-Path -LiteralPath $reportTemp) { Remove-Item -LiteralPath $reportTemp }
}
