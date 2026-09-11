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
    $OutputDirectory = Join-Path $reportRoot 'artifacts/weekly-report-verification'
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$reportTemp = Join-Path ([IO.Path]::GetTempPath()) ('weekly-report-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($reportTemp)
$reportCopy = Join-Path $reportTemp 'weekly-report-test.xlsm'
Copy-Item -LiteralPath $WorkbookPath -Destination $reportCopy
$reportExcel = $null
$reportBook = $null
$reportPhase = 'Excel startup'
try {
    $reportExcel = New-Object -ComObject Excel.Application
    $reportExcel.Visible = $false
    $reportExcel.DisplayAlerts = $false
    $reportExcel.EnableEvents = $false
    $reportExcel.AskToUpdateLinks = $false
    $reportExcel.AutomationSecurity = 1
    $reportPhase = 'open temporary workbook'
    $reportBook = $reportExcel.Workbooks.Open($reportCopy, 0, $false)
    foreach ($reportModule in @('modGanttConfig', 'modHoliday', 'modWeeklyPptReport')) {
        $reportPhase = 'import ' + $reportModule
        $reportSource = [IO.File]::ReadAllText((Join-Path $reportRoot ('vba-files/Module/' + $reportModule + '.bas')))
        $reportSource = $reportSource -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        if ($reportModule -eq 'modWeeklyPptReport') {
            # Append to the production module to exercise its private grouping/rendering functions.
            $reportSource += "`r`n" + [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'WeeklyReportRegression.bas'))
        }
        $reportCode = $reportBook.VBProject.VBComponents.Item($reportModule).CodeModule
        if ($reportCode.CountOfLines -gt 0) { $reportCode.DeleteLines(1, $reportCode.CountOfLines) }
        $reportCode.AddFromString($reportSource)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reportCode)
    }
    $reportOutput = Join-Path $OutputDirectory 'weekly-report-regression.pptx'
    $reportPhase = 'run VBA regression'
    $reportResult = [string]$reportExcel.Run("'weekly-report-test.xlsm'!RunWeeklyReportRegression", $reportOutput)
    if (-not $reportResult.StartsWith('PASS:')) { throw $reportResult }
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'result.txt'), $reportResult, [Text.UTF8Encoding]::new($false))
    Write-Output $reportResult
    Write-Output ('PPT: ' + $reportOutput)
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
    # Only remove the exact temporary workbook created by this run.
    if (Test-Path -LiteralPath $reportCopy) { Remove-Item -LiteralPath $reportCopy -Force }
    if (Test-Path -LiteralPath $reportTemp) { Remove-Item -LiteralPath $reportTemp }
}
