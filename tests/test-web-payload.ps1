[CmdletBinding()]
param([string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
$webRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $webConfig = Get-Content -LiteralPath (Join-Path $webRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $webRoot $webConfig.excel_file
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$webTestDirectory = Join-Path ([IO.Path]::GetTempPath()) ('web-payload-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($webTestDirectory)
$webTestCopy = Join-Path $webTestDirectory 'web-payload-test.xlsm'
Copy-Item -LiteralPath $WorkbookPath -Destination $webTestCopy
$webExcel = $null
$webBook = $null
$webPhase = 'Excel startup'

function Read-WebPayloadSource([string]$SourcePath) {
    $webBytes = [IO.File]::ReadAllBytes($SourcePath)
    try { $webSource = [Text.UTF8Encoding]::new($false, $true).GetString($webBytes) }
    catch { $webSource = [Text.Encoding]::GetEncoding(949).GetString($webBytes) }
    return ($webSource.TrimStart([char]0xFEFF) -replace '(?m)^Attribute .*\r?\n', '')
}

try {
    $webExcel = New-Object -ComObject Excel.Application
    $webExcel.Visible = $false
    $webExcel.DisplayAlerts = $false
    $webExcel.EnableEvents = $false
    $webExcel.AskToUpdateLinks = $false
    $webExcel.AutomationSecurity = 1
    $webPhase = 'open disposable workbook'
    $webBook = $webExcel.Workbooks.Open($webTestCopy, 0, $false)
    foreach ($webFile in Get-ChildItem -LiteralPath (Join-Path $webRoot 'vba-files/Module') -Filter '*.bas') {
        $webPhase = 'import ' + $webFile.BaseName
        $webSource = Read-WebPayloadSource $webFile.FullName
        if ($webFile.BaseName -eq 'modWebPayload') {
            $webSource += "`r`n" + (Read-WebPayloadSource (Join-Path $PSScriptRoot 'WebPayloadRegression.bas'))
        }
        $webComponent = $null
        try { $webComponent = $webBook.VBProject.VBComponents.Item($webFile.BaseName) } catch {}
        if ($null -eq $webComponent) {
            $webComponent = $webBook.VBProject.VBComponents.Add(1)
            $webComponent.Name = $webFile.BaseName
        }
        $webCode = $webComponent.CodeModule
        if ($webCode.CountOfLines -gt 0) { $webCode.DeleteLines(1, $webCode.CountOfLines) }
        $webCode.AddFromString($webSource)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webCode)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webComponent)
    }
    $webPhase = 'run VBA regression'
    $webResult = [string]$webExcel.Run("'web-payload-test.xlsm'!RunWebPayloadRegression")
    if (-not $webResult.StartsWith('PASS:')) { throw $webResult }
    Write-Output $webResult
}
catch {
    Write-Error ('Web payload regression failed during ' + $webPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
    throw
}
finally {
    if ($null -ne $webBook) {
        try { $webBook.Close($false) } catch { Write-Warning ('Test workbook cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webBook)
    }
    if ($null -ne $webExcel) {
        try { $webExcel.Quit() } catch { Write-Warning ('Test Excel cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # Both paths were created above under the resolved temporary test directory.
    if (Test-Path -LiteralPath $webTestCopy) { Remove-Item -LiteralPath $webTestCopy -Force }
    if (Test-Path -LiteralPath $webTestDirectory) { Remove-Item -LiteralPath $webTestDirectory }
}
