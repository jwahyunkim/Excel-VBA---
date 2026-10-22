[CmdletBinding()]
param([string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
$webRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $webConfig = Get-Content -LiteralPath (Join-Path $webRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $webRoot $webConfig.excel_file
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$webTestDirectory = Join-Path ([IO.Path]::GetTempPath()) ('web-upload-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($webTestDirectory)
$webTestCopy = Join-Path $webTestDirectory 'web-upload-test.xlsm'
Copy-Item -LiteralPath $WorkbookPath -Destination $webTestCopy
$webExcel = $null
$webBook = $null
$webPhase = 'Excel startup'

function Read-WebUploadSource([string]$SourcePath) {
    $webBytes = [IO.File]::ReadAllBytes($SourcePath)
    try { $webSource = [Text.UTF8Encoding]::new($false, $true).GetString($webBytes) }
    catch { $webSource = [Text.Encoding]::GetEncoding(949).GetString($webBytes) }
    $webSource = $webSource.TrimStart([char]0xFEFF) -replace '(?m)^Attribute .*\r?\n', ''
    return ($webSource -replace '(?s)^VERSION 1\.0 CLASS\r?\nBEGIN\r?\n.*?\r?\nEND\r?\n', '')
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
        $webSource = Read-WebUploadSource $webFile.FullName
        if ($webFile.BaseName -eq 'modWebUpload') {
            $webSource += "`r`n" + (Read-WebUploadSource (Join-Path $PSScriptRoot 'WebUploadRegression.bas'))
        }
        if ($webFile.BaseName -eq 'modWebTransport') {
            # Replace only the public transport boundary in this temporary project.
            # The production transport source remains present under Original names.
            foreach ($webBoundary in @('WebOpenSession', 'WebLoadRegistration', 'WebPostRegistration', 'WebLoadExisting', 'WebReadExistingValues', 'WebPostUpdate')) {
                $webSource = $webSource -replace ('\b' + $webBoundary + '\b'), ($webBoundary + 'Original')
            }
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
    $webPhase = 'import offline transport stub'
    $webComponent = $webBook.VBProject.VBComponents.Add(1)
    $webComponent.Name = 'WebUploadTransportStub'
    $webComponent.CodeModule.AddFromString((Read-WebUploadSource (Join-Path $PSScriptRoot 'WebUploadTransportStub.bas')))
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webComponent)
    $webPhase = 'import production workbook events'
    $webSource = Read-WebUploadSource (Join-Path $webRoot 'vba-files/Class/현재_통합_문서.cls')
    $webComponent = $webBook.VBProject.VBComponents.Item($webBook.CodeName)
    $webCode = $webComponent.CodeModule
    if ($webCode.CountOfLines -gt 0) { $webCode.DeleteLines(1, $webCode.CountOfLines) }
    $webCode.AddFromString($webSource)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webCode)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webComponent)
    $webPhase = 'run VBA regression'
    $webResult = [string]$webExcel.Run("'web-upload-test.xlsm'!RunWebUploadRegression")
    if (-not $webResult.StartsWith('PASS:')) { throw $webResult }
    Write-Output $webResult
    # 2026-10-22 김좌현 수정 : 이력 삭제로 해제한 상태를 임시 파일에 저장하고 재열기 후 검증
    $webPhase = 'save history-reconciled disposable workbook'
    $webExcel.EnableEvents = $false
    $webBook.Save()
    $webBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($webBook)
    $webBook = $null
    $webPhase = 'reopen and verify deleted history stays released'
    $webBook = $webExcel.Workbooks.Open($webTestCopy, 0, $false)
    $webResult = [string]$webExcel.Run("'web-upload-test.xlsm'!RunSavedWebHistoryRegression")
    if (-not $webResult.StartsWith('PASS:')) { throw $webResult }
    Write-Output $webResult
}
catch {
    Write-Error ('Web upload integration regression failed during ' + $webPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
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
