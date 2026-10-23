[CmdletBinding()]
param([string]$WorkbookPath)

# 2026-10-23 김좌현 수정 : 실제 파일 선택 UI만 우회한 임시 통합문서 데이터·웹 이력 이관 및 저장 재열기 검증
$ErrorActionPreference = 'Stop'
$migrationRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $migrationConfig = Get-Content -LiteralPath (Join-Path $migrationRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $migrationRoot $migrationConfig.excel_file
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$migrationTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$migrationDirectory = Join-Path $migrationTempRoot ('version-migration-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($migrationDirectory)
$migrationTestCopy = Join-Path $migrationDirectory 'version-migration-test.xlsm'
$migrationSourceCopy = Join-Path $migrationDirectory 'migration-source.xlsx'
Copy-Item -LiteralPath $WorkbookPath -Destination $migrationTestCopy
$migrationExcel = $null
$migrationBook = $null
$migrationPhase = 'Excel startup'

function Read-MigrationSource([string]$SourcePath) {
    $migrationBytes = [IO.File]::ReadAllBytes($SourcePath)
    try { $migrationSource = [Text.UTF8Encoding]::new($false, $true).GetString($migrationBytes) }
    catch { $migrationSource = [Text.Encoding]::GetEncoding(949).GetString($migrationBytes) }
    $migrationSource = $migrationSource.TrimStart([char]0xFEFF) -replace '(?m)^Attribute .*\r?\n', ''
    return ($migrationSource -replace '(?s)^VERSION 1\.0 CLASS\r?\nBEGIN\r?\n.*?\r?\nEND\r?\n', '')
}

try {
    $migrationExcel = New-Object -ComObject Excel.Application
    $migrationExcel.Visible = $false
    $migrationExcel.DisplayAlerts = $false
    $migrationExcel.EnableEvents = $false
    $migrationExcel.AskToUpdateLinks = $false
    $migrationExcel.AutomationSecurity = 1
    $migrationPhase = 'open disposable workbook'
    $migrationBook = $migrationExcel.Workbooks.Open($migrationTestCopy, 0, $false)
    foreach ($migrationFile in Get-ChildItem -LiteralPath (Join-Path $migrationRoot 'vba-files/Module') -Filter '*.bas') {
        $migrationPhase = 'import ' + $migrationFile.BaseName
        $migrationSource = Read-MigrationSource $migrationFile.FullName
        if ($migrationFile.BaseName -eq 'modVersionMigration') {
            $migrationHarness = Read-MigrationSource (Join-Path $PSScriptRoot 'VersionMigrationRegression.bas')
            $migrationDeclarations = [regex]::Matches($migrationHarness, '(?m)^Private Const MIG_.*\r?$')
            $migrationDeclarationText = ($migrationDeclarations | ForEach-Object { $_.Value.TrimEnd("`r") }) -join "`r`n"
            $migrationHarness = [regex]::Replace($migrationHarness, '(?m)^Private Const MIG_.*\r?\n', '')
            $migrationSource = $migrationSource.Replace('Option Explicit', "Option Explicit`r`n" + $migrationDeclarationText)
            $migrationSource += "`r`n" + $migrationHarness
        }
        if ($migrationFile.BaseName -eq 'modWebTransport') {
            # Every public network boundary is replaced only in this disposable project.
            foreach ($migrationBoundary in @('WebOpenSession', 'WebLoadRegistration', 'WebPostRegistration', 'WebLoadExisting', 'WebReadExistingValues', 'WebPostUpdate')) {
                $migrationSource = $migrationSource -replace ('\b' + $migrationBoundary + '\b'), ($migrationBoundary + 'Original')
            }
        }
        $migrationComponent = $null
        try { $migrationComponent = $migrationBook.VBProject.VBComponents.Item($migrationFile.BaseName) } catch {}
        if ($null -eq $migrationComponent) {
            $migrationComponent = $migrationBook.VBProject.VBComponents.Add(1)
            $migrationComponent.Name = $migrationFile.BaseName
        }
        $migrationCode = $migrationComponent.CodeModule
        if ($migrationCode.CountOfLines -gt 0) { $migrationCode.DeleteLines(1, $migrationCode.CountOfLines) }
        $migrationCode.AddFromString($migrationSource)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationCode)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationComponent)
    }
    $migrationPhase = 'import offline transport stub'
    $migrationComponent = $migrationBook.VBProject.VBComponents.Add(1)
    $migrationComponent.Name = 'WebUploadTransportStub'
    $migrationComponent.CodeModule.AddFromString((Read-MigrationSource (Join-Path $PSScriptRoot 'WebUploadTransportStub.bas')))
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationComponent)
    $migrationPhase = 'import production workbook events'
    $migrationSource = Read-MigrationSource (Join-Path $migrationRoot 'vba-files/Class/현재_통합_문서.cls')
    $migrationComponent = $migrationBook.VBProject.VBComponents.Item($migrationBook.CodeName)
    $migrationCode = $migrationComponent.CodeModule
    if ($migrationCode.CountOfLines -gt 0) { $migrationCode.DeleteLines(1, $migrationCode.CountOfLines) }
    $migrationCode.AddFromString($migrationSource)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationCode)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationComponent)
    $migrationPhase = 'run current and legacy migration regression'
    $migrationResult = [string]$migrationExcel.Run("'version-migration-test.xlsm'!RunVersionMigrationRegression")
    if (-not $migrationResult.StartsWith('PASS:')) { throw $migrationResult }
    Write-Output $migrationResult
    $migrationPhase = 'save and reopen migrated disposable workbook'
    $migrationExcel.EnableEvents = $false
    $migrationBook.Save()
    $migrationBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationBook)
    $migrationBook = $null
    $migrationBook = $migrationExcel.Workbooks.Open($migrationTestCopy, 0, $false)
    $migrationResult = [string]$migrationExcel.Run("'version-migration-test.xlsm'!RunSavedVersionMigrationRegression")
    if (-not $migrationResult.StartsWith('PASS:')) { throw $migrationResult }
    Write-Output $migrationResult
}
catch {
    Write-Error ('Version migration regression failed during ' + $migrationPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
    throw
}
finally {
    if ($null -ne $migrationBook) {
        try { $migrationBook.Close($false) } catch { Write-Warning ('Migration test workbook cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationBook)
    }
    if ($null -ne $migrationExcel) {
        try {
            while ($migrationExcel.Workbooks.Count -gt 0) { $migrationExcel.Workbooks.Item(1).Close($false) }
            $migrationExcel.Quit()
        } catch { Write-Warning ('Migration test Excel cleanup: ' + $_.Exception.Message) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($migrationExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # Delete only the two exact files created above inside this unique temporary directory.
    foreach ($migrationCleanup in @($migrationTestCopy, $migrationSourceCopy)) {
        if ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($migrationCleanup)) -ne [IO.Path]::GetFullPath($migrationDirectory)) {
            throw 'Migration cleanup path escaped the test directory.'
        }
        if (Test-Path -LiteralPath $migrationCleanup) { Remove-Item -LiteralPath $migrationCleanup -Force }
    }
    if (Test-Path -LiteralPath $migrationDirectory) { Remove-Item -LiteralPath $migrationDirectory }
}
