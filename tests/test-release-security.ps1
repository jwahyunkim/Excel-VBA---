[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$securityScript = Join-Path $repositoryRoot 'release-security.ps1'
$parseTokens = $null
$parseErrors = $null
$syntax = [Management.Automation.Language.Parser]::ParseFile($securityScript, [ref]$parseTokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors.Message -join '; ') }

# Load function definitions only: never execute the action dispatcher or read local credentials.
$definitions = $syntax.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
foreach ($definition in $definitions) {
    . ([ScriptBlock]::Create($definition.Extent.Text))
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$fixtureRoot = Join-Path $tempRoot ('release-security-test-' + [guid]::NewGuid().ToString('N'))
$null = [IO.Directory]::CreateDirectory($fixtureRoot)
$ScriptRoot = Join-Path $fixtureRoot 'workspace'
$null = [IO.Directory]::CreateDirectory($ScriptRoot)
$ConfigPath = Join-Path $ScriptRoot 'config.json'
$SecurityModulePath = Join-Path $fixtureRoot 'security.bas'
$WorkbookClassPath = Join-Path $fixtureRoot 'workbook.cls'
$Date = [DateTime]::Today.ToString('yyyy-MM-dd')
$OutputPath = ''
$ResultPath = Join-Path $fixtureRoot 'result.txt'
$ReleaseUser = ''
$UsageDays = ''
$RenewalDays = ''
$CodeModulus = 1679616
$script:testConfig = [PSCustomObject]@{
    WorkbookPath = Join-Path $ScriptRoot 'source v3.2.1.xlsm'
    UsageDays = 30
    RenewalDays = 15
    ProjectPassword = 'MOCK-PASSWORD-DO-NOT-PRINT'
    RenewalSecret = 'MOCK-SECRET-DO-NOT-PRINT'
    DistributionFolder = Join-Path $ScriptRoot 'releases'
}
$script:failUpdate = $false
$script:updateCalls = 0
$script:readCalls = 0
$script:saveCalls = 0
$script:expectedExistingContent = $null
$script:expectedValidationContent = ''
$script:checks = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
    $script:checks++
}

function Assert-Throws {
    param([ScriptBlock]$Action, [string]$Message)
    $failed = $false
    try { & $Action 6>$null } catch { $failed = $true }
    Assert-True $failed $Message
}

function Assert-NoStagedFiles {
    $staged = @(Get-ChildItem -LiteralPath $fixtureRoot -File -Recurse -Force | Where-Object { $_.Name -match '^\.release-build-' })
    Assert-True ($staged.Count -eq 0) 'Temporary build files are removed.'
}

function Get-LocalConfig { return $script:testConfig }

function Invoke-WorkbookUpdate {
    param($WorkbookPath, [switch]$ConfigureRelease, $ReleaseDate, $ExpiryDate, $ReleaseVersion, $RenewalDays, $RenewalSecret, $ReleaseUser, $AdminPassword)
    $script:updateCalls++
    Assert-True ($WorkbookPath -ne $script:testConfig.WorkbookPath -and $WorkbookPath -ne $script:finalPath) 'Excel only receives a staged copy.'
    Assert-True ([IO.Path]::GetFileName($WorkbookPath) -match '^\.release-build-[0-9a-f]{32}\.xlsm$') 'Stage has a unique workbook filename.'
    Assert-True ([IO.File]::ReadAllText($WorkbookPath) -eq 'development-source') 'Stage begins as a source copy.'
    Assert-True ($ConfigureRelease -and $ReleaseVersion -eq '3.2.1') 'Release configuration reaches the update.'
    if ($null -ne $script:expectedExistingContent) {
        Assert-True ([IO.File]::ReadAllText($script:finalPath) -eq $script:expectedExistingContent) 'Existing output survives until configuration finishes.'
    }
    [IO.File]::WriteAllText($WorkbookPath, 'configured-release')
    if ($script:failUpdate) { throw 'Simulated Excel configuration failure.' }
}

function Read-ReleaseWorkbookState {
    param($WorkbookPath, [bool]$MacrosEnabled)
    $script:readCalls++
    Assert-True ([IO.File]::ReadAllText($WorkbookPath) -eq $script:expectedValidationContent) 'Validation selects the intended artifact.'
    Assert-True ($WorkbookPath -ne $script:finalPath) 'Validation uses a copy.'
    $guideName = [string][char]0xC0AC + [char]0xC6A9 + [char]0xC548 + [char]0xB0B4
    return [PSCustomObject]@{
        VisibleSheets = @(if ($MacrosEnabled) { 'Business' } else { $guideName })
        InfoVisibility = 2
        InfoProtected = $true
        Marker = 'RELEASE_SECURITY_V1'
        ReleaseUser = 'Test User'
        RenewalDays = 15
        ExpiryDate = [DateTime]::Today.AddDays(30)
        Saved = $true
    }
}

function Test-ReleaseSaveCycle {
    param($WorkbookPath)
    $script:saveCalls++
    Assert-True ([IO.File]::ReadAllText($WorkbookPath) -eq $script:expectedValidationContent) 'Save validation uses the intended copy.'
}

try {
    foreach ($fixturePath in @($testConfig.WorkbookPath, $SecurityModulePath, $WorkbookClassPath)) {
        [IO.File]::WriteAllText($fixturePath, 'development-source')
    }
    $null = [IO.Directory]::CreateDirectory($testConfig.DistributionFolder)
    $script:finalPath = Get-ReleaseOutputPath -Config $testConfig -ReleaseDate (Resolve-TargetDate)
    [IO.File]::WriteAllText($finalPath, 'previous-release')
    $script:expectedExistingContent = 'previous-release'

    $buildLog = Build-ReleaseWorkbook 6>&1 | Out-String
    Assert-True ([IO.File]::ReadAllText($finalPath) -eq 'configured-release') 'Successful build replaces the old output.'
    Assert-True ([IO.File]::ReadAllText($ResultPath, [Text.Encoding]::UTF8) -eq $finalPath) 'Build returns the exact absolute artifact path as UTF-8.'
    Assert-True ([IO.Path]::IsPathRooted([IO.File]::ReadAllText($ResultPath))) 'Result path is absolute.'
    Assert-True (-not $buildLog.Contains($testConfig.ProjectPassword) -and -not $buildLog.Contains($testConfig.RenewalSecret)) 'Build output contains no credentials.'
    $renewalCode = New-RenewalCode -TargetDate (Resolve-TargetDate) -Secret $testConfig.RenewalSecret
    Assert-True (-not $buildLog.Contains($renewalCode)) 'Build output contains no renewal code.'
    Assert-NoStagedFiles

    $script:failUpdate = $true
    $script:expectedExistingContent = 'configured-release'
    Assert-Throws { Build-ReleaseWorkbook } 'Excel failure is reported.'
    Assert-True ([IO.File]::ReadAllText($finalPath) -eq 'configured-release') 'Failed build preserves the previous release.'
    Assert-True ([IO.File]::ReadAllText($ResultPath) -eq '') 'Failed build clears an earlier result.'
    Assert-NoStagedFiles

    $script:failUpdate = $false
    $lockedOutput = [IO.File]::Open($finalPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        Assert-Throws { Build-ReleaseWorkbook } 'A locked output prevents publication.'
    }
    finally { $lockedOutput.Dispose() }
    Assert-True ([IO.File]::ReadAllText($finalPath) -eq 'configured-release') 'Publication failure preserves the previous output.'
    Assert-True ([IO.File]::ReadAllText($ResultPath) -eq '') 'Publication failure returns no successful result.'
    Assert-NoStagedFiles

    $script:failUpdate = $true
    $OutputPath = Join-Path $ScriptRoot 'custom\new-release.xlsm'
    $script:finalPath = $OutputPath
    $script:expectedExistingContent = $null
    Assert-Throws { Build-ReleaseWorkbook } 'Failure before first publication is reported.'
    Assert-True (-not (Test-Path -LiteralPath $OutputPath)) 'Failed first build creates no final output.'
    Assert-NoStagedFiles
    $script:failUpdate = $false
    $ResultPath = ''
    Build-ReleaseWorkbook 6>$null
    Assert-True ([IO.File]::ReadAllText($OutputPath) -eq 'configured-release') 'Custom output directory is created and published.'
    Assert-NoStagedFiles

    $callsBefore = $script:updateCalls
    $OutputPath = $testConfig.WorkbookPath.ToUpperInvariant()
    Assert-Throws { Build-ReleaseWorkbook } 'Source/output collision is rejected case-insensitively.'
    Assert-True ($script:updateCalls -eq $callsBefore) 'Source collision never reaches Excel.'
    Assert-True ([IO.File]::ReadAllText($testConfig.WorkbookPath) -eq 'development-source') 'Source remains unchanged.'

    $OutputPath = ''
    $script:finalPath = Get-ReleaseOutputPath -Config $testConfig -ReleaseDate (Resolve-TargetDate)
    foreach ($collisionPath in @($testConfig.WorkbookPath, $finalPath, $ConfigPath)) {
        $ResultPath = $collisionPath
        Assert-Throws { Build-ReleaseWorkbook } 'Result/source, output, or config collision is rejected.'
    }
    $ResultPath = Join-Path $fixtureRoot 'result.txt'
    $otherVersion = Join-Path $testConfig.DistributionFolder 'source v9.9.9_newest.xlsm'
    [IO.File]::WriteAllText($otherVersion, 'wrong-version')
    [IO.File]::SetLastWriteTime($otherVersion, [DateTime]::Now.AddDays(1))
    $script:expectedValidationContent = 'configured-release'
    Test-ReleaseWorkbook 6>$null
    Assert-True ($script:readCalls -eq 3 -and $script:saveCalls -eq 1) 'All validation stages run against the current standard artifact.'

    $Date = [DateTime]::Today.AddDays(-1).ToString('yyyy-MM-dd')
    $readsBefore = $script:readCalls
    Assert-Throws { Test-ReleaseWorkbook } 'Missing standard artifact does not fall back to another version or date.'
    Assert-True ($script:readCalls -eq $readsBefore) 'Missing artifact fails before Excel opens.'
    $OutputPath = Join-Path $ScriptRoot 'custom\new-release.xlsm'
    Test-ReleaseWorkbook 6>$null
    Assert-True ($script:readCalls -eq ($readsBefore + 3)) 'Explicit output overrides the default date selection.'
    Write-Output ("PASS: release-security regression ({0} assertions; Excel mocked)." -f $script:checks)
}
finally {
    $resolvedFixtureRoot = [IO.Path]::GetFullPath($fixtureRoot)
    if ($resolvedFixtureRoot.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedFixtureRoot) -match '^release-security-test-[0-9a-f]{32}$' -and
        (Test-Path -LiteralPath $resolvedFixtureRoot)) {
        Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force
    }
}
