[CmdletBinding()]
param(
    [ValidateSet(
        "Menu", "Start", "StartVersion", "CreateBranch", "MergeBranch", "Release", "CancelVersion", "Status",
        "Security", "SecurityConfig", "SecurityStatus", "SecuritySync", "SecurityBuild",
        "SecurityValidate", "SecurityAll", "SecurityCode")]
    [string]$Action = "Menu",

    [string]$Value = "",
    [string]$BranchName = "",
    [string]$BaseBranch = "",
    [string]$SourceBranch = "",
    [string]$TargetBranch = "",
    [string]$MergeTitle = "",
    [string]$MergeBody = ""
)

# 사용 예시:
#   패치 버전 개발 시작:  .\version-workflow.ps1 -Action StartVersion patch
#   마이너 버전 시작:     .\version-workflow.ps1 -Action StartVersion minor
#   현재 브랜치의 자식:   .\version-workflow.ps1 -Action CreateBranch -BranchName feature/report
#   지정 부모의 자식:     .\version-workflow.ps1 -Action CreateBranch -BaseBranch develop/v3.2.13 -BranchName feature/report
#   기록된 부모로 병합:   .\version-workflow.ps1 -Action MergeBranch
#   지정 부모로 병합:     .\version-workflow.ps1 -Action MergeBranch -SourceBranch feature/report -TargetBranch develop/v3.2.13
#   현재 개발 완료/릴리즈: .\version-workflow.ps1 -Action Release
#   현재 버전 개발 취소:  .\version-workflow.ps1 -Action CancelVersion

# ============================================================================
# 사용자 설정
# ============================================================================
$ConfigFile = "config.json"
$ConfigField = "excel_file"
$ReleaseSecurityScript = "release-security.ps1"
$DefaultVersionBaseBranch = "main"

# 자식 브랜치 기본값입니다. 비워두면 실행할 때 입력할 수 있습니다.
$DefaultChildBaseBranch = ""
$DefaultChildBranch = ""
$DefaultChildMergeTitle = ""
$DefaultChildMergeBody = ""
# ============================================================================

$ErrorActionPreference = "Stop"
$ScriptRelativePath = Split-Path -Leaf $PSCommandPath
$WorkflowSupportFiles = @(
    $ScriptRelativePath,
    $ReleaseSecurityScript,
    "package.json",
    "WORKFLOW_GUIDE.md",
    "RELEASE_SECURITY_GUIDE.md"
)

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Host "> $Command $($Arguments -join ' ')" -ForegroundColor DarkGray
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "명령 실행 실패(exit $LASTEXITCODE): $Command $($Arguments -join ' ')"
    }
}

function Invoke-Git {
    param([string[]]$Arguments)
    Invoke-CheckedCommand -Command "git" -Arguments $Arguments
}

function Get-GitOutput {
    param([string[]]$Arguments)

    $result = & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Git 명령 실행 실패: git $($Arguments -join ' ')"
    }

    return ($result | Out-String).Trim()
}

function Get-GhCommand {
    $ghCommand = Get-Command "gh" -ErrorAction SilentlyContinue
    if ($null -ne $ghCommand) {
        return $ghCommand.Source
    }

    $fallback = Join-Path $env:ProgramFiles "GitHub CLI\gh.exe"
    if (Test-Path -LiteralPath $fallback) {
        return $fallback
    }

    throw "GitHub CLI(gh)를 찾을 수 없습니다. GitHub CLI를 설치하거나 PATH를 확인하세요."
}

function Assert-GitRepository {
    & git rev-parse --is-inside-work-tree *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "현재 위치가 Git 저장소가 아닙니다."
    }
}

function Get-CurrentBranch {
    return Get-GitOutput -Arguments @("branch", "--show-current")
}

function Get-LatestSemanticVersion {
    $tags = & git tag --list "v*" --sort=-version:refname
    if ($LASTEXITCODE -ne 0) {
        throw "버전 태그를 조회하지 못했습니다."
    }

    foreach ($tag in $tags) {
        if ($tag -match '^v(\d+)\.(\d+)\.(\d+)$') {
            return [PSCustomObject]@{
                Tag = $tag
                Version = "$($Matches[1]).$($Matches[2]).$($Matches[3])"
                Major = [int64]$Matches[1]
                Minor = [int64]$Matches[2]
                Patch = [int64]$Matches[3]
            }
        }
    }

    throw "vA.B.C 형식의 기존 버전 태그를 찾을 수 없습니다."
}

function Get-NextSemanticVersion {
    param(
        [Parameter(Mandatory = $true)]$CurrentVersion,
        [Parameter(Mandatory = $true)][string]$Part
    )

    switch ($Part.ToLowerInvariant()) {
        "patch" { return "$($CurrentVersion.Major).$($CurrentVersion.Minor).$($CurrentVersion.Patch + 1)" }
        "minor" { return "$($CurrentVersion.Major).$($CurrentVersion.Minor + 1).0" }
        "major" { return "$($CurrentVersion.Major + 1).0.0" }
        default { throw "버전 증가 단위는 patch, minor, major 중 하나여야 합니다: $Part" }
    }
}

function Resolve-VersionPart {
    param(
        [string]$RequestedPart,
        [Parameter(Mandatory = $true)]$CurrentVersion
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPart)) {
        $normalizedPart = $RequestedPart.Trim().ToLowerInvariant()
        if ($normalizedPart -notin @("patch", "minor", "major")) {
            throw "버전 증가 단위는 patch, minor, major 중 하나여야 합니다: $RequestedPart"
        }
        return $normalizedPart
    }

    $nextPatch = Get-NextSemanticVersion -CurrentVersion $CurrentVersion -Part "patch"
    $nextMinor = Get-NextSemanticVersion -CurrentVersion $CurrentVersion -Part "minor"
    $nextMajor = Get-NextSemanticVersion -CurrentVersion $CurrentVersion -Part "major"

    Write-Host ""
    Write-Host "현재 버전: $($CurrentVersion.Tag)"
    Write-Host "[1] patch (c 증가) : v$nextPatch"
    Write-Host "[2] minor (b 증가) : v$nextMinor  - c는 0으로 초기화"
    Write-Host "[3] major (a 증가) : v$nextMajor  - b와 c는 0으로 초기화"
    $selection = (Read-Host "올릴 버전 단위 선택").Trim().ToLowerInvariant()

    switch ($selection) {
        { $_ -in @("1", "patch", "c") } { return "patch" }
        { $_ -in @("2", "minor", "b") } { return "minor" }
        { $_ -in @("3", "major", "a") } { return "major" }
        default { throw "1(patch), 2(minor), 3(major) 중 하나를 선택하세요." }
    }
}

function Get-VersionFromDevelopmentBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    if ($Name -match '^(?:develop|version)/v(\d+\.\d+\.\d+)$') {
        return $Matches[1]
    }

    throw "릴리즈할 브랜치는 'develop/vA.B.C' 형식이어야 합니다: $Name"
}

function Get-ConfiguredParentBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    $parent = & git config --local --get "branch.$Name.workflowParent" 2> $null
    if ($LASTEXITCODE -ne 0) {
        return ""
    }

    return ($parent | Out-String).Trim()
}

function Set-ConfiguredParentBranch {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Parent
    )

    Invoke-Git -Arguments @("config", "--local", "branch.$Name.workflowParent", $Parent)
}

function Resolve-InputValue {
    param(
        [string]$Value,
        [string]$DefaultValue,
        [Parameter(Mandatory = $true)][string]$Prompt
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        return $Value.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) {
        return $DefaultValue.Trim()
    }

    $answer = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($answer)) {
        throw "$Prompt 값은 비워둘 수 없습니다."
    }

    return $answer.Trim()
}

function Assert-ValidBranchName {
    param([Parameter(Mandatory = $true)][string]$Name)

    & git check-ref-format --branch $Name *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "올바르지 않은 브랜치 이름입니다: $Name"
    }
}

function Test-LocalBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    & git show-ref --verify --quiet "refs/heads/$Name"
    return ($LASTEXITCODE -eq 0)
}

function Test-RemoteTrackingBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    & git show-ref --verify --quiet "refs/remotes/origin/$Name"
    return ($LASTEXITCODE -eq 0)
}

function Switch-ToUpdatedBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    Invoke-Git -Arguments @("fetch", "origin", "--prune")

    if (Test-LocalBranch -Name $Name) {
        if ((Get-CurrentBranch) -ne $Name) {
            Invoke-Git -Arguments @("switch", $Name)
        }
    }
    elseif (Test-RemoteTrackingBranch -Name $Name) {
        Invoke-Git -Arguments @("switch", "--track", "-c", $Name, "origin/$Name")
    }
    else {
        throw "부모 브랜치를 찾을 수 없습니다: $Name"
    }

    if (Test-RemoteTrackingBranch -Name $Name) {
        Invoke-Git -Arguments @("pull", "--ff-only", "origin", $Name)
    }
}

function Get-WorkingTreeStatus {
    $result = & git status --porcelain=v1
    if ($LASTEXITCODE -ne 0) {
        throw "Git 작업 트리 상태를 조회하지 못했습니다."
    }
    return ($result -join "`n")
}

function Assert-CleanWorkingTree {
    $status = Get-WorkingTreeStatus
    if (-not [string]::IsNullOrWhiteSpace($status)) {
        Write-Host $status
        throw "커밋되지 않은 변경 사항이 있습니다. 먼저 개발 내용을 커밋하거나 정리하세요."
    }
}

function Assert-StartWorkingTreeIsSafe {
    $status = Get-WorkingTreeStatus
    if ([string]::IsNullOrWhiteSpace($status)) {
        return
    }

    $unsafeLines = @(
        $status -split "`r?`n" |
            Where-Object {
                $linePath = if ($_.Length -gt 3) { $_.Substring(3).Trim('"') } else { "" }
                $linePath -notin $WorkflowSupportFiles
            }
    )

    if ($unsafeLines.Count -gt 0) {
        Write-Host $status
        throw "워크플로 스크립트 외에 변경 사항이 있습니다. 작업 트리를 깨끗하게 만든 후 다시 실행하세요."
    }

    Write-Host "알림: 워크플로 파일 변경 사항은 새 버전 브랜치로 함께 이동합니다." -ForegroundColor Yellow
}

function Get-ConfigFieldValue {
    $configPath = Join-Path (Get-Location) $ConfigFile
    if (-not (Test-Path -LiteralPath $configPath)) {
        throw "설정 파일을 찾을 수 없습니다: $ConfigFile"
    }

    $text = [System.IO.File]::ReadAllText($configPath)
    try {
        $json = $text | ConvertFrom-Json
    }
    catch {
        throw "$ConfigFile JSON 파싱 실패: $($_.Exception.Message)"
    }

    $property = $json.PSObject.Properties[$ConfigField]
    if ($null -eq $property) {
        throw "$ConfigFile 파일에 '$ConfigField' 필드가 없습니다."
    }

    return [string]$property.Value
}

function Set-ConfigFieldValue {
    param([Parameter(Mandatory = $true)][string]$Value)

    $configPath = Join-Path (Get-Location) $ConfigFile
    $text = [System.IO.File]::ReadAllText($configPath)
    $null = $text | ConvertFrom-Json

    $escapedField = [System.Text.RegularExpressions.Regex]::Escape($ConfigField)
    $pattern = '(?m)(^\s*"' + $escapedField + '"\s*:\s*)"(?:\\.|[^"\\])*"'
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern)
    $matches = $regex.Matches($text)
    if ($matches.Count -ne 1) {
        throw "$ConfigFile 파일에서 '$ConfigField' 문자열 필드를 정확히 하나 찾지 못했습니다."
    }

    $encodedValue = ConvertTo-Json -InputObject $Value -Compress
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
        param($match)
        return $match.Groups[1].Value + $encodedValue
    }
    $updatedText = $regex.Replace($text, $evaluator, 1)
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($configPath, $updatedText, $utf8WithoutBom)

    $actualValue = Get-ConfigFieldValue
    if ($actualValue -ne $Value) {
        throw "$ConfigFile 업데이트 검증에 실패했습니다."
    }
}

function Get-DevelopmentVersionForBranch {
    param([Parameter(Mandatory = $true)][string]$Name)

    $candidate = $Name
    $visited = @{}
    for ($depth = 0; $depth -lt 10; $depth++) {
        if ($candidate -match '^(?:develop|version)/v(\d+\.\d+\.\d+)$') {
            return $Matches[1]
        }
        if ($visited.ContainsKey($candidate)) {
            throw "브랜치 부모 설정이 순환합니다: $candidate"
        }
        $visited[$candidate] = $true

        $parent = Get-ConfiguredParentBranch -Name $candidate
        if ([string]::IsNullOrWhiteSpace($parent)) { return "" }
        $candidate = $parent
    }

    throw "브랜치 부모 설정이 너무 깊습니다: $Name"
}

function Set-DevelopmentWorkbookVersion {
    param(
        [Parameter(Mandatory = $true)][string]$TargetVersion,
        [string]$ExpectedCurrentVersion = ""
    )

    $oldExcelFile = Get-ConfigFieldValue
    $versionPattern = 'v\d+\.\d+\.\d+(?=\.xlsm$)'
    $versionMatches = [regex]::Matches(
        $oldExcelFile,
        $versionPattern,
        [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($versionMatches.Count -ne 1) {
        throw "$ConfigFile 의 '$ConfigField' 파일명에서 vA.B.C 버전을 정확히 하나 찾지 못했습니다: $oldExcelFile"
    }

    $currentVersion = $versionMatches[0].Value.TrimStart('v', 'V')
    if ($currentVersion -eq $TargetVersion) {
        if (-not (Test-Path -LiteralPath $oldExcelFile)) {
            throw "config의 개발 원본을 찾을 수 없습니다: $oldExcelFile"
        }
        return
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedCurrentVersion) -and
        $currentVersion -ne $ExpectedCurrentVersion) {
        throw "개발 원본 버전(v$currentVersion)이 예상 버전(v$ExpectedCurrentVersion)과 다릅니다."
    }
    if (-not (Test-Path -LiteralPath $oldExcelFile)) {
        throw "버전을 변경할 개발 원본을 찾을 수 없습니다: $oldExcelFile"
    }

    $newExcelFile = [regex]::Replace(
        $oldExcelFile,
        $versionPattern,
        "v$TargetVersion",
        [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (Test-Path -LiteralPath $newExcelFile) {
        throw "변경할 버전의 개발 원본이 이미 존재합니다: $newExcelFile"
    }

    Move-Item -LiteralPath $oldExcelFile -Destination $newExcelFile
    Set-ConfigFieldValue -Value $newExcelFile
    Write-Host "개발 원본/config 버전 정렬: v$currentVersion -> v$TargetVersion" -ForegroundColor Green
}

function Assert-MainMatchesRemote {
    param([Parameter(Mandatory = $true)][string]$BranchName)

    if (-not (Test-RemoteTrackingBranch -Name $BranchName)) { return }

    $counts = Get-GitOutput -Arguments @(
        "rev-list", "--left-right", "--count", "$BranchName...origin/$BranchName")
    $parts = @($counts -split '\s+' | Where-Object { $_ -ne "" })
    if ($parts.Count -ne 2) {
        throw "main/origin 동기화 상태를 해석하지 못했습니다: $counts"
    }

    $localOnly = [int]$parts[0]
    $remoteOnly = [int]$parts[1]
    if ($localOnly -ne 0 -or $remoteOnly -ne 0) {
        throw "$BranchName 브랜치가 origin과 다릅니다(local +$localOnly, remote +$remoteOnly). 먼저 동기화하세요."
    }
}

function Start-AutomaticPatchDevelopment {
    param([Parameter(Mandatory = $true)][string]$Reason)

    $currentBranch = Get-CurrentBranch
    if ($currentBranch -ne $DefaultVersionBaseBranch) {
        throw "자동 patch 개발 시작은 $DefaultVersionBaseBranch 브랜치에서만 가능합니다: $currentBranch"
    }

    $unmerged = Get-GitOutput -Arguments @("diff", "--name-only", "--diff-filter=U")
    if (-not [string]::IsNullOrWhiteSpace($unmerged)) {
        Write-Host $unmerged
        throw "충돌이 해결되지 않은 파일이 있어 개발 브랜치를 만들 수 없습니다."
    }

    Invoke-Git -Arguments @("fetch", "origin", "--tags", "--prune")
    Assert-MainMatchesRemote -BranchName $DefaultVersionBaseBranch

    $currentVersion = Get-LatestSemanticVersion
    $nextVersion = Get-NextSemanticVersion -CurrentVersion $currentVersion -Part "patch"
    $developmentBranch = "develop/v$nextVersion"

    if ((Test-LocalBranch -Name $developmentBranch) -or
        (Test-RemoteTrackingBranch -Name $developmentBranch)) {
        throw "다음 patch 개발 브랜치가 이미 존재합니다: $developmentBranch`n해당 브랜치로 전환한 뒤 다시 실행하세요."
    }

    $status = Get-WorkingTreeStatus
    if (-not [string]::IsNullOrWhiteSpace($status)) {
        Write-Host "현재 변경 사항을 새 개발 브랜치로 함께 이동합니다:" -ForegroundColor Yellow
        Write-Host $status
    }

    Invoke-Git -Arguments @("switch", "-c", $developmentBranch)
    Set-ConfiguredParentBranch -Name $developmentBranch -Parent $DefaultVersionBaseBranch
    Set-DevelopmentWorkbookVersion `
        -TargetVersion $nextVersion `
        -ExpectedCurrentVersion $currentVersion.Version
    Invoke-Git -Arguments @("push", "-u", "origin", $developmentBranch)

    Write-Host ""
    Write-Host "$Reason 작업을 위해 $developmentBranch patch 개발을 시작했습니다." -ForegroundColor Green
    Write-Host "변경 사항은 스테이징하지 않았습니다." -ForegroundColor Yellow
    return $nextVersion
}

function Ensure-SecurityMutationDevelopmentContext {
    param([Parameter(Mandatory = $true)][string]$Reason)

    $currentBranch = Get-CurrentBranch
    $developmentVersion = Get-DevelopmentVersionForBranch -Name $currentBranch
    if (-not [string]::IsNullOrWhiteSpace($developmentVersion)) {
        Set-DevelopmentWorkbookVersion -TargetVersion $developmentVersion
        Write-Host "$Reason 적용 대상: $currentBranch (v$developmentVersion)" -ForegroundColor DarkGray
        return $developmentVersion
    }

    if ($currentBranch -eq $DefaultVersionBaseBranch) {
        return Start-AutomaticPatchDevelopment -Reason $Reason
    }
    if ($currentBranch -match '^release/v') {
        throw "릴리즈 브랜치에서는 새 보안 변경을 시작할 수 없습니다: $currentBranch"
    }

    throw ("현재 브랜치의 개발 버전을 확인할 수 없습니다: $currentBranch`n" +
           "main에서 다시 실행하거나 workflowParent가 설정된 개발/기능 브랜치를 사용하세요.")
}

function Show-WorkflowStatus {
    $latestVersion = Get-LatestSemanticVersion
    $nextPatch = Get-NextSemanticVersion -CurrentVersion $latestVersion -Part "patch"
    $currentBranch = Get-CurrentBranch

    Write-Host ""
    Write-Host "최신 릴리즈     : $($latestVersion.Tag)"
    Write-Host "다음 patch      : v$nextPatch"
    Write-Host "Config          : $ConfigFile -> $ConfigField"
    Write-Host "현재 Excel      : $(Get-ConfigFieldValue)"
    Write-Host "현재 브랜치     : $currentBranch"
    $developmentVersion = Get-DevelopmentVersionForBranch -Name $currentBranch
    if (-not [string]::IsNullOrWhiteSpace($developmentVersion)) {
        Write-Host "작업 버전       : v$developmentVersion"
        Write-Host "변경 작업       : 현재 개발 버전에 적용"
    }
    elseif ($currentBranch -eq $DefaultVersionBaseBranch) {
        Write-Host "변경 작업       : develop/v$nextPatch 자동 시작"
    }
    elseif ($currentBranch -match '^release/v') {
        Write-Host "변경 작업       : 릴리즈 브랜치에서는 시작 불가"
    }
    else {
        Write-Host "변경 작업       : 개발 부모 브랜치 설정 필요"
    }
    Write-Host ""
    Invoke-Git -Arguments @("status", "--short", "--branch")
}

function Invoke-ReleaseSecurityCommand {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Build", "Code", "Status", "SyncDev", "Validate")]
        [string]$ReleaseAction,
        [string]$TargetDate = ""
    )

    $arguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", $ReleaseSecurityScript,
        "-Action", $ReleaseAction
    )
    if (-not [string]::IsNullOrWhiteSpace($TargetDate)) {
        $arguments += @("-Date", $TargetDate.Trim())
    }

    & powershell @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "배포 보안 작업에 실패했습니다: $ReleaseAction"
    }
}

function Read-PositiveIntegerSetting {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][int]$CurrentValue
    )

    $answer = (Read-Host "$Prompt [$CurrentValue]").Trim()
    if ([string]::IsNullOrWhiteSpace($answer)) { return $CurrentValue }

    $parsed = 0
    if (-not [int]::TryParse($answer, [ref]$parsed) -or $parsed -lt 1) {
        throw "$Prompt 값은 1 이상의 정수여야 합니다."
    }
    return $parsed
}

function Edit-ReleaseSecurityConfig {
    $configPath = Join-Path (Get-Location) $ConfigFile
    $config = [IO.File]::ReadAllText($configPath) | ConvertFrom-Json
    if ($null -eq $config.release_security) {
        throw "$ConfigFile 파일에 release_security 설정이 없습니다."
    }

    Write-Host ""
    Write-Host "빈 값으로 입력하면 현재 설정을 유지합니다." -ForegroundColor DarkGray

    $requestedUsageDays = Read-PositiveIntegerSetting `
        -Prompt "기본 사용 기간(일)" `
        -CurrentValue ([int]$config.release_security.usage_days)
    $requestedRenewalDays = Read-PositiveIntegerSetting `
        -Prompt "1회 연장 기간(일)" `
        -CurrentValue ([int]$config.release_security.renewal_days)

    $projectPassword = (Read-Host "VBA 프로젝트 암호 [Enter=현재 암호 유지]").Trim()
    $requestedProjectPassword = if ([string]::IsNullOrWhiteSpace($projectPassword)) {
        [string]$config.release_security.vba_project_password
    }
    else {
        $projectPassword
    }

    $renewalSecret = (Read-Host "연장코드 비밀키 [Enter=현재 비밀키 유지]").Trim()
    if (-not [string]::IsNullOrWhiteSpace($renewalSecret)) {
        if ($renewalSecret -match '[^\x20-\x7E]') {
            throw "연장코드 비밀키는 영문, 숫자, ASCII 특수문자만 사용할 수 있습니다."
        }
    }
    $requestedRenewalSecret = if ([string]::IsNullOrWhiteSpace($renewalSecret)) {
        [string]$config.release_security.renewal_secret
    }
    else {
        $renewalSecret
    }

    $distributionFolder = (Read-Host "배포 폴더 [$($config.release_security.distribution_folder)]").Trim()
    $requestedDistributionFolder = if ([string]::IsNullOrWhiteSpace($distributionFolder)) {
        [string]$config.release_security.distribution_folder
    }
    else {
        $distributionFolder
    }

    $hasChanges = (
        $requestedUsageDays -ne [int]$config.release_security.usage_days -or
        $requestedRenewalDays -ne [int]$config.release_security.renewal_days -or
        $requestedProjectPassword -ne [string]$config.release_security.vba_project_password -or
        $requestedRenewalSecret -ne [string]$config.release_security.renewal_secret -or
        $requestedDistributionFolder -ne [string]$config.release_security.distribution_folder)
    if (-not $hasChanges) {
        Write-Host "변경된 보안 설정이 없습니다. 버전과 파일을 변경하지 않습니다." -ForegroundColor Yellow
        return
    }

    $null = Ensure-SecurityMutationDevelopmentContext -Reason "보안 설정 수정"
    $config = [IO.File]::ReadAllText($configPath) | ConvertFrom-Json
    $config.release_security.usage_days = $requestedUsageDays
    $config.release_security.renewal_days = $requestedRenewalDays
    $config.release_security.vba_project_password = $requestedProjectPassword
    $config.release_security.renewal_secret = $requestedRenewalSecret
    $config.release_security.distribution_folder = $requestedDistributionFolder

    $json = $config | ConvertTo-Json -Depth 10
    $utf8WithoutBom = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($configPath, $json + [Environment]::NewLine, $utf8WithoutBom)

    Write-Host "config.json 보안 설정을 저장했습니다." -ForegroundColor Green
    Invoke-ReleaseSecurityCommand -ReleaseAction Status
}

function Invoke-SecuritySyncWorkflow {
    $null = Ensure-SecurityMutationDevelopmentContext -Reason "개발본 보안 VBA 동기화"
    Invoke-ReleaseSecurityCommand -ReleaseAction SyncDev
}

function Invoke-SecurityBuildWorkflow {
    param([string]$TargetDate = "")

    $null = Ensure-SecurityMutationDevelopmentContext -Reason "배포본 생성"
    Invoke-ReleaseSecurityCommand -ReleaseAction Build -TargetDate $TargetDate
}

function Invoke-SecurityAllWorkflow {
    param([string]$TargetDate = "")

    $null = Ensure-SecurityMutationDevelopmentContext -Reason "보안 동기화/배포"
    Invoke-ReleaseSecurityCommand -ReleaseAction SyncDev
    Invoke-ReleaseSecurityCommand -ReleaseAction Build -TargetDate $TargetDate
    Invoke-ReleaseSecurityCommand -ReleaseAction Validate
}

function Show-ReleaseSecurityMenu {
    while ($true) {
        Write-Host ""
        Write-Host "=== 배포 보안 관리 ===" -ForegroundColor Cyan
        Write-Host "[1] 보안 설정 확인"
        Write-Host "[2] 보안 설정 수정"
        Write-Host "[3] 개발본 보안 VBA 동기화"
        Write-Host "[4] 배포본 생성"
        Write-Host "[5] 배포본 자동 검증"
        Write-Host "[6] 동기화 + 생성 + 자동 검증"
        Write-Host "[7] 날짜별 기간 연장코드 생성"
        Write-Host "[0] 종료"
        $selection = (Read-Host "선택").Trim()

        switch ($selection) {
            "1" { Invoke-ReleaseSecurityCommand -ReleaseAction Status }
            "2" { Edit-ReleaseSecurityConfig }
            "3" { Invoke-SecuritySyncWorkflow }
            "4" {
                $targetDate = Read-Host "배포 기준일(yyyy-MM-dd, Enter=오늘)"
                Invoke-SecurityBuildWorkflow -TargetDate $targetDate
            }
            "5" { Invoke-ReleaseSecurityCommand -ReleaseAction Validate }
            "6" {
                $targetDate = Read-Host "배포 기준일(yyyy-MM-dd, Enter=오늘)"
                Invoke-SecurityAllWorkflow -TargetDate $targetDate
            }
            "7" {
                $targetDate = Read-Host "코드 날짜(yyyy-MM-dd, Enter=오늘)"
                Invoke-ReleaseSecurityCommand -ReleaseAction Code -TargetDate $targetDate
            }
            "0" { return }
            default { Write-Host "0~7 중 하나를 선택하세요." -ForegroundColor Yellow }
        }
    }
}

function Create-WorkBranch {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedBaseBranch,
        [Parameter(Mandatory = $true)][string]$RequestedBranchName,
        [switch]$AllowScriptOnlyChanges
    )

    Assert-ValidBranchName -Name $RequestedBaseBranch
    Assert-ValidBranchName -Name $RequestedBranchName

    if ($RequestedBaseBranch -eq $RequestedBranchName) {
        throw "부모 브랜치와 자식 브랜치 이름은 달라야 합니다."
    }

    if ($AllowScriptOnlyChanges) {
        Assert-StartWorkingTreeIsSafe
    }
    else {
        Assert-CleanWorkingTree
    }

    Switch-ToUpdatedBranch -Name $RequestedBaseBranch

    if (Test-LocalBranch -Name $RequestedBranchName) {
        throw "로컬 브랜치가 이미 존재합니다: $RequestedBranchName"
    }

    if (Test-RemoteTrackingBranch -Name $RequestedBranchName) {
        throw "원격 브랜치가 이미 존재합니다: $RequestedBranchName"
    }

    Invoke-Git -Arguments @("switch", "-c", $RequestedBranchName)
    Invoke-Git -Arguments @("push", "-u", "origin", $RequestedBranchName)
    Set-ConfiguredParentBranch -Name $RequestedBranchName -Parent $RequestedBaseBranch

    Write-Host ""
    Write-Host "$RequestedBaseBranch -> $RequestedBranchName 자식 브랜치를 만들었습니다." -ForegroundColor Green
}

function Ensure-ReleaseBranch {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedBaseBranch,
        [Parameter(Mandatory = $true)][string]$RequestedBranchName
    )

    if (-not (Test-LocalBranch -Name $RequestedBranchName) -and
        -not (Test-RemoteTrackingBranch -Name $RequestedBranchName)) {
        Create-WorkBranch `
            -RequestedBaseBranch $RequestedBaseBranch `
            -RequestedBranchName $RequestedBranchName
        return
    }

    Assert-CleanWorkingTree
    Switch-ToUpdatedBranch -Name $RequestedBranchName

    $configuredParent = Get-ConfiguredParentBranch -Name $RequestedBranchName
    if (-not [string]::IsNullOrWhiteSpace($configuredParent) -and
        $configuredParent -ne $RequestedBaseBranch) {
        throw "기존 릴리즈 브랜치의 부모가 다릅니다: $RequestedBranchName -> $configuredParent"
    }

    if ([string]::IsNullOrWhiteSpace($configuredParent)) {
        Set-ConfiguredParentBranch -Name $RequestedBranchName -Parent $RequestedBaseBranch
    }

    Write-Host "기존 릴리즈 브랜치를 사용해 중단 지점부터 계속합니다: $RequestedBranchName" -ForegroundColor Yellow
}

function Start-VersionDevelopment {
    Invoke-Git -Arguments @("fetch", "origin", "--tags", "--prune")

    $currentVersion = Get-LatestSemanticVersion
    $versionPart = Resolve-VersionPart -RequestedPart $Value -CurrentVersion $currentVersion
    $nextVersion = Get-NextSemanticVersion -CurrentVersion $currentVersion -Part $versionPart
    $developmentBranch = "develop/v$nextVersion"

    $resolvedBaseBranch = if ([string]::IsNullOrWhiteSpace($BaseBranch)) {
        $DefaultVersionBaseBranch
    }
    else {
        $BaseBranch.Trim()
    }

    Create-WorkBranch `
        -RequestedBaseBranch $resolvedBaseBranch `
        -RequestedBranchName $developmentBranch `
        -AllowScriptOnlyChanges

    Set-DevelopmentWorkbookVersion `
        -TargetVersion $nextVersion `
        -ExpectedCurrentVersion $currentVersion.Version

    Write-Host ""
    Write-Host "$($currentVersion.Tag) -> v$nextVersion ($versionPart)" -ForegroundColor Green
    Write-Host "개발 브랜치: $developmentBranch" -ForegroundColor Green
    Write-Host "이제 기능별 자식 브랜치를 만들거나 개발을 진행하세요."
}

function Start-ChildBranch {
    $currentBranch = Get-CurrentBranch
    $requestedBranchName = if (-not [string]::IsNullOrWhiteSpace($BranchName)) {
        $BranchName
    }
    else {
        $Value
    }

    $resolvedBaseBranch = Resolve-InputValue `
        -Value $BaseBranch `
        -DefaultValue $(if ([string]::IsNullOrWhiteSpace($DefaultChildBaseBranch)) { $currentBranch } else { $DefaultChildBaseBranch }) `
        -Prompt "부모 브랜치 이름"

    $resolvedBranchName = Resolve-InputValue `
        -Value $requestedBranchName `
        -DefaultValue $DefaultChildBranch `
        -Prompt "새 자식 브랜치 이름(예: feature/report-validation)"

    Create-WorkBranch `
        -RequestedBaseBranch $resolvedBaseBranch `
        -RequestedBranchName $resolvedBranchName
}

function Prepare-ReleaseCommit {
    param(
        [Parameter(Mandatory = $true)][string]$PreviousVersion,
        [Parameter(Mandatory = $true)][string]$ReleaseVersion
    )

    $oldExcelFile = Get-ConfigFieldValue
    $versionPattern = 'v\d+\.\d+\.\d+(?=\.xlsm$)'
    $versionMatches = [regex]::Matches($oldExcelFile, $versionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($versionMatches.Count -ne 1) {
        throw "$ConfigFile 의 '$ConfigField' 파일명에서 vA.B.C 버전을 정확히 하나 찾지 못했습니다: $oldExcelFile"
    }

    $fileVersion = $versionMatches[0].Value.TrimStart('v', 'V')
    if ($fileVersion -eq $ReleaseVersion) {
        if (-not (Test-Path -LiteralPath $oldExcelFile)) {
            throw "Config에는 새 버전이 반영됐지만 Excel 파일을 찾을 수 없습니다: $oldExcelFile"
        }

        Write-Host "버전 파일명과 로컬 config 설정이 이미 준비되어 있습니다." -ForegroundColor Yellow
        Invoke-ReleaseSecurityCommand -ReleaseAction Build
        Invoke-ReleaseSecurityCommand -ReleaseAction Validate
        Invoke-Git -Arguments @("add", "-A", "--", $ConfigFile, "workbooks/dev", "dist")
        Invoke-Git -Arguments @("diff", "--check", "--cached")
        & git diff --cached --quiet
        if ($LASTEXITCODE -eq 1) {
            Invoke-Git -Arguments @("commit", "-m", "v$ReleaseVersion 배포 준비")
        }
        elseif ($LASTEXITCODE -ne 0) {
            throw "배포 파일 변경 여부를 확인하지 못했습니다."
        }
        return
    }

    if ($fileVersion -ne $PreviousVersion) {
        throw "최신 태그(v$PreviousVersion)와 Excel 파일 버전(v$fileVersion)이 일치하지 않습니다."
    }

    $newExcelFile = [regex]::Replace(
        $oldExcelFile,
        $versionPattern,
        "v$ReleaseVersion",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $oldExists = Test-Path -LiteralPath $oldExcelFile
    $newExists = Test-Path -LiteralPath $newExcelFile

    if ($oldExists -and -not $newExists) {
        Move-Item -LiteralPath $oldExcelFile -Destination $newExcelFile
        Set-ConfigFieldValue -Value $newExcelFile
        Invoke-ReleaseSecurityCommand -ReleaseAction Build
        Invoke-ReleaseSecurityCommand -ReleaseAction Validate
        Invoke-Git -Arguments @("add", "-A", "--", $ConfigFile, "workbooks/dev", "dist")
        Invoke-Git -Arguments @("diff", "--check", "--cached")
        Invoke-Git -Arguments @("commit", "-m", "v$ReleaseVersion 배포 준비")
        Write-Host "개발 원본을 v$ReleaseVersion으로 변경하고 배포본을 함께 커밋했습니다." -ForegroundColor Green
        return
    }

    if (-not $oldExists -and $newExists) {
        $configValue = Get-ConfigFieldValue
        if ($configValue -ne $newExcelFile) {
            throw "신규 Excel 파일은 있지만 $ConfigFile 값이 일치하지 않습니다: $configValue"
        }

        Invoke-ReleaseSecurityCommand -ReleaseAction Build
        Invoke-ReleaseSecurityCommand -ReleaseAction Validate
        Invoke-Git -Arguments @("add", "-A", "--", $ConfigFile, "workbooks/dev", "dist")
        Invoke-Git -Arguments @("diff", "--check", "--cached")
        & git diff --cached --quiet
        if ($LASTEXITCODE -eq 1) {
            Invoke-Git -Arguments @("commit", "-m", "v$ReleaseVersion 배포 준비")
        }
        elseif ($LASTEXITCODE -ne 0) {
            throw "배포 파일 변경 여부를 확인하지 못했습니다."
        }
        Write-Host "버전 파일명과 config 설정이 이미 준비되어 배포본을 반영했습니다." -ForegroundColor Yellow
        return
    }

    if ($oldExists -and $newExists) {
        throw "기존 Excel 파일과 신규 Excel 파일이 모두 존재합니다. 어느 파일이 릴리즈 대상인지 확인하세요."
    }

    throw "기존 Excel 파일을 찾을 수 없습니다: $oldExcelFile"
}

function Ensure-PullRequest {
    param(
        [Parameter(Mandatory = $true)][string]$GhCommand,
        [Parameter(Mandatory = $true)][string]$HeadBranch,
        [Parameter(Mandatory = $true)][string]$BaseBranchName,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Body = ""
    )

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "SilentlyContinue"
        $existingUrl = & $GhCommand pr view $HeadBranch --json url --jq .url 2> $null
        $viewExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($viewExitCode -eq 0) {
        Write-Host "기존 PR을 사용합니다: $existingUrl" -ForegroundColor Yellow
        return
    }

    Invoke-CheckedCommand -Command $GhCommand -Arguments @(
        "pr", "create",
        "--base", $BaseBranchName,
        "--head", $HeadBranch,
        "--title", $Title,
        "--body", $Body
    )
}

function Merge-WorkBranch {
    param(
        [Parameter(Mandatory = $true)][string]$HeadBranch,
        [Parameter(Mandatory = $true)][string]$BaseBranchName,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Body = ""
    )

    Assert-ValidBranchName -Name $HeadBranch
    Assert-ValidBranchName -Name $BaseBranchName
    if ($HeadBranch -eq $BaseBranchName) {
        throw "병합할 자식 브랜치와 부모 브랜치는 달라야 합니다."
    }

    Assert-CleanWorkingTree
    Switch-ToUpdatedBranch -Name $HeadBranch
    Invoke-Git -Arguments @("push", "origin", $HeadBranch)

    $gh = Get-GhCommand
    Ensure-PullRequest `
        -GhCommand $gh `
        -HeadBranch $HeadBranch `
        -BaseBranchName $BaseBranchName `
        -Title $Title `
        -Body $Body

    Invoke-CheckedCommand -Command $gh -Arguments @(
        "pr", "merge", $HeadBranch,
        "--merge",
        "--subject", $Title,
        "--delete-branch"
    )

    Switch-ToUpdatedBranch -Name $BaseBranchName

    if (Test-LocalBranch -Name $HeadBranch) {
        Invoke-Git -Arguments @("branch", "-d", $HeadBranch)
    }

    Write-Host ""
    Write-Host "$HeadBranch -> $BaseBranchName 일반 Merge를 완료하고 자식 브랜치를 삭제했습니다." -ForegroundColor Green
}

function Merge-ChildBranch {
    $defaultSourceBranch = if ([string]::IsNullOrWhiteSpace($DefaultChildBranch)) {
        Get-CurrentBranch
    }
    else {
        $DefaultChildBranch
    }

    $resolvedSourceBranch = Resolve-InputValue `
        -Value $SourceBranch `
        -DefaultValue $defaultSourceBranch `
        -Prompt "병합할 자식 브랜치 이름"

    $configuredParentBranch = Get-ConfiguredParentBranch -Name $resolvedSourceBranch
    $defaultTargetBranch = if (-not [string]::IsNullOrWhiteSpace($DefaultChildBaseBranch)) {
        $DefaultChildBaseBranch
    }
    else {
        $configuredParentBranch
    }

    $resolvedTargetBranch = Resolve-InputValue `
        -Value $TargetBranch `
        -DefaultValue $defaultTargetBranch `
        -Prompt "병합 대상 부모 브랜치 이름"

    $defaultTitle = if ([string]::IsNullOrWhiteSpace($DefaultChildMergeTitle)) {
        "merge: $resolvedSourceBranch -> $resolvedTargetBranch"
    }
    else {
        $DefaultChildMergeTitle
    }

    $resolvedTitle = if ([string]::IsNullOrWhiteSpace($MergeTitle)) {
        $defaultTitle
    }
    else {
        $MergeTitle.Trim()
    }

    $resolvedBody = if ([string]::IsNullOrWhiteSpace($MergeBody)) {
        $DefaultChildMergeBody
    }
    else {
        $MergeBody
    }

    if ((Get-CurrentBranch) -ne $resolvedSourceBranch) {
        Assert-CleanWorkingTree
        Switch-ToUpdatedBranch -Name $resolvedSourceBranch
    }
    Commit-CurrentBranchChanges `
        -ExpectedBranch $resolvedSourceBranch `
        -CommitMessage "work: $resolvedSourceBranch 변경 사항 반영"
    Invoke-Git -Arguments @("push", "origin", $resolvedSourceBranch)

    Merge-WorkBranch `
        -HeadBranch $resolvedSourceBranch `
        -BaseBranchName $resolvedTargetBranch `
        -Title $resolvedTitle `
        -Body $resolvedBody
}

function Commit-CurrentBranchChanges {
    param(
        [Parameter(Mandatory = $true)][string]$ExpectedBranch,
        [Parameter(Mandatory = $true)][string]$CommitMessage
    )

    if ((Get-CurrentBranch) -ne $ExpectedBranch) {
        throw "자동 커밋은 대상 브랜치에서 실행해야 합니다: $ExpectedBranch"
    }

    $conflictedFiles = & git diff --name-only --diff-filter=U
    if ($LASTEXITCODE -ne 0) {
        throw "병합 충돌 여부를 확인하지 못했습니다."
    }
    if (-not [string]::IsNullOrWhiteSpace(($conflictedFiles | Out-String))) {
        throw "해결되지 않은 병합 충돌이 있습니다. 충돌을 먼저 해결하세요."
    }

    Invoke-Git -Arguments @("diff", "--check")
    Invoke-Git -Arguments @("add", "-A", "--", ".")
    Invoke-Git -Arguments @("diff", "--cached", "--check")

    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 1) {
        Invoke-Git -Arguments @("commit", "-m", $CommitMessage)
    }
    elseif ($LASTEXITCODE -ne 0) {
        throw "자동 커밋할 변경 여부를 확인하지 못했습니다."
    }
    else {
        Write-Host "자동 커밋할 변경 사항이 없습니다." -ForegroundColor Yellow
    }
}

function Get-ChildBranches {
    param([Parameter(Mandatory = $true)][string]$ParentBranch)

    $children = @()
    $localBranches = & git for-each-ref --format='%(refname:short)' refs/heads
    if ($LASTEXITCODE -ne 0) {
        throw "로컬 브랜치 목록을 조회하지 못했습니다."
    }

    foreach ($localBranch in $localBranches) {
        $branchName = ([string]$localBranch).Trim()
        if (-not [string]::IsNullOrWhiteSpace($branchName) -and
            (Get-ConfiguredParentBranch -Name $branchName) -eq $ParentBranch) {
            $children += $branchName
        }
    }

    return @($children)
}

function Assert-NoChildBranches {
    param([Parameter(Mandatory = $true)][string]$ParentBranch)

    $children = @(Get-ChildBranches -ParentBranch $ParentBranch)
    if ($children.Count -gt 0) {
        throw "먼저 자식 브랜치를 병합하거나 정리하세요: $($children -join ', ')"
    }
}

function Commit-DevelopmentChanges {
    param([Parameter(Mandatory = $true)][string]$DevelopmentBranch)

    $currentBranch = Get-CurrentBranch
    if ($currentBranch -ne $DevelopmentBranch) {
        throw "개발 변경 자동 커밋은 대상 개발 브랜치에서 실행해야 합니다: $DevelopmentBranch"
    }

    $releaseVersion = Get-VersionFromDevelopmentBranch -Name $DevelopmentBranch
    Commit-CurrentBranchChanges `
        -ExpectedBranch $DevelopmentBranch `
        -CommitMessage "develop: v$releaseVersion 변경 사항 반영"

    Invoke-Git -Arguments @("push", "origin", $DevelopmentBranch)
}

function Complete-VersionDevelopmentAndRelease {
    $developmentBranch = Get-CurrentBranch
    $null = Get-VersionFromDevelopmentBranch -Name $developmentBranch
    if (-not [string]::IsNullOrWhiteSpace($SourceBranch) -and
        $SourceBranch.Trim() -ne $developmentBranch) {
        throw "Release는 현재 개발 브랜치만 대상으로 실행할 수 있습니다: $developmentBranch"
    }
    Assert-NoChildBranches -ParentBranch $developmentBranch

    Write-Host "개발 변경 사항을 프로젝트 워크플로가 자동 스테이징·커밋합니다." -ForegroundColor Yellow
    Commit-DevelopmentChanges -DevelopmentBranch $developmentBranch
    Release-Version
}

function Cancel-VersionDevelopment {
    $developmentBranch = Get-CurrentBranch
    $releaseVersion = Get-VersionFromDevelopmentBranch -Name $developmentBranch
    $releaseBranch = "release/v$releaseVersion"
    $tagName = "v$releaseVersion"

    Invoke-Git -Arguments @("fetch", "origin", "--tags", "--prune")
    Assert-NoChildBranches -ParentBranch $developmentBranch

    if ((Test-LocalBranch -Name $releaseBranch) -or (Test-RemoteTrackingBranch -Name $releaseBranch)) {
        throw "릴리즈 브랜치가 이미 존재하므로 자동 취소할 수 없습니다: $releaseBranch"
    }

    & git show-ref --verify --quiet "refs/tags/$tagName"
    if ($LASTEXITCODE -eq 0) {
        throw "동일 버전 태그가 이미 존재하므로 자동 취소할 수 없습니다: $tagName"
    }

    $configuredParent = Get-ConfiguredParentBranch -Name $developmentBranch
    $parentBranch = if ([string]::IsNullOrWhiteSpace($configuredParent)) {
        $DefaultVersionBaseBranch
    }
    else {
        $configuredParent
    }
    Assert-ValidBranchName -Name $parentBranch

    $parentReference = if (Test-RemoteTrackingBranch -Name $parentBranch) {
        "origin/$parentBranch"
    }
    elseif (Test-LocalBranch -Name $parentBranch) {
        $parentBranch
    }
    else {
        throw "복귀할 부모 브랜치를 찾을 수 없습니다: $parentBranch"
    }

    Write-Host ""
    Write-Host "취소 대상: $developmentBranch" -ForegroundColor Yellow
    Write-Host "복귀 브랜치: $parentBranch"
    Write-Host "미커밋 변경은 stash에, 개발 커밋은 로컬 백업 태그에 보관합니다."

    $confirmation = if ([string]::IsNullOrWhiteSpace($Value)) {
        (Read-Host "취소하려면 브랜치 이름 '$developmentBranch' 을(를) 다시 입력").Trim()
    }
    else {
        $Value.Trim()
    }
    if ($confirmation -ne $developmentBranch) {
        throw "브랜치 이름이 일치하지 않아 개발 취소를 중단했습니다."
    }

    $stashReference = ""
    if (-not [string]::IsNullOrWhiteSpace((Get-WorkingTreeStatus))) {
        $stashMessage = "workflow cancel backup: $developmentBranch $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Invoke-Git -Arguments @("stash", "push", "--include-untracked", "-m", $stashMessage)
        $stashReference = Get-GitOutput -Arguments @("stash", "list", "-1", "--format=%gd")
    }

    $uniqueCommitCount = [int](Get-GitOutput -Arguments @("rev-list", "--count", "$parentReference..$developmentBranch"))
    $backupTag = ""
    if ($uniqueCommitCount -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupTag = "cancel-backup/v$releaseVersion-$timestamp"
        $suffix = 1
        while ($true) {
            & git show-ref --verify --quiet "refs/tags/$backupTag"
            if ($LASTEXITCODE -ne 0) {
                break
            }
            $backupTag = "cancel-backup/v$releaseVersion-$timestamp-$suffix"
            $suffix++
        }
        Invoke-Git -Arguments @("tag", "-a", $backupTag, $developmentBranch, "-m", "$developmentBranch 취소 백업")
    }

    Switch-ToUpdatedBranch -Name $parentBranch

    if (Test-RemoteTrackingBranch -Name $developmentBranch) {
        Invoke-Git -Arguments @("push", "origin", "--delete", $developmentBranch)
    }
    if (Test-LocalBranch -Name $developmentBranch) {
        Invoke-Git -Arguments @("branch", "-D", $developmentBranch)
    }

    Write-Host ""
    Write-Host "$developmentBranch 개발을 취소하고 $parentBranch 로 돌아왔습니다." -ForegroundColor Green
    if (-not [string]::IsNullOrWhiteSpace($backupTag)) {
        Write-Host "커밋 복구용 로컬 태그: $backupTag" -ForegroundColor Yellow
    }
    if (-not [string]::IsNullOrWhiteSpace($stashReference)) {
        Write-Host "미커밋 변경 복구용 stash: $stashReference (git stash apply $stashReference)" -ForegroundColor Yellow
    }
}

function Release-Version {
    $currentBranch = Get-CurrentBranch
    $developmentBranch = if ([string]::IsNullOrWhiteSpace($SourceBranch)) {
        $currentBranch
    }
    else {
        $SourceBranch.Trim()
    }

    if ($currentBranch -ne $developmentBranch) {
        throw "Release는 개발 브랜치에서 실행해야 합니다. 지정 브랜치: $developmentBranch, 현재 브랜치: $currentBranch"
    }

    $releaseVersion = Get-VersionFromDevelopmentBranch -Name $developmentBranch
    $tagName = "v$releaseVersion"
    $releaseBranch = "release/$tagName"
    $releaseTitle = "release: $tagName"

    $configuredParentBranch = Get-ConfiguredParentBranch -Name $developmentBranch
    $releaseBaseBranch = if (-not [string]::IsNullOrWhiteSpace($TargetBranch)) {
        $TargetBranch.Trim()
    }
    elseif (-not [string]::IsNullOrWhiteSpace($configuredParentBranch)) {
        $configuredParentBranch
    }
    else {
        $DefaultVersionBaseBranch
    }

    Assert-ValidBranchName -Name $developmentBranch
    Assert-ValidBranchName -Name $releaseBaseBranch
    Assert-CleanWorkingTree

    Invoke-Git -Arguments @("fetch", "origin", "--tags", "--prune")
    $previousVersion = Get-LatestSemanticVersion
    $validNextVersions = @(
        Get-NextSemanticVersion -CurrentVersion $previousVersion -Part "patch"
        Get-NextSemanticVersion -CurrentVersion $previousVersion -Part "minor"
        Get-NextSemanticVersion -CurrentVersion $previousVersion -Part "major"
    )
    if ($releaseVersion -notin $validNextVersions) {
        throw "릴리즈 버전 v$releaseVersion 은 최신 태그 $($previousVersion.Tag)의 올바른 다음 major/minor/patch 버전이 아닙니다."
    }

    & git show-ref --verify --quiet "refs/tags/$tagName"
    if ($LASTEXITCODE -eq 0) {
        throw "태그가 이미 존재합니다: $tagName"
    }

    Ensure-ReleaseBranch `
        -RequestedBaseBranch $releaseBaseBranch `
        -RequestedBranchName $releaseBranch

    Merge-WorkBranch `
        -HeadBranch $developmentBranch `
        -BaseBranchName $releaseBranch `
        -Title "develop: $tagName" `
        -Body "$tagName 개발 내용을 릴리즈 준비 브랜치로 병합"

    Prepare-ReleaseCommit `
        -PreviousVersion $previousVersion.Version `
        -ReleaseVersion $releaseVersion

    Merge-WorkBranch `
        -HeadBranch $releaseBranch `
        -BaseBranchName $releaseBaseBranch `
        -Title $releaseTitle `
        -Body "$tagName 릴리즈"

    Invoke-Git -Arguments @("tag", "-a", $tagName, "-m", "$tagName release")
    Invoke-Git -Arguments @("push", "origin", $tagName)

    Write-Host ""
    Write-Host "$tagName 릴리즈가 완료되었습니다." -ForegroundColor Green
    Write-Host "$developmentBranch -> $releaseBranch -> $releaseBaseBranch 이력이 보존되었습니다." -ForegroundColor Green
    Invoke-Git -Arguments @("log", "--first-parent", "--oneline", "--decorate", "-5", $releaseBaseBranch)
}

function Select-MenuAction {
    Write-Host ""
    Write-Host "[1] 새 버전 개발 시작"
    Write-Host "[2] 개발 완료 + 자동 커밋 + 릴리즈"
    Write-Host "[3] 현재 버전 개발 취소"
    Write-Host "[4] 설정 및 Git 상태 확인"
    Write-Host "[5] 현재/지정 브랜치에서 자식 브랜치 생성"
    Write-Host "[6] 자식 브랜치를 부모 브랜치로 병합"
    Write-Host "[7] 배포 보안 관리"
    Write-Host "[0] 종료"
    $selection = Read-Host "선택"

    switch ($selection) {
        "1" { return "Start" }
        "2" { return "Release" }
        "3" { return "CancelVersion" }
        "4" { return "Status" }
        "5" { return "CreateBranch" }
        "6" { return "MergeBranch" }
        "7" { return "Security" }
        "0" { return "Exit" }
        default { throw "0~7 중 올바른 메뉴 번호를 선택하세요." }
    }
}

Push-Location $PSScriptRoot
try {
    Assert-GitRepository

    if ($Action -eq "Menu") {
        $Action = Select-MenuAction
    }

    switch ($Action) {
        { $_ -in @("Start", "StartVersion") } { Start-VersionDevelopment }
        "CreateBranch" { Start-ChildBranch }
        "MergeBranch" { Merge-ChildBranch }
        "Release" { Complete-VersionDevelopmentAndRelease }
        "CancelVersion" { Cancel-VersionDevelopment }
        "Status" { Show-WorkflowStatus }
        "Security" { Show-ReleaseSecurityMenu }
        "SecurityConfig" { Edit-ReleaseSecurityConfig }
        "SecurityStatus" { Invoke-ReleaseSecurityCommand -ReleaseAction Status }
        "SecuritySync" { Invoke-SecuritySyncWorkflow }
        "SecurityBuild" { Invoke-SecurityBuildWorkflow -TargetDate $Value }
        "SecurityValidate" { Invoke-ReleaseSecurityCommand -ReleaseAction Validate }
        "SecurityAll" { Invoke-SecurityAllWorkflow -TargetDate $Value }
        "SecurityCode" { Invoke-ReleaseSecurityCommand -ReleaseAction Code -TargetDate $Value }
        "Exit" { Write-Host "워크플로를 종료합니다." }
    }
}
catch {
    Write-Host ""
    Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
