[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$workflowTestRoot = Split-Path -Parent $PSScriptRoot
$workflowScriptPath = Join-Path $workflowTestRoot 'version-workflow.ps1'
$workflowTokens = $null
$workflowParseErrors = $null
$workflowAst = [Management.Automation.Language.Parser]::ParseFile(
    $workflowScriptPath, [ref]$workflowTokens, [ref]$workflowParseErrors)
if ($workflowParseErrors.Count -gt 0) {
    throw ('Workflow syntax errors: ' + ($workflowParseErrors.Message -join '; '))
}

# Import definitions without running the menu, changing the real repository, or
# invoking Excel / GitHub. Each case runs these definitions in a fresh scope.
$workflowFunctionDefinitions = @($workflowAst.EndBlock.Statements |
    Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] })
$workflowFixtureRoot = Join-Path ([IO.Path]::GetTempPath()) (
    'version-workflow-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($workflowFixtureRoot)
$workflowCases = [Collections.Generic.List[object]]::new()

function Add-WorkflowCase {
    param([string]$Name, [scriptblock]$Body)
    $workflowCases.Add([pscustomobject]@{ Name = $Name; Body = $Body })
}

function Assert-WorkflowEqual {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -cne $Expected) {
        throw ($Message + '. Expected <' + $Expected + '>; actual <' + $Actual + '>.')
    }
}

function Assert-WorkflowTrue {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Invoke-FixtureGit {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    $fixturePreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $fixtureOutput = & git @Arguments 2>&1
        $fixtureExitCode = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = $fixturePreference }
    if ($fixtureExitCode -ne 0) {
        throw ('Fixture Git failed: git ' + ($Arguments -join ' ') + ': ' + ($fixtureOutput -join ' '))
    }
    return ($fixtureOutput -join "`n").Trim()
}

function New-WorkflowRepository {
    param([string]$DistributionFolder = 'dist')
    $fixturePath = Join-Path $workflowFixtureRoot ([guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($fixturePath)
    Push-Location $fixturePath
    try {
        $null = Invoke-FixtureGit init --quiet -b main
        $null = Invoke-FixtureGit config user.name 'Workflow Regression'
        $null = Invoke-FixtureGit config user.email 'workflow-test@example.invalid'
        $null = Invoke-FixtureGit config commit.gpgsign false
        $null = Invoke-FixtureGit config core.autocrlf false
        $null = Invoke-FixtureGit config core.hooksPath (Join-Path $fixturePath 'no-hooks')
        [void][IO.Directory]::CreateDirectory((Join-Path $fixturePath 'workbooks/dev'))
        [void][IO.Directory]::CreateDirectory((Join-Path $fixturePath 'dist'))
        [IO.File]::WriteAllText((Join-Path $fixturePath 'dist/.gitkeep'), '')
        [IO.File]::WriteAllText((Join-Path $fixturePath 'workbooks/dev/Report v1.0.0.xlsm'), 'fixture workbook')
        $fixtureConfig = [ordered]@{
            excel_file = 'workbooks/dev/Report v1.0.0.xlsm'
            release_security = [ordered]@{ distribution_folder = $DistributionFolder }
        } | ConvertTo-Json -Depth 4
        [IO.File]::WriteAllText((Join-Path $fixturePath 'config.json'), $fixtureConfig.Replace("`r`n", "`n"))
        $null = Invoke-FixtureGit add --all
        $null = Invoke-FixtureGit commit --quiet -m 'fixture: base release'
        $null = Invoke-FixtureGit tag v1.0.0
        return $fixturePath
    }
    finally { Pop-Location }
}

function Invoke-FixtureSecurity {
    param([string]$ReleaseAction, [string]$OutputPath)
    $script:workflowSecurityCalls.Add($ReleaseAction)
    $fixtureConfig = [IO.File]::ReadAllText((Join-Path (Get-Location) 'config.json')) | ConvertFrom-Json
    $fixtureOutputFolder = Join-Path (Get-Location) $fixtureConfig.release_security.distribution_folder
    $fixtureArtifact = Join-Path $fixtureOutputFolder 'release-test.xlsm'
    switch ($ReleaseAction) {
        'Build' {
            Assert-WorkflowTrue (Test-Path -LiteralPath $fixtureConfig.excel_file) 'Build must use an existing configured workbook'
            [void][IO.Directory]::CreateDirectory($fixtureOutputFolder)
            [IO.File]::WriteAllText($fixtureArtifact, ('built from ' + $fixtureConfig.excel_file))
            return $fixtureArtifact
        }
        'Validate' {
            Assert-WorkflowTrue (Test-Path -LiteralPath $fixtureArtifact) 'Validate must run after Build'
            Assert-WorkflowEqual $OutputPath $fixtureArtifact 'Validate must use the artifact returned by this build'
        }
        default { throw ('Unexpected security operation: ' + $ReleaseAction) }
    }
}

function Get-FixtureOption {
    param([string[]]$Arguments, [string]$Name)
    $optionIndex = [array]::IndexOf($Arguments, $Name)
    if ($optionIndex -lt 0 -or $optionIndex + 1 -ge $Arguments.Count) { return '' }
    return $Arguments[$optionIndex + 1]
}

function New-FixturePullRequest {
    param([string]$HeadBranch = 'develop/v1.0.1', [string]$BaseBranch = 'release/v1.0.1',
        [string]$State = 'OPEN', [int]$Number = 101)
    return [pscustomobject]@{
        number = $Number
        url = 'https://example.invalid/pull/' + $Number
        state = $State
        headRefName = $HeadBranch
        baseRefName = $BaseBranch
        headRefOid = Invoke-FixtureGit rev-parse $HeadBranch
        mergeCommit = $null
    }
}

function New-WorkflowMergeRepository {
    $fixturePath = New-WorkflowRepository
    Push-Location $fixturePath
    try {
        $null = Invoke-FixtureGit branch release/v1.0.1
        $null = Invoke-FixtureGit switch --quiet -c develop/v1.0.1
        [IO.File]::WriteAllText((Join-Path $fixturePath 'development.txt'), 'the new version must include this change')
        $null = Invoke-FixtureGit add development.txt
        $null = Invoke-FixtureGit commit --quiet -m 'fixture: development change'
        $originPath = Join-Path $workflowFixtureRoot (([guid]::NewGuid().ToString('N')) + '.git')
        $null = Invoke-FixtureGit init --bare --quiet $originPath
        $null = Invoke-FixtureGit remote add origin $originPath
        $null = Invoke-FixtureGit push --quiet --all origin
        $null = Invoke-FixtureGit push --quiet --tags origin
        return $fixturePath
    }
    finally { Pop-Location }
}

$workflowGhMockDefinitions = {
    $script:workflowGhRequests = [Collections.Generic.List[object]]::new()
    $script:workflowGhQueries = [Collections.Generic.List[object]]::new()
    $script:workflowGhCommands = [Collections.Generic.List[object]]::new()
    $script:workflowGhFailQuery = $false
    $script:workflowGhMergeMode = 'merge'
    $script:workflowGhBody = ''
    $script:workflowGhBodyPath = ''
    function Get-GhCommand { return 'FixtureGh' }
    function Get-GhJson {
        param([string]$GhCommand, [string[]]$Arguments)
        $script:workflowGhQueries.Add(@($Arguments))
        if ($script:workflowGhFailQuery) { throw 'Simulated GitHub connection failure' }
        if ($Arguments[0] -ne 'pr') { throw 'Unexpected GitHub command' }
        switch ($Arguments[1]) {
            'list' {
                $head = Get-FixtureOption $Arguments '--head'
                $base = Get-FixtureOption $Arguments '--base'
                $state = Get-FixtureOption $Arguments '--state'
                Assert-WorkflowTrue (-not [string]::IsNullOrWhiteSpace($head)) 'PR list must select a head branch'
                Assert-WorkflowTrue (-not [string]::IsNullOrWhiteSpace($base)) 'PR list must select a base branch'
                Assert-WorkflowTrue ($state -in @('open', 'merged')) 'PR list must select a supported state'
                return @($script:workflowGhRequests | Where-Object {
                    $_.headRefName -eq $head -and $_.baseRefName -eq $base -and $_.state -eq $state
                })
            }
            'view' {
                $matchingRequests = @($script:workflowGhRequests | Where-Object { [string]$_.number -eq $Arguments[2] })
                Assert-WorkflowEqual $matchingRequests.Count 1 'PR view must use the exact PR number'
                return $matchingRequests[0]
            }
            default { throw ('Unexpected GitHub query: ' + ($Arguments -join ' ')) }
        }
    }
    function Invoke-CheckedCommand {
        param([string]$Command, [string[]]$Arguments)
        if ($Command -eq 'git') {
            $null = Invoke-FixtureGit @Arguments
            return
        }
        Assert-WorkflowEqual $Command 'FixtureGh' 'Only mocked GitHub may run in the test'
        $script:workflowGhCommands.Add(@($Arguments))
        switch ($Arguments[1]) {
            'create' {
                $bodyPath = Get-FixtureOption $Arguments '--body-file'
                Assert-WorkflowTrue (Test-Path -LiteralPath $bodyPath) 'PR body must be written to a real temporary file'
                $script:workflowGhBody = [IO.File]::ReadAllText($bodyPath)
                $script:workflowGhBodyPath = $bodyPath
                $script:workflowGhRequests.Add((New-FixturePullRequest `
                    -HeadBranch (Get-FixtureOption $Arguments '--head') `
                    -BaseBranch (Get-FixtureOption $Arguments '--base') `
                    -Number (201 + $script:workflowGhRequests.Count)))
            }
            'merge' {
                $request = @($script:workflowGhRequests | Where-Object { [string]$_.number -eq $Arguments[2] })[0]
                Assert-WorkflowTrue ($null -ne $request) 'Merge must identify an existing PR by number'
                Assert-WorkflowEqual (Get-FixtureOption $Arguments '--match-head-commit') $request.headRefOid `
                    'Merge must protect the reviewed head commit'
                Assert-WorkflowTrue ($Arguments -notcontains '--delete-branch') 'GitHub merge must retain branches until verification'
                if ($script:workflowGhMergeMode -eq 'pending') { return }
                $oldBranch = Invoke-FixtureGit branch --show-current
                $null = Invoke-FixtureGit switch --quiet $request.baseRefName
                if ($script:workflowGhMergeMode -eq 'merge') {
                    $null = Invoke-FixtureGit merge --quiet --no-ff $request.headRefName -m 'fixture: merged PR'
                    $request.mergeCommit = [pscustomobject]@{ oid = Invoke-FixtureGit rev-parse HEAD }
                    $null = Invoke-FixtureGit push --quiet origin $request.baseRefName
                }
                else {
                    $request.mergeCommit = [pscustomobject]@{ oid = Invoke-FixtureGit rev-parse $request.headRefName }
                }
                $null = Invoke-FixtureGit switch --quiet $oldBranch
                $request.state = 'MERGED'
            }
            default { throw ('Unexpected GitHub mutation: ' + ($Arguments -join ' ')) }
        }
    }
}

Add-WorkflowCase 'npm positional values bind after the named Action parameter' {
    $bindingProbe = [scriptblock]::Create('[CmdletBinding()]' + [Environment]::NewLine +
        $workflowAst.ParamBlock.Extent.Text + [Environment]::NewLine +
        '[pscustomobject]@{ Action = $Action; Value = $Value; BranchName = $BranchName }')
    $start = & $bindingProbe -Action StartVersion patch
    Assert-WorkflowEqual $start.Action 'StartVersion' 'version:start action'
    Assert-WorkflowEqual $start.Value 'patch' 'version:start positional version part'
    $child = & $bindingProbe -Action CreateBranch feature/report
    Assert-WorkflowEqual $child.Value 'feature/report' 'branch:create positional branch'
    $build = & $bindingProbe -Action SecurityBuild 2026-09-11
    Assert-WorkflowEqual $build.Value '2026-09-11' 'release:build positional date'
}

Add-WorkflowCase 'release preparation renames, builds, validates, and commits the workbook' {
    $fixturePath = New-WorkflowRepository
    Push-Location $fixturePath
    try {
        function Invoke-ReleaseSecurityCommand { param([string]$ReleaseAction, [string]$OutputPath) Invoke-FixtureSecurity $ReleaseAction $OutputPath }
        Prepare-ReleaseCommit -PreviousVersion '1.0.0' -ReleaseVersion '1.0.1'
        Assert-WorkflowEqual (Get-ConfigFieldValue) 'workbooks/dev/Report v1.0.1.xlsm' 'Release config version'
        Assert-WorkflowTrue (Test-Path -LiteralPath 'workbooks/dev/Report v1.0.1.xlsm') 'New workbook is missing'
        Assert-WorkflowTrue (-not (Test-Path -LiteralPath 'workbooks/dev/Report v1.0.0.xlsm')) 'Old workbook was not renamed'
        Assert-WorkflowEqual ($script:workflowSecurityCalls -join ',') 'Build,Validate' 'Build / validation order'
        Assert-WorkflowEqual (Invoke-FixtureGit rev-list --count HEAD) '2' 'Release preparation creates one commit'
        Assert-WorkflowEqual (Invoke-FixtureGit status --porcelain) '' 'Release preparation must leave a clean worktree'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'release preparation recovers a rename interrupted before config update' {
    $fixturePath = New-WorkflowRepository
    Push-Location $fixturePath
    try {
        function Invoke-ReleaseSecurityCommand { param([string]$ReleaseAction, [string]$OutputPath) Invoke-FixtureSecurity $ReleaseAction $OutputPath }
        Move-Item -LiteralPath (Join-Path $fixturePath 'workbooks/dev/Report v1.0.0.xlsm') `
            -Destination (Join-Path $fixturePath 'workbooks/dev/Report v1.0.1.xlsm')
        Prepare-ReleaseCommit -PreviousVersion '1.0.0' -ReleaseVersion '1.0.1'
        Assert-WorkflowEqual (Get-ConfigFieldValue) 'workbooks/dev/Report v1.0.1.xlsm' 'Interrupted rename must repair config'
        Assert-WorkflowEqual ($script:workflowSecurityCalls -join ',') 'Build,Validate' 'Recovered release must be validated'
        Assert-WorkflowEqual (Invoke-FixtureGit status --porcelain) '' 'Recovered rename must be committed'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'release preparation commits artifacts in the configured distribution folder' {
    $fixturePath = New-WorkflowRepository -DistributionFolder 'deliverables'
    Push-Location $fixturePath
    try {
        function Invoke-ReleaseSecurityCommand { param([string]$ReleaseAction, [string]$OutputPath) Invoke-FixtureSecurity $ReleaseAction $OutputPath }
        Prepare-ReleaseCommit -PreviousVersion '1.0.0' -ReleaseVersion '1.0.1'
        Assert-WorkflowEqual (Invoke-FixtureGit ls-files -- 'deliverables/release-test.xlsm') `
            'deliverables/release-test.xlsm' 'Configured distribution artifact must be tracked'
        Assert-WorkflowEqual (Invoke-FixtureGit status --porcelain) '' 'Custom output folder must not block the following merge'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'repeating release preparation does not create an empty commit' {
    $fixturePath = New-WorkflowRepository
    Push-Location $fixturePath
    try {
        function Invoke-ReleaseSecurityCommand { param([string]$ReleaseAction, [string]$OutputPath) Invoke-FixtureSecurity $ReleaseAction $OutputPath }
        Prepare-ReleaseCommit -PreviousVersion '1.0.0' -ReleaseVersion '1.0.1'
        $preparedCommit = Invoke-FixtureGit rev-parse HEAD
        Prepare-ReleaseCommit -PreviousVersion '1.0.0' -ReleaseVersion '1.0.1'
        Assert-WorkflowEqual (Invoke-FixtureGit rev-parse HEAD) $preparedCommit 'Unchanged retry should retain the preparation commit'
        Assert-WorkflowEqual ($script:workflowSecurityCalls -join ',') 'Build,Validate,Build,Validate' 'Retry still validates its artifact'
        Assert-WorkflowEqual (Invoke-FixtureGit status --porcelain) '' 'Repeated preparation must leave a clean worktree'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'PR lookup excludes closed requests and requests for a different base' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhRequests.Add((New-FixturePullRequest -State 'CLOSED' -Number 11))
        $script:workflowGhRequests.Add((New-FixturePullRequest -BaseBranch 'main' -Number 12))
        $script:workflowGhRequests.Add((New-FixturePullRequest -Number 13))
        $request = Ensure-PullRequest -GhCommand FixtureGh -HeadBranch develop/v1.0.1 `
            -BaseBranchName release/v1.0.1 -Title 'fixture release'
        Assert-WorkflowEqual $request.number 13 'Only the open request for the exact head and base may be reused'
        Assert-WorkflowEqual $script:workflowGhCommands.Count 0 'An existing exact PR must not be created again'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'PR creation preserves multiline text and removes the temporary body file' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhRequests.Add((New-FixturePullRequest -State 'MERGED' -Number 11))
        $body = 'First line with literal $() and `backticks`.' + "`n`n" + 'Second line: "quotes" and \n.'
        $request = Ensure-PullRequest -GhCommand FixtureGh -HeadBranch develop/v1.0.1 `
            -BaseBranchName release/v1.0.1 -Title 'fixture release' -Body $body
        Assert-WorkflowEqual $request.number 202 'A past merged PR must not be reused for new changes'
        Assert-WorkflowEqual $script:workflowGhBody $body 'PR body bytes must retain actual newlines and literal characters'
        Assert-WorkflowTrue (-not (Test-Path -LiteralPath $script:workflowGhBodyPath)) 'PR body temporary file was not cleaned up'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'a GitHub lookup failure stops without attempting PR creation' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhFailQuery = $true
        $caught = $null
        try {
            Ensure-PullRequest -GhCommand FixtureGh -HeadBranch develop/v1.0.1 `
                -BaseBranchName release/v1.0.1 -Title 'fixture release'
        }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'A connection failure must be reported'
        Assert-WorkflowEqual $script:workflowGhCommands.Count 0 'A failed lookup must not trigger PR creation'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'pending PR merge leaves its source branch available for retry' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhRequests.Add((New-FixturePullRequest))
        $script:workflowGhMergeMode = 'pending'
        $originalBase = Invoke-FixtureGit rev-parse release/v1.0.1
        $caught = $null
        try { Merge-WorkBranch -HeadBranch develop/v1.0.1 -BaseBranchName release/v1.0.1 -Title 'fixture release' }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'An OPEN PR after merge request must stop the workflow'
        Assert-WorkflowEqual (Get-CurrentBranch) 'develop/v1.0.1' 'Pending merge must not switch to the release branch'
        Assert-WorkflowEqual (Invoke-FixtureGit rev-parse release/v1.0.1) $originalBase 'Pending merge must not change the release branch'
        Assert-WorkflowTrue (Test-LocalBranch develop/v1.0.1) 'Pending merge must preserve local development branch'
        Assert-WorkflowTrue (Test-RemoteTrackingBranch develop/v1.0.1) 'Pending merge must preserve remote development branch'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'confirmed merge checks real ancestry and keeps release recovery branches' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhRequests.Add((New-FixturePullRequest))
        $developmentCommit = Invoke-FixtureGit rev-parse develop/v1.0.1
        $request = Merge-WorkBranch -HeadBranch develop/v1.0.1 -BaseBranchName release/v1.0.1 `
            -Title 'fixture release' -KeepBranch
        Assert-WorkflowEqual $request.state 'MERGED' 'Return only a confirmed merged PR'
        Assert-WorkflowEqual (Get-CurrentBranch) 'release/v1.0.1' 'Successful merge switches to its target'
        Assert-WorkflowTrue (Test-GitAncestor $developmentCommit release/v1.0.1) 'Target must contain the exact development commit'
        Assert-WorkflowTrue (Test-LocalBranch develop/v1.0.1) 'KeepBranch must preserve local source'
        Assert-WorkflowTrue (Test-RemoteTrackingBranch develop/v1.0.1) 'KeepBranch must preserve remote source'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'a MERGED API response without Git ancestry must not delete source branches' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $script:workflowGhRequests.Add((New-FixturePullRequest))
        $script:workflowGhMergeMode = 'not-reflected'
        $caught = $null
        try { Merge-WorkBranch -HeadBranch develop/v1.0.1 -BaseBranchName release/v1.0.1 -Title 'fixture release' }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'A merge missing from the target branch must be rejected'
        Assert-WorkflowTrue (Test-LocalBranch develop/v1.0.1) 'Unverified ancestry must preserve local source'
        Assert-WorkflowTrue (Test-RemoteTrackingBranch develop/v1.0.1) 'Unverified ancestry must preserve remote source'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'stale release includes development and resumes a failed tag push from main' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        function Invoke-ReleaseSecurityCommand {
            param([string]$ReleaseAction, [string]$OutputPath)
            if ($ReleaseAction -eq 'Build') {
                Assert-WorkflowEqual (Get-CurrentBranch) 'release/v1.0.1' 'Build must run on the release branch'
                Assert-WorkflowTrue (Test-GitAncestor develop/v1.0.1 HEAD) 'Outstanding development must be merged before Build'
                Assert-WorkflowTrue (Test-Path -LiteralPath 'development.txt') 'The release must contain the new development file'
            }
            Invoke-FixtureSecurity $ReleaseAction $OutputPath
        }
        $fixtureCheckedCommand = (Get-Item Function:Invoke-CheckedCommand).ScriptBlock
        function Invoke-CheckedCommand {
            param([string]$Command, [string[]]$Arguments)
            if ($Command -eq 'git' -and $Arguments[0] -eq 'push' -and $Arguments -contains 'refs/tags/v1.0.1') {
                throw 'Simulated tag push interruption'
            }
            & $fixtureCheckedCommand -Command $Command -Arguments $Arguments
        }
        $null = Invoke-FixtureGit switch --quiet release/v1.0.1
        $pending = Set-PendingVersionRelease -Version '1.0.1' -Base main -Development develop/v1.0.1
        $caught = $null
        try { Invoke-PendingVersionRelease -Pending $pending }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'The simulated tag push failure must stop the workflow'
        Assert-WorkflowEqual $caught.Exception.Message 'Simulated tag push interruption' 'Release must reach the tag push before stopping'
        $savedPending = Get-PendingVersionRelease
        Assert-WorkflowTrue ($savedPending.Commit -match '^[0-9a-f]{40,64}$') 'The verified release merge commit must be persisted before tag push'
        Assert-WorkflowEqual (Get-CurrentBranch) 'main' 'The interrupted final stage should be recoverable from main'
        Assert-WorkflowTrue (Test-LocalBranch develop/v1.0.1) 'Failed tag push must retain development branch'
        Assert-WorkflowTrue (Test-LocalBranch release/v1.0.1) 'Failed tag push must retain release branch'
        Assert-WorkflowEqual (Invoke-FixtureGit ls-remote origin refs/tags/v1.0.1) '' 'The failed tag push must not be considered published'

        # Another main commit after the release merge must not move the release tag.
        [IO.File]::WriteAllText((Join-Path $fixturePath 'later-main.txt'), 'later unrelated main change')
        $null = Invoke-FixtureGit add later-main.txt
        $null = Invoke-FixtureGit commit --quiet -m 'fixture: later main update'
        $null = Invoke-FixtureGit push --quiet origin main
        Set-Item Function:Invoke-CheckedCommand -Value $fixtureCheckedCommand
        Complete-VersionDevelopmentAndRelease
        Assert-WorkflowEqual (Invoke-FixtureGit rev-parse 'v1.0.1^{commit}') $savedPending.Commit `
            'Recovered tag must target the saved release merge, even after main advances'
        Assert-WorkflowTrue ((Invoke-FixtureGit rev-parse HEAD) -ne $savedPending.Commit) 'The fixture must include a later main commit'
        $remoteTag = Invoke-FixtureGit ls-remote origin 'refs/tags/v1.0.1^{}'
        Assert-WorkflowTrue ($remoteTag.StartsWith($savedPending.Commit)) 'The exact release commit must be verified on the remote'
        Assert-WorkflowEqual ($script:workflowSecurityCalls -join ',') 'Build,Validate' 'Tag-only recovery must not rebuild the release'
        Assert-WorkflowTrue ($null -eq (Get-PendingVersionRelease)) 'Verified tag publication must clear pending state'
        Assert-WorkflowTrue (-not (Test-LocalBranch develop/v1.0.1)) 'Completed release should clean its development branch'
        Assert-WorkflowTrue (-not (Test-LocalBranch release/v1.0.1)) 'Completed release should clean its release branch'
        Assert-WorkflowTrue (-not (Test-RemoteTrackingBranch develop/v1.0.1)) 'Completed release should clean remote development branch'
        Assert-WorkflowTrue (-not (Test-RemoteTrackingBranch release/v1.0.1)) 'Completed release should clean remote release branch'
        Assert-WorkflowEqual (Invoke-FixtureGit status --porcelain) '' 'Completed recovery must leave the working tree clean'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'missing development branch requires verified merge history before building' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        function Invoke-ReleaseSecurityCommand {
            param([string]$ReleaseAction, [string]$OutputPath)
            Invoke-FixtureSecurity $ReleaseAction $OutputPath
        }
        $null = Invoke-FixtureGit switch --quiet release/v1.0.1
        # Legacy version/* names are supported; this missing legacy source has no
        # merged PR proving its contents are present in the release branch.
        $pending = Set-PendingVersionRelease -Version '1.0.1' -Base main -Development version/v1.0.1
        $caught = $null
        try { Invoke-PendingVersionRelease -Pending $pending }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'Missing development without merge evidence must stop release recovery'
        Assert-WorkflowEqual $script:workflowSecurityCalls.Count 0 'Unverified development must not be packaged'
        Assert-WorkflowEqual $script:workflowGhCommands.Count 0 'Unverified development must not trigger a release PR'
        Assert-WorkflowEqual (Invoke-FixtureGit tag --list v1.0.1) '' 'Unverified development must not be tagged'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'a branch advanced after validation is rejected before push or PR activity' {
    $fixturePath = New-WorkflowMergeRepository
    Push-Location $fixturePath
    try {
        . $workflowGhMockDefinitions
        $validatedCommit = Invoke-FixtureGit rev-parse develop/v1.0.1
        $originalRemote = Invoke-FixtureGit ls-remote origin refs/heads/develop/v1.0.1
        $originalBase = Invoke-FixtureGit rev-parse release/v1.0.1
        [IO.File]::WriteAllText((Join-Path $fixturePath 'unvalidated-change.txt'), 'committed after validation')
        $null = Invoke-FixtureGit add unvalidated-change.txt
        $null = Invoke-FixtureGit commit --quiet -m 'fixture: unvalidated change'
        $advancedCommit = Invoke-FixtureGit rev-parse HEAD
        $caught = $null
        try {
            Merge-WorkBranch -HeadBranch develop/v1.0.1 -BaseBranchName release/v1.0.1 `
                -Title 'fixture release' -ExpectedHeadCommit $validatedCommit
        }
        catch { $caught = $_ }
        Assert-WorkflowTrue ($null -ne $caught) 'A changed head must invalidate the earlier validation'
        Assert-WorkflowEqual (Invoke-FixtureGit ls-remote origin refs/heads/develop/v1.0.1) $originalRemote `
            'The unvalidated head must not be pushed'
        Assert-WorkflowEqual $script:workflowGhQueries.Count 0 'Head mismatch must stop before GitHub PR lookup'
        Assert-WorkflowEqual $script:workflowGhCommands.Count 0 'Head mismatch must stop before GitHub PR mutation'
        Assert-WorkflowEqual (Invoke-FixtureGit rev-parse develop/v1.0.1) $advancedCommit 'The new local commit must be preserved'
        Assert-WorkflowEqual (Invoke-FixtureGit rev-parse release/v1.0.1) $originalBase 'The release target must stay unchanged'
    }
    finally { Pop-Location }
}

Add-WorkflowCase 'a new release clears an orphaned merge commit without a pending version' {
    $fixturePath = New-WorkflowRepository
    Push-Location $fixturePath
    try {
        $orphanedCommit = Invoke-FixtureGit rev-parse HEAD
        $null = Invoke-FixtureGit config --local workflow.releaseCommit $orphanedCommit
        $null = Invoke-FixtureGit config --local workflow.releaseBase previous-base
        $null = Invoke-FixtureGit config --local workflow.releaseDevelopment develop/v0.9.9
        Assert-WorkflowTrue ($null -eq (Get-PendingVersionRelease)) 'An absent version must mean there is no active release'
        $pending = Set-PendingVersionRelease -Version '1.0.1' -Base main -Development develop/v1.0.1
        Assert-WorkflowEqual $pending.Version '1.0.1' 'The new release version must be recorded'
        Assert-WorkflowEqual $pending.Base 'main' 'The stale release target must be replaced'
        Assert-WorkflowEqual $pending.Development 'develop/v1.0.1' 'The stale development source must be replaced'
        Assert-WorkflowEqual $pending.Commit '' 'A new release must not inherit the previous release merge commit'
        Assert-WorkflowEqual (Get-WorkflowSetting -Name releaseCommit) '' 'The orphaned commit must be removed from local Git config'
    }
    finally { Pop-Location }
}

$workflowFailures = [Collections.Generic.List[string]]::new()
$workflowPassed = 0
try {
    foreach ($workflowCase in $workflowCases) {
        try {
            & {
                foreach ($workflowFunction in $workflowFunctionDefinitions) {
                    . ([scriptblock]::Create($workflowFunction.Extent.Text))
                }
                $ConfigFile = 'config.json'
                $ConfigField = 'excel_file'
                $DefaultVersionBaseBranch = 'main'
                $SourceBranch = ''
                $TargetBranch = ''
                $script:workflowSecurityCalls = [Collections.Generic.List[string]]::new()
                & $workflowCase.Body
            } 6> $null | Out-Null
            $workflowPassed++
            Write-Host ('PASS: ' + $workflowCase.Name)
        }
        catch {
            $workflowFailure = $workflowCase.Name + ': ' + $_.Exception.Message
            $workflowFailures.Add($workflowFailure)
            Write-Host ('FAIL: ' + $workflowFailure) -ForegroundColor Red
        }
    }
}
finally {
    # Delete only this run's checked absolute temporary root, never the real repo.
    $resolvedFixtureRoot = [IO.Path]::GetFullPath($workflowFixtureRoot)
    $allowedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedFixtureRoot.StartsWith($allowedTempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($resolvedFixtureRoot) -notmatch '^version-workflow-test-[a-f0-9]{32}$') {
        throw ('Refusing to remove unexpected fixture path: ' + $resolvedFixtureRoot)
    }
    if (Test-Path -LiteralPath $resolvedFixtureRoot) {
        Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force
    }
}

Write-Host ("Workflow regression: $workflowPassed passed, $($workflowFailures.Count) failed.")
if ($workflowFailures.Count -gt 0) { exit 1 }
