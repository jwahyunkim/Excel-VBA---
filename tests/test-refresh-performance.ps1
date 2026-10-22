[CmdletBinding()]
param([string]$WorkbookPath, [string]$BaselineRef)

# 2026-10-22 김좌현 수정 : 격리 Excel 사본의 간트 계산/레이아웃 회귀와 선택적 Git 기준본 비교
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
if ($BaselineRef -and $BaselineRef -notmatch '^[A-Za-z0-9_./^~:-]+$') { throw 'Invalid Git baseline reference.' }

function ConvertFrom-TaskVbaBytes([byte[]]$Bytes) {
    try { $taskSource = [Text.UTF8Encoding]::new($false, $true).GetString($Bytes) }
    catch { $taskSource = [Text.Encoding]::GetEncoding(949).GetString($Bytes) }
    return $taskSource.TrimStart([char]0xFEFF)
}
function Read-TaskVba([string]$Path, [string]$SourceRef) {
    if (-not $SourceRef) { return ConvertFrom-TaskVbaBytes ([IO.File]::ReadAllBytes($Path)) }
    $taskRelative = $Path.Substring($taskRoot.Length + 1).Replace('\', '/')
    $taskStart = [Diagnostics.ProcessStartInfo]::new()
    $taskStart.FileName = 'git'
    $taskStart.Arguments = 'show "' + $SourceRef + ':' + $taskRelative + '"'
    $taskStart.WorkingDirectory = $taskRoot
    $taskStart.UseShellExecute = $false
    $taskStart.CreateNoWindow = $true
    $taskStart.RedirectStandardOutput = $true
    $taskStart.RedirectStandardError = $true
    $taskProcess = [Diagnostics.Process]::Start($taskStart)
    $taskBuffer = [IO.MemoryStream]::new()
    try {
        $taskProcess.StandardOutput.BaseStream.CopyTo($taskBuffer)
        $taskError = $taskProcess.StandardError.ReadToEnd()
        $taskProcess.WaitForExit()
        if ($taskProcess.ExitCode -ne 0) { throw "Cannot read $SourceRef/$taskRelative : $taskError" }
        return ConvertFrom-TaskVbaBytes ($taskBuffer.ToArray())
    }
    finally { $taskBuffer.Dispose(); $taskProcess.Dispose() }
}
$taskHarnessSource = (Read-TaskVba (Join-Path $PSScriptRoot 'RefreshPerformanceRegression.bas') '') -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
function Invoke-RefreshRegression([string]$SourceRef) {
    $taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('refresh-regression-' + [guid]::NewGuid().ToString('N') + '.xlsm')
    Copy-Item -LiteralPath $WorkbookPath -Destination $taskTemp
    $taskExcel = $null
    $taskBook = $null
    $taskPhase = 'Excel startup'
    try {
        $taskExcel = New-Object -ComObject Excel.Application
        $taskExcel.Visible = $false
        $taskExcel.DisplayAlerts = $false
        $taskExcel.EnableEvents = $false
        $taskExcel.AutomationSecurity = 1
        $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
        foreach ($taskFile in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
            $taskPhase = 'import ' + $taskFile.BaseName
            $taskSource = (Read-TaskVba $taskFile.FullName $SourceRef) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
            try { $taskComponent = $taskBook.VBProject.VBComponents.Item($taskFile.BaseName) }
            catch { $taskComponent = $taskBook.VBProject.VBComponents.Add(1); $taskComponent.Name = $taskFile.BaseName }
            $taskCode = $taskComponent.CodeModule
            if ($taskCode.CountOfLines -gt 0) { [void]$taskCode.DeleteLines(1, $taskCode.CountOfLines) }
            [void]$taskCode.AddFromString($taskSource)
        }
        $taskPhase = 'import regression harness'
        try { $taskComponent = $taskBook.VBProject.VBComponents.Item('RefreshPerformanceRegression') }
        catch { $taskComponent = $taskBook.VBProject.VBComponents.Add(1); $taskComponent.Name = 'RefreshPerformanceRegression' }
        $taskSource = $taskHarnessSource
        $taskCode = $taskComponent.CodeModule
        if ($taskCode.CountOfLines -gt 0) { [void]$taskCode.DeleteLines(1, $taskCode.CountOfLines) }
        [void]$taskCode.AddFromString($taskSource)
        $taskPhase = 'run regression'
        $taskResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunRefreshPerformanceRegression")
        Write-Host (($SourceRef, 'working tree' | Where-Object { $_ } | Select-Object -First 1) + ': ' + $taskResult)
        if (-not $taskResult.StartsWith('PASS:')) { throw $taskResult }
        $taskSnapshot = [string]$taskExcel.Run("'$($taskBook.Name)'!GetRefreshScenarioSnapshots")
        return @{ Result = $taskResult; Snapshot = $taskSnapshot }
    }
    catch { throw "$taskPhase : $($_.Exception.Message)" }
    finally {
        if ($taskExcel) { try { $taskExcel.EnableEvents = $false } catch {} }
        if ($taskBook) { try { $taskBook.Close($false) } catch {} }
        if ($taskExcel) { try { $taskExcel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
        if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp -Force }
    }
}

$taskCurrent = Invoke-RefreshRegression ''
if ($BaselineRef) {
    $taskBaseline = Invoke-RefreshRegression $BaselineRef
    if ($taskCurrent.Snapshot -cne $taskBaseline.Snapshot) {
        $taskCurrentLines = $taskCurrent.Snapshot -split '\r?\n'
        $taskBaselineLines = $taskBaseline.Snapshot -split '\r?\n'
        $taskDifference = 'snapshot length differs'
        for ($taskIndex = 0; $taskIndex -lt [Math]::Min($taskCurrentLines.Count, $taskBaselineLines.Count); $taskIndex++) {
            if ($taskCurrentLines[$taskIndex] -cne $taskBaselineLines[$taskIndex]) {
                $taskDifference = "line $($taskIndex + 1): baseline=$($taskBaselineLines[$taskIndex]); current=$($taskCurrentLines[$taskIndex])"
                break
            }
        }
        throw "Git baseline comparison failed: $taskDifference"
    }
    Write-Output "PASS: working tree and $BaselineRef have identical scenario snapshots, including geometry and four-edge/gradient visuals."
}
