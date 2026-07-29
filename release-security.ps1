[CmdletBinding()]
param(
    [ValidateSet("Build", "Code", "Status", "SyncDev", "Validate")]
    [string]$Action = "Status",

    [string]$Date = "",
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
$ScriptRoot = $PSScriptRoot
$ConfigPath = Join-Path $ScriptRoot "config.json"
$SecurityModulePath = Join-Path $ScriptRoot "vba-files\Module\modReleaseSecurity.bas"
$WorkbookClassPath = Join-Path $ScriptRoot "vba-files\Class\현재_통합_문서.cls"
$CodeModulus = 1679616

function Resolve-WorkspacePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $candidate = if ([System.IO.Path]::IsPathRooted($Path)) {
        $Path
    }
    else {
        Join-Path $ScriptRoot $Path
    }

    $resolved = [System.IO.Path]::GetFullPath($candidate)
    $rootWithSeparator = [System.IO.Path]::GetFullPath($ScriptRoot).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "워크스페이스 밖의 경로는 사용할 수 없습니다: $resolved"
    }

    return $resolved
}

function Get-LocalConfig {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "config.json이 없습니다. config.example.json을 복사하고 로컬 값을 설정하세요."
    }

    try {
        $config = [System.IO.File]::ReadAllText($ConfigPath) | ConvertFrom-Json
    }
    catch {
        throw "config.json 파싱 실패: $($_.Exception.Message)"
    }

    if ([string]::IsNullOrWhiteSpace([string]$config.excel_file)) {
        throw "config.json의 excel_file이 비어 있습니다."
    }
    if ($null -eq $config.release_security) {
        throw "config.json의 release_security 설정이 없습니다."
    }

    $security = $config.release_security
    $usageDays = [int]$security.usage_days
    $renewalDays = [int]$security.renewal_days
    $projectPassword = [string]$security.vba_project_password
    $renewalSecret = [string]$security.renewal_secret
    $distributionFolder = [string]$security.distribution_folder

    if ($usageDays -lt 1) { throw "usage_days는 1 이상이어야 합니다." }
    if ($renewalDays -lt 1) { throw "renewal_days는 1 이상이어야 합니다." }
    if ([string]::IsNullOrWhiteSpace($projectPassword)) { throw "vba_project_password가 비어 있습니다." }
    if ([string]::IsNullOrWhiteSpace($renewalSecret)) { throw "renewal_secret이 비어 있습니다." }
    if ($renewalSecret -match '[^\x20-\x7E]') {
        throw "PowerShell과 VBA가 같은 코드를 만들도록 renewal_secret은 ASCII 문자만 사용하세요."
    }
    if ([string]::IsNullOrWhiteSpace($distributionFolder)) { throw "distribution_folder가 비어 있습니다." }

    return [PSCustomObject]@{
        WorkbookPath = Resolve-WorkspacePath -Path ([string]$config.excel_file)
        UsageDays = $usageDays
        RenewalDays = $renewalDays
        ProjectPassword = $projectPassword
        RenewalSecret = $renewalSecret
        DistributionFolder = Resolve-WorkspacePath -Path $distributionFolder
    }
}

function Resolve-TargetDate {
    if ([string]::IsNullOrWhiteSpace($Date)) {
        return [DateTime]::Today
    }

    $parsedDate = [DateTime]::MinValue
    $format = "yyyy-MM-dd"
    if (-not [DateTime]::TryParseExact(
            $Date,
            $format,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$parsedDate)) {
        throw "Date는 yyyy-MM-dd 형식이어야 합니다: $Date"
    }

    return $parsedDate.Date
}

function ConvertTo-Base36 {
    param(
        [Parameter(Mandatory = $true)][long]$Value,
        [Parameter(Mandatory = $true)][int]$Width
    )

    $digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $result = ""
    do {
        $digitValue = [int]($Value % 36)
        $result = $digits[$digitValue] + $result
        $Value = [Math]::Floor($Value / 36)
    } while ($Value -gt 0)

    return $result.PadLeft($Width, '0').Substring([Math]::Max(0, $result.PadLeft($Width, '0').Length - $Width))
}

function New-RenewalCode {
    param(
        [Parameter(Mandatory = $true)][DateTime]$TargetDate,
        [Parameter(Mandatory = $true)][string]$Secret
    )

    $rotatedSecret = if ($Secret.Length -le 1) {
        $Secret
    }
    else {
        $Secret.Substring(1) + $Secret.Substring(0, 1)
    }

    $payload = $rotatedSecret + "|" + $TargetDate.ToString("ddyyMM", [Globalization.CultureInfo]::InvariantCulture)
    [long]$accumulatorA = 7919
    [long]$accumulatorB = 104729

    for ($index = 1; $index -le $payload.Length; $index++) {
        [int]$characterCode = [int][char]$payload[$index - 1]
        $accumulatorA = (($accumulatorA * 33) + $characterCode + ($index * 17)) % $CodeModulus
        $accumulatorB = (($accumulatorB * 37) + ($characterCode * 7) + ($index * 13)) % $CodeModulus
    }

    return "$(ConvertTo-Base36 -Value $accumulatorA -Width 4)-$(ConvertTo-Base36 -Value $accumulatorB -Width 4)"
}

function Get-ReleaseVersion {
    param([Parameter(Mandatory = $true)][string]$WorkbookPath)

    $fileName = [System.IO.Path]::GetFileName($WorkbookPath)
    $matches = [regex]::Matches($fileName, 'v(\d+\.\d+\.\d+)', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($matches.Count -ne 1) {
        throw "개발본 파일명에서 vA.B.C 버전을 정확히 하나 찾지 못했습니다: $fileName"
    }
    return $matches[0].Groups[1].Value
}

function Get-ClassCode {
    $lines = [IO.File]::ReadAllLines($WorkbookClassPath)
    $startIndex = -1
    for ($index = 0; $index -lt $lines.Length; $index++) {
        if ($lines[$index].Trim() -eq "Option Explicit") {
            $startIndex = $index
            break
        }
    }
    if ($startIndex -lt 0) {
        throw "현재_통합_문서.cls에서 Option Explicit 시작점을 찾지 못했습니다."
    }

    return [string]::Join("`r`n", $lines[$startIndex..($lines.Length - 1)])
}

function Get-ModuleCode {
    $lines = [IO.File]::ReadAllLines($SecurityModulePath)
    $startIndex = -1
    for ($index = 0; $index -lt $lines.Length; $index++) {
        if ($lines[$index].Trim() -eq "Option Explicit") {
            $startIndex = $index
            break
        }
    }
    if ($startIndex -lt 0) {
        throw "modReleaseSecurity.bas에서 Option Explicit 시작점을 찾지 못했습니다."
    }

    return [string]::Join("`r`n", $lines[$startIndex..($lines.Length - 1)])
}

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Set-WorkbookVbaSource {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)]$Components
    )

    $existingModule = $null
    try {
        $existingModule = $Components.Item("modReleaseSecurity")
    }
    catch {
        $existingModule = $null
    }

    if ($null -ne $existingModule) {
        $Components.Remove($existingModule)
        Release-ComObject $existingModule
    }

    $moduleComponent = $null
    $moduleCode = $null
    try {
        $moduleComponent = $Components.Add(1)
        $moduleComponent.Name = "modReleaseSecurity"
        $moduleCode = $moduleComponent.CodeModule
        $moduleCode.AddFromString((Get-ModuleCode))
    }
    finally {
        Release-ComObject $moduleCode
        Release-ComObject $moduleComponent
    }

    $workbookComponent = $null
    $codeModule = $null
    try {
        $workbookComponent = $Components.Item([string]$Workbook.CodeName)
        $codeModule = $workbookComponent.CodeModule
        if ($codeModule.CountOfLines -gt 0) {
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
        }
        $codeModule.AddFromString((Get-ClassCode))
    }
    finally {
        Release-ComObject $codeModule
        Release-ComObject $workbookComponent
    }
}

function Invoke-WorkbookUpdate {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [switch]$ConfigureRelease,
        [DateTime]$ReleaseDate,
        [DateTime]$ExpiryDate,
        [string]$ReleaseVersion,
        [int]$RenewalDays,
        [string]$RenewalSecret
    )

    $excel = $null
    $workbook = $null
    $project = $null
    $components = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 1

        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
        try {
            $project = $workbook.VBProject
            $components = $project.VBComponents
        }
        catch {
            throw "Excel의 'VBA 프로젝트 개체 모델에 안전하게 액세스' 설정이 필요합니다. Excel 옵션 > 보안 센터 > 매크로 설정에서 해당 항목을 켜세요. 원인: $($_.Exception.Message)"
        }

        Set-WorkbookVbaSource -Workbook $workbook -Components $components

        if ($ConfigureRelease) {
            $macroName = "'$($workbook.Name.Replace("'", "''"))'!ConfigureReleaseSecurity"
            $vbaCode = [string]$excel.Run(
                $macroName,
                $ReleaseVersion,
                [double]$ReleaseDate.ToOADate(),
                [double]$ExpiryDate.ToOADate(),
                [int]$RenewalDays,
                $RenewalSecret)
            $expectedCode = New-RenewalCode -TargetDate $ReleaseDate -Secret $RenewalSecret
            if ($vbaCode -ne $expectedCode) {
                throw "PowerShell/VBA 연장코드 계산 불일치: PowerShell=$expectedCode, VBA=$vbaCode"
            }

            $visibleSheets = @()
            for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
                $sheet = $null
                try {
                    $sheet = $workbook.Worksheets.Item($index)
                    if ([int]$sheet.Visible -eq -1) { $visibleSheets += [string]$sheet.Name }
                }
                finally {
                    Release-ComObject $sheet
                }
            }
            if ($visibleSheets.Count -ne 1 -or $visibleSheets[0] -ne "사용안내") {
                throw "배포본 잠금 상태가 올바르지 않습니다. 표시 시트: $($visibleSheets -join ', ')"
            }
        }

        $workbook.Save()
        $workbook.Close($false)
        Release-ComObject $workbook
        $workbook = $null
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
        }
        if ($null -ne $excel) {
            try { $excel.EnableEvents = $true } catch {}
            try { $excel.Quit() } catch {}
        }
        Release-ComObject $components
        Release-ComObject $project
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Show-Status {
    $config = Get-LocalConfig
    $targetDate = Resolve-TargetDate
    $workbookExists = Test-Path -LiteralPath $config.WorkbookPath

    Write-Host ""
    Write-Host "개발 원본       : $($config.WorkbookPath)"
    Write-Host "개발 원본 존재  : $workbookExists"
    Write-Host "배포 폴더       : $($config.DistributionFolder)"
    Write-Host "기본 사용 기간  : $($config.UsageDays)일"
    Write-Host "1회 연장 기간   : $($config.RenewalDays)일"
    Write-Host "프로젝트 암호   : 설정됨(config.json)"
    Write-Host "오늘 연장 코드  : $(New-RenewalCode -TargetDate $targetDate -Secret $config.RenewalSecret)"
}

function Show-Code {
    $config = Get-LocalConfig
    $targetDate = Resolve-TargetDate
    Write-Host "날짜: $($targetDate.ToString('yyyy-MM-dd'))"
    Write-Host "연장 코드: $(New-RenewalCode -TargetDate $targetDate -Secret $config.RenewalSecret)" -ForegroundColor Green
}

function Sync-DevelopmentWorkbook {
    $config = Get-LocalConfig
    if (-not (Test-Path -LiteralPath $config.WorkbookPath)) {
        throw "개발 원본을 찾을 수 없습니다: $($config.WorkbookPath)"
    }
    if (-not (Test-Path -LiteralPath $SecurityModulePath)) { throw "보안 VBA 모듈이 없습니다: $SecurityModulePath" }
    if (-not (Test-Path -LiteralPath $WorkbookClassPath)) { throw "통합문서 클래스 소스가 없습니다: $WorkbookClassPath" }

    Write-Host "[1/2] 개발 원본에 보안 모듈과 통합문서 이벤트를 동기화합니다."
    Invoke-WorkbookUpdate -WorkbookPath $config.WorkbookPath
    Write-Host "[2/2] 완료: 보안 표시자가 없으므로 개발 원본의 사용 방식은 바뀌지 않습니다." -ForegroundColor Green
}

function Build-ReleaseWorkbook {
    $config = Get-LocalConfig
    $releaseDate = Resolve-TargetDate
    if (-not (Test-Path -LiteralPath $config.WorkbookPath)) {
        throw "개발 원본을 찾을 수 없습니다: $($config.WorkbookPath)"
    }
    if (-not (Test-Path -LiteralPath $SecurityModulePath)) { throw "보안 VBA 모듈이 없습니다: $SecurityModulePath" }
    if (-not (Test-Path -LiteralPath $WorkbookClassPath)) { throw "통합문서 클래스 소스가 없습니다: $WorkbookClassPath" }

    $releaseVersion = Get-ReleaseVersion -WorkbookPath $config.WorkbookPath
    $expiryDate = $releaseDate.AddDays($config.UsageDays)
    $distributionPath = if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $baseName = [IO.Path]::GetFileNameWithoutExtension($config.WorkbookPath)
        Join-Path $config.DistributionFolder "${baseName}_배포_$($releaseDate.ToString('yyyyMMdd')).xlsm"
    }
    else {
        Resolve-WorkspacePath -Path $OutputPath
    }

    if ([IO.Path]::GetExtension($distributionPath) -ne ".xlsm") {
        throw "배포 파일 확장자는 .xlsm이어야 합니다: $distributionPath"
    }

    Write-Host "[1/6] 설정 검증 완료"
    if (-not (Test-Path -LiteralPath $config.DistributionFolder)) {
        $null = New-Item -ItemType Directory -Path $config.DistributionFolder
    }
    Write-Host "[2/6] 배포 폴더 준비 완료"
    Copy-Item -LiteralPath $config.WorkbookPath -Destination $distributionPath -Force
    Write-Host "[3/6] 개발 원본을 배포본으로 복사 완료"

    Invoke-WorkbookUpdate `
        -WorkbookPath $distributionPath `
        -ConfigureRelease `
        -ReleaseDate $releaseDate `
        -ExpiryDate $expiryDate `
        -ReleaseVersion $releaseVersion `
        -RenewalDays $config.RenewalDays `
        -RenewalSecret $config.RenewalSecret
    Write-Host "[4/6] 보안 VBA/만료정보 적용 및 계산 교차검증 완료"
    Write-Host "[5/6] 잠금 상태 검증 완료(디스크에는 사용안내 시트만 표시)"
    Write-Host "[6/6] 배포본 생성 완료" -ForegroundColor Green
    Write-Host ""
    Write-Host "배포본: $distributionPath" -ForegroundColor Green
    Write-Host "사용 기간: $($releaseDate.ToString('yyyy-MM-dd')) ~ $($expiryDate.ToString('yyyy-MM-dd'))"
    Write-Host "오늘 연장 코드: $(New-RenewalCode -TargetDate $releaseDate -Secret $config.RenewalSecret)"
    Write-Host ""
    Write-Host "마지막 수동 단계(VBA 소스 열람 방지):" -ForegroundColor Yellow
    Write-Host "1) 배포본을 열고 Alt+F11"
    Write-Host "2) 도구 > VBAProject 속성 > 보호"
    Write-Host "3) '보기 위해 프로젝트 잠금' 체크 후 아래 암호를 두 번 입력"
    Write-Host "   $($config.ProjectPassword)"
    Write-Host "4) 저장하고 Excel을 완전히 닫은 뒤 다시 열어 잠금을 확인"
}

function Read-ReleaseWorkbookState {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][bool]$MacrosEnabled
    )

    $excel = $null
    $workbook = $null
    $securitySheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = if ($MacrosEnabled) { 1 } else { 3 }
        $excel.EnableEvents = $MacrosEnabled

        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
        $visibleSheets = @()
        $guideVisibility = $null
        $securityVisibility = $null
        $marker = ""
        $expiryDate = [DateTime]::MinValue

        for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
            $sheet = $null
            try {
                $sheet = $workbook.Worksheets.Item($index)
                $sheetName = [string]$sheet.Name
                if ([int]$sheet.Visible -eq -1) { $visibleSheets += $sheetName }
                if ($sheetName -eq "사용안내") { $guideVisibility = [int]$sheet.Visible }
                if ($sheetName -eq "__RELEASE_SECURITY") {
                    $securityVisibility = [int]$sheet.Visible
                    $marker = [string]$sheet.Range("A1").Value2
                    $expiryDate = [DateTime]::FromOADate([double]$sheet.Range("B4").Value2).Date
                }
            }
            finally {
                Release-ComObject $sheet
            }
        }

        return [PSCustomObject]@{
            VisibleSheets = $visibleSheets
            GuideVisibility = $guideVisibility
            SecurityVisibility = $securityVisibility
            Marker = $marker
            ExpiryDate = $expiryDate
            Saved = [bool]$workbook.Saved
        }
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
        }
        Release-ComObject $securitySheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Test-ReleaseSaveCycle {
    param([Parameter(Mandatory = $true)][string]$WorkbookPath)

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 1
        $excel.EnableEvents = $true

        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
        $workbook.Save()

        $visibleSheets = @()
        for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
            $sheet = $null
            try {
                $sheet = $workbook.Worksheets.Item($index)
                if ([int]$sheet.Visible -eq -1) { $visibleSheets += [string]$sheet.Name }
            }
            finally {
                Release-ComObject $sheet
            }
        }

        $businessSheets = @($visibleSheets | Where-Object { $_ -notin @("사용안내", "__RELEASE_SECURITY") })
        if ($businessSheets.Count -lt 1) { throw "저장 후 업무 시트가 다시 표시되지 않았습니다." }
        if ($visibleSheets -contains "사용안내" -or $visibleSheets -contains "__RELEASE_SECURITY") {
            throw "저장 후 보안용 시트가 표시되었습니다: $($visibleSheets -join ', ')"
        }
        if (-not [bool]$workbook.Saved) { throw "저장 후 통합문서가 변경 상태로 남았습니다." }
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
        }
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Test-ReleaseWorkbook {
    $config = Get-LocalConfig
    $releasePath = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        Resolve-WorkspacePath -Path $OutputPath
    }
    else {
        $latest = Get-ChildItem -LiteralPath $config.DistributionFolder -Filter "*.xlsm" -File |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($null -eq $latest) { throw "검증할 배포본이 없습니다: $($config.DistributionFolder)" }
        $latest.FullName
    }

    if (-not (Test-Path -LiteralPath $releasePath)) { throw "배포본을 찾을 수 없습니다: $releasePath" }

    Write-Host "[1/4] 매크로 차단 상태에서 디스크 잠금을 확인합니다."
    $lockedBefore = Read-ReleaseWorkbookState -WorkbookPath $releasePath -MacrosEnabled $false
    if ($lockedBefore.Marker -ne "RELEASE_SECURITY_V1") { throw "배포 보안 표시자가 없습니다." }
    if ($lockedBefore.VisibleSheets.Count -ne 1 -or $lockedBefore.VisibleSheets[0] -ne "사용안내") {
        throw "디스크 잠금 상태가 아닙니다: $($lockedBefore.VisibleSheets -join ', ')"
    }
    if ($lockedBefore.ExpiryDate -lt [DateTime]::Today) {
        throw "이 자동 검증은 만료 전 배포본만 실행할 수 있습니다. 만료일: $($lockedBefore.ExpiryDate.ToString('yyyy-MM-dd'))"
    }

    Write-Host "[2/4] 매크로 허용 상태에서 정상 잠금 해제를 확인합니다."
    $unlocked = Read-ReleaseWorkbookState -WorkbookPath $releasePath -MacrosEnabled $true
    $businessSheets = @($unlocked.VisibleSheets | Where-Object { $_ -notin @("사용안내", "__RELEASE_SECURITY") })
    if ($businessSheets.Count -lt 1) { throw "정상 업무 시트가 표시되지 않았습니다." }
    if ($unlocked.VisibleSheets -contains "사용안내" -or $unlocked.VisibleSheets -contains "__RELEASE_SECURITY") {
        throw "잠금 해제 후 보안용 시트가 표시되었습니다: $($unlocked.VisibleSheets -join ', ')"
    }
    if (-not $unlocked.Saved) { throw "정상 개봉 직후 통합문서가 불필요한 변경 상태입니다." }

    Write-Host "[3/4] 정상 저장 후 화면 복원을 확인합니다."
    Test-ReleaseSaveCycle -WorkbookPath $releasePath

    Write-Host "[4/4] 파일을 다시 열어 디스크 잠금이 그대로인지 확인합니다."
    $lockedAfter = Read-ReleaseWorkbookState -WorkbookPath $releasePath -MacrosEnabled $false
    if ($lockedAfter.VisibleSheets.Count -ne 1 -or $lockedAfter.VisibleSheets[0] -ne "사용안내") {
        throw "검증 후 디스크 잠금이 유지되지 않았습니다: $($lockedAfter.VisibleSheets -join ', ')"
    }

    Write-Host "배포본 자동 검증 통과: $releasePath" -ForegroundColor Green
    Write-Host "표시 업무 시트 수: $($businessSheets.Count), 만료일: $($lockedAfter.ExpiryDate.ToString('yyyy-MM-dd'))"
}

Push-Location $ScriptRoot
try {
    switch ($Action) {
        "Build" { Build-ReleaseWorkbook }
        "Code" { Show-Code }
        "Status" { Show-Status }
        "SyncDev" { Sync-DevelopmentWorkbook }
        "Validate" { Test-ReleaseWorkbook }
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
