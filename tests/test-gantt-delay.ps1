[CmdletBinding()]
param([string]$WorkbookPath, [string]$PreviewPath, [switch]$UseWorkbookVba)

# 2026-10-21 김좌현 수정 : 개발본 복사본에서 지연 계산/분홍 구간 회귀 검증
# 2026-10-22 김좌현 수정 : 민트 오늘 및 지연 완료 빗금의 저장/재열기 유지 검사
# 2026-10-22 김좌현 수정 : 저장된 개발본/스테이지 검증에서는 소스를 재주입하지 않고 내장 VBA 일치 확인
# 2026-10-22 김좌현 수정 : 셀별 단일 사선과 빈 막대 및 설명 행 연분홍의 저장 유지 검사
# 2026-10-22 김좌현 수정 : 도형/메모/셀 값 없는 배경 삼각형의 저장 유지 검사
# 2026-10-22 김좌현 수정 : 확대된 기본 노란색 배경 모서리의 저장 유지 검사
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('gantt-delay-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath $WorkbookPath -Destination $taskTemp
$taskExcel = $null
$taskBook = $null
function Read-TaskVbaSource([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    try { $source = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
    catch { $source = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
    return ($source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute VB_Name = .*\r?\n', '')
}
function ConvertTo-ComparableVba([string]$Source) {
    $result = [Text.StringBuilder]::new()
    $inString = $false
    $inComment = $false
    for ($index = 0; $index -lt $Source.Length; $index++) {
        $character = $Source[$index]
        if ($inComment) {
            if ($character -eq "`r" -or $character -eq "`n") { $inComment = $false }
            continue
        }
        if ($character -eq '"') {
            [void]$result.Append($character)
            if ($inString -and $index + 1 -lt $Source.Length -and $Source[$index + 1] -eq '"') {
                [void]$result.Append($Source[++$index])
            }
            else { $inString = -not $inString }
        }
        elseif (-not $inString -and $character -eq "'") { $inComment = $true }
        elseif ($inString) { [void]$result.Append($character) }
        elseif (-not [char]::IsWhiteSpace($character)) { [void]$result.Append([char]::ToLowerInvariant($character)) }
    }
    return $result.ToString()
}
try {
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskExcel.AutomationSecurity = 1
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    foreach ($name in @('modGanttDate', 'modGanttSheet', 'modGanttBars')) {
        $code = $taskBook.VBProject.VBComponents.Item($name).CodeModule
        $source = Read-TaskVbaSource (Join-Path $taskRoot "vba-files/Module/$name.bas")
        if ($UseWorkbookVba) {
            $savedSource = [string]$code.Lines(1, $code.CountOfLines)
            if ((ConvertTo-ComparableVba $savedSource) -cne (ConvertTo-ComparableVba $source)) {
                throw "Saved workbook VBA does not match source: $name"
            }
        }
        else {
            $code.DeleteLines(1, $code.CountOfLines)
            $code.AddFromString($source)
        }
    }
    if ($UseWorkbookVba) { Write-Output 'PASS: saved workbook VBA matches source; production modules were not replaced' }
    foreach ($name in @('GanttBarsRegression', 'GanttDelayRegression')) {
        $harness = $taskBook.VBProject.VBComponents.Add(1)
        $harness.Name = $name
        $harness.CodeModule.AddFromString((Read-TaskVbaSource (Join-Path $PSScriptRoot "$name.bas")))
    }
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    foreach ($procedure in @('RunGanttBarsRegression', 'RunGanttDelayRegression')) {
        $result = [string]$taskExcel.Run("'$($taskBook.Name)'!$procedure")
        if (-not $result.StartsWith('PASS:')) { throw $result }
        Write-Output $result
    }
    $completedDelayText = [string]$taskBook.Worksheets.Item('GanttDelayPreview').Range('AF6').Text
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $sheet = $taskBook.Worksheets.Item('GanttDelayPreview')
    $todayColumn = 24 + [int]((Get-Date).Date - [datetime]'2026-10-12').TotalDays
    if ($sheet.Range('AF6').Borders.Item(6).LineStyle -ne -4142 -or
        $sheet.Range('AF6').Borders.Item(5).LineStyle -ne -4142 -or
        $sheet.Range('AF6').Interior.Pattern -ne 4000 -or
        $completedDelayText.Length -ne 0 -or
        [string]$sheet.Range('AF6').Text -cne '' -or
        $sheet.Shapes.Count -ne 0 -or $sheet.Comments.Count -ne 0 -or
        $sheet.Range('AF6').Interior.Gradient.Degree -ne 45 -or
        $sheet.Range('AF6').Interior.Gradient.ColorStops.Count -ne 4 -or
        $sheet.Range('AF6').Interior.Gradient.ColorStops.Item(1).Color -ne 65535 -or
        $sheet.Range('AF6').Interior.Gradient.ColorStops.Item(2).Position -le 0.11 -or
        $sheet.Range('AF6').Interior.Gradient.ColorStops.Item(4).Color -ne 5454628 -or
        $sheet.Cells.Item(3, $todayColumn).Interior.Color -ne 12571693 -or
        $sheet.Cells.Item(8, $todayColumn).Interior.Color -ne 15461375) {
        throw 'Saved native background corner, empty cells or delay formatting changed.'
    }
    Write-Output 'PASS: saved/reopened enlarged yellow background triangle, original completion color, zero shapes/notes/cell values, mint today and pale pink preserved'
    if ($PreviewPath) {
        $taskExcel.Visible = $true
        $sheet = $taskBook.Worksheets.Item('GanttDelayPreview')
        $sheet.Activate()
        $sheet.Range('H3:AP21').CopyPicture(1, 2)
        $chart = $sheet.ChartObjects().Add(0, 0, $sheet.Range('H3:AP21').Width, $sheet.Range('H3:AP21').Height)
        $chart.Activate()
        $chart.Chart.Paste()
        if (-not $chart.Chart.Export($PreviewPath, 'PNG')) { throw 'Preview PNG export failed.' }
        Write-Output ('Preview: ' + $PreviewPath)
    }
}
finally {
    if ($null -ne $taskBook) { $taskBook.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook) }
    if ($null -ne $taskExcel) { $taskExcel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
    Remove-Item -LiteralPath $taskTemp -ErrorAction SilentlyContinue
}
