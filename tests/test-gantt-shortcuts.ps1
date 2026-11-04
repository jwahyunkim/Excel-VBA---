[CmdletBinding()]
param([string]$WorkbookPath, [switch]$Baseline)

# 2026-11-04 김좌현 수정 : 실제 Excel 이벤트로 닫기 취소·비활성 저장·창 전환·다른 이름 저장 단축키 회귀 검증
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('gantt-shortcuts-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $taskTemp
$excel = $null
$observer = $null
$bookA = $null
$bookB = $null
$failures = [Collections.Generic.List[string]]::new()

function Read-TaskVba([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    try { $value = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
    catch { $value = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
    return $value.TrimStart([char]0xFEFF)
}
function Check([bool]$ok, [string]$label) {
    if ($ok) { Write-Output ('PASS: ' + $label); return }
    if ($Baseline) { $failures.Add($label); Write-Output ('BASELINE FAILURE: ' + $label); return }
    throw $label
}
function Get-Binding([int]$row = 1) {
    $value = [string]$observer.Worksheets.Item(1).Cells.Item($row, 1).Value2
    if ($value.StartsWith('binding:')) { return $value.Substring(8) }
    return $value
}
function Get-ExpectedBinding($book, [string]$macro = '칸트차트_새로고침') {
    return "'" + $book.Name.Replace("'", "''") + "'!" + $macro
}
function Clear-TestBinding {
    $excel.OnKey('^+r')
    $observer.Worksheets.Item(1).Cells.Item(1, 1).Value2 = ''
}
function Prepare-TestBook([string]$path) {
    $book = $excel.Workbooks.Open($path, 0, $false)
    if ($Baseline) {
        $classSource = (& git -C $taskRoot show 'HEAD:vba-files/Class/현재_통합_문서.cls') -join "`r`n"
        if ($LASTEXITCODE -ne 0) { throw 'Cannot load baseline workbook class.' }
    } else {
        $classSource = Read-TaskVba (Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls')
    }
    $classSource = $classSource.Substring($classSource.IndexOf('Option Explicit'))
    # Record assignments while forwarding them to the real Application.OnKey API.
    $classSource = $classSource.Replace('Application.OnKey ', 'RecordTestShortcut ')
    $cancelBody = @'
    If ShortcutRegression.CancelNextClose Then
        Cancel = True
        ShortcutRegression.CancelNextClose = False
    End If
'@
    $closePattern = '(?s)(Private Sub Workbook_BeforeClose\(Cancel As Boolean\).*?)(End Sub)'
    if ([regex]::IsMatch($classSource, $closePattern)) {
        $classSource = [regex]::Replace($classSource, $closePattern, { param($match) $match.Groups[1].Value + $cancelBody + "`r`nEnd Sub" })
    } else {
        $classSource += "`r`nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)`r`n" + $cancelBody + "`r`nEnd Sub`r`n"
    }
    $observerName = $observer.Name.Replace('"', '""')
    $classSource += @"

Private Sub RecordTestShortcut(ByVal key As String, Optional ByVal procedure As Variant)
    Dim row As Long
    row = IIf(key = "^+r", 1, 2)
    If IsMissing(procedure) Then
        Application.Workbooks("$observerName").Worksheets(1).Cells(row, 1).Value2 = ""
        Application.OnKey key
    Else
        Application.Workbooks("$observerName").Worksheets(1).Cells(row, 1).Value2 = "binding:" & CStr(procedure)
        Application.OnKey key, CStr(procedure)
    End If
End Sub

Public Sub RegressionRegisterShortcuts()
    RegisterGanttShortcuts
End Sub
"@
    $code = $book.VBProject.VBComponents.Item([string]$book.CodeName).CodeModule
    $code.DeleteLines(1, $code.CountOfLines)
    $code.AddFromString($classSource)
    $main = Read-TaskVba (Join-Path $taskRoot 'vba-files/Module/modGanttMain.bas')
    $main = $main -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    $entryPattern = '(?s)(Public Sub 칸트차트_새로고침\(\)\s*)RefreshGanttSheet(\s*End Sub)'
    if (-not [regex]::IsMatch($main, $entryPattern)) { throw 'Cannot instrument refresh entry point.' }
    $main = [regex]::Replace($main, $entryPattern, '${1}ShortcutRegression.RefreshCalls = ShortcutRegression.RefreshCalls + 1' + "`r`n    RefreshGanttSheet False" + '${2}')
    $code = $book.VBProject.VBComponents.Item('modGanttMain').CodeModule
    $code.DeleteLines(1, $code.CountOfLines)
    $code.AddFromString($main)
    $harness = $book.VBProject.VBComponents.Add(1)
    $harness.Name = 'ShortcutRegression'
    $harness.CodeModule.AddFromString(@'
Option Explicit
Public CancelNextClose As Boolean
Public RefreshCalls As Long
Public Sub ArmCloseCancellation()
    CancelNextClose = True
End Sub
Public Function RefreshCount() As Long
    RefreshCount = RefreshCalls
End Function
Public Sub CreateFixture()
    Dim ws As Worksheet, events As Boolean
    events = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "ShortcutTask"
    SetupDataHeaders ws
    ws.Cells(5, COL_LEVEL).Value2 = 1
    ws.Cells(5, COL_REQUEST).Value2 = "Shortcut regression task"
    ws.Cells(5, COL_PLAN_START).Value = DateSerial(2026, 11, 2)
    ws.Cells(5, COL_PLAN_END).Value = DateSerial(2026, 11, 6)
    ws.Cells(5, COL_ACTUAL_START).Value = DateSerial(2026, 11, 2)
    ws.Cells(5, COL_PROGRESS).Value2 = 0.5
    ws.Activate
    RefreshGanttSheet False
    RefreshCalls = 0
    Application.EnableEvents = events
    Exit Sub
Failed:
    Application.EnableEvents = events
    Err.Raise Err.Number, "ShortcutRegression.CreateFixture", Err.Description
End Sub
'@)
    $book.Save()
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
}

try {
    $pathA = Join-Path $taskTemp 'Shortcut-A.xlsm'
    $pathB = Join-Path $taskTemp 'Shortcut-B.xlsm'
    Copy-Item -LiteralPath $WorkbookPath -Destination $pathA
    Copy-Item -LiteralPath $WorkbookPath -Destination $pathB
    $excel = New-Object -ComObject Excel.Application
    # Hidden Excel omits close-time deactivation. Exercise real window lifecycle.
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $observer = $excel.Workbooks.Add()
    Prepare-TestBook $pathA
    Prepare-TestBook $pathB
    $excel.EnableEvents = $true
    $bookA = $excel.Workbooks.Open($pathA, 0, $false)
    Write-Verbose ('Initial binding=' + (Get-Binding) + '; expected=' + (Get-ExpectedBinding $bookA) + '; events=' + $excel.EnableEvents + '; active=' + $excel.ActiveWorkbook.Name)
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'opening workbook registers Ctrl+Shift+R'
    Check ((Get-Binding 2) -eq (Get-ExpectedBinding $bookA '칸트차트_완료숨김')) 'Ctrl+Shift+H registration preserved'

    $null = $excel.Run("'$($bookA.Name)'!ArmCloseCancellation")
    $bookA.Close($false)
    Check ($excel.ActiveWorkbook.Name -eq $bookA.Name) 'cancelled close leaves workbook active'
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'cancelled close retains refresh shortcut'
    $null = $excel.Run("'$($bookA.Name)'!$($bookA.CodeName).RegressionRegisterShortcuts")

    $bookB = $excel.Workbooks.Open($pathB, 0, $false)
    $bookA.Activate()
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'switching back restores active workbook shortcut'
    $bookB.Save()
    Check ($excel.ActiveWorkbook.Name -eq $bookA.Name) 'saving inactive workbook keeps active workbook'
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'inactive workbook save cannot overwrite active shortcut'
    if ($Baseline) {
        if ($failures.Count -lt 2) { throw 'Expected baseline failures were not reproduced.' }
        Write-Output ('BASELINE REPRODUCED: ' + $failures.Count + ' shortcut lifecycle failures')
        return
    }

    $bookB.Close($false)
    $bookB = $null
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'closing inactive workbook preserves active shortcut'
    $null = $excel.Run("'$($bookA.Name)'!CreateFixture")
    $bookA.Worksheets.Item('ShortcutTask').Activate()
    Clear-TestBinding
    [void]$bookA.Worksheets.Item('ShortcutTask').Range('R6').Select()
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'cell selection restores an overwritten shortcut'
    Clear-TestBinding
    $bookA.Worksheets.Item('config').Activate()
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'sheet activation restores an overwritten shortcut'
    $bookA.Worksheets.Item('ShortcutTask').Activate()

    $firstWindow = $bookA.Windows.Item(1)
    $extraWindow = $bookA.NewWindow()
    Clear-TestBinding
    [void]$firstWindow.Activate()
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'same-workbook window activation restores shortcut'
    [void]$extraWindow.Close()
    $observer.Activate()
    Check ((Get-Binding) -eq '') 'leaving task workbook clears its shortcut'
    $bookA.Activate()
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'returning from ordinary workbook restores shortcut'

    $renamedPath = Join-Path $taskTemp "업무 'Shortcut Renamed.xlsm"
    $bookA.SaveAs($renamedPath, 52)
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'Save As updates escaped Korean workbook name'
    $bookA.Worksheets.Item('ShortcutTask').Activate()
    $null = $excel.Run((Get-Binding))
    Check ([int]$excel.Run("'" + $bookA.Name.Replace("'", "''") + "'!RefreshCount") -eq 1) 'registered refresh target executes common refresh exactly once'
    Check ([double]$bookA.Worksheets.Item('ShortcutTask').Range('AF5').Value2 -gt 0) 'registered refresh calculates task results'
    Write-Verbose ('Before final save events=' + $excel.EnableEvents)
    $bookA.Save()
    Write-Verbose ('Before final close events=' + $excel.EnableEvents)
    $bookA.Close($false)
    $bookA = $null
    Write-Verbose ('After close binding=' + (Get-Binding) + '; active=' + $excel.ActiveWorkbook.Name + '; workbooks=' + $excel.Workbooks.Count)
    Check ((Get-Binding) -eq '' -and (Get-Binding 2) -eq '') 'successful close releases both shortcuts'
    $bookA = $excel.Workbooks.Open($renamedPath, 0, $false)
    Check ((Get-Binding) -eq (Get-ExpectedBinding $bookA)) 'saved workbook reopens with current shortcut'
    Write-Output 'PASS: shortcut lifecycle and registered refresh dispatch in isolated Excel instance'
} finally {
    if ($excel) { try { $excel.EnableEvents = $false } catch {}; try { $excel.OnKey('^+r'); $excel.OnKey('^+h') } catch {} }
    foreach ($book in @($bookB, $bookA, $observer)) { if ($book) { try { $book.Close($false) } catch {} } }
    if ($excel) { try { $excel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    # Only this run's explicitly named temporary workbook files are removed.
    foreach ($name in @('Shortcut-A.xlsm', 'Shortcut-B.xlsm', "업무 'Shortcut Renamed.xlsm")) {
        $path = Join-Path $taskTemp $name
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp }
}
