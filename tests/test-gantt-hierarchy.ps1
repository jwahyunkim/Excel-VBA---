[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath, [switch]$ViewOnly, [switch]$SelectedColumnsOnly)

$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('gantt-hierarchy-' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item -LiteralPath $WorkbookPath -Destination $taskTemp
$taskExcel = $null
$taskBook = $null
function Open-TaskTestWorkbook($Excel, [string]$Path) {
    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        try { return $Excel.Workbooks.Open($Path, 0, $false) }
        catch {
            $openError = $_.Exception
            while ($null -ne $openError -and $openError.HResult -ne -2146777998) {
                $openError = $openError.InnerException
            }
            if ($null -eq $openError -or $attempt -eq 4) { throw }
            # Excel may still be completing Close; retry only 0x800AC472.
            Start-Sleep -Milliseconds 250
        }
    }
}
function Assert-TaskExpandedView($Window, [string]$Context, [int]$FirstColumn = 1) {
    if ($Window.FreezePanes -or $Window.SplitColumn -ne 0 -or $Window.SplitRow -ne 0 -or
        $Window.Panes.Count -ne 1 -or $Window.ScrollColumn -ne $FirstColumn) {
        throw "$Context did not restore the expanded view without splits at column $FirstColumn"
    }
}
function Ensure-TaskWorkbookEventCall($Code, [string]$Name, [string]$Signature,
                                      [string]$Statement, [string]$BeforeStatement = '') {
    try { $start = $Code.ProcStartLine($Name, 0) }
    catch {
        $Code.AddFromString($Signature + "`r`n    " + $Statement + "`r`nEnd Sub")
        return
    }
    $count = $Code.ProcCountLines($Name, 0)
    if (([string]$Code.Lines($start, $count)).Contains($Statement)) { return }
    for ($line = $start; $line -lt $start + $count; $line++) {
        $current = ([string]$Code.Lines($line, 1)).Trim()
        if (($BeforeStatement -and $current -eq $BeforeStatement) -or
            (-not $BeforeStatement -and $current -eq 'End Sub')) {
            $Code.InsertLines($line, '    ' + $Statement)
            return
        }
    }
    throw "Could not patch workbook event $Name"
}
function Assert-TaskRequestFreeze($Window, [string]$Context, [int]$FirstDetailColumn = 9, [int]$AnchorColumn = 8) {
    if (-not $Window.FreezePanes -or $Window.SplitColumn -ne 1 -or $Window.SplitRow -ne 4 -or
        $Window.Panes.Count -ne 4 -or $Window.Panes.Item(1).ScrollColumn -ne $AnchorColumn -or
        $Window.Panes.Item(3).ScrollColumn -ne $AnchorColumn) {
        throw "$Context did not retain the collapsed display start $AnchorColumn and four-row header freeze (fold starts at column $FirstDetailColumn)"
    }
}
function Assert-TaskDetailOutline($Sheet, [int]$FirstColumn, [int]$LastColumn, [string]$Context) {
    for ($column = 9; $column -le [Math]::Max(24, $LastColumn + 1); $column++) {
        $expectedLevel = if ($column -ge $FirstColumn -and $column -le $LastColumn) { 2 } else { 1 }
        if ($Sheet.Columns.Item($column).OutlineLevel -ne $expectedLevel) {
            throw "$Context has an incorrect outline level at column $column"
        }
    }
    if ($Sheet.Columns.Item(8).OutlineLevel -ne 1 -or
        $Sheet.Rows.Item(53).OutlineLevel -ne 2 -or $Sheet.Rows.Item(61).OutlineLevel -ne 3) {
        throw "$Context changed the request/Gantt boundaries or row outlines"
    }
}
function Assert-TaskSelectedColumns($Window, $Sheet, [int[]]$Selected, [string]$Context) {
    if (-not $Window.FreezePanes -or $Window.SplitRow -ne 4 -or $Window.SplitColumn -le 0) {
        $states = 1..24 | ForEach-Object { "${_}:$($Sheet.Columns.Item($_).Hidden)/$($Sheet.Columns.Item($_).OutlineLevel)" }
        throw "$Context did not freeze: freeze=$($Window.FreezePanes), split=$($Window.SplitColumn)/$($Window.SplitRow); $($states -join ' ')"
    }
    $frozenRange = $Window.Panes.Item(1).VisibleRange
    $frozenColumns = @()
    for ($column = $frozenRange.Column; $column -lt $frozenRange.Column + $frozenRange.Columns.Count; $column++) {
        if (-not $Sheet.Columns.Item($column).Hidden) { $frozenColumns += $column }
    }
    if (($frozenColumns -join ',') -ne ($Selected -join ',')) {
        throw "$Context has incorrect visible frozen columns: $($frozenColumns -join ',') (split=$($Window.SplitColumn))"
    }
}
try {
    $taskExcel = New-Object -ComObject Excel.Application
    # CommandBars.OnUpdate and freeze setup require a visible, normal Excel window.
    $taskExcel.WindowState = -4143
    $taskExcel.Visible = $true
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskExcel.AutomationSecurity = 1
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskInitialWindow = $taskBook.Windows.Item(1)
    $taskOriginalScrollRow = [int]$taskInitialWindow.Panes.Item($taskInitialWindow.Panes.Count).ScrollRow
    $taskComponents = $taskBook.VBProject.VBComponents
    foreach ($taskModuleName in @('modGanttBars', 'modGanttHierarchy')) {
        try { $component = $taskComponents.Item($taskModuleName) }
        catch { $component = $taskComponents.Add(1); $component.Name = $taskModuleName }
        $source = [IO.File]::ReadAllText((Join-Path $taskRoot ('vba-files/Module/' + $taskModuleName + '.bas')))
        $source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        $code = $component.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }

    # Retain the workbook's config/holiday implementation; patch only new settings.
    $code = $taskComponents.Item('modGanttConfig').CodeModule
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'vba-files/Module/modGanttConfig.bas'))
    $foldDeclarations = [regex]::Matches($source, '(?m)^Public Const (TASK_FOLD_\w+) As String = [^\r\n]+')
    foreach ($declaration in $foldDeclarations) {
        $declarationFound = $false
        for ($line = 1; $line -le $code.CountOfLines; $line++) {
            $current = ([string]$code.Lines($line, 1)).Trim()
            if ($current.StartsWith('Public Const ' + $declaration.Groups[1].Value + ' ')) {
                if ($current -ne $declaration.Value) { $code.ReplaceLine($line, $declaration.Value) }
                $declarationFound = $true
                break
            }
        }
        if (-not $declarationFound) {
            for ($line = 1; $line -le $code.CountOfLines; $line++) {
                if (([string]$code.Lines($line, 1)).Trim() -eq 'Option Explicit') {
                    $code.InsertLines($line + 1, $declaration.Value)
                    break
                }
            }
        }
    }
    $code = $taskComponents.Item('modHoliday').CodeModule
    $start = $code.ProcStartLine('EnsureConfigSheet', 0)
    $count = $code.ProcCountLines('EnsureConfigSheet', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('EnsureTaskDetailFoldConfig ws')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'End Sub') {
                $code.InsertLines($line, '    EnsureTaskDetailFoldConfig ws')
                break
            }
        }
    }

    try { $component = $taskComponents.Item('CGanttOutlineView') }
    catch { $component = $taskComponents.Add(2); $component.Name = 'CGanttOutlineView' }
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'vba-files/Class/CGanttOutlineView.cls'))
    $classCodeStart = $source.IndexOf('Option Explicit')
    if ($classCodeStart -lt 0) { throw 'CGanttOutlineView is missing its VBA code section' }
    $source = $source.Substring($classCodeStart)
    $code = $component.CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($source)

    # Patch only the changed routines; retain the existing workbook's VBA text
    # and unrelated modules, including modules with legacy Korean encodings.
    $code = $taskComponents.Item('modGanttSheet').CodeModule
    $start = $code.ProcStartLine('ClearGanttArea', 0)
    $count = $code.ProcCountLines('ClearGanttArea', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('ClearTaskBarVisuals ws')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim().StartsWith('startCol =')) {
                $code.InsertLines($line, '    ClearTaskBarVisuals ws')
                break
            }
        }
    }
    for ($line = $code.CountOfLines; $line -ge 1; $line--) {
        $text = [string]$code.Lines($line, 1)
        if ($text.Contains('BuildTaskDisplayText(baseTaskText, levelValue)')) {
            $code.ReplaceLine($line, $text.Replace('BuildTaskDisplayText(baseTaskText, levelValue)', 'baseTaskText'))
        }
    }
    $start = $code.ProcStartLine('FormatBaseArea', 0)
    $count = $code.ProcCountLines('FormatBaseArea', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('ApplyTaskHierarchyAppearance ws, lastRow, startCol, endCol')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'ActiveWindow.DisplayGridlines = False') {
                $code.InsertLines($line, '    ApplyTaskHierarchyAppearance ws, lastRow, startCol, endCol')
                break
            }
        }
    }
    $count = $code.ProcCountLines('FormatBaseArea', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('Dim detailColumnHidden As Variant')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'Dim baseTaskText As String') {
                $code.InsertLines($line + 1, '    Dim detailColumnHidden As Variant')
                break
            }
        }
    }
    $count = $code.ProcCountLines('FormatBaseArea', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('detailColumnHidden = CaptureTaskDetailColumnHidden(ws)')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'dataLastCol = COL_STATUS') {
                $code.InsertLines($line, '    detailColumnHidden = CaptureTaskDetailColumnHidden(ws)')
                break
            }
        }
    }
    $count = $code.ProcCountLines('FormatBaseArea', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('RestoreTaskDetailColumnHidden ws, detailColumnHidden')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'ApplyTaskHierarchyAppearance ws, lastRow, startCol, endCol') {
                $code.InsertLines($line + 1, '    RestoreTaskDetailColumnHidden ws, detailColumnHidden')
                break
            }
        }
    }
    $start = $code.ProcStartLine('ApplyCalculatedColumnsProtection', 0)
    $count = $code.ProcCountLines('ApplyCalculatedColumnsProtection', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('ws.EnableOutlining = True')) {
        for ($line = $start + $count - 1; $line -ge $start; $line--) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'End Sub') {
                $code.InsertLines($line, '    ws.EnableOutlining = True')
                break
            }
        }
    }
    $code = $taskComponents.Item('modGanttMain').CodeModule
    $start = $code.ProcStartLine('RefreshGanttSheet', 0)
    $count = $code.ProcCountLines('RefreshGanttSheet', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('Dim detailColumnHidden As Variant')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'Dim errorText As String') {
                $code.InsertLines($line + 1, '    Dim detailColumnHidden As Variant')
                break
            }
        }
    }
    $count = $code.ProcCountLines('RefreshGanttSheet', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('Dim detailScrollColumn As Long, detailScrollRow As Long')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'Dim detailColumnHidden As Variant') {
                $code.InsertLines($line + 1, '    Dim detailScrollColumn As Long, detailScrollRow As Long')
                break
            }
        }
    }
    $count = $code.ProcCountLines('RefreshGanttSheet', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('detailColumnHidden = CaptureTaskDetailColumnHidden(ws)')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'Set ws = ActiveSheet') {
                $code.InsertLines($line + 1, @"
    detailColumnHidden = CaptureTaskDetailColumnHidden(ws)
    If Not ActiveWindow Is Nothing Then
        detailScrollColumn = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn
        detailScrollRow = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow
    End If
"@)
                break
            }
        }
    }
    $count = $code.ProcCountLines('RefreshGanttSheet', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('RestoreTaskDetailColumnHidden ws, detailColumnHidden')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'ApplyDisplayTaskRowFilter ws, lastRow, chartStartDate, chartEndDate') {
                $code.InsertLines($line + 1, @"
    RestoreTaskDetailColumnHidden ws, detailColumnHidden
    If detailScrollColumn > 0 And Not ActiveWindow Is Nothing Then
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = detailScrollColumn
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow = detailScrollRow
    End If
"@)
                break
            }
        }
    }
    $code = $taskComponents.Item([string]$taskBook.CodeName).CodeModule
    Ensure-TaskWorkbookEventCall $code 'Workbook_Open' 'Private Sub Workbook_Open()' 'InitializeTaskHierarchyOutlines'
    Ensure-TaskWorkbookEventCall $code 'Workbook_Open' 'Private Sub Workbook_Open()' 'StartTaskOutlineViewMonitor'
    $start = $code.ProcStartLine('Workbook_SheetActivate', 0)
    $count = $code.ProcCountLines('Workbook_SheetActivate', 0)
    for ($line = $start; $line -lt $start + $count; $line++) {
        if (([string]$code.Lines($line, 1)).Trim() -eq 'EnsureTaskRequestFreeze Sh') {
            $code.ReplaceLine($line, '    StartTaskOutlineViewMonitor')
        }
    }
    Ensure-TaskWorkbookEventCall $code 'Workbook_SheetActivate' 'Private Sub Workbook_SheetActivate(ByVal Sh As Object)' 'StartTaskOutlineViewMonitor' 'If Sh.Name <> WEEKLY_REPORT_CONFIG_SHEET_NAME Then Exit Sub'
    Ensure-TaskWorkbookEventCall $code 'Workbook_Activate' 'Private Sub Workbook_Activate()' 'StartTaskOutlineViewMonitor'
    Ensure-TaskWorkbookEventCall $code 'Workbook_Deactivate' 'Private Sub Workbook_Deactivate()' 'StopTaskOutlineViewMonitor'
    Ensure-TaskWorkbookEventCall $code 'Workbook_BeforeClose' 'Private Sub Workbook_BeforeClose(Cancel As Boolean)' 'StopTaskOutlineViewMonitor'
    Ensure-TaskWorkbookEventCall $code 'Workbook_BeforeSave' 'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)' 'StopTaskOutlineViewMonitor' 'PrepareReleaseForSave SaveAsUI, Cancel'
    Ensure-TaskWorkbookEventCall $code 'Workbook_BeforeSave' 'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)' 'If Cancel Then StartTaskOutlineViewMonitor'
    Ensure-TaskWorkbookEventCall $code 'Workbook_AfterSave' 'Private Sub Workbook_AfterSave(ByVal Success As Boolean)' 'StartTaskOutlineViewMonitor'
    Ensure-TaskWorkbookEventCall $code 'Workbook_SheetSelectionChange' 'Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)' 'StartTaskOutlineViewMonitor'

    $start = $code.ProcStartLine('Workbook_SheetChange', 0)
    $count = $code.ProcCountLines('Workbook_SheetChange', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('ApplyTaskDetailColumnSettings')) {
        for ($line = $start; $line -lt $start + $count; $line++) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'If Sh.Name = CONFIG_SHEET_NAME And _') {
                $code.InsertLines($line, @"
        If Sh.Name = CONFIG_SHEET_NAME Then
            If Not Intersect(Target, Sh.Range(TASK_FOLD_SETTING_VALUE_RANGE)) Is Nothing Then
                Application.EnableEvents = False
                ApplyTaskDetailColumnSettings
                GoTo SafeExit
            End If
        End If
"@)
                break
            }
        }
    }

    $harness = $taskComponents.Add(1)
    $harness.Name = 'GanttHierarchyRegression'
    $harness.CodeModule.AddFromString(@'
Option Explicit
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 920, "GanttHierarchyRegression", label
End Sub
Public Function PrepareTaskFoldRegression() As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    EnsureTaskDetailFoldConfig ws
    ws.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    ws.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "H"
    PrepareTaskFoldRegression = "PASS"
    Exit Function
Failed:
    PrepareTaskFoldRegression = "FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function RunGanttHierarchyRegression() As String
    Dim ws As Worksheet, original As Worksheet, r As Long, c As Long
    Dim levels As Variant, dataBefore As Variant, hiddenBefore() As Boolean
    Dim holidays As Object, workdays As Object, firstDate As Date, lastDate As Date
    Dim lastRow As Long, lastCol As Long, taskCol As Long
    Dim existingShape As Shape, unrelatedShapeCount As Long
    Dim edge As Variant, edgeColor As Long
    Dim foldConfig As Worksheet, foldFirst As Long, foldLast As Long
    On Error GoTo Failed
    Application.EnableEvents = False
    Set original = ThisWorkbook.Worksheets("Sheet1")
    Set foldConfig = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    EnsureTaskDetailFoldConfig foldConfig
    foldConfig.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    foldConfig.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    foldConfig.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "H"
    Check GetTaskDetailColumnBounds(original, foldFirst, foldLast), "default fold bounds invalid"
    Check foldFirst = 9 And foldLast = 23, "default fold bounds changed"
    foldConfig.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "W"
    foldConfig.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "I"
    Check Not GetTaskDetailColumnBounds(original, foldFirst, foldLast), "reversed fold bounds accepted"
    foldConfig.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "L1"
    foldConfig.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    Check Not GetTaskDetailColumnBounds(original, foldFirst, foldLast), "cell reference accepted as a fold column"
    foldConfig.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "XFE"
    Check Not GetTaskDetailColumnBounds(original, foldFirst, foldLast), "out-of-range fold column accepted"
    foldConfig.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    EnsureTaskDetailFoldConfig foldConfig
    Set ws = ThisWorkbook.Worksheets.Add
    levels = Array(1, 2, 3, 3, 2, 1, 1, 3)
    For r = 5 To 12
        ws.Cells(r, COL_LEVEL).Value = levels(r - 5)
        ws.Cells(r, COL_REQUEST).Value = "Hierarchy test " & r
    Next r
    ws.Range("X5:X12").ClearFormats
    ws.Range("Y5").Interior.Color = RGB(180, 180, 180)
    ws.Range("Z5").Interior.Color = RGB(102, 255, 51)
    ws.Range("AA6").Interior.Color = RGB(245, 245, 245)
    ws.Range("AB6").Interior.Color = RGB(255, 199, 206)
    ws.Range("B9:V9").Interior.Color = RGB(255, 235, 235)
    ws.Cells(9, COL_STATUS).Value = STATUS_DELAY
    ApplyTaskHierarchyAppearance ws, 12, 24, 28
    ApplyTaskHierarchyAppearance ws, 12, 24, 28
    Check ws.Rows(5).OutlineLevel = 1, "root included in its own outline"
    Check ws.Rows(6).OutlineLevel = 2, "child outline"
    Check ws.Rows(7).OutlineLevel = 3 And ws.Rows(8).OutlineLevel = 3, "grandchild outline"
    Check ws.Rows(9).OutlineLevel = 2, "sibling group boundary"
    Check ws.Rows(10).OutlineLevel = 1, "next root grouped"
    Check ws.Rows(12).OutlineLevel = 2, "skipped level group"
    For c = 9 To 23
        Check ws.Columns(c).OutlineLevel = 2, "detail column outline depth " & c
    Next c
    Check ws.Columns("H").OutlineLevel = 1 And ws.Columns("X").OutlineLevel = 1, "request or Gantt grouped"
    Check ws.Outline.SummaryColumn = xlRight, "column summary position"
    Check ws.Range("H5").Font.Bold And ws.Range("H6").Font.Bold, "summary emphasis"
    Check Not ws.Range("H7").Font.Bold, "leaf emphasis"
    Check ws.Range("H7").IndentLevel = 2 And ws.Range("I7").IndentLevel = 2, "content indentation"
    Check ws.Range("Y5").Interior.Color = RGB(180, 180, 180), "plan bar overwritten"
    Check ws.Range("Z5").Interior.Color = RGB(102, 255, 51), "actual bar overwritten"
    Check ws.Range("AA6").Interior.Color = RGB(245, 245, 245), "holiday overwritten"
    Check ws.Range("AB6").Interior.Color = RGB(255, 199, 206), "today overwritten"
    Check ws.Range("H9").Interior.Color = RGB(255, 235, 235), "warning overwritten"
    ApplyCalculatedColumnsProtection ws, 12
    Check ws.EnableOutlining And ws.ProtectContents, "protected outline disabled"
    ws.Columns("X").ShowDetail = False
    For c = 9 To 23
        Check ws.Columns(c).Hidden, "protected column collapse failed " & c
    Next c
    Check Not ws.Columns("H").Hidden And Not ws.Columns("X").Hidden, "column collapse hid request or Gantt"
    Check Not ws.Rows(6).Hidden And ws.Rows(7).OutlineLevel = 3, "column collapse changed row outline"
    ws.Outline.ShowLevels 1
    Check ws.Rows(6).Hidden And Not ws.Rows(5).Hidden, "collapse failed"
    ws.Outline.ShowLevels 3
    Check Not ws.Rows(7).Hidden, "expand failed"
    For c = 9 To 23
        Check ws.Columns(c).Hidden, "row outline expansion changed column collapse " & c
    Next c
    ws.Columns("X").ShowDetail = True
    For c = 9 To 23
        Check Not ws.Columns(c).Hidden, "protected column expansion failed " & c
    Next c
    UnprotectTaskSheet ws
    ws.Cells(6, COL_LEVEL).Value = 1
    RebuildTaskHierarchyOutline ws, 12
    Check ws.Rows(6).OutlineLevel = 1 And ws.Rows(7).OutlineLevel = 2, "changed hierarchy retained old groups"
    ws.Delete
    Set ws = Nothing

    original.Activate
    For Each existingShape In original.Shapes
        If Left$(existingShape.Name, 6) <> "gtBar_" And Left$(existingShape.Name, 6) <> "gtKey_" Then
            unrelatedShapeCount = unrelatedShapeCount + 1
        End If
    Next existingShape
    lastRow = GetLastDataRow(original)
    dataBefore = original.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_STATUS & lastRow).Value2
    ReDim hiddenBefore(DATA_START_ROW To lastRow)
    For r = DATA_START_ROW To lastRow
        hiddenBefore(r) = original.Rows(r).Hidden
    Next r
    UnprotectTaskSheet original
    LoadHolidaySettings holidays, workdays
    taskCol = original.Range(COL_GANTT_START & "1").Column
    lastCol = original.Cells(GANTT_HEADER_ROW_DATE, original.Columns.Count).End(xlToLeft).Column
    firstDate = original.Cells(GANTT_HEADER_ROW_DATE, taskCol).Value
    lastDate = original.Cells(GANTT_HEADER_ROW_DATE, lastCol).Value
    original.Outline.ShowLevels ColumnLevels:=2
    FormatBaseArea original, lastRow, firstDate, lastDate, holidays, workdays
    Check Not ActiveWindow.FreezePanes And ActiveWindow.SplitColumn = 0 And ActiveWindow.SplitRow = 0, "expanded task view retained freeze/splits"
    Check ActiveWindow.ScrollColumn = 1, "expanded task view does not start at A"
    For r = DATA_START_ROW To lastRow
        For c = 1 To UBound(dataBefore, 2)
            If c + 2 <> original.Range(COL_REQUEST & "1").Column Then
                Check CStr(original.Cells(r, c + 2).Value2) = CStr(dataBefore(r - DATA_START_ROW + 1, c)), _
                    "input changed row " & r & " col " & c + 2
            Else
                Check LTrim$(CStr(original.Cells(r, c + 2).Value2)) = LTrim$(CStr(dataBefore(r - DATA_START_ROW + 1, c))), _
                    "request changed row " & r
            End If
        Next c
        original.Rows(r).Hidden = hiddenBefore(r)
    Next r
    ApplyCalculatedColumnsProtection original, lastRow
    Check original.Range("H52").Font.Bold, "52 root emphasis"
    Check original.Rows(53).OutlineLevel = 2, "52 group children"
    Check original.Rows(61).OutlineLevel = 3, "59 group grandchildren"
    Check original.Range("AL52").Interior.Color = RGB(36, 59, 83), "52 summary cell"
    Check original.Range("AL53").Interior.Color = RGB(52, 120, 212), "53 child cell"
    Check original.Range("AL61").Interior.Color = RGB(0, 140, 134), "61 detail cell"
    Check original.Range("AJ52").Interior.Color = RGB(203, 213, 225), "52 plan cell"
    Check Application.CountA(original.Range("X5:BA" & lastRow)) = 0, "timeline captions remain"
    Check Application.CountA(original.Range("L3:W3")) = 0, "legend remains"
    Check original.Range("AL52").Borders(xlEdgeRight).Color = RGB(22, 163, 74), "52 right status"
    Check original.Range("AL52").Borders(xlEdgeBottom).Color <> RGB(22, 163, 74), "52 bottom status remains"
    Check original.Range("AJ59").Borders(xlEdgeTop).Color = RGB(99, 102, 241), "combined group top"
    Check original.Range("AK59").Borders(xlEdgeTop).Color = RGB(99, 102, 241), "combined frame bridges plan/actual gap"
    Check original.Range("AK74").Borders(xlEdgeBottom).Color = RGB(99, 102, 241), "combined group bottom bridges gap"
    Check original.Range("AJ64").Borders(xlEdgeLeft).Color = RGB(99, 102, 241), "combined group left"
    For r = 60 To 73
        Check original.Cells(r, "AK").Borders(xlEdgeTop).Color <> RGB(99, 102, 241), "internal child top frame remains row " & r
        Check original.Cells(r, "AK").Borders(xlEdgeBottom).Color <> RGB(99, 102, 241), "internal child bottom frame remains row " & r
        Check original.Cells(r, "AL").Borders(xlEdgeRight).Color = RGB(22, 163, 74), "descendant status right lost row " & r
    Next r
    Check original.Range("AL61").Borders(xlEdgeLeft).Color <> RGB(99, 102, 241), "separate actual outline remains"
    ' The removed table boxes must not return in task columns or unused dates.
    For r = 52 To 74
        For c = 2 To lastCol
            If c < taskCol Or c <= original.Range("AH1").Column Or c >= original.Range("AN1").Column Then
                For Each edge In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
                    edgeColor = original.Cells(r, c).Borders(edge).Color
                    Check edgeColor <> RGB(124, 58, 237) And edgeColor <> RGB(6, 182, 212) And _
                          edgeColor <> RGB(13, 148, 136) And edgeColor <> RGB(99, 102, 241), "bar outline outside its dates row " & r & " col " & c
                Next edge
            End If
        Next c
    Next r
    Check original.Shapes.Count = unrelatedShapeCount, "new drawing objects remain or existing buttons lost"
    RunGanttHierarchyRegression = "PASS: protected column collapse/expand, independent row outlines, one frame per level-1 task, right status, no legend/captions and input preservation"
    Exit Function
Failed:
    RunGanttHierarchyRegression = "FAIL: " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
End Function
'@)
    $barRegression = $taskComponents.Add(1)
    $barRegression.Name = 'GanttBarsRegression'
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'tests/GanttBarsRegression.bas'))
    $source = $source -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
    $barRegression.CodeModule.AddFromString($source)
    # Reload the edited VBA before execution so Excel rebuilds its compile cache.
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('GanttHierarchyRegression')
    $barRegression = $taskComponents.Item('GanttBarsRegression')
    $prepareResult = [string]$taskExcel.Run("'$($taskBook.Name)'!PrepareTaskFoldRegression")
    Write-Output ('Fold settings preparation: ' + $prepareResult)
    if ($prepareResult -ne 'PASS') { throw $prepareResult }
    if ($ViewOnly -or $SelectedColumnsOnly) {
        Write-Output 'Running native view regression; previously verified bar/hierarchy synthetic cases skipped'
    }
    else {
        Write-Output 'Running native hierarchy and bar regression'
        $barResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunGanttBarsRegression")
        if (-not $barResult.StartsWith('PASS:')) { throw $barResult }
        Write-Output $barResult
        $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunGanttHierarchyRegression")
        if (-not $result.StartsWith('PASS:')) { throw $result }
        Write-Output $result
    }
    $taskComponents.Remove($harness)
    $taskComponents.Remove($barRegression)
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskExcel.EnableEvents = $true
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    if (-not $taskBook.Saved) {
        throw 'Opening the workbook dirtied it during hierarchy initialization'
    }
    $taskWindow = $taskExcel.ActiveWindow
    Assert-TaskExpandedView $taskWindow 'Initial expanded reopen'
    $taskFoldConfig = $taskBook.Worksheets.Item('config')
    if (-not $SelectedColumnsOnly) {
    # Expanded sheet activation must clear a previous request/header freeze.
    $taskWindow.FreezePanes = $false
    $taskWindow.SplitColumn = 0
    $taskWindow.SplitRow = 0
    $taskWindow.ScrollColumn = 8
    $taskWindow.SplitColumn = 1
    $taskWindow.SplitRow = 4
    $taskWindow.FreezePanes = $true
    $taskBook.Saved = $true
    $taskBook.Worksheets.Item('config').Activate()
    $taskSheet.Activate()
    Assert-TaskExpandedView $taskWindow 'Expanded task sheet activation'
    if (-not $taskBook.Saved) {
        throw 'Task sheet activation dirtied a clean workbook while restoring the expanded view'
    }
    # An actual unsaved input edit must remain unsaved after initialization.
    $taskExcel.EnableEvents = $false
    $taskOwnerBefore = $taskSheet.Range('J52').Value2
    $taskSheet.Range('J52').Value2 = [string]$taskOwnerBefore + ' regression edit'
    $taskExcel.Run("'$($taskBook.Name)'!InitializeTaskHierarchyOutlines")
    if ($taskBook.Saved) {
        throw 'Hierarchy initialization hid an existing unsaved input edit'
    }
    $taskSheet.Range('J52').Value2 = $taskOwnerBefore
    $taskBook.Saved = $true
    $taskExcel.EnableEvents = $true
    if (-not $taskSheet.ProtectContents -or -not $taskSheet.EnableOutlining) {
        throw 'Saved workbook did not restore protected outlines on open'
    }
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if ($taskSheet.Columns.Item($taskColumn).OutlineLevel -ne 2) {
            throw "Reopening changed the detail column outline depth at column $taskColumn"
        }
    }
    if ($taskSheet.Columns.Item(8).OutlineLevel -ne 1 -or
        $taskSheet.Columns.Item(24).OutlineLevel -ne 1 -or $taskSheet.Outline.SummaryColumn -ne -4152) {
        throw 'Request/Gantt columns or column summary position are incorrect'
    }
    # Native outline levels update command state; the UI event synchronizes the view.
    $taskSheet.Outline.ShowLevels([Type]::Missing, 1)
    Start-Sleep -Seconds 2
    Assert-TaskRequestFreeze $taskWindow 'Native column level 1 collapse'
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if (-not $taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Protected column collapse failed at column $taskColumn"
        }
    }
    if ($taskSheet.Columns.Item(8).Hidden -or $taskSheet.Columns.Item(24).Hidden) {
        throw 'Column collapse hid the request or Gantt column'
    }
    $taskWindow.Panes.Item(4).ScrollColumn = 30
    $taskWindow.Panes.Item(4).ScrollRow = 59
    Assert-TaskRequestFreeze $taskWindow 'Collapsed horizontal scroll'
    $taskScrollColumn = $taskWindow.Panes.Item(4).ScrollColumn
    $taskScrollRow = $taskWindow.Panes.Item(4).ScrollRow
    $taskSheet.Rows.Item(59).ShowDetail = $false
    if (-not $taskSheet.Rows.Item(60).Hidden -or $taskSheet.Rows.Item(59).Hidden) {
        throw 'Collapse after reopening failed'
    }
    $taskSheet.Rows.Item(59).ShowDetail = $true
    $taskSheet.Activate()
    $taskExcel.Run("'$($taskBook.Name)'!RefreshGanttSheet", $false)
    Assert-TaskRequestFreeze $taskWindow 'Collapsed refresh'
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne $taskScrollColumn -or
        $taskWindow.Panes.Item(4).ScrollRow -ne $taskScrollRow) {
        throw 'Refresh moved the scrolling Gantt pane away from its current date or task row'
    }
    if ($taskSheet.Rows.Item(53).OutlineLevel -ne 2 -or $taskSheet.Rows.Item(61).OutlineLevel -ne 3) {
        throw 'Refresh changed the outline depth'
    }
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if (-not $taskSheet.Columns.Item($taskColumn).Hidden -or
            $taskSheet.Columns.Item($taskColumn).OutlineLevel -ne 2) {
            throw "Refresh expanded or duplicated the detail column group at column $taskColumn"
        }
    }
    if ($taskSheet.Columns.Item(8).Hidden -or $taskSheet.Columns.Item(24).Hidden) {
        throw 'Refresh hid the request or Gantt column'
    }
    if ($taskExcel.WorksheetFunction.CountA($taskSheet.Range('L3:W3')) -ne 0 -or
        $taskExcel.WorksheetFunction.CountA($taskSheet.Range('X5:BA74')) -ne 0) {
        throw 'Refresh restored a legend or timeline caption'
    }
    if ($taskSheet.Range('AL52').Borders.Item(10).Color -ne 4891414 -or
        $taskSheet.Range('AK59').Borders.Item(8).Color -ne 15820387 -or
        $taskSheet.Range('AK74').Borders.Item(9).Color -ne 15820387) {
        throw 'Refresh changed status or combined plan/actual frames'
    }
    for ($taskRow = 60; $taskRow -le 73; $taskRow++) {
        if ($taskSheet.Cells.Item($taskRow, 37).Borders.Item(8).Color -eq 15820387 -or
            $taskSheet.Cells.Item($taskRow, 37).Borders.Item(9).Color -eq 15820387 -or
            $taskSheet.Cells.Item($taskRow, 38).Borders.Item(10).Color -ne 4891414) {
            throw "Refresh restored a child frame or lost status at row $taskRow"
        }
    }
    foreach ($taskShape in $taskSheet.Shapes) {
        if ($taskShape.Name.StartsWith('gtBar_') -or $taskShape.Name.StartsWith('gtKey_')) {
            throw 'Refresh inserted a Gantt drawing object'
        }
    }
    # Persist a collapsed column outline and exercise Workbook_Open again.
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    if (-not $taskBook.Saved -or -not $taskSheet.ProtectContents -or -not $taskSheet.EnableOutlining) {
        throw 'Collapsed reopen did not retain clean state and protected outlines'
    }
    $taskWindow = $taskExcel.ActiveWindow
    Assert-TaskRequestFreeze $taskWindow 'Collapsed reopen'
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne $taskScrollColumn -or
        $taskWindow.Panes.Item(4).ScrollRow -ne $taskScrollRow) {
        throw 'Collapsed reopen moved the scrolling Gantt pane away from its saved position'
    }
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if (-not $taskSheet.Columns.Item($taskColumn).Hidden -or
            $taskSheet.Columns.Item($taskColumn).OutlineLevel -ne 2) {
            throw "Collapsed reopen expanded or duplicated the detail group at column $taskColumn"
        }
    }
    if ($taskSheet.Columns.Item(8).Hidden -or $taskSheet.Columns.Item(24).Hidden -or
        $taskSheet.Rows.Item(53).OutlineLevel -ne 2 -or $taskSheet.Rows.Item(61).OutlineLevel -ne 3) {
        throw 'Collapsed reopen changed request/Gantt visibility or row outlines'
    }
    # Native outline level 2 must clear all freeze/splits and return to A once.
    $taskSheet.Outline.ShowLevels([Type]::Missing, 2)
    Start-Sleep -Seconds 2
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if ($taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Column expansion after reopening failed at column $taskColumn"
        }
    }
    Assert-TaskExpandedView $taskWindow 'Native column level 2 expansion'
    if ($taskWindow.ScrollRow -ne $taskScrollRow) {
        throw 'Expansion changed the visible task row'
    }
    $taskWindow.ScrollColumn = 30
    $taskExcel.Run("'$($taskBook.Name)'!TaskOutlineViewTick")
    Assert-TaskExpandedView $taskWindow 'Expanded manual horizontal scroll' 30
    # Editing the column settings applies the group and fixed column immediately.
    $taskFoldConfig = $taskBook.Worksheets.Item('config')
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J13').Value2 = ' j '
    $taskFoldConfig.Range('J14').Value2 = 'v'
    $taskSheet.Activate()
    Assert-TaskDetailOutline $taskSheet 10 22 'Custom J:V group'
    if ($taskFoldConfig.Range('J13').Value2 -cne 'J' -or $taskFoldConfig.Range('J14').Value2 -cne 'V') {
        throw 'Fold column settings were not trimmed and normalized to uppercase'
    }
    $taskSheet.Outline.ShowLevels([Type]::Missing, 1)
    Start-Sleep -Seconds 2
    Assert-TaskRequestFreeze $taskWindow 'Custom J:V collapsed view' 10
    if ($taskSheet.Columns.Item(9).Hidden -or $taskSheet.Columns.Item(23).Hidden -or
        -not $taskSheet.Columns.Item(10).Hidden -or -not $taskSheet.Columns.Item(22).Hidden) {
        throw 'Custom J:V fold retained hidden columns outside its new range'
    }
    # Display start is independent of folding and resets the right pane only on changes.
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J15').Value2 = 'F'
    $taskSheet.Activate()
    Assert-TaskRequestFreeze $taskWindow 'F display start with J:V fold' 10 6
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne 7 -or $taskSheet.Columns.Item(6).Hidden -or
        $taskSheet.Columns.Item(7).Hidden -or $taskSheet.Columns.Item(8).Hidden) {
        throw 'F display start did not show F/G/H consecutively beside the folded group'
    }
    Assert-TaskDetailOutline $taskSheet 10 22 'F display start preserved fold outline'
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J15').Value2 = 'G'
    $taskSheet.Activate()
    Assert-TaskRequestFreeze $taskWindow 'G display start with J:V fold' 10 7
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne 8 -or $taskSheet.Columns.Item(8).Hidden) {
        throw 'G display start did not show H immediately beside it'
    }
    # A display start inside the hidden range advances to its next visible column.
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J15').Value2 = 'L'
    $taskSheet.Activate()
    Assert-TaskRequestFreeze $taskWindow 'Hidden L display start advances to W' 10 23
    Assert-TaskDetailOutline $taskSheet 10 22 'Hidden display start preserved folding'
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne 24 -or
        -not $taskSheet.Columns.Item(10).Hidden -or -not $taskSheet.Columns.Item(22).Hidden) {
        throw 'Hidden display start changed the fold range or failed to advance to visible columns'
    }
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J15').Value2 = 'F'
    $taskSheet.Activate()
    Assert-TaskRequestFreeze $taskWindow 'Restored F display start' 10 6
    # A changed range preserves the collapsed state and removes the old outline.
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J13').Value2 = 'L'
    $taskFoldConfig.Range('J14').Value2 = 'W'
    $taskSheet.Activate()
    Assert-TaskDetailOutline $taskSheet 12 23 'Custom L:W group'
    Assert-TaskRequestFreeze $taskWindow 'Custom L:W collapsed view with F start' 12 6
    if ($taskSheet.Columns.Item(10).Hidden -or $taskSheet.Columns.Item(11).Hidden -or
        -not $taskSheet.Columns.Item(12).Hidden -or -not $taskSheet.Columns.Item(23).Hidden) {
        throw 'Changing the fold range lost its collapsed state or retained old hidden columns'
    }
    # A range reaching into calendar columns must survive the date-header redraw.
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J14').Value2 = 'AA'
    $taskSheet.Activate()
    Assert-TaskDetailOutline $taskSheet 12 27 'Calendar-inclusive L:AA group'
    Assert-TaskRequestFreeze $taskWindow 'Calendar-inclusive collapsed view' 12 6
    $taskWindow.Panes.Item(4).ScrollColumn = 30
    $taskWindow.Panes.Item(4).ScrollRow = 59
    $taskCalendarScrollColumn = $taskWindow.Panes.Item(4).ScrollColumn
    $taskCalendarScrollRow = $taskWindow.Panes.Item(4).ScrollRow
    $taskExcel.Run("'$($taskBook.Name)'!RefreshGanttSheet", $false)
    Assert-TaskDetailOutline $taskSheet 12 27 'Calendar-inclusive refresh outline'
    Assert-TaskRequestFreeze $taskWindow 'Calendar-inclusive collapsed refresh' 12 6
    for ($taskColumn = 12; $taskColumn -le 27; $taskColumn++) {
        if (-not $taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Calendar redraw expanded a configured fold column at $taskColumn"
        }
    }
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne $taskCalendarScrollColumn -or
        $taskWindow.Panes.Item(4).ScrollRow -ne $taskCalendarScrollRow) {
        throw 'Calendar-inclusive refresh moved the current date or task row'
    }
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J14').Value2 = 'W'
    $taskSheet.Activate()
    Assert-TaskDetailOutline $taskSheet 12 23 'Restored L:W group after calendar fold'
    for ($taskColumn = 24; $taskColumn -le 27; $taskColumn++) {
        if ($taskSheet.Columns.Item($taskColumn).Hidden -or $taskSheet.Columns.Item($taskColumn).OutlineLevel -ne 1) {
            throw "Restoring L:W retained a former calendar fold column at $taskColumn"
        }
    }
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    $taskFoldConfig = $taskBook.Worksheets.Item('config')
    $taskWindow = $taskExcel.ActiveWindow
    Assert-TaskDetailOutline $taskSheet 12 23 'Custom range reopen'
    Assert-TaskRequestFreeze $taskWindow 'Custom collapsed reopen with saved F start' 12 6
    if (-not $taskBook.Saved -or $taskFoldConfig.Range('J13').Value2 -cne 'L' -or
        $taskFoldConfig.Range('J14').Value2 -cne 'W' -or $taskFoldConfig.Range('J15').Value2 -cne 'F' -or
        $taskWindow.Panes.Item(4).ScrollColumn -ne 7) {
        throw 'Custom fold range reopen changed its settings or dirtied the workbook'
    }
    # Invalid/reversed config pairs keep the last valid outline intact.
    $taskFoldConfig.Activate()
    $taskFoldConfig.Range('J13').Value2 = 'X'
    Assert-TaskDetailOutline $taskSheet 12 23 'Invalid range preserved outline'
    if ([string]::IsNullOrWhiteSpace([string]$taskFoldConfig.Range('I17').Value2) -or
        -not $taskSheet.Columns.Item(12).Hidden -or -not $taskSheet.Columns.Item(23).Hidden) {
        throw 'Invalid fold settings lost their error guide or changed the existing collapsed group'
    }
    }
    # 2026-10-19 김좌현 수정 : noncontiguous selected columns, including a column inside I:W.
    Write-Output 'Testing multiple fixed column configuration'
    $taskFoldConfig.Range('J13').Value2 = 'I'
    $taskFoldConfig.Range('J14').Value2 = 'W'
    $taskFoldConfig.Range('J15').Value2 = ' g, a, b, k, b '
    Write-Output 'Multiple fixed column settings applied'
    $taskSheet.Activate()
    if ($taskFoldConfig.Range('J15').Value2 -cne 'A,B,G,K') {
        throw 'Multiple fixed columns were not sorted, normalized and deduplicated'
    }
    $taskSheet.Outline.ShowLevels([Type]::Missing, 1)
    Write-Output 'Multiple fixed column groups collapsed'
    Start-Sleep -Seconds 2
    $taskExcel.Run("'$($taskBook.Name)'!TaskOutlineViewTick")
    Assert-TaskSelectedColumns $taskWindow $taskSheet @(1, 2, 7, 11) 'A,B,G,K collapse'
    if (-not $taskWindow.FreezePanes -or
        $taskWindow.SplitRow -ne 4 -or $taskWindow.Panes.Item(1).ScrollColumn -ne 1 -or
        $taskWindow.Panes.Item(4).ScrollColumn -ne 24) {
        throw "A,B,G,K view mismatch: freeze=$($taskWindow.FreezePanes), split=$($taskWindow.SplitColumn)/$($taskWindow.SplitRow), left=$($taskWindow.Panes.Item(1).ScrollColumn), right=$($taskWindow.Panes.Item($taskWindow.Panes.Count).ScrollColumn)"
    }
    for ($taskColumn = 1; $taskColumn -le 23; $taskColumn++) {
        $selected = $taskColumn -in @(1, 2, 7, 11)
        if ($taskSheet.Columns.Item($taskColumn).Hidden -eq $selected -or
            $taskSheet.Columns.Item($taskColumn).OutlineLevel -ne $(if ($selected) { 1 } else { 2 })) {
            throw "Multiple fixed columns have incorrect visibility or grouping at $taskColumn"
        }
    }
    $taskWindow.Panes.Item(4).ScrollColumn = 30
    $selectedScrollColumn = $taskWindow.Panes.Item(4).ScrollColumn
    $taskExcel.Run("'$($taskBook.Name)'!RefreshGanttSheet", $false)
    Assert-TaskSelectedColumns $taskWindow $taskSheet @(1, 2, 7, 11) 'A,B,G,K refresh'
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne $selectedScrollColumn -or $taskSheet.Columns.Item(11).Hidden) {
        throw "Refreshing multiple fixed columns: expected scroll=$selectedScrollColumn, actual=$($taskWindow.Panes.Item(4).ScrollColumn), K hidden=$($taskSheet.Columns.Item(11).Hidden)"
    }
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    $taskFoldConfig = $taskBook.Worksheets.Item('config')
    $taskWindow = $taskExcel.ActiveWindow
    Assert-TaskSelectedColumns $taskWindow $taskSheet @(1, 2, 7, 11) 'A,B,G,K reopen'
    if (-not $taskBook.Saved -or -not $taskWindow.FreezePanes -or
        $taskWindow.Panes.Item(4).ScrollColumn -ne $selectedScrollColumn -or
        $taskSheet.Columns.Item(11).Hidden -or -not $taskSheet.Columns.Item(3).Hidden) {
        throw 'Reopening multiple fixed columns changed the saved view or expanded the group'
    }
    $taskSheet.Outline.ShowLevels([Type]::Missing, 2)
    Start-Sleep -Seconds 2
    $taskExcel.Run("'$($taskBook.Name)'!TaskOutlineViewTick")
    Assert-TaskExpandedView $taskWindow 'Multiple fixed columns expansion'
    for ($taskColumn = 1; $taskColumn -le 24; $taskColumn++) {
        if ($taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Expansion did not restore column $taskColumn"
        }
    }
    # Invalid lists must retain the previous groups and display.
    foreach ($invalidSelection in @('A,,G', 'A,B,X', 'A,1,G')) {
        $taskFoldConfig.Activate()
        $taskFoldConfig.Range('J15').Value2 = $invalidSelection
        if ($taskSheet.Columns.Item(3).OutlineLevel -ne 2 -or
            $taskSheet.Columns.Item(11).OutlineLevel -ne 1) {
            throw "Invalid selection '$invalidSelection' changed the previous groups"
        }
    }
    $taskFoldConfig.Range('J15').Value2 = 'A,B,G'
    $taskSheet.Activate()
    $taskSheet.Outline.ShowLevels([Type]::Missing, 1)
    Start-Sleep -Seconds 2
    $taskExcel.Run("'$($taskBook.Name)'!TaskOutlineViewTick")
    Assert-TaskSelectedColumns $taskWindow $taskSheet @(1, 2, 7) 'A,B,G replacement'
    if ($taskWindow.Panes.Item(4).ScrollColumn -ne 24 -or
        -not $taskSheet.Columns.Item(11).Hidden) {
        throw 'Changing the selected fixed columns retained K or the previous freeze boundary'
    }
    Write-Output 'PASS: multiple fixed columns A,B,G,K, in-range selection, normalization, native levels, refresh, saved reopen and invalid input'
    # Deliver the original default I:W setting, expanded from A.
    $taskFoldConfig.Range('J13').Value2 = 'I'
    $taskFoldConfig.Range('J14').Value2 = 'W'
    $taskFoldConfig.Range('J15').Value2 = 'H'
    $taskSheet.Activate()
    Assert-TaskDetailOutline $taskSheet 9 23 'Restored default group'
    $taskSheet.Outline.ShowLevels([Type]::Missing, 2)
    Start-Sleep -Seconds 2
    Assert-TaskExpandedView $taskWindow 'Restored default expanded view'
    # Restore the original view and verify the saved expanded state on reopen.
    $taskWindow.ScrollColumn = 1
    $taskWindow.ScrollRow = [Math]::Max(1, $taskOriginalScrollRow)
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskBook = Open-TaskTestWorkbook $taskExcel $taskTemp
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    $taskWindow = $taskExcel.ActiveWindow
    Assert-TaskExpandedView $taskWindow 'Final expanded reopen'
    if (-not $taskBook.Saved -or $taskWindow.ScrollRow -ne [Math]::Max(1, $taskOriginalScrollRow)) {
        throw 'Expanded reopen did not retain clean state and original task row'
    }
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    Start-Sleep -Seconds 3
    if ($taskExcel.Workbooks.Count -ne 0) {
        throw 'Outline event cleanup left a workbook open after closing'
    }
    if ($SelectedColumnsOnly) {
        Write-Output 'PASS: selected column view, default outline restoration, clean expanded reopen and event close cleanup'
    }
    else {
        Write-Output 'PASS: independent fold ranges and collapsed display start, normalization, invalid input, old outline cleanup, custom reopen; native levels 1/2 synchronize collapse freeze and expanded A view, manual expanded scrolling, clean reopen, saved row/date positions, protected outlines, full refresh and outline event close cleanup'
    }
    if ($SavePath) {
        Copy-Item -LiteralPath $taskTemp -Destination ([IO.Path]::GetFullPath($SavePath))
        Write-Output ('UPDATED COPY: ' + [IO.Path]::GetFullPath($SavePath))
    }
}
catch {
    Write-Output ('REGRESSION FAILURE: ' + $_.Exception.ToString())
    Write-Output ('SCRIPT STACK: ' + $_.ScriptStackTrace)
    throw
}
finally {
    if ($null -ne $taskBook) {
        try { $taskBook.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook) } catch {}
        $taskBook = $null
    }
    if ($null -ne $taskExcel) {
        try { $taskExcel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) } catch {}
        $taskExcel = $null
    }
    [GC]::Collect()
    if (Test-Path -LiteralPath $taskTemp) {
        try { Remove-Item -LiteralPath $taskTemp -Force }
        catch { Write-Warning ('Could not remove the temporary workbook: ' + $taskTemp + ' (' + $_.Exception.Message + ')') }
    }
}
