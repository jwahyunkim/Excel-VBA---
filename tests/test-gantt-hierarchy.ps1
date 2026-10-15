[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath)

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
try {
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskExcel.AutomationSecurity = 1
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
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
    $code = $taskComponents.Item([string]$taskBook.CodeName).CodeModule
    $start = $code.ProcStartLine('Workbook_Open', 0)
    $count = $code.ProcCountLines('Workbook_Open', 0)
    if (-not ([string]$code.Lines($start, $count)).Contains('InitializeTaskHierarchyOutlines')) {
        for ($line = $start + $count - 1; $line -ge $start; $line--) {
            if (([string]$code.Lines($line, 1)).Trim() -eq 'End Sub') {
                $code.InsertLines($line, '    InitializeTaskHierarchyOutlines')
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
Public Function RunGanttHierarchyRegression() As String
    Dim ws As Worksheet, original As Worksheet, r As Long, c As Long
    Dim levels As Variant, dataBefore As Variant, hiddenBefore() As Boolean
    Dim holidays As Object, workdays As Object, firstDate As Date, lastDate As Date
    Dim lastRow As Long, lastCol As Long, taskCol As Long
    Dim existingShape As Shape, unrelatedShapeCount As Long
    Dim edge As Variant, edgeColor As Long
    On Error GoTo Failed
    Application.EnableEvents = False
    Set original = ThisWorkbook.Worksheets("Sheet1")
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
    FormatBaseArea original, lastRow, firstDate, lastDate, holidays, workdays
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
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('GanttHierarchyRegression')
    $barRegression = $taskComponents.Item('GanttBarsRegression')
    Write-Output 'Running native hierarchy and bar regression'
    $barResult = [string]$taskExcel.Run("'$($taskBook.Name)'!RunGanttBarsRegression")
    if (-not $barResult.StartsWith('PASS:')) { throw $barResult }
    Write-Output $barResult
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunGanttHierarchyRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    $taskComponents.Remove($harness)
    $taskComponents.Remove($barRegression)
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    $taskExcel.EnableEvents = $true
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    if (-not $taskBook.Saved) {
        throw 'Opening the workbook dirtied it during hierarchy initialization'
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
    $taskSheet.Columns.Item(24).ShowDetail = $false
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if (-not $taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Protected column collapse failed at column $taskColumn"
        }
    }
    if ($taskSheet.Columns.Item(8).Hidden -or $taskSheet.Columns.Item(24).Hidden) {
        throw 'Column collapse hid the request or Gantt column'
    }
    $taskSheet.Rows.Item(59).ShowDetail = $false
    if (-not $taskSheet.Rows.Item(60).Hidden -or $taskSheet.Rows.Item(59).Hidden) {
        throw 'Collapse after reopening failed'
    }
    $taskSheet.Rows.Item(59).ShowDetail = $true
    $taskSheet.Activate()
    $taskExcel.Run("'$($taskBook.Name)'!RefreshGanttSheet", $false)
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
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskSheet = $taskBook.Worksheets.Item('Sheet1')
    if (-not $taskBook.Saved -or -not $taskSheet.ProtectContents -or -not $taskSheet.EnableOutlining) {
        throw 'Collapsed reopen did not retain clean state and protected outlines'
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
    $taskSheet.Columns.Item(24).ShowDetail = $true
    for ($taskColumn = 9; $taskColumn -le 23; $taskColumn++) {
        if ($taskSheet.Columns.Item($taskColumn).Hidden) {
            throw "Column expansion after reopening failed at column $taskColumn"
        }
    }
    # Deliver the workbook expanded, matching its original view.
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    Write-Output 'PASS: clean and collapsed reopen, preserved unsaved edits, protected row/column collapse/expand and collapsed full Gantt refresh'
    if ($SavePath) {
        Copy-Item -LiteralPath $taskTemp -Destination ([IO.Path]::GetFullPath($SavePath))
        Write-Output ('UPDATED COPY: ' + [IO.Path]::GetFullPath($SavePath))
    }
}
finally {
    if ($null -ne $taskBook) { try { $taskBook.Close($false) } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook) }
    if ($null -ne $taskExcel) { try { $taskExcel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
    [GC]::Collect()
    if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp -Force }
}
