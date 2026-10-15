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
        if ($text.Trim() -eq 'ActiveWindow.DisplayGridlines = False' -and
            -not ([string]$code.Lines($line - 1, 1)).Contains('ApplyTaskHierarchyAppearance')) {
            $code.InsertLines($line, '    ApplyTaskHierarchyAppearance ws, lastRow, startCol, endCol')
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
    ws.Outline.ShowLevels 1
    Check ws.Rows(6).Hidden And Not ws.Rows(5).Hidden, "collapse failed"
    ws.Outline.ShowLevels 3
    Check Not ws.Rows(7).Hidden, "expand failed"
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
    Check CStr(original.Range("AL52").Value2) = "1", "52 level caption"
    Check CStr(original.Range("AL53").Value2) = "2", "53 level caption"
    Check CStr(original.Range("AL61").Value2) = "3", "61 level caption"
    Check original.Shapes.Count = unrelatedShapeCount, "new drawing objects remain or existing buttons lost"
    RunGanttHierarchyRegression = "PASS: nested outlines, protected collapse/expand, input preservation, cell hierarchy, calendar and no inserted drawing objects"
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
    if (-not $taskSheet.ProtectContents -or -not $taskSheet.EnableOutlining) {
        throw 'Saved workbook did not restore protected outlines on open'
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
    foreach ($taskShape in $taskSheet.Shapes) {
        if ($taskShape.Name.StartsWith('gtBar_') -or $taskShape.Name.StartsWith('gtKey_')) {
            throw 'Refresh inserted a Gantt drawing object'
        }
    }
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $null
    Write-Output 'PASS: reopen, protected collapse/expand and full Gantt refresh'
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
