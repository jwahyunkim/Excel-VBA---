[CmdletBinding()]
param([string]$WorkbookPath, [string]$SavePath)

$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
if (-not $WorkbookPath) {
    $taskConfig = Get-Content -LiteralPath (Join-Path $taskRoot 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $WorkbookPath = Join-Path $taskRoot $taskConfig.excel_file
}
$taskTemp = Join-Path ([IO.Path]::GetTempPath()) ('task-layout-' + [guid]::NewGuid().ToString('N') + '.xlsm')
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
    $taskPhase = 'open temporary workbook'
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    foreach ($obsolete in @('modDailyProgressReport', 'modWeeklyPptReport')) {
        try { $component = $taskComponents.Item($obsolete) } catch { $component = $null }
        if ($null -ne $component) { $taskComponents.Remove($component) }
    }
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $taskRoot 'vba-files/Module') -Filter '*.bas') {
        $taskPhase = 'import ' + $file.BaseName
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        try { $source = [Text.UTF8Encoding]::new($false, $true).GetString($bytes) }
        catch { $source = [Text.Encoding]::GetEncoding(949).GetString($bytes) }
        $source = $source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        try { $component = $taskComponents.Item($file.BaseName) }
        catch { $component = $taskComponents.Add(1); $component.Name = $file.BaseName }
        $code = $component.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString($source)
    }
    $taskPhase = 'import workbook class'
    $source = [IO.File]::ReadAllText((Join-Path $taskRoot 'vba-files/Class/현재_통합_문서.cls'))
    $source = $source.Substring($source.IndexOf('Option Explicit'))
    $code = $taskComponents.Item([string]$taskBook.CodeName).CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($source)
    $harness = $taskComponents.Add(1)
    $harness.Name = 'TaskLayoutRegression'
    $harness.CodeModule.AddFromString(@'
Option Explicit
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 910, "TaskLayoutRegression", label
End Sub
Public Function RunTaskLayoutRegression() As String
    Dim ws As Worksheet, r As Long, c As Long, lastRow As Long
    Dim requests As Variant, modifications As Variant, details As Variant, manual As Variant
    Dim name As Variant, headerBefore As String, wasLegacy As Boolean
    Dim startDate As Date, endDate As Date, reportWs As Worksheet, retiredWs As Worksheet
    Dim configuredStart As Variant, configuredEnd As Variant, configuredModification As Variant
    Dim reversedPeriodError As Long
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    UnprotectTaskSheet ws
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    headerBefore = CStr(ws.Range("I4").Value2)
    wasLegacy = (headerBefore <> "수정 내용")
    requests = ws.Range("H5:H" & lastRow).Value2
    If wasLegacy Then
        details = ws.Range("I5:N" & lastRow).Value2
        manual = ws.Range("Q5:S" & lastRow).Value2
    Else
        modifications = ws.Range("I5:I" & lastRow).Value2
        details = ws.Range("J5:O" & lastRow).Value2
        manual = ws.Range("R5:T" & lastRow).Value2
    End If
    SetupDataHeaders ws
    SetupDataHeaders ws
    Check CStr(ws.Range("H4").Value2) = "요청 내용", "request header"
    Check CStr(ws.Range("I4").Value2) = "수정 내용", "modification header"
    Check CStr(ws.Range("J4").Value2) = "담당", "owner header shifted twice"
    Check CStr(ws.Range("K4").Value2) = "비고", "note header shifted twice"
    Check CStr(ws.Range("L4").Value2) = "계획 시작일", "date header shifted twice"
    For r = 1 To UBound(requests, 1)
        Check CStr(ws.Cells(r + 4, "H").Value2) = CStr(requests(r, 1)), "request lost row " & r
        If wasLegacy Then Check Len(CStr(ws.Cells(r + 4, "I").Value2)) = 0, "new modification not blank"
        If Not wasLegacy Then Check CStr(ws.Cells(r + 4, "I").Value2) = CStr(modifications(r, 1)), "modification lost row " & r
        For c = 1 To 6
            Check CStr(ws.Cells(r + 4, c + 9).Value2) = CStr(details(r, c)), "owner/note/date lost row " & r & " column " & c
        Next c
        For c = 1 To 3
            Check CStr(ws.Cells(r + 4, c + 17).Value2) = CStr(manual(r, c)), "manual/status lost row " & r
        Next c
    Next r
    For Each name In Array("WeeklyPptTemplate", "_일별진행현황템플릿", "_일별진척률이력", "config_일일현황")
        Set retiredWs = Nothing
        On Error Resume Next
        Set retiredWs = ThisWorkbook.Worksheets(CStr(name))
        On Error GoTo Failed
        If Not retiredWs Is Nothing Then
            retiredWs.Visible = xlSheetVisible
            retiredWs.Delete
        End If
    Next name
    EnsureConfigSheet
    Set reportWs = ThisWorkbook.Worksheets(OUTPUT_CONFIG_SHEET_NAME)
    configuredStart = reportWs.Range(OUTPUT_REPORT_START_CELL).Value
    configuredEnd = reportWs.Range(OUTPUT_REPORT_END_CELL).Value
    configuredModification = reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "N"
    EnsureWeeklyReportConfigSheet
    Check Not GetOutputShowModificationFlag(), "modification toggle reset during refresh"
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "Y"
    Check GetOutputShowModificationFlag(), "modification toggle did not enable"
    reportWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = configuredModification
    GetOutputReportPeriod startDate, endDate
    Check startDate <= endDate, "default output period"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = DateSerial(2026, 9, 1)
    reportWs.Range(OUTPUT_REPORT_END_CELL).Value = DateSerial(2026, 9, 30)
    EnsureWeeklyReportConfigSheet
    GetOutputReportPeriod startDate, endDate
    Check startDate = DateSerial(2026, 9, 1) And endDate = DateSerial(2026, 9, 30), "period reset during refresh"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = DateSerial(2026, 10, 1)
    On Error Resume Next
    GetOutputReportPeriod startDate, endDate
    reversedPeriodError = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check reversedPeriodError <> 0, "reversed output period accepted"
    reportWs.Range(OUTPUT_REPORT_START_CELL).Value = configuredStart
    reportWs.Range(OUTPUT_REPORT_END_CELL).Value = configuredEnd
    ApplyTaskInputValidation ws
    Check InStr(ws.Range("H5").Validation.Formula1, "LEN(H5)") > 0, "request length validation"
    Check InStr(ws.Range("I5").Validation.Formula1, "LEN(I5)") > 0, "modification length validation"
    ws.Activate
    RefreshGanttSheet False
    UpgradeTaskReportButtons ws
    버튼_생성_선택
    버튼_생성_취소
    Check ws.Shapes("btnPeriodTextReport").OnAction = "기간보고TXT_생성", "TXT button action"
    Check ws.Shapes("btnPeriodTextReport").TextFrame2.TextRange.Text = "TXT 출력", "TXT button caption"
    Check ws.Range("I5").Locked = False, "modification input locked"
    Check ws.ProtectContents, "task sheet protection disabled"
    RunTaskLayoutRegression = "PASS: data preservation, idempotent columns, period validation/persistence, input validation, Gantt refresh, TXT button, protection"
    Exit Function
Failed:
    RunTaskLayoutRegression = "FAIL: " & Err.Description
End Function
'@)
    $taskPhase = 'save and reload imported VBA'
    $taskBook.Save()
    $taskBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)
    $taskBook = $taskExcel.Workbooks.Open($taskTemp, 0, $false)
    $taskComponents = $taskBook.VBProject.VBComponents
    $harness = $taskComponents.Item('TaskLayoutRegression')
    $taskPhase = 'run VBA layout regression'
    $result = [string]$taskExcel.Run("'$($taskBook.Name)'!RunTaskLayoutRegression")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    $taskComponents.Remove($harness)
    if ($SavePath) {
        $taskPhase = 'save updated copy'
        $taskBook.SaveAs([IO.Path]::GetFullPath($SavePath), 52)
        Write-Output ('UPDATED COPY: ' + [IO.Path]::GetFullPath($SavePath))
    }
}
catch {
    Write-Error ('Failed during ' + $taskPhase + ': ' + $_.Exception.Message) -ErrorAction Continue
    throw
}
finally {
    if ($null -ne $taskBook) { try { $taskBook.Close($false) } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook) }
    if ($null -ne $taskExcel) { try { $taskExcel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $taskTemp) { Remove-Item -LiteralPath $taskTemp -Force }
}
