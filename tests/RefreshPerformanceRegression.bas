Attribute VB_Name = "RefreshPerformanceRegression"
Option Explicit
' 2026-10-29 김좌현 수정 : 통합 새로고침 경로의 반복 실행·기존 결과 보존 검증
' 2026-10-23 김좌현 수정 : 중복 등록 버튼 제거에 맞춰 웹 버튼 보기 비교 범위 정리
' 2026-10-22 김좌현 수정 : 간트 성능 개선 전후의 계산/막대/레이아웃/보기 보존 회귀

Private Const LAST_FIXTURE_ROW As Long = 14
Private mMonday As Date
Private mFirstDate As Date
Private mLastDate As Date
Private mScenarioSnapshots As String
Private mStage As String
Private mExpectedWebMode As Boolean
Private mExpectedModuleHidden As Boolean

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 9290, "RefreshPerformanceRegression", message
End Sub

Private Function DayCell(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal targetDate As Date) As Range
    Set DayCell = ws.Cells(rowNumber, ws.Range(COL_GANTT_START & "1").Column + CLng(targetDate - mFirstDate))
End Function

Private Sub Seed(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal level As Long, _
                 ByVal caption As String, ByVal planStart As Date, ByVal planEnd As Date, _
                 ByVal actualStart As Variant, ByVal actualEnd As Variant, ByVal progress As Double)
    ws.Cells(rowNumber, COL_LEVEL).Value2 = level
    ws.Cells(rowNumber, COL_REQUEST).Value2 = caption
    ws.Cells(rowNumber, COL_MINOR_CATEGORY).Value2 = "Regression"
    ws.Cells(rowNumber, COL_PLAN_START).Value = planStart
    ws.Cells(rowNumber, COL_PLAN_END).Value = planEnd
    If IsDate(actualStart) Then ws.Cells(rowNumber, COL_ACTUAL_START).Value = CDate(actualStart)
    If IsDate(actualEnd) Then ws.Cells(rowNumber, COL_ACTUAL_END).Value = CDate(actualEnd)
    ws.Cells(rowNumber, COL_PROGRESS).Value2 = progress
End Sub

Private Function CellVisual(ByVal cell As Range) As String
    Dim result As String, stopIndex As Long, edge As Variant
    result = CStr(cell.Interior.Pattern) & ":" & CStr(cell.Interior.Color)
    If cell.Interior.Pattern = xlPatternLinearGradient Then
        result = result & ":degree=" & CStr(cell.Interior.Gradient.Degree)
        For stopIndex = 1 To cell.Interior.Gradient.ColorStops.Count
            With cell.Interior.Gradient.ColorStops(stopIndex)
                result = result & ":stop=" & CStr(.Position) & "," & CStr(.Color)
            End With
        Next stopIndex
    End If
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalUp, xlDiagonalDown)
        With cell.Borders(CLng(edge))
            result = result & ":edge=" & CStr(edge) & "," & CStr(.LineStyle)
            If .LineStyle <> xlNone Then result = result & "," & CStr(.Color) & "," & CStr(.Weight)
        End With
    Next edge
    CellVisual = result & ":font=" & CStr(cell.Font.Color) & "," & CStr(cell.Font.Bold) & _
                 ":indent=" & CStr(cell.IndentLevel) & ":format=" & CStr(cell.NumberFormat)
End Function

Private Function Snapshot(ByVal ws As Worksheet) As Object
    Dim result As Object, rowNumber As Long, columnNumber As Long, lastColumn As Long
    Dim cell As Range, name As Variant, value As Variant
    Set result = CreateObject("Scripting.Dictionary")
    lastColumn = ws.Range(COL_GANTT_START & "1").Column + CLng(mLastDate - mFirstDate)
    For rowNumber = DATA_START_ROW To LAST_FIXTURE_ROW
        For columnNumber = ws.Range(COL_NO & "1").Column To lastColumn
            Set cell = ws.Cells(rowNumber, columnNumber)
            If columnNumber <> ws.Range(COL_WEB_TASK_ID & "1").Column Then
                value = cell.Value2
                If IsError(value) Then value = "#ERROR"
                result.Add "value:" & cell.Address(False, False), CStr(value)
                result.Add "visual:" & cell.Address(False, False), CellVisual(cell)
            End If
        Next columnNumber
        result.Add "row:" & rowNumber, CStr(ws.Rows(rowNumber).OutlineLevel) & ":" & _
                   CStr(ws.Rows(rowNumber).Hidden) & ":" & Format$(ws.Rows(rowNumber).RowHeight, "0.00")
    Next rowNumber
    For columnNumber = ws.Range(COL_NO & "1").Column To lastColumn
        result.Add "column:" & columnNumber, CStr(ws.Columns(columnNumber).OutlineLevel) & ":" & _
                   CStr(ws.Columns(columnNumber).Hidden) & ":" & Format$(ws.Columns(columnNumber).ColumnWidth, "0.000")
    Next columnNumber
    For Each name In Array(COL_MINOR_CATEGORY, COL_WEB_ENABLED, COL_WEB_LEVEL, COL_WEB_RESULT, COL_REQUEST, COL_MODIFICATION, COL_ACTUAL_START, COL_ACTUAL_END, COL_WORK_MINUTES)
        result.Add "header:" & CStr(name), CellVisual(ws.Cells(HEADER_ROW, CStr(name)))
    Next name
    For Each name In Array("btnWebPreview", "btnWebConfig")
        result.Add "button:" & CStr(name), CStr(ws.Shapes(CStr(name)).Visible)
    Next name
    result.Add "protection", CStr(ws.ProtectContents) & ":" & CStr(ws.Cells(DATA_START_ROW, COL_WEB_TASK_ID).Locked) & ":" & _
               CStr(ws.Cells(DATA_START_ROW, COL_WEB_RESULT).Locked) & ":" & CStr(ws.Cells(DATA_START_ROW, COL_REQUEST).Locked)
    result.Add "mode", CStr(IsWebRegistrationMode(ws))
    Set Snapshot = result
End Function

Private Function SnapshotText(ByVal snapshotValue As Object) As String
    Dim key As Variant, result As String
    For Each key In snapshotValue.Keys
        result = result & CStr(key) & "=" & CStr(snapshotValue(key)) & vbLf
    Next key
    SnapshotText = result
End Function

Private Sub CheckView(ByVal ws As Worksheet)
    Check ws.ProtectContents, "sheet protection lost"
    Check ws.Columns(COL_WEB_TASK_ID).Hidden, "private ID exposed"
    Check ws.Cells(DATA_START_ROW, COL_WEB_TASK_ID).Locked, "private ID unlocked"
    Check ws.Cells(DATA_START_ROW, COL_WEB_RESULT).Locked, "result unlocked"
    Check Not ws.Cells(DATA_START_ROW, COL_REQUEST).Locked, "request input locked"
    Check IsWebRegistrationMode(ws) = mExpectedWebMode, "web mode lost"
    Check CBool(ws.Columns(COL_WEB_MODULE).Hidden) = mExpectedModuleHidden, "manual web-column hide lost"
    Check ws.Rows(10).Hidden, "manual task-row hide lost"
    If mExpectedWebMode Then
        Check ws.Cells(HEADER_ROW, COL_REQUEST).Interior.Color = RGB(255, 235, 156), "web header lost yellow"
    Else
        Check ws.Cells(HEADER_ROW, COL_REQUEST).Interior.Color = RGB(242, 242, 242), "normal header color lost"
    End If
End Sub

Private Sub CheckRepeatedRefresh(ByVal ws As Worksheet, ByVal label As String)
    Dim first As Object, complete As Object, key As Variant
    mStage = label & " first"
    RefreshGanttSheet False
    CheckView ws
    Set first = Snapshot(ws)
    mStage = label & " repeated refresh"
    RefreshGanttSheet False
    CheckView ws
    Set complete = Snapshot(ws)
    Check first.Count = complete.Count, label & " snapshot key count"
    For Each key In first.Keys
        Check complete.Exists(key), label & " missing key " & CStr(key)
        Check CStr(first(key)) = CStr(complete(key)), label & " mismatch " & CStr(key) & ": first=" & CStr(first(key)) & ", full=" & CStr(complete(key))
    Next key
    mScenarioSnapshots = mScenarioSnapshots & "[" & label & "]" & vbLf & SnapshotText(complete)
End Sub

Private Sub CheckCorner(ByVal cell As Range)
    Check cell.Interior.Pattern = xlPatternLinearGradient, "completed delay corner lost"
    Check cell.Interior.Gradient.Degree = 45, "completed delay angle changed"
    Check cell.Interior.Gradient.ColorStops.Count = 4, "completed delay stops changed"
    Check cell.Interior.Gradient.ColorStops(1).Color = RGB(255, 255, 0), "completed delay yellow changed"
    Check cell.Interior.Gradient.ColorStops(4).Color = RGB(0, 140, 134), "level 3 actual color changed"
End Sub

Public Function RunRefreshPerformanceRegression() As String
    Dim ws As Worksheet, config As Worksheet, rowNumber As Long, baseline As Object, current As Object
    Dim key As Variant, progressBefore As Double, columnNumber As Long
    On Error GoTo Failed
    mScenarioSnapshots = vbNullString
    mExpectedWebMode = True
    mExpectedModuleHidden = True
    Application.EnableEvents = False
    EnsureConfigSheet
    Set config = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    mMonday = Date - Weekday(Date, vbMonday) + 1
    mFirstDate = DateSerial(Year(Date), Month(Date), 1) - 3
    mLastDate = DateSerial(Year(Date), Month(Date) + 1, 1) + 3
    If mMonday - 7 < mFirstDate Then mFirstDate = mMonday - 7
    If mMonday + 11 > mLastDate Then mLastDate = mMonday + 11
    config.Range(DISPLAY_SETTING_START_VALUE_CELL).Value = mFirstDate
    config.Range(DISPLAY_SETTING_END_VALUE_CELL).Value = mLastDate
    config.Range(TASK_FOLD_START_VALUE_CELL).Value2 = COL_MODIFICATION
    config.Range(TASK_FOLD_END_VALUE_CELL).Value2 = COL_STATUS
    config.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = COL_MINOR_CATEGORY
    config.Range(HOLIDAY_COL_DATE & HOLIDAY_DATA_START_ROW & ":" & HOLIDAY_COL_DESC & config.Rows.Count).ClearContents
    config.Cells(HOLIDAY_DATA_START_ROW, HOLIDAY_COL_DATE).Value = mMonday - 4
    config.Cells(HOLIDAY_DATA_START_ROW, HOLIDAY_COL_TYPE).Value2 = HOLIDAY_TYPE_HOLIDAY
    config.Cells(HOLIDAY_DATA_START_ROW + 1, HOLIDAY_COL_DATE).Value = mMonday - 2
    config.Cells(HOLIDAY_DATA_START_ROW + 1, HOLIDAY_COL_TYPE).Value2 = HOLIDAY_TYPE_WORKDAY
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "RefreshPerformanceFixture"
    SetupDataHeaders ws
    Seed ws, 5, 1, "Root group", mMonday - 7, mMonday + 9, Empty, Empty, 0
    Seed ws, 6, 2, "Branch group", mMonday - 7, mMonday + 4, Empty, Empty, 0
    Seed ws, 7, 3, "Completed delayed child", mMonday - 7, mMonday - 5, mMonday - 7, mMonday - 2, 0
    Seed ws, 8, 3, "Open child", mMonday, mMonday + 4, mMonday, Empty, 0.25
    Seed ws, 9, 2, "Future sibling", mMonday + 7, mMonday + 9, Empty, Empty, 0
    Seed ws, 10, 1, "Standalone completed", mMonday - 7, mMonday - 5, mMonday - 7, mMonday - 5, 0
    Seed ws, 11, 1, "Future standalone", mMonday + 7, mMonday + 11, Empty, Empty, 0
    ' Row 12 intentionally remains blank between tasks.
    Seed ws, 13, 1, "Invalid planned period", mMonday + 1, mMonday, Empty, Empty, 0
    Seed ws, 14, 1, "Manual hold task", mMonday - 7, mMonday - 5, Empty, Empty, 0
    ws.Cells(14, COL_MANUAL_STATUS).Value2 = STATUS_HOLD
    mStage = "initial full"
    RefreshGanttSheet False
    ws.Outline.ShowLevels ColumnLevels:=3
    RefreshWebModeView ws
    ws.Columns(COL_WEB_MODULE).Hidden = True
    ws.Rows(10).Hidden = True
    RefreshWebModeView ws
    CheckView ws
    Check ws.Rows(7).OutlineLevel = 3 And ws.Rows(8).OutlineLevel = 3, "initial hierarchy outline"
    Check CDbl(ws.Cells(7, COL_PROGRESS).Value2) = 1, "actual end did not normalize progress"
    Check CStr(ws.Cells(7, COL_STATUS).Value2) = STATUS_DONE, "completed status changed"
    CheckCorner DayCell(ws, 7, mMonday - 4)
    Check DayCell(ws, 7, mMonday - 2).Borders(xlEdgeRight).Color = RGB(22, 163, 74), "completed status right border changed"
    Check DayCell(ws, 5, mMonday - 7).Borders(xlEdgeTop).Color = RGB(99, 102, 241), "root frame missing"
    Check DayCell(ws, 11, mMonday - 1).Interior.Color = RGB(245, 245, 245), "Sunday background changed"
    Check DayCell(ws, 11, Date).Interior.Color = RGB(224, 252, 244), "today background changed"
    Check Len(CStr(ws.Cells(5, COL_ACTUAL_END).Value2)) = 0, "unfinished parent received actual end"
    Check Not HasTaskContent(ws, 12), "blank fixture row gained task content"
    Check CStr(ws.Cells(13, COL_STATUS).Value2) = STATUS_ERROR, "reversed plan did not retain error status"
    Check ws.Cells(13, COL_REQUEST).Interior.Color = RGB(255, 242, 204), "error row warning fill changed"
    Check CStr(ws.Cells(14, COL_STATUS).Value2) = STATUS_HOLD, "manual hold status changed"
    Check ws.Cells(14, COL_REQUEST).Interior.Color = RGB(217, 210, 233), "manual hold fill changed"
    Set baseline = Snapshot(ws)
    mStage = "unchanged first"
    RefreshGanttSheet False
    Set current = Snapshot(ws)
    For Each key In baseline.Keys
        Check CStr(baseline(key)) = CStr(current(key)), "unchanged first changed " & CStr(key)
    Next key
    CheckRepeatedRefresh ws, "unchanged"
    ws.Cells(7, COL_ACTUAL_END).Value = mMonday - 1
    CheckRepeatedRefresh ws, "actual date"
    CheckCorner DayCell(ws, 7, mMonday - 1)
    progressBefore = CDbl(ws.Cells(5, COL_PROGRESS).Value2)
    ws.Cells(8, COL_PROGRESS).Value2 = 0.75
    CheckRepeatedRefresh ws, "progress"
    Check CDbl(ws.Cells(5, COL_PROGRESS).Value2) > progressBefore, "parent progress stale"
    ws.Cells(8, COL_REQUEST).Value2 = "Updated request caption with additional context"
    ws.Cells(8, COL_MODIFICATION).Value2 = "Updated correction notes"
    CheckRepeatedRefresh ws, "text"
    Check ws.Cells(8, COL_REQUEST).Value2 = "Updated request caption with additional context", "request text changed"
    ws.Rows(6).ShowDetail = False
    CheckRepeatedRefresh ws, "collapsed branch"
    Check ws.Rows(7).Hidden And ws.Rows(8).Hidden, "collapsed descendants exposed"
    ws.Rows(6).ShowDetail = True
    ws.Cells(8, COL_LEVEL).Value2 = 2
    CheckRepeatedRefresh ws, "hierarchy level"
    Check ws.Rows(7).OutlineLevel = 3 And ws.Rows(8).OutlineLevel = 2, "promoted row outline stale"
    Check ws.Cells(8, COL_REQUEST).IndentLevel = 1, "promoted row indentation stale"
    config.Cells(HOLIDAY_DATA_START_ROW + 2, HOLIDAY_COL_DATE).Value = mMonday + 8
    config.Cells(HOLIDAY_DATA_START_ROW + 2, HOLIDAY_COL_TYPE).Value2 = HOLIDAY_TYPE_HOLIDAY
    CheckRepeatedRefresh ws, "holiday"
    Check DayCell(ws, 9, mMonday + 8).Interior.Color = RGB(245, 245, 245), "new holiday background stale"
    ws.Outline.ShowLevels ColumnLevels:=2
    RefreshWebModeView ws
    mExpectedWebMode = False
    mExpectedModuleHidden = True
    ws.Cells(8, COL_REQUEST).Value2 = "Hidden web columns text change"
    CheckRepeatedRefresh ws, "stage2 hidden web columns and text"
    ws.Outline.ShowLevels ColumnLevels:=3
    RefreshWebModeView ws
    mExpectedWebMode = True
    mExpectedModuleHidden = False
    CheckRepeatedRefresh ws, "stage3 restored web column widths"
    RunRefreshPerformanceRegression = "PASS: unchanged/date/progress/text/level/calendar repeated refresh preserves calculations, fills, gradients, borders, outlines, hidden view, widths, heights, web mode and protection"
    Exit Function
Failed:
    RunRefreshPerformanceRegression = "FAIL: " & mStage & ": " & Err.Description
End Function

Public Function GetRefreshScenarioSnapshots() As String
    GetRefreshScenarioSnapshots = mScenarioSnapshots
End Function
