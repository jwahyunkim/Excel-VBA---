Attribute VB_Name = "GanttBarsRegression"
Option Explicit

Private Const FIRST_ROW As Long = 5
Private Const LAST_ROW As Long = 17
Private Const FIRST_COL As Long = 24
Private Const LAST_COL As Long = 35

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 952, "GanttBarsRegression", label
End Sub

Public Function RunGanttBarsRegression() As String
    Dim ws As Worksheet, previousSheet As Object
    Dim holidays As Object, workdays As Object
    Dim firstDate As Date, lastDate As Date
    Dim oldEvents As Boolean, oldAlerts As Boolean
    Dim preserved As Shape, legacy As Shape
    Dim r As Long, c As Long, shapeCount As Long
    Dim dataBefore As Variant, dataAfter As Variant
    Dim rowHeight As Double, timelineBefore As String, result As String
    On Error GoTo Failed

    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Set previousSheet = ActiveSheet
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "GanttBarsTest_" & Format$(Now, "hhmmss")
    firstDate = DateSerial(2026, 10, 12)
    lastDate = DateSerial(2026, 10, 23)
    Set holidays = CreateObject("Scripting.Dictionary")
    Set workdays = CreateObject("Scripting.Dictionary")
    holidays.Add NormalizeDateKey(DateSerial(2026, 10, 14)), True
    workdays.Add NormalizeDateKey(DateSerial(2026, 10, 17)), True

    For c = FIRST_COL To LAST_COL
        ws.Cells(4, c).Value = firstDate + c - FIRST_COL
        ws.Columns(c).ColumnWidth = 4.5
    Next c
    ws.Range("C4").Value = "Level"
    ws.Range("H4").Value = "Request"
    ws.Range("A1").Value = "Native cell regression"
    ws.Range("A20").Value = "After timeline"
    ws.Range("K3").Value = "Legend left neighbor"
    ws.Range("X3").Value = "Legend right neighbor"
    For r = FIRST_ROW To LAST_ROW
        ws.Cells(r, "C").Value = 3
        ws.Cells(r, "H").Value = "Task"
        ws.Cells(r, "L").Value = firstDate
        ws.Cells(r, "M").Value = lastDate
        ws.Cells(r, "N").Value = firstDate
        ws.Cells(r, "O").Value = lastDate
        ws.Cells(r, "W").Value = STATUS_DONE
        ws.Rows(r).RowHeight = 23
    Next r
    ws.Cells(5, "C").Value = 1
    ws.Cells(6, "C").Value = 2
    ws.Rows(5).RowHeight = 30
    ws.Rows(6).RowHeight = 26
    ws.Cells(6, "W").Value = STATUS_HOLD
    ws.Cells(7, "W").Value = STATUS_CAUTION

    ' Plan only, ongoing actual, and future actual with no end.
    ws.Range("N8:O8").ClearContents
    ws.Cells(8, "W").Value = STATUS_NORMAL
    ws.Cells(9, "O").ClearContents
    ws.Cells(9, "W").Value = STATUS_NORMAL
    ws.Cells(10, "N").Value = Date + 2
    ws.Cells(10, "O").ClearContents
    ws.Cells(10, "W").Value = STATUS_NORMAL
    ' Invalid dates suppress both planned and actual fills.
    ws.Cells(11, "L").Value = firstDate + 8
    ws.Cells(11, "M").Value = firstDate
    ws.Cells(11, "W").Value = STATUS_ERROR
    ' This range crosses both display boundaries.
    ws.Cells(12, "C").Value = 1
    ws.Rows(12).RowHeight = 30
    ws.Range("L12").Value = firstDate - 4
    ws.Range("M12").Value = lastDate + 2
    ws.Range("N12").Value = firstDate - 4
    ws.Range("O12").Value = lastDate + 2
    ws.Cells(12, "W").Value = STATUS_DELAY
    ' Single days, offscreen period, holiday-only period, empty request.
    ws.Cells(13, "C").Value = 2
    ws.Rows(13).RowHeight = 26
    ws.Range("L13:O13").Value = DateSerial(2026, 10, 15)
    ws.Range("L14:O14").Value = DateSerial(2026, 10, 13)
    ws.Cells(14, "W").Value = STATUS_HOLD
    ws.Range("L15:N15").Value = lastDate + 1
    ws.Range("M15").Value = lastDate + 5
    ws.Range("O15").Value = lastDate + 5
    ws.Range("L16:O16").Value = DateSerial(2026, 10, 14)
    ws.Cells(17, "H").ClearContents

    ' Only fixtures insert objects: the renderer must remove legacy objects
    ' and retain the unrelated button without adding replacement shapes.
    Set preserved = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("A1").Left, _
                                      ws.Range("A1").Top, 30, 16)
    preserved.Name = "btnGanttRegressionPreserved"
    preserved.AlternativeText = "Unrelated action button"
    Set legacy = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("X5").Left, _
                                   ws.Range("X5").Top, 20, 10)
    legacy.Name = "gtBar_legacy_fixture"
    Set legacy = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("L3").Left, _
                                   ws.Range("L3").Top, 20, 10)
    legacy.Name = "gtKey_legacy_fixture"
    Check ws.Shapes.Count = 3, "shape fixture setup"

    dataBefore = ws.Range("C5:W17").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Shapes.Count = 1, "cell renderer added objects or left legacy objects"
    VerifyPreservedButton ws, preserved.Name
    VerifyLegend ws
    VerifySurroundingCells ws
    VerifyAllRows ws, firstDate, lastDate
    dataAfter = ws.Range("C5:W17").Value2
    VerifyInputValues dataBefore, dataAfter

    shapeCount = ws.Shapes.Count
    timelineBefore = TimelineSignature(ws, FIRST_ROW, LAST_ROW)
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Shapes.Count = shapeCount, "repeated cell render added objects"
    Check TimelineSignature(ws, FIRST_ROW, LAST_ROW) = timelineBefore, "repeated render changed cell visuals"
    VerifyAllRows ws, firstDate, lastDate

    ' Native text and fills must survive hide/unhide and outline collapse.
    timelineBefore = TimelineSignature(ws, 6, 7)
    rowHeight = ws.Rows(6).RowHeight
    ws.Rows(6).Hidden = True
    DoEvents
    Check ws.Rows(6).Hidden, "row hide failed"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "row hide changed native cell visuals"
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Rows(6).Hidden, "render expanded a hidden row"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "render changed hidden row cell visuals"
    ws.Rows(6).Hidden = False
    Check ws.Rows(6).RowHeight = rowHeight, "unhide changed row height"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "unhide changed native cell visuals"

    ws.Outline.SummaryRow = xlAbove
    ws.Rows("6:7").Group
    ws.Rows(5).ShowDetail = False
    DoEvents
    Check ws.Rows(6).Hidden And ws.Rows(7).Hidden, "outline collapse did not hide children"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "outline collapse changed child cell visuals"
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Rows(6).Hidden And ws.Rows(7).Hidden, "render expanded a collapsed group"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "render changed collapsed child cell visuals"
    ws.Rows(5).ShowDetail = True
    DoEvents
    Check Not ws.Rows(6).Hidden And Not ws.Rows(7).Hidden, "outline expansion did not restore children"
    Check TimelineSignature(ws, 6, 7) = timelineBefore, "outline expansion changed child cell visuals"

    ' Move both periods: old fills and their former marker must disappear.
    ws.Cells(6, "L").Value = firstDate + 8
    ws.Cells(6, "M").Value = lastDate
    ws.Cells(6, "N").Value = firstDate + 7
    ws.Cells(6, "O").Value = firstDate + 8
    dataBefore = ws.Range("C5:W17").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    VerifyExpectedRow ws, 6, firstDate + 8, lastDate, firstDate + 7, firstDate + 8, True
    Check Len(CStr(ws.Cells(6, FIRST_COL).Value2)) = 0, "old start marker survived date change"
    Check Not IsBarFill(ws.Cells(6, FIRST_COL).Interior.Color), "old period fill survived date change"
    Check ws.Shapes.Count = shapeCount, "date change render added objects"
    dataAfter = ws.Range("C5:W17").Value2
    VerifyInputValues dataBefore, dataAfter
    VerifyLegend ws
    VerifySurroundingCells ws

    ClearTaskBarVisuals ws
    Check ws.Shapes.Count = 1, "cleanup deleted unrelated objects or left legacy objects"
    VerifyPreservedButton ws, preserved.Name
    VerifyLegendCleared ws
    VerifySurroundingCells ws
    ClearTaskBarVisuals ws
    Check ws.Shapes.Count = 1, "repeated cleanup changed unrelated object"
    VerifyLegendCleared ws
    result = "PASS: native cell hierarchy, calendar coverage, overlap, clipping, markers, status borders, input preservation, stale-fill cleanup, hide/outline restore and no inserted objects"
    GoTo Cleanup
Failed:
    result = "FAIL: " & Err.Description
Cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    RunGanttBarsRegression = result
End Function

Private Sub VerifyAllRows(ByVal ws As Worksheet, ByVal firstDate As Date, ByVal lastDate As Date)
    VerifyExpectedRow ws, 5, firstDate, lastDate, firstDate, lastDate, True
    VerifyExpectedRow ws, 6, firstDate, lastDate, firstDate, lastDate, True
    VerifyExpectedRow ws, 7, firstDate, lastDate, firstDate, lastDate, True
    VerifyExpectedRow ws, 8, firstDate, lastDate, Empty, Empty, True
    VerifyExpectedRow ws, 9, firstDate, lastDate, firstDate, Date, True
    VerifyExpectedRow ws, 10, firstDate, lastDate, Empty, Empty, True
    VerifyExpectedRow ws, 11, Empty, Empty, Empty, Empty, False
    VerifyExpectedRow ws, 12, firstDate - 4, lastDate + 2, firstDate - 4, lastDate + 2, True
    VerifyExpectedRow ws, 13, DateSerial(2026, 10, 15), DateSerial(2026, 10, 15), _
                      DateSerial(2026, 10, 15), DateSerial(2026, 10, 15), True
    VerifyExpectedRow ws, 14, DateSerial(2026, 10, 13), DateSerial(2026, 10, 13), _
                      DateSerial(2026, 10, 13), DateSerial(2026, 10, 13), True
    VerifyExpectedRow ws, 15, lastDate + 1, lastDate + 5, lastDate + 1, lastDate + 5, True
    VerifyExpectedRow ws, 16, DateSerial(2026, 10, 14), DateSerial(2026, 10, 14), _
                      DateSerial(2026, 10, 14), DateSerial(2026, 10, 14), True
    VerifyExpectedRow ws, 17, Empty, Empty, Empty, Empty, False
End Sub

' Expected calendar is independent of the production IsWorkingDay helper.
Private Function ExpectedWorkingDay(ByVal d As Date) As Boolean
    If d = DateSerial(2026, 10, 17) Then
        ExpectedWorkingDay = True
    ElseIf d = DateSerial(2026, 10, 14) Then
        ExpectedWorkingDay = False
    Else
        ExpectedWorkingDay = (Weekday(d, vbMonday) <= 5)
    End If
End Function

Private Function InExpectedPeriod(ByVal d As Date, ByVal startDate As Variant, _
                                  ByVal endDate As Variant) As Boolean
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Function
    If CDate(startDate) > CDate(endDate) Then Exit Function
    InExpectedPeriod = (d >= CDate(startDate) And d <= CDate(endDate))
End Function

Private Function ExpectedLevelColor(ByVal level As Long) As Long
    Select Case level
        Case 1: ExpectedLevelColor = RGB(36, 59, 83)
        Case 2: ExpectedLevelColor = RGB(52, 120, 212)
        Case Else: ExpectedLevelColor = RGB(0, 140, 134)
    End Select
End Function

Private Function ExpectedStatusColor(ByVal status As String) As Long
    Select Case status
        Case STATUS_DONE: ExpectedStatusColor = RGB(22, 163, 74)
        Case STATUS_HOLD: ExpectedStatusColor = RGB(139, 92, 246)
        Case STATUS_CAUTION: ExpectedStatusColor = RGB(245, 158, 11)
        Case STATUS_DELAY: ExpectedStatusColor = RGB(220, 38, 38)
        Case STATUS_NORMAL: ExpectedStatusColor = RGB(37, 99, 235)
        Case Else: Err.Raise vbObjectError + 952, "GanttBarsRegression", "unknown test status"
    End Select
End Function

Private Function IsBarFill(ByVal color As Long) As Boolean
    IsBarFill = (color = RGB(36, 59, 83) Or color = RGB(52, 120, 212) Or _
                 color = RGB(0, 140, 134) Or color = RGB(203, 213, 225))
End Function

Private Sub VerifyExpectedRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                              ByVal planStart As Variant, ByVal planEnd As Variant, _
                              ByVal actualStart As Variant, ByVal actualEnd As Variant, _
                              ByVal validTask As Boolean)
    Dim c As Long, level As Long, firstActualCol As Long, firstPlanCol As Long
    Dim markerCol As Long, markerCount As Long, actualCount As Long, planCount As Long
    Dim d As Date, actual As Boolean, planned As Boolean, label As String
    Dim cell As Range
    level = CLng(ws.Cells(rowNum, "C").Value2)
    ' The marker is based on visible working dates, not stored start dates.
    For c = FIRST_COL To LAST_COL
        d = ws.Cells(4, c).Value
        If validTask And ExpectedWorkingDay(d) Then
            If InExpectedPeriod(d, actualStart, actualEnd) And firstActualCol = 0 Then firstActualCol = c
            If InExpectedPeriod(d, planStart, planEnd) And firstPlanCol = 0 Then firstPlanCol = c
        End If
    Next c
    markerCol = firstActualCol
    If markerCol = 0 Then markerCol = firstPlanCol

    For c = FIRST_COL To LAST_COL
        Set cell = ws.Cells(rowNum, c)
        d = ws.Cells(4, c).Value
        label = "row " & rowNum & " date " & Format$(d, "yyyy-mm-dd")
        actual = validTask And ExpectedWorkingDay(d) And InExpectedPeriod(d, actualStart, actualEnd)
        planned = validTask And ExpectedWorkingDay(d) And InExpectedPeriod(d, planStart, planEnd)
        Check Not cell.MergeCells, "merged timeline cell " & label
        If actual Then
            actualCount = actualCount + 1
            Check cell.Interior.Color = ExpectedLevelColor(level), "actual fill " & label
            Check cell.Borders(xlEdgeBottom).LineStyle <> xlNone, "missing status border " & label
            Check cell.Borders(xlEdgeBottom).Color = ExpectedStatusColor(CStr(ws.Cells(rowNum, "W").Value2)), _
                  "status border color " & label
        ElseIf planned Then
            planCount = planCount + 1
            Check cell.Interior.Color = RGB(203, 213, 225), "plan fill or actual overlap priority " & label
        Else
            Check Not IsBarFill(cell.Interior.Color), "unexpected or stale period fill " & label
            If Not ExpectedWorkingDay(d) Then
                Check cell.Interior.Color = RGB(245, 245, 245), "nonworking background " & label
            ElseIf d = Date Then
                Check cell.Interior.Color = RGB(255, 199, 206), "today background outside period " & label
            End If
        End If
        If c = markerCol Then
            Check CStr(cell.Value2) = CStr(level), "missing or misplaced level marker " & label
            Check cell.NumberFormat = "@", "marker is not text " & label
            Check cell.Font.Bold, "marker is not bold " & label
            Check cell.HorizontalAlignment = xlCenter And cell.VerticalAlignment = xlCenter, _
                  "marker is not centered " & label
            If firstActualCol > 0 Then
                Check cell.Font.Color = RGB(255, 255, 255), "actual marker lacks white text " & label
            Else
                Check cell.Font.Color <> RGB(255, 255, 255), "plan marker lacks dark text " & label
            End If
        Else
            Check Len(CStr(cell.Value2)) = 0, "extra text outside first marker " & label
        End If
        If Len(CStr(cell.Value2)) > 0 Then markerCount = markerCount + 1
    Next c
    Check markerCount = IIf(markerCol > 0, 1, 0), "level marker count row " & rowNum
    If rowNum = 5 Then Check actualCount = 10, "working calendar coverage including Saturday override"
    If rowNum = 8 Then Check planCount = 10 And actualCount = 0, "planned-only calendar coverage"
End Sub

Private Sub VerifyLegend(ByVal ws As Worksheet)
    Dim c As Long
    Check CStr(ws.Range("L3").Value2) = "1 상위", "level 1 native legend"
    Check CStr(ws.Range("N3").Value2) = "2 하위", "level 2 native legend"
    Check CStr(ws.Range("P3").Value2) = "3 세부", "level 3 native legend"
    Check CStr(ws.Range("R3").Value2) = "계획", "planned native legend"
    Check CStr(ws.Range("T3").Value2) = "상태: 아래 테두리", "status native legend"
    Check ws.Range("L3").Interior.Color = RGB(36, 59, 83), "level 1 legend fill"
    Check ws.Range("N3").Interior.Color = RGB(52, 120, 212), "level 2 legend fill"
    Check ws.Range("P3").Interior.Color = RGB(0, 140, 134), "level 3 legend fill"
    Check ws.Range("R3").Interior.Color = RGB(203, 213, 225), "plan legend fill"
    For c = 12 To 23
        Check Not ws.Cells(3, c).MergeCells, "legend inserted merged cells"
    Next c
End Sub

Private Sub VerifyLegendCleared(ByVal ws As Worksheet)
    Dim c As Long
    For c = 12 To 23
        Check Len(CStr(ws.Cells(3, c).Value2)) = 0, "cleanup left native legend text"
        Check ws.Cells(3, c).Interior.Pattern = xlNone, "cleanup left native legend fill"
    Next c
End Sub

Private Sub VerifySurroundingCells(ByVal ws As Worksheet)
    Check CStr(ws.Range("A1").Value2) = "Native cell regression", "renderer changed unrelated cell"
    Check CStr(ws.Range("A20").Value2) = "After timeline", "renderer inserted rows or changed outside cell"
    Check CStr(ws.Range("K3").Value2) = "Legend left neighbor", "renderer changed left legend neighbor"
    Check CStr(ws.Range("X3").Value2) = "Legend right neighbor", "renderer changed right legend neighbor"
    Check ws.Rows.Count = 1048576, "renderer changed worksheet row count"
End Sub

Private Sub VerifyPreservedButton(ByVal ws As Worksheet, ByVal expectedName As String)
    Check ws.Shapes.Count = 1, "unexpected inserted object count"
    Check ws.Shapes(1).Name = expectedName, "renderer removed unrelated button"
    Check ws.Shapes(1).AlternativeText = "Unrelated action button", "renderer changed unrelated button"
End Sub

Private Sub VerifyInputValues(ByVal dataBefore As Variant, ByVal dataAfter As Variant)
    Dim r As Long, c As Long
    For r = 1 To UBound(dataBefore, 1)
        For c = 1 To UBound(dataBefore, 2)
            Check CStr(dataBefore(r, c)) = CStr(dataAfter(r, c)), "renderer changed task input"
        Next c
    Next r
End Sub

Private Function TimelineSignature(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim r As Long, c As Long, cell As Range, signature As String
    For r = firstRow To lastRow
        For c = FIRST_COL To LAST_COL
            Set cell = ws.Cells(r, c)
            signature = signature & "|" & CStr(cell.Value2) & ":" & cell.Interior.Color & ":" & _
                        cell.Borders(xlEdgeBottom).LineStyle & ":" & cell.Borders(xlEdgeBottom).Color & ":" & _
                        cell.Font.Color & ":" & cell.Font.Bold & ":" & cell.NumberFormat
        Next c
    Next r
    TimelineSignature = signature
End Function
