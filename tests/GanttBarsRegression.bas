Attribute VB_Name = "GanttBarsRegression"
Option Explicit

Private Const FIRST_ROW As Long = 5
Private Const LAST_ROW As Long = 19
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
    SeedLegacyLegend ws
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
    ' The final visible working date must carry the status when the end is a
    ' holiday or Sunday. Saturday is explicitly a workday in this calendar.
    ws.Range("L18").Value = DateSerial(2026, 10, 12)
    ws.Range("M18").Value = DateSerial(2026, 10, 14)
    ws.Range("N18").Value = DateSerial(2026, 10, 12)
    ws.Range("O18").Value = DateSerial(2026, 10, 14)
    ws.Range("L19").Value = DateSerial(2026, 10, 16)
    ws.Range("M19").Value = DateSerial(2026, 10, 18)
    ws.Range("N19").Value = DateSerial(2026, 10, 16)
    ws.Range("O19").Value = DateSerial(2026, 10, 18)
    ' Old bottom status lines must be cleaned when upgrading the renderer.
    With ws.Range("X6:AI6").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(139, 92, 246)
    End With
    SeedDetachedBarBorders ws

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

    dataBefore = ws.Range("C5:W19").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Shapes.Count = 1, "cell renderer added objects or left legacy objects"
    VerifyPreservedButton ws, preserved.Name
    VerifyLegendCleared ws
    VerifySurroundingCells ws
    VerifyAllRows ws, firstDate, lastDate
    VerifyCombinedFrameEdges ws
    dataAfter = ws.Range("C5:W19").Value2
    VerifyInputValues dataBefore, dataAfter

    shapeCount = ws.Shapes.Count
    timelineBefore = TimelineSignature(ws, FIRST_ROW, LAST_ROW)
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check ws.Shapes.Count = shapeCount, "repeated cell render added objects"
    Check TimelineSignature(ws, FIRST_ROW, LAST_ROW) = timelineBefore, "repeated render changed cell visuals"
    VerifyAllRows ws, firstDate, lastDate
    VerifyCombinedFrameEdges ws

    ' Native fills and right status borders survive hide/unhide and outline collapse.
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

    ' Move both periods: old fills and their former right status must disappear.
    ws.Cells(6, "L").Value = firstDate + 8
    ws.Cells(6, "M").Value = lastDate
    ws.Cells(6, "N").Value = firstDate + 7
    ws.Cells(6, "O").Value = firstDate + 8
    dataBefore = ws.Range("C5:W19").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    VerifyExpectedRow ws, 6, firstDate + 8, lastDate, firstDate + 7, firstDate + 8, True
    VerifyCombinedFrameEdges ws
    Check Not IsBarFill(ws.Cells(6, FIRST_COL).Interior.Color), "old period fill survived date change"
    Check Not HasStatusBorder(ws.Cells(6, LAST_COL), xlEdgeRight), "old right status survived date change"
    Check ws.Shapes.Count = shapeCount, "date change render added objects"
    dataAfter = ws.Range("C5:W19").Value2
    VerifyInputValues dataBefore, dataAfter
    VerifyLegendCleared ws
    VerifySurroundingCells ws

    ' Removing actual dates must clear the latest right status while retaining plan fill.
    ws.Range("N6:O6").ClearContents
    dataBefore = ws.Range("C5:W19").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    VerifyExpectedRow ws, 6, firstDate + 8, lastDate, Empty, Empty, True
    VerifyCombinedFrameEdges ws
    Check ws.Shapes.Count = shapeCount, "actual removal render added objects"
    dataAfter = ws.Range("C5:W19").Value2
    VerifyInputValues dataBefore, dataAfter
    VerifyLegendCleared ws
    VerifySurroundingCells ws

    VerifyCombinedHierarchyFrames ws, holidays, workdays

    ClearTaskBarVisuals ws
    Check ws.Shapes.Count = 1, "cleanup deleted unrelated objects or left legacy objects"
    VerifyPreservedButton ws, preserved.Name
    VerifyLegendCleared ws
    VerifySurroundingCells ws
    ClearTaskBarVisuals ws
    Check ws.Shapes.Count = 1, "repeated cleanup changed unrelated object"
    VerifyLegendCleared ws
    result = "PASS: one combined planned/actual rectangle per Level 1 subtree, no inner Level 2 or Level 3 outlines, bridged calendar gaps, right status priority, all-edge stale cleanup, no legend or markers, input preservation, hide/outline restore and no inserted objects"
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
    VerifyExpectedRow ws, 18, DateSerial(2026, 10, 12), DateSerial(2026, 10, 14), _
                      DateSerial(2026, 10, 12), DateSerial(2026, 10, 14), True
    VerifyExpectedRow ws, 19, DateSerial(2026, 10, 16), DateSerial(2026, 10, 18), _
                      DateSerial(2026, 10, 16), DateSerial(2026, 10, 18), True
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
    Dim c As Long, level As Long, lastActualCol As Long
    Dim statusCount As Long, actualCount As Long, planCount As Long
    Dim d As Date, actual As Boolean, planned As Boolean, label As String
    Dim cell As Range
    level = CLng(ws.Cells(rowNum, "C").Value2)
    ' The status ends at the final visible working date, not a holiday,
    ' weekend or actual date beyond the current display boundary.
    For c = FIRST_COL To LAST_COL
        d = ws.Cells(4, c).Value
        If validTask And ExpectedWorkingDay(d) Then
            If InExpectedPeriod(d, actualStart, actualEnd) Then lastActualCol = c
        End If
    Next c

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
        Check Len(CStr(cell.Value2)) = 0, "timeline contains level number or symbol " & label
        Check Not HasStatusBorder(cell, xlEdgeBottom), "bottom status border remains " & label
        If c = lastActualCol Then
            Check cell.Borders(xlEdgeRight).LineStyle = xlContinuous, "missing right status border " & label
            Check cell.Borders(xlEdgeRight).Weight = xlMedium, "right status border weight " & label
            Check cell.Borders(xlEdgeRight).Color = ExpectedStatusColor(CStr(ws.Cells(rowNum, "W").Value2)), _
                  "right status border color " & label
            statusCount = statusCount + 1
        Else
            Check Not HasStatusBorder(cell, xlEdgeRight), "extra or stale right status border " & label
        End If
    Next c
    Check statusCount = IIf(lastActualCol > 0, 1, 0), "right status border count row " & rowNum
    If rowNum = 5 Then Check actualCount = 10, "working calendar coverage including Saturday override"
    If rowNum = 8 Then Check planCount = 10 And actualCount = 0, "planned-only calendar coverage"
End Sub

Private Function HasStatusBorder(ByVal cell As Range, ByVal edge As XlBordersIndex) As Boolean
    If cell.Borders(edge).LineStyle = xlNone Then Exit Function
    Select Case cell.Borders(edge).Color
        Case RGB(22, 163, 74), RGB(37, 99, 235), RGB(245, 158, 11), _
             RGB(139, 92, 246), RGB(220, 38, 38), RGB(234, 88, 12), RGB(100, 116, 139)
            HasStatusBorder = True
    End Select
End Function

Private Function HasOwnedBarBorder(ByVal cell As Range, ByVal edge As XlBordersIndex) As Boolean
    If cell.Borders(edge).LineStyle = xlNone Then Exit Function
    If HasStatusBorder(cell, edge) Then
        HasOwnedBarBorder = True
        Exit Function
    End If
    Select Case cell.Borders(edge).Color
        Case RGB(124, 58, 237), RGB(6, 182, 212), RGB(13, 148, 136), RGB(99, 102, 241)
            HasOwnedBarBorder = True
    End Select
End Function

Private Sub SeedDetachedBarBorders(ByVal ws As Worksheet)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With ws.Range("AI17").Borders(CLng(edge))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(124, 58, 237)
        End With
        With ws.Range("Z17").Borders(CLng(edge))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(6, 182, 212)
        End With
        With ws.Range("AD17").Borders(CLng(edge))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(13, 148, 136)
        End With
        With ws.Range("AE17").Borders(CLng(edge))
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(99, 102, 241)
        End With
    Next edge
End Sub

' A frame deliberately crosses holidays and gaps between plan and actual.
' Check its geometric perimeter, including Excel's shared adjacent edges,
' rather than requiring every colored edge to touch a filled cell.
Private Sub VerifyCombinedFrameEdges(ByVal ws As Worksheet, Optional ByVal hierarchyFixture As Boolean = False)
    Dim r As Long, c As Long, frameIndex As Long, frameRow As Long, frameCol As Long
    Dim edge As Variant, frames As Variant, frame As Variant, permitted As Boolean
    Dim minCols() As Long, maxCols() As Long, actualCol As Long
    If hierarchyFixture Then
        frames = Array(Array(5, 10), Array(11, 11))
    Else
        frames = Array(Array(5, 11), Array(12, 19))
    End If
    ReDim minCols(0 To UBound(frames))
    ReDim maxCols(0 To UBound(frames))
    For frameIndex = 0 To UBound(frames)
        frame = frames(frameIndex)
        For frameRow = CLng(frame(0)) To CLng(frame(1))
            For frameCol = FIRST_COL To LAST_COL
                If IsBarFill(ws.Cells(frameRow, frameCol).Interior.Color) Then
                    If minCols(frameIndex) = 0 Or frameCol < minCols(frameIndex) Then minCols(frameIndex) = frameCol
                    If frameCol > maxCols(frameIndex) Then maxCols(frameIndex) = frameCol
                End If
            Next frameCol
        Next frameRow
    Next frameIndex

    For r = FIRST_ROW To LAST_ROW
        For c = FIRST_COL To LAST_COL
            For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
                If HasOwnedBarBorder(ws.Cells(r, c), CLng(edge)) Then
                    permitted = False
                    If HasStatusBorder(ws.Cells(r, c), CLng(edge)) Then
                        actualCol = 0
                        If CLng(edge) = xlEdgeRight Then actualCol = c
                        If CLng(edge) = xlEdgeLeft And c > FIRST_COL Then actualCol = c - 1
                        If actualCol > 0 Then
                            permitted = (ws.Cells(r, actualCol).Interior.Color = _
                                         ExpectedLevelColor(CLng(ws.Cells(r, "C").Value2)))
                        End If
                    ElseIf ws.Cells(r, c).Borders(CLng(edge)).Color = RGB(99, 102, 241) Then
                        For frameIndex = 0 To UBound(frames)
                            If minCols(frameIndex) > 0 Then
                                frame = frames(frameIndex)
                                permitted = permitted Or EdgeTouchesRectangle(r, c, CLng(edge), _
                                    CLng(frame(0)), CLng(frame(1)), minCols(frameIndex), maxCols(frameIndex))
                            End If
                        Next frameIndex
                    End If
                    Check permitted, "border outside combined rectangle " & ws.Cells(r, c).Address & " edge " & CLng(edge)
                End If
            Next edge
        Next c
    Next r
End Sub

Private Function EdgeTouchesRectangle(ByVal rowNum As Long, ByVal colNum As Long, _
                                      ByVal edge As XlBordersIndex, ByVal firstRow As Long, _
                                      ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As Boolean
    Select Case edge
        Case xlEdgeTop
            EdgeTouchesRectangle = (rowNum = firstRow Or rowNum = lastRow + 1) And colNum >= firstCol And colNum <= lastCol
        Case xlEdgeBottom
            EdgeTouchesRectangle = (rowNum = lastRow Or rowNum = firstRow - 1) And colNum >= firstCol And colNum <= lastCol
        Case xlEdgeLeft
            EdgeTouchesRectangle = (colNum = firstCol Or colNum = lastCol + 1) And rowNum >= firstRow And rowNum <= lastRow
        Case xlEdgeRight
            EdgeTouchesRectangle = (colNum = lastCol Or colNum = firstCol - 1) And rowNum >= firstRow And rowNum <= lastRow
    End Select
End Function

Private Sub CheckBarBorder(ByVal cell As Range, ByVal edge As XlBordersIndex, _
                           ByVal color As Long, ByVal weight As XlBorderWeight, ByVal label As String)
    Check cell.Borders(edge).LineStyle = xlContinuous, label & " line style"
    Check cell.Borders(edge).Color = color, label & " color"
    Check cell.Borders(edge).Weight = weight, label & " weight"
End Sub

Private Sub CheckNoOwnedBorders(ByVal cell As Range, ByVal label As String)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        Check Not HasOwnedBarBorder(cell, CLng(edge)), label & " edge " & CLng(edge)
    Next edge
End Sub

Private Sub VerifyCombinedHierarchyFrames(ByVal ws As Worksheet, ByVal holidays As Object, ByVal workdays As Object)
    Dim r As Long, c As Long, levels As Variant, dataBefore As Variant, dataAfter As Variant
    Dim frameColor As Long, signature As String
    frameColor = RGB(99, 102, 241)
    ws.Range("C5:W19").ClearContents
    levels = Array(1, 2, 3, 3, 2, 3, 1)
    For r = 5 To 11
        ws.Cells(r, "C").Value2 = levels(r - 5)
        ws.Cells(r, "H").Value2 = "Combined frame fixture"
        ws.Cells(r, "W").Value2 = STATUS_DONE
    Next r
    SetFixturePlan ws, 5, DateSerial(2026, 10, 15), DateSerial(2026, 10, 21)
    SetFixturePlan ws, 6, DateSerial(2026, 10, 16), DateSerial(2026, 10, 20)
    ' Planned X:Y and actual AE:AF are separated by an unfilled date gap.
    SetFixturePlan ws, 7, DateSerial(2026, 10, 12), DateSerial(2026, 10, 13)
    ws.Range("N7").Value = DateSerial(2026, 10, 19)
    ws.Range("O7").Value = DateSerial(2026, 10, 20)
    SetFixturePlan ws, 8, DateSerial(2026, 10, 16), DateSerial(2026, 10, 19)
    SetFixturePlan ws, 9, DateSerial(2026, 10, 12), DateSerial(2026, 10, 13)
    SetFixturePlan ws, 10, DateSerial(2026, 10, 13), DateSerial(2026, 10, 13)
    ' A separate Level 1 task must enclose plan and actual in its own rectangle.
    SetFixturePlan ws, 11, DateSerial(2026, 10, 12), DateSerial(2026, 10, 13)
    ws.Range("N11").Value = DateSerial(2026, 10, 22)
    ws.Range("O11").Value = DateSerial(2026, 10, 23)
    dataBefore = ws.Range("C5:W19").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    dataAfter = ws.Range("C5:W19").Value2
    VerifyInputValues dataBefore, dataAfter

    CheckBarBorder ws.Range("X5"), xlEdgeTop, frameColor, xlMedium, "root frame includes earlier child plan"
    CheckBarBorder ws.Range("Z5"), xlEdgeTop, frameColor, xlMedium, "root top bridges holiday before root fill"
    CheckBarBorder ws.Range("AG5"), xlEdgeRight, frameColor, xlMedium, "root right frame"
    CheckBarBorder ws.Range("AD10"), xlEdgeBottom, frameColor, xlMedium, "root bottom bridges Sunday"
    CheckBarBorder ws.Range("X7"), xlEdgeLeft, frameColor, xlMedium, "root left encloses Level 3 rows"
    CheckBarBorder ws.Range("AG8"), xlEdgeRight, frameColor, xlMedium, "root right encloses shorter child periods"
    CheckBarBorder ws.Range("AH11"), xlEdgeTop, frameColor, xlMedium, "separate Level 1 top bridges plan actual gap"
    CheckBarBorder ws.Range("Z11"), xlEdgeBottom, frameColor, xlMedium, "separate Level 1 bottom bridges holiday"
    CheckBarBorder ws.Range("AG11"), xlEdgeBottom, frameColor, xlMedium, "separate Level 1 bottom bridges empty date"
    CheckBarBorder ws.Range("X11"), xlEdgeLeft, frameColor, xlMedium, "separate Level 1 left starts at plan"
    CheckBarBorder ws.Range("AF7"), xlEdgeRight, RGB(22, 163, 74), xlMedium, "child right status survives outer frame"
    CheckBarBorder ws.Range("AI11"), xlEdgeRight, RGB(22, 163, 74), xlMedium, "standalone status right"
    Check Not HasOwnedBarBorder(ws.Range("Y7"), xlEdgeRight), "plan end gained a separate vertical outline"
    Check Not HasOwnedBarBorder(ws.Range("AE7"), xlEdgeLeft), "actual start gained a separate vertical outline"
    Check Not HasOwnedBarBorder(ws.Range("X6"), xlEdgeTop), "Level 2 gained its own top outline"
    Check Not HasOwnedBarBorder(ws.Range("Z6"), xlEdgeTop), "Level 2 gained a top outline across holiday"
    Check Not HasOwnedBarBorder(ws.Range("AD8"), xlEdgeBottom), "Level 2 gained its own bottom outline"
    Check Not HasOwnedBarBorder(ws.Range("AF8"), xlEdgeRight), "Level 2 gained its own right outline"
    Check Not HasOwnedBarBorder(ws.Range("X9"), xlEdgeTop), "second Level 2 gained its own top outline"
    Check Not HasOwnedBarBorder(ws.Range("AB7"), xlEdgeTop), "grouped leaf gained its own top outline"
    Check Not HasOwnedBarBorder(ws.Range("AD7"), xlEdgeTop), "interior Sunday gained a top outline"
    Check Not HasOwnedBarBorder(ws.Range("AD7"), xlEdgeBottom), "interior Sunday gained a bottom outline"
    Check Not HasOwnedBarBorder(ws.Range("Y11"), xlEdgeRight), "standalone plan end gained a separate outline"
    Check Not HasOwnedBarBorder(ws.Range("AH11"), xlEdgeLeft), "standalone actual start gained a separate outline"
    CheckNoOwnedBorders ws.Range("Z13"), "holiday outside combined frames"
    CheckNoOwnedBorders ws.Range("AI5"), "group enclosed dates beyond every period"
    VerifyCombinedFrameEdges ws, True
    Check Application.CountA(ws.Range("X5:AI19")) = 0, "frame inserted numbers or symbols"
    VerifyLegendCleared ws
    VerifySurroundingCells ws
    signature = TimelineSignature(ws, FIRST_ROW, LAST_ROW)
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    Check TimelineSignature(ws, FIRST_ROW, LAST_ROW) = signature, "frame redraw changed four-edge visuals"

    ' Move an isolated bar to verify all four former edges and its status clear.
    ws.Range("L11:O11").Value = DateSerial(2026, 10, 12)
    ws.Range("M11").Value = DateSerial(2026, 10, 13)
    ws.Range("O11").Value = DateSerial(2026, 10, 13)
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    CheckNoOwnedBorders ws.Range("AH11"), "moved bar former first date"
    CheckNoOwnedBorders ws.Range("AI11"), "moved bar former final date"
    CheckBarBorder ws.Range("Y11"), xlEdgeRight, RGB(22, 163, 74), xlMedium, "moved status right"
    VerifyCombinedFrameEdges ws, True

    ' Removing periods from every subtree must remove every owned edge.
    ws.Range("L5:O19").ClearContents
    dataBefore = ws.Range("C5:W19").Value2
    RenderHierarchyTaskBars ws, LAST_ROW, FIRST_COL, LAST_COL, holidays, workdays
    dataAfter = ws.Range("C5:W19").Value2
    VerifyInputValues dataBefore, dataAfter
    For r = FIRST_ROW To LAST_ROW
        For c = FIRST_COL To LAST_COL
            CheckNoOwnedBorders ws.Cells(r, c), "deleted subtree dates " & ws.Cells(r, c).Address
            Check Not IsBarFill(ws.Cells(r, c).Interior.Color), "deleted subtree fill remains"
        Next c
    Next r
    Check ws.Shapes.Count = 1, "frames inserted objects"
    VerifyLegendCleared ws
    VerifySurroundingCells ws
End Sub

Private Sub SetFixturePlan(ByVal ws As Worksheet, ByVal rowNum As Long, _
                           ByVal first As Date, ByVal last As Date)
    ws.Cells(rowNum, "L").Value = first
    ws.Cells(rowNum, "M").Value = last
End Sub

Private Sub SeedLegacyLegend(ByVal ws As Worksheet)
    Dim addresses As Variant, captions As Variant, i As Long, cell As Range
    addresses = Array("L3", "N3", "P3", "R3", "T3")
    captions = Array("1 상위", "2 하위", "3 세부", "계획", "상태: 아래 테두리")
    For i = 0 To UBound(addresses)
        Set cell = ws.Range(addresses(i))
        cell.Value2 = captions(i)
        cell.Interior.Color = RGB(203, 213, 225)
        cell.Font.Bold = True
    Next i
End Sub

Private Sub VerifyLegendCleared(ByVal ws As Worksheet)
    Dim c As Long
    For c = 12 To 23
        Check Len(CStr(ws.Cells(3, c).Value2)) = 0, "cleanup left native legend text"
        Check ws.Cells(3, c).Interior.Pattern = xlNone, "cleanup left native legend fill"
        Check Not ws.Cells(3, c).MergeCells, "renderer inserted merged legend cells"
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
                        cell.Borders(xlEdgeLeft).LineStyle & ":" & cell.Borders(xlEdgeLeft).Color & ":" & _
                        cell.Borders(xlEdgeLeft).Weight & ":" & _
                        cell.Borders(xlEdgeTop).LineStyle & ":" & cell.Borders(xlEdgeTop).Color & ":" & _
                        cell.Borders(xlEdgeTop).Weight & ":" & _
                        cell.Borders(xlEdgeBottom).LineStyle & ":" & cell.Borders(xlEdgeBottom).Color & ":" & _
                        cell.Borders(xlEdgeBottom).Weight & ":" & _
                        cell.Borders(xlEdgeRight).LineStyle & ":" & cell.Borders(xlEdgeRight).Color & ":" & _
                        cell.Borders(xlEdgeRight).Weight & ":" & _
                        cell.Font.Color & ":" & cell.Font.Bold & ":" & cell.NumberFormat
        Next c
    Next r
    TimelineSignature = signature
End Function
