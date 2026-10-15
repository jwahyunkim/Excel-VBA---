Attribute VB_Name = "modGanttBars"
Option Explicit

Private Const BAR_PREFIX As String = "gtBar_"
Private Const KEY_PREFIX As String = "gtKey_"

' Remove objects left by the previous implementation. All current output
' uses native cells and formatting; no drawing objects are created.
Public Sub ClearTaskBarVisuals(ByVal ws As Worksheet)
    Dim i As Long, shapeName As String
    For i = ws.Shapes.Count To 1 Step -1
        shapeName = ws.Shapes(i).Name
        If Left$(shapeName, Len(BAR_PREFIX)) = BAR_PREFIX Or _
           Left$(shapeName, Len(KEY_PREFIX)) = KEY_PREFIX Then ws.Shapes(i).Delete
    Next i
    ClearCellLegend ws
End Sub

Public Sub RenderHierarchyTaskBars(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                    ByVal startCol As Long, ByVal endCol As Long, _
                                    Optional ByVal holidayDict As Object = Nothing, _
                                    Optional ByVal workdayDict As Object = Nothing)
    Dim holidays As Object, workdays As Object
    Dim firstDate As Date, lastDate As Date, d As Date
    Dim r As Long, c As Long, level As Long
    Dim inRootGroup As Boolean, hasChildren As Boolean
    Dim planStart As Variant, planEnd As Variant
    Dim actualStart As Variant, actualEnd As Variant, drawEnd As Variant
    Dim actualMarker As Variant, planMarker As Variant
    Dim blankFill As Long
    Dim cell As Range, timeline As Range

    ClearTaskBarVisuals ws
    If lastRow < DATA_START_ROW Or endCol < startCol Then Exit Sub
    If Not IsDate(ws.Cells(GANTT_HEADER_ROW_DATE, startCol).Value) Then Exit Sub
    If Not IsDate(ws.Cells(GANTT_HEADER_ROW_DATE, endCol).Value) Then Exit Sub
    firstDate = DateValue(ws.Cells(GANTT_HEADER_ROW_DATE, startCol).Value)
    lastDate = DateValue(ws.Cells(GANTT_HEADER_ROW_DATE, endCol).Value)
    If holidayDict Is Nothing Or workdayDict Is Nothing Then
        LoadHolidaySettings holidays, workdays
    Else
        Set holidays = holidayDict
        Set workdays = workdayDict
    End If
    ws.Range(ws.Cells(1, startCol), ws.Cells(lastRow, endCol)).EntireColumn.ColumnWidth = 4.5

    For r = DATA_START_ROW To lastRow
        level = GetTaskLevel(ws, r)
        hasChildren = HasChildTask(ws, r, lastRow)
        If HasTaskContent(ws, r) Then
            If level = 1 Then inRootGroup = hasChildren
        End If
        blankFill = TimelineBackground(level, hasChildren, inRootGroup)
        If Not HasTaskContent(ws, r) Then blankFill = RGB(255, 255, 255)
        Set timeline = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
        timeline.ClearContents
        With timeline
            .NumberFormat = "@"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .IndentLevel = 0
            .WrapText = False
            .Font.Bold = False
            .Font.Size = 11
            .Font.Color = RGB(51, 65, 85)
        End With
        For c = startCol To endCol
            d = firstDate + c - startCol
            Set cell = ws.Cells(r, c)
            cell.Interior.Color = blankFill
            ' Preserve structural group/month rules while clearing stale status borders.
            If IsStatusBorder(cell.Borders(xlEdgeBottom).Color) Then
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(221, 229, 237)
                End With
            End If
            If Not IsWorkingDay(d, holidays, workdays) Then
                cell.Interior.Color = RGB(245, 245, 245)
            ElseIf d = Date Then
                cell.Interior.Color = RGB(255, 199, 206)
            End If
        Next c
        If Not HasTaskContent(ws, r) Then GoTo NextTask
        If HasTaskError(ws, r) Then GoTo NextTask

        planStart = ws.Cells(r, COL_PLAN_START).Value
        planEnd = ws.Cells(r, COL_PLAN_END).Value
        actualStart = ws.Cells(r, COL_ACTUAL_START).Value
        actualEnd = ws.Cells(r, COL_ACTUAL_END).Value
        drawEnd = Empty
        If IsDate(actualStart) Then
            If IsDate(actualEnd) Then
                drawEnd = DateValue(actualEnd)
            ElseIf DateValue(actualStart) <= Date Then
                drawEnd = Date
            End If
        End If
        actualMarker = FirstVisibleWorkingDate(actualStart, drawEnd, firstDate, lastDate, holidays, workdays)
        planMarker = FirstVisibleWorkingDate(planStart, planEnd, firstDate, lastDate, holidays, workdays)

        For c = startCol To endCol
            d = firstDate + c - startCol
            If Not IsWorkingDay(d, holidays, workdays) Then GoTo NextDate
            Set cell = ws.Cells(r, c)
            If DateFallsInPeriod(d, planStart, planEnd) Then
                cell.Interior.Color = RGB(203, 213, 225)
                cell.Font.Color = RGB(71, 85, 105)
                If IsEmpty(actualMarker) And Not IsEmpty(planMarker) Then
                    If d = planMarker Then SetLevelCaption cell, level, False
                End If
            End If
            If DateFallsInPeriod(d, actualStart, drawEnd) Then
                cell.Interior.Color = LevelBarColor(level)
                cell.Font.Color = RGB(255, 255, 255)
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = StatusBorderColor(CStr(ws.Cells(r, COL_STATUS).Value))
                End With
                If Not IsEmpty(actualMarker) Then
                    If d = actualMarker Then SetLevelCaption cell, level, True
                End If
            End If
NextDate:
        Next c
NextTask:
    Next r
    DrawCellLegend ws
End Sub

Private Function DateFallsInPeriod(ByVal d As Date, ByVal first As Variant, ByVal last As Variant) As Boolean
    If Not IsDate(first) Or Not IsDate(last) Then Exit Function
    DateFallsInPeriod = (d >= DateValue(first) And d <= DateValue(last))
End Function

Private Function FirstVisibleWorkingDate(ByVal periodStart As Variant, ByVal periodEnd As Variant, _
                                         ByVal firstDate As Date, ByVal lastDate As Date, _
                                         ByVal holidays As Object, ByVal workdays As Object) As Variant
    Dim d As Date, first As Date, last As Date
    FirstVisibleWorkingDate = Empty
    If Not IsDate(periodStart) Or Not IsDate(periodEnd) Then Exit Function
    first = DateValue(periodStart)
    last = DateValue(periodEnd)
    If first < firstDate Then first = firstDate
    If last > lastDate Then last = lastDate
    If first > last Then Exit Function
    For d = first To last
        If IsWorkingDay(d, holidays, workdays) Then
            FirstVisibleWorkingDate = d
            Exit Function
        End If
    Next d
End Function

Private Sub SetLevelCaption(ByVal cell As Range, ByVal level As Long, ByVal actual As Boolean)
    cell.Value2 = CStr(level)
    cell.Font.Bold = True
    cell.Font.Size = 14
    If actual Then cell.Font.Color = RGB(255, 255, 255)
End Sub

Private Function TimelineBackground(ByVal level As Long, ByVal hasChildren As Boolean, _
                                     ByVal inRootGroup As Boolean) As Long
    TimelineBackground = RGB(255, 255, 255)
    If inRootGroup Then TimelineBackground = RGB(243, 247, 251)
    If inRootGroup And level = 3 Then TimelineBackground = RGB(250, 252, 254)
    If level = 1 And hasChildren Then TimelineBackground = RGB(226, 234, 242)
    If level = 2 And hasChildren Then TimelineBackground = RGB(236, 242, 248)
End Function

Private Function LevelBarColor(ByVal level As Long) As Long
    Select Case level
        Case 1: LevelBarColor = RGB(36, 59, 83)
        Case 2: LevelBarColor = RGB(52, 120, 212)
        Case Else: LevelBarColor = RGB(0, 140, 134)
    End Select
End Function

Private Function StatusBorderColor(ByVal status As String) As Long
    If status = STATUS_DONE Or status = STATUS_DONE & STATUS_DONE_WITH_HOLD_SUFFIX Then
        StatusBorderColor = RGB(22, 163, 74)
        Exit Function
    End If
    Select Case status
        Case STATUS_NORMAL: StatusBorderColor = RGB(37, 99, 235)
        Case STATUS_CAUTION: StatusBorderColor = RGB(245, 158, 11)
        Case STATUS_HOLD: StatusBorderColor = RGB(139, 92, 246)
        Case STATUS_DELAY: StatusBorderColor = RGB(220, 38, 38)
        Case STATUS_ERROR: StatusBorderColor = RGB(234, 88, 12)
        Case Else: StatusBorderColor = RGB(100, 116, 139)
    End Select
End Function

Private Function IsStatusBorder(ByVal color As Long) As Boolean
    Select Case color
        Case RGB(22, 163, 74), RGB(37, 99, 235), RGB(245, 158, 11), _
             RGB(139, 92, 246), RGB(220, 38, 38), RGB(234, 88, 12), RGB(100, 116, 139)
            IsStatusBorder = True
    End Select
End Function

Private Function LegendCaption(ByVal index As Long) As String
    Select Case index
        Case 1: LegendCaption = "1 상위"
        Case 2: LegendCaption = "2 하위"
        Case 3: LegendCaption = "3 세부"
        Case 4: LegendCaption = "계획"
        Case 5: LegendCaption = "상태: 아래 테두리"
    End Select
End Function

Private Sub ClearCellLegend(ByVal ws As Worksheet)
    Dim addresses As Variant, i As Long, cell As Range
    addresses = Array("L3", "N3", "P3", "R3", "T3")
    For i = 0 To 4
        Set cell = ws.Range(addresses(i))
        If CStr(cell.Value2) = LegendCaption(i + 1) Then
            cell.ClearContents
            cell.ClearFormats
        End If
    Next i
End Sub

Private Sub DrawCellLegend(ByVal ws As Worksheet)
    Dim addresses As Variant, i As Long, cell As Range
    If Application.CountA(ws.Range("L3:W3")) > 0 Then Exit Sub
    addresses = Array("L3", "N3", "P3", "R3", "T3")
    For i = 0 To 4
        Set cell = ws.Range(addresses(i))
        With cell
            .NumberFormat = "@"
            .Value2 = LegendCaption(i + 1)
            .Font.Size = 9
            .Font.Bold = (i < 3)
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .WrapText = False
        End With
        If i < 3 Then
            cell.Interior.Color = LevelBarColor(i + 1)
            cell.Font.Color = RGB(255, 255, 255)
        ElseIf i = 3 Then
            cell.Interior.Color = RGB(203, 213, 225)
            cell.Font.Color = RGB(71, 85, 105)
        Else
            cell.HorizontalAlignment = xlLeft
            cell.Font.Color = RGB(51, 65, 85)
        End If
    Next i
    If ws.Rows(3).RowHeight < 24 Then ws.Rows(3).RowHeight = 24
End Sub
