Attribute VB_Name = "modGanttHierarchy"
Option Explicit

Public Sub InitializeTaskHierarchyOutlines()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." And _
           CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Level" Then
            ' UserInterfaceOnly and EnableOutlining are session settings;
            ' restore them when a saved, protected workbook is opened.
            UnprotectTaskSheet ws
            ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
        End If
    Next ws
End Sub

' Visual hierarchy and native outlines share the same subtree boundaries as
' the task summary calculation. Build them before applying display filters.
Public Sub ApplyTaskHierarchyAppearance(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                        ByVal startCol As Long, ByVal endCol As Long)
    Dim r As Long
    Dim c As Long
    Dim taskLevel As Long
    Dim groupEnd As Long
    Dim hasChildren As Boolean
    Dim inRootGroup As Boolean
    Dim isWarning As Boolean
    Dim statusText As String
    Dim rowFill As Long
    Dim chartFill As Long
    Dim rowRange As Range
    Dim taskRange As Range
    Dim chartCell As Range

    If lastRow < DATA_START_ROW Or endCol < startCol Then Exit Sub

    With ws.Range(COL_NO & DATA_START_ROW & ":" & COL_STATUS & lastRow).Font
        .Bold = False
        .Color = RGB(40, 50, 62)
    End With

    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then GoTo NextTask

        taskLevel = GetTaskLevel(ws, r)
        hasChildren = HasChildTask(ws, r, lastRow)
        If taskLevel = 1 Then inRootGroup = hasChildren
        statusText = CStr(ws.Cells(r, COL_STATUS).Value)
        isWarning = (statusText = STATUS_HOLD Or statusText = STATUS_DELAY Or _
                     statusText = STATUS_ERROR)

        Set rowRange = ws.Range(ws.Cells(r, ws.Range(COL_NO & "1").Column), _
                                ws.Cells(r, endCol))
        Set taskRange = ws.Range(COL_REQUEST & r & ":" & COL_MODIFICATION & r)
        taskRange.IndentLevel = taskLevel - 1
        taskRange.Font.Bold = (taskLevel < 3 And (inRootGroup Or hasChildren))

        rowFill = RGB(255, 255, 255)
        chartFill = rowFill
        If inRootGroup Then
            rowFill = RGB(243, 247, 251)
            chartFill = rowFill
            If taskLevel = 3 Then
                rowFill = RGB(250, 252, 254)
                chartFill = rowFill
            End If
        End If
        If taskLevel = 1 And hasChildren Then
            rowFill = RGB(226, 234, 242)
            chartFill = rowFill
        ElseIf taskLevel = 2 And hasChildren Then
            rowFill = RGB(222, 233, 243)
            chartFill = RGB(236, 242, 248)
        End If

        ' Warning rows and status cells retain their existing status colors.
        If Not isWarning Then
            ws.Range(COL_NO & r & ":" & COL_ACTUAL_DAYS & r).Interior.Color = rowFill
            If taskLevel = 1 And hasChildren Then
                With ws.Range(COL_NO & r & ":" & COL_MODIFICATION & r)
                    .Interior.Color = RGB(44, 72, 98)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With
            End If
        End If

        ' Only unfilled cells receive a group band. Date bars, holidays and
        ' today's empty-cell highlight have already been drawn by the caller.
        For c = startCol To endCol
            Set chartCell = ws.Cells(r, c)
            If chartCell.Interior.Pattern = xlNone Then
                chartCell.Interior.Color = chartFill
            End If
        Next c

        With rowRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(221, 229, 237)
        End With
        If taskLevel = 1 And hasChildren Then
            If ws.Rows(r).RowHeight < 30 Then ws.Rows(r).RowHeight = 30
            SetHierarchyRule rowRange, xlEdgeTop, RGB(68, 96, 122), xlMedium
            SetHierarchyRule rowRange, xlEdgeBottom, RGB(126, 150, 173), xlThin
        ElseIf taskLevel = 2 Then
            If ws.Rows(r).RowHeight < 26 Then ws.Rows(r).RowHeight = 26
        Else
            If ws.Rows(r).RowHeight < 23 Then ws.Rows(r).RowHeight = 23
        End If

NextTask:
    Next r

    ' Fit the final bold text and native indentation, not the old text prefixes.
    ws.Range(COL_REQUEST & HEADER_ROW & ":" & COL_MODIFICATION & lastRow).Columns.AutoFit

    ' Draw enclosing rails after all individual row borders are in place.
    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then GoTo NextGroup
        If Not HasChildTask(ws, r, lastRow) Then GoTo NextGroup
        groupEnd = GetVisibleTaskSubtreeEnd(ws, r, lastRow)
        taskLevel = GetTaskLevel(ws, r)

        If taskLevel = 1 Then
            Set rowRange = ws.Range(COL_NO & r & ":" & COL_NO & groupEnd)
            SetHierarchyRule rowRange, xlEdgeLeft, RGB(68, 96, 122), xlMedium
            Set rowRange = ws.Range(ws.Cells(groupEnd, ws.Range(COL_NO & "1").Column), _
                                    ws.Cells(groupEnd, endCol))
            SetHierarchyRule rowRange, xlEdgeBottom, RGB(68, 96, 122), xlMedium
        ElseIf taskLevel = 2 Then
            Set rowRange = ws.Range(COL_REQUEST & r & ":" & COL_REQUEST & groupEnd)
            SetHierarchyRule rowRange, xlEdgeLeft, RGB(143, 166, 187), xlMedium
            Set rowRange = ws.Range(ws.Cells(r, ws.Range(COL_REQUEST & "1").Column), _
                                    ws.Cells(r, endCol))
            SetHierarchyRule rowRange, xlEdgeTop, RGB(164, 183, 201), xlThin
        End If
NextGroup:
    Next r

    RebuildTaskHierarchyOutline ws, lastRow
    RenderHierarchyTaskBars ws, lastRow, startCol, endCol
End Sub

Private Sub SetHierarchyRule(ByVal target As Range, ByVal edge As XlBordersIndex, _
                             ByVal color As Long, ByVal weight As XlBorderWeight)
    With target.Borders(edge)
        .LineStyle = xlContinuous
        .Weight = weight
        .Color = color
    End With
End Sub

Private Function GetVisibleTaskSubtreeEnd(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                          ByVal lastRow As Long) As Long
    Dim r As Long
    r = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    Do While r > rowNum
        If HasTaskContent(ws, r) Then Exit Do
        r = r - 1
    Loop
    GetVisibleTaskSubtreeEnd = r
End Function

Public Sub RebuildTaskHierarchyOutline(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim taskLevel As Long
    Dim r As Long
    Dim groupEnd As Long
    Dim outlineLastRow As Long

    If lastRow < DATA_START_ROW Then Exit Sub
    ' Include rows left behind when the last tasks have been removed.
    outlineLastRow = Application.Max(lastRow, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(outlineLastRow)).ClearOutline
    ws.Outline.SummaryRow = xlAbove

    ' Inner groups first; the outer group adds one level to all descendants.
    For taskLevel = 2 To 1 Step -1
        For r = DATA_START_ROW To lastRow
            If HasTaskContent(ws, r) Then
                If GetTaskLevel(ws, r) = taskLevel And HasChildTask(ws, r, lastRow) Then
                    groupEnd = GetVisibleTaskSubtreeEnd(ws, r, lastRow)
                    If groupEnd > r Then ws.Rows(CStr(r + 1) & ":" & CStr(groupEnd)).Group
                End If
            End If
        Next r
    Next taskLevel
    ws.EnableOutlining = True
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayOutline = True
End Sub
