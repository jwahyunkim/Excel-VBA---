Attribute VB_Name = "modGanttHierarchy"
Option Explicit

Private mTaskOutlineViewEvents As CGanttOutlineView
Private mTaskOutlineViewStarting As Boolean

Public Sub InitializeTaskHierarchyOutlines()
    Dim ws As Worksheet
    Dim wasSaved As Boolean
    wasSaved = ThisWorkbook.Saved
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." And _
           CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Level" Then
            ' UserInterfaceOnly and EnableOutlining are session settings;
            ' restore them when a saved, protected workbook is opened.
            UnprotectTaskSheet ws
            EnsureTaskDetailColumnOutline ws
            EnsureTaskRequestFreeze ws
            ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
        End If
    Next ws
    ' Session-only protection restoration must not dirty a saved workbook.
    If wasSaved Then ThisWorkbook.Saved = True
End Sub

Public Sub EnsureTaskRequestFreeze(ByVal ws As Worksheet)
    Dim taskWindow As Window, requestCol As Long, previousRow As Long, c As Long
    Dim wasSaved As Boolean, collapsed As Boolean
    Dim detailColumns As Range, wasScreenUpdating As Boolean
    Dim viewError As Long, viewErrorText As String
    If ActiveWindow Is Nothing Then Exit Sub
    If Not (ActiveWorkbook Is ws.Parent) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) <> "No." Or _
       CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) <> "Level" Then Exit Sub

    Set taskWindow = ActiveWindow
    requestCol = ws.Range(COL_REQUEST & "1").Column
    collapsed = True
    For c = requestCol + 1 To ws.Range(COL_GANTT_START & "1").Column - 1
        If Not ws.Columns(c).Hidden Then
            collapsed = False
            Exit For
        End If
    Next c
    With taskWindow
        If collapsed Then
            If .FreezePanes And .SplitColumn = 1 And .SplitRow = HEADER_ROW Then
                If .Panes(1).ScrollColumn = requestCol Then Exit Sub
            End If
        ElseIf Not .FreezePanes And .SplitColumn = 0 And .SplitRow = 0 Then
            ' Restore A only on the transition; leave normal expanded scrolling free.
            Exit Sub
        End If
        wasSaved = ws.Parent.Saved
        wasScreenUpdating = Application.ScreenUpdating
        previousRow = .Panes(.Panes.Count).ScrollRow
        On Error GoTo Failed
        Application.ScreenUpdating = False
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0
        If collapsed Then
            ' Excel counts hidden columns differently when creating a split.
            ' Build the H-only freeze with details visible, then restore the fold.
            Set detailColumns = ws.Range(ws.Cells(1, requestCol + 1), _
                ws.Cells(1, ws.Range(COL_GANTT_START & "1").Column - 1)).EntireColumn
            detailColumns.Hidden = False
            .ScrollColumn = requestCol
            .ScrollRow = 1
            .SplitColumn = 1
            .SplitRow = HEADER_ROW
            .FreezePanes = True
            detailColumns.Hidden = True
            .Panes(.Panes.Count).ScrollRow = Application.Max(DATA_START_ROW, previousRow)
        Else
            .ScrollColumn = 1
            .ScrollRow = Application.Max(1, previousRow)
        End If
    End With
    If wasSaved Then ws.Parent.Saved = True
    Application.ScreenUpdating = wasScreenUpdating
    Exit Sub
Failed:
    viewError = Err.Number
    viewErrorText = Err.Description
    On Error Resume Next
    If Not detailColumns Is Nothing Then detailColumns.Hidden = True
    Application.ScreenUpdating = wasScreenUpdating
    If wasSaved Then ws.Parent.Saved = True
    On Error GoTo 0
    Err.Raise viewError, "EnsureTaskRequestFreeze", viewErrorText
End Sub

Private Function ActiveTaskOutlineSheet() As Worksheet
    If ActiveWindow Is Nothing Or ActiveWorkbook Is Nothing Then Exit Function
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Function
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Function
    If CStr(ActiveSheet.Cells(HEADER_ROW, COL_NO).Value2) <> "No." Or _
       CStr(ActiveSheet.Cells(HEADER_ROW, COL_LEVEL).Value2) <> "Level" Then Exit Function
    Set ActiveTaskOutlineSheet = ActiveSheet
End Function

Public Sub StartTaskOutlineViewMonitor()
    Dim ws As Worksheet
    If mTaskOutlineViewStarting Then Exit Sub
    mTaskOutlineViewStarting = True
    On Error GoTo Failed
    Set ws = ActiveTaskOutlineSheet()
    If ws Is Nothing Then
        StopTaskOutlineViewMonitor
        GoTo Finished
    End If
    If mTaskOutlineViewEvents Is Nothing Then
        Set mTaskOutlineViewEvents = New CGanttOutlineView
        mTaskOutlineViewEvents.StartListening
    End If
    EnsureTaskRequestFreeze ws
Finished:
    mTaskOutlineViewStarting = False
    Exit Sub
Failed:
    StopTaskOutlineViewMonitor
    Resume Finished
End Sub

Public Sub StopTaskOutlineViewMonitor()
    If Not mTaskOutlineViewEvents Is Nothing Then
        mTaskOutlineViewEvents.StopListening
        Set mTaskOutlineViewEvents = Nothing
    End If
End Sub

Public Sub TaskOutlineViewTick()
    Dim ws As Worksheet
    If mTaskOutlineViewStarting Then Exit Sub
    Set ws = ActiveTaskOutlineSheet()
    If Not ws Is Nothing Then EnsureTaskRequestFreeze ws
End Sub

Public Function CaptureTaskDetailColumnHidden(ByVal ws As Worksheet) As Variant
    Dim firstCol As Long, lastCol As Long, c As Long
    Dim hidden() As Boolean
    firstCol = ws.Range(COL_REQUEST & "1").Column + 1
    lastCol = ws.Range(COL_GANTT_START & "1").Column - 1
    If lastCol < firstCol Then Exit Function
    ReDim hidden(firstCol To lastCol)
    For c = firstCol To lastCol
        hidden(c) = ws.Columns(c).Hidden
    Next c
    CaptureTaskDetailColumnHidden = hidden
End Function

Public Sub RestoreTaskDetailColumnHidden(ByVal ws As Worksheet, ByVal hidden As Variant)
    Dim c As Long
    If Not IsArray(hidden) Then Exit Sub
    For c = LBound(hidden) To UBound(hidden)
        If ws.Columns(c).Hidden <> CBool(hidden(c)) Then
            ws.Columns(c).Hidden = CBool(hidden(c))
        End If
    Next c
    EnsureTaskRequestFreeze ws
End Sub

Private Sub UngroupTaskOutline(ByVal target As Range)
    Dim pass As Long, errorNumber As Long, errorText As String
    ' ClearOutline also clears the other axis. Ungroup respects full rows/columns.
    ' Excel supports at most eight levels, including the ungrouped base level.
    For pass = 1 To 7
        On Error Resume Next
        target.Ungroup
        errorNumber = Err.Number
        errorText = Err.Description
        Err.Clear
        On Error GoTo 0
        If errorNumber <> 0 Then
            If errorNumber <> 1004 Then Err.Raise errorNumber, "UngroupTaskOutline", errorText
            Exit For
        End If
    Next pass
End Sub

Public Sub EnsureTaskDetailColumnOutline(ByVal ws As Worksheet)
    Dim firstCol As Long, lastCol As Long, c As Long
    Dim needsGroup As Boolean, hidden As Variant, detailColumns As Range
    firstCol = ws.Range(COL_REQUEST & "1").Column + 1
    lastCol = ws.Range(COL_GANTT_START & "1").Column - 1
    If lastCol < firstCol Then Exit Sub
    Set detailColumns = ws.Range(ws.Cells(1, firstCol), ws.Cells(1, lastCol)).EntireColumn
    For c = firstCol To lastCol
        If ws.Columns(c).OutlineLevel <> 2 Then needsGroup = True
    Next c
    If needsGroup Then
        hidden = CaptureTaskDetailColumnHidden(ws)
        UngroupTaskOutline detailColumns
        detailColumns.Group
        RestoreTaskDetailColumnHidden ws, hidden
    End If
    ws.Outline.SummaryColumn = xlRight
    ws.EnableOutlining = True
    If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayOutline = True
End Sub

' Visual hierarchy and native outlines share the same subtree boundaries as
' the task summary calculation. Build them before applying display filters.
Public Sub ApplyTaskHierarchyAppearance(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                        ByVal startCol As Long, ByVal endCol As Long)
    Dim r As Long
    Dim c As Long
    Dim taskLevel As Long
    Dim hasChildren As Boolean
    Dim inRootGroup As Boolean
    Dim isWarning As Boolean
    Dim statusText As String
    Dim rowFill As Long
    Dim chartFill As Long
    Dim rowRange As Range
    Dim taskRange As Range
    Dim chartCell As Range

    EnsureTaskDetailColumnOutline ws
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
    UngroupTaskOutline ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(outlineLastRow))
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
