Attribute VB_Name = "modGanttHierarchy"
Option Explicit

Private mTaskOutlineViewEvents As CGanttOutlineView
Private mTaskOutlineViewStarting As Boolean
Private mTaskViewResetSheets As Object
Private Const TASK_FOLD_RANGE_NAME As String = "_GanttDetailFoldRange"

Private Function TaskFoldColumnNumber(ByVal value As Variant, ByVal columnCount As Long) As Long
    Dim letters As String, i As Long, code As Long, result As Long
    If IsError(value) Then Exit Function
    letters = UCase$(Trim$(CStr(value)))
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    For i = 1 To Len(letters)
        code = AscW(Mid$(letters, i, 1))
        If code < 65 Or code > 90 Then Exit Function
        result = result * 26 + code - 64
    Next i
    If result <= columnCount Then TaskFoldColumnNumber = result
End Function

Public Function GetTaskDetailColumnBounds(ByVal ws As Worksheet, ByRef firstCol As Long, _
                                          ByRef lastCol As Long) As Boolean
    Dim configWs As Worksheet
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If configWs Is Nothing Then
        firstCol = ws.Range(COL_REQUEST & "1").Column + 1
        lastCol = ws.Range(COL_GANTT_START & "1").Column - 1
    Else
        firstCol = TaskFoldColumnNumber(configWs.Range(TASK_FOLD_START_VALUE_CELL).Value2, ws.Columns.Count)
        lastCol = TaskFoldColumnNumber(configWs.Range(TASK_FOLD_END_VALUE_CELL).Value2, ws.Columns.Count)
    End If
    GetTaskDetailColumnBounds = (firstCol > 0 And firstCol <= lastCol And lastCol < ws.Columns.Count - 1)
End Function

Private Function GetTaskViewStartColumn(ByVal ws As Worksheet, ByVal firstCol As Long, _
                                        ByVal lastCol As Long) As Long
    Dim configWs As Worksheet, viewCol As Long
    If firstCol < 1 Or lastCol < firstCol Then Exit Function
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If configWs Is Nothing Then
        viewCol = ws.Range(COL_REQUEST & "1").Column
    Else
        viewCol = TaskFoldColumnNumber(configWs.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2, ws.Columns.Count - 1)
    End If
    If viewCol < 1 Then Exit Function
    ' Visibility and grouping are independent: skip columns hidden by the fold.
    If viewCol >= firstCol And viewCol <= lastCol Then viewCol = lastCol + 1
    If viewCol < ws.Columns.Count Then GetTaskViewStartColumn = viewCol
End Function

Private Sub UpdateTaskFoldHint(ByVal configWs As Worksheet, ByVal valid As Boolean)
    With configWs.Range(TASK_FOLD_HINT_CELL)
        If valid Then
            .Value2 = "화면 시작 열과 접기 범위는 별도 설정" & vbLf & "숨겨진 열은 건너뜀 · 펼치면 A열"
            .Font.Color = RGB(89, 89, 89)
        Else
            .Value2 = "접기 시작 ≤ 끝 · 열 문자를 확인하세요." & vbLf & "화면 시작 열도 열 문자로 입력 (예: F)"
            .Font.Color = RGB(192, 0, 0)
        End If
    End With
End Sub

Public Sub EnsureTaskDetailFoldConfig(ByVal ws As Worksheet)
    Dim cell As Range, formula As String, firstCol As Long, lastCol As Long
    If ws.Range("I15").MergeArea.Address = "$I$15:$J$16" Then
        ws.Range("I15:J16").UnMerge
        ws.Range("I15:J16").Clear
    End If
    ws.Range(TASK_FOLD_TITLE_CELL).Value2 = "간트 - 열 접기 설정"
    ws.Range("I13").Value2 = "접기 시작 열"
    ws.Range("I14").Value2 = "접기 끝 열"
    ws.Range("I15").Value2 = "접힌 화면 시작 열"
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_START_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    End If
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_END_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    End If
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = "H"
    End If
    With ws.Range("I12:J12")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("I13:J15").Borders.LineStyle = xlContinuous
    ws.Range("I13:J15").Font.Size = 11
    ws.Range("I13:J15").WrapText = False
    With ws.Range(TASK_FOLD_SETTING_VALUE_RANGE)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 249, 219)
        .Font.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("I17:J18")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Size = 9
    End With
    For Each cell In ws.Range(TASK_FOLD_SETTING_VALUE_RANGE).Cells
        formula = "=IFERROR(SUBSTITUTE(ADDRESS(1,COLUMN(INDIRECT(UPPER(TRIM(" & _
                  cell.Address(False, False) & "))&""1"")),4),""1"","""")=UPPER(TRIM(" & _
                  cell.Address(False, False) & ")),FALSE)"
        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formula
        cell.Validation.IgnoreBlank = False
        cell.Validation.InputTitle = "접을 열 범위"
        cell.Validation.InputMessage = "열 문자만 입력하세요. 예: I, W, AA"
        cell.Validation.ErrorTitle = "열 문자 입력 오류"
        cell.Validation.ErrorMessage = "숫자나 셀 주소 대신 열 문자를 입력하세요. 예: I, W"
        cell.Validation.ShowError = True
    Next cell
    If GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then
        UpdateTaskFoldHint ws, (GetTaskViewStartColumn(ws, firstCol, lastCol) > 0)
    Else
        UpdateTaskFoldHint ws, False
    End If
End Sub

Public Sub ApplyTaskDetailColumnSettings()
    Dim configWs As Worksheet, ws As Worksheet, cell As Range
    Dim firstCol As Long, lastCol As Long, number As Long
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean
    Dim settingError As Long, settingErrorText As String
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For Each cell In configWs.Range(TASK_FOLD_SETTING_VALUE_RANGE).Cells
        number = TaskFoldColumnNumber(cell.Value2, configWs.Columns.Count)
        If number > 0 Then cell.Value2 = UCase$(Trim$(CStr(cell.Value2)))
    Next cell
    If Not GetTaskDetailColumnBounds(configWs, firstCol, lastCol) Then
        UpdateTaskFoldHint configWs, False
        GoTo Finished
    End If
    UpdateTaskFoldHint configWs, (GetTaskViewStartColumn(configWs, firstCol, lastCol) > 0)
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." And _
           CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Level" Then
            UnprotectTaskSheet ws
            EnsureTaskDetailColumnOutline ws
            EnsureTaskRequestFreeze ws
            ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
        End If
    Next ws
Finished:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
Failed:
    settingError = Err.Number
    settingErrorText = Err.Description
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Err.Raise settingError, "ApplyTaskDetailColumnSettings", settingErrorText
End Sub

Public Sub InitializeTaskHierarchyOutlines()
    Dim ws As Worksheet
    Dim configWs As Worksheet
    Dim wasSaved As Boolean
    wasSaved = ThisWorkbook.Saved
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If Not configWs Is Nothing Then EnsureTaskDetailFoldConfig configWs
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
    Dim firstCol As Long, lastCol As Long, freezeColumns As Long, nextVisibleCol As Long
    Dim wasSaved As Boolean, collapsed As Boolean, resetView As Boolean
    Dim detailColumns As Range, wasScreenUpdating As Boolean
    Dim viewError As Long, viewErrorText As String
    If ActiveWindow Is Nothing Then Exit Sub
    If Not (ActiveWorkbook Is ws.Parent) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) <> "No." Or _
       CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) <> "Level" Then Exit Sub

    Set taskWindow = ActiveWindow
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Sub
    requestCol = GetTaskViewStartColumn(ws, firstCol, lastCol)
    If requestCol = 0 Then Exit Sub
    Do While requestCol < ws.Columns.Count
        If Not ws.Columns(requestCol).Hidden Then Exit Do
        requestCol = requestCol + 1
    Loop
    If requestCol >= ws.Columns.Count Then Exit Sub
    If Not mTaskViewResetSheets Is Nothing Then resetView = mTaskViewResetSheets.Exists(ws.CodeName)
    ' Start the collapsed viewport here; the fold range is configured separately.
    freezeColumns = 1
    collapsed = True
    For c = firstCol To lastCol
        If Not ws.Columns(c).Hidden Then
            collapsed = False
            Exit For
        End If
    Next c
    With taskWindow
        If collapsed Then
            If Not resetView Then
                If .FreezePanes And .SplitColumn = freezeColumns And .SplitRow = HEADER_ROW Then
                    If .Panes(1).ScrollColumn = requestCol Then Exit Sub
                End If
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
            ' Build the freeze with details visible, then restore the fold.
            Set detailColumns = ws.Range(ws.Cells(1, firstCol), ws.Cells(1, lastCol)).EntireColumn
            For c = firstCol To lastCol
                If ws.Columns(c).Hidden Then ws.Columns(c).Hidden = False
            Next c
            .ScrollColumn = requestCol
            .ScrollRow = 1
            .SplitColumn = freezeColumns
            .SplitRow = HEADER_ROW
            .FreezePanes = True
            For c = firstCol To lastCol
                If Not ws.Columns(c).Hidden Then ws.Columns(c).Hidden = True
            Next c
            nextVisibleCol = requestCol + 1
            Do While nextVisibleCol < ws.Columns.Count
                If Not ws.Columns(nextVisibleCol).Hidden Then Exit Do
                nextVisibleCol = nextVisibleCol + 1
            Loop
            .Panes(.Panes.Count).ScrollColumn = nextVisibleCol
            .Panes(.Panes.Count).ScrollRow = Application.Max(DATA_START_ROW, previousRow)
        Else
            .ScrollColumn = 1
            .ScrollRow = Application.Max(1, previousRow)
        End If
    End With
    If resetView Then mTaskViewResetSheets.Remove ws.CodeName
    If wasSaved Then ws.Parent.Saved = True
    Application.ScreenUpdating = wasScreenUpdating
    Exit Sub
Failed:
    viewError = Err.Number
    viewErrorText = Err.Description
    On Error Resume Next
    If Not detailColumns Is Nothing Then
        For c = firstCol To lastCol
            ws.Columns(c).Hidden = True
        Next c
    End If
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
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Function
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
    Dim needsGroup As Boolean, hidden As Variant, detailColumns As Range, oldColumns As Range
    Dim boundsChanged As Boolean, wasCollapsed As Boolean, rangeName As Name
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Sub
    Set detailColumns = ws.Range(ws.Cells(1, firstCol), ws.Cells(1, lastCol)).EntireColumn
    On Error Resume Next
    Set rangeName = ws.Names(TASK_FOLD_RANGE_NAME)
    If Not rangeName Is Nothing Then Set oldColumns = rangeName.RefersToRange
    On Error GoTo 0
    If oldColumns Is Nothing Then Set oldColumns = ws.Range("I:W")
    boundsChanged = (oldColumns.Column <> firstCol Or oldColumns.Columns.Count <> lastCol - firstCol + 1)
    needsGroup = boundsChanged
    For c = firstCol To lastCol
        If ws.Columns(c).OutlineLevel <> 2 Then needsGroup = True
    Next c
    If needsGroup Then
        If boundsChanged Then
            wasCollapsed = True
            For c = oldColumns.Column To oldColumns.Column + oldColumns.Columns.Count - 1
                If Not ws.Columns(c).Hidden Then wasCollapsed = False
            Next c
            For c = oldColumns.Column To oldColumns.Column + oldColumns.Columns.Count - 1
                If ws.Columns(c).Hidden Then ws.Columns(c).Hidden = False
            Next c
        Else
            hidden = CaptureTaskDetailColumnHidden(ws)
        End If
        UngroupTaskOutline oldColumns
        If boundsChanged Then UngroupTaskOutline detailColumns
        detailColumns.Group
        If boundsChanged Then
            For c = firstCol To lastCol
                If ws.Columns(c).Hidden <> wasCollapsed Then ws.Columns(c).Hidden = wasCollapsed
            Next c
        Else
            RestoreTaskDetailColumnHidden ws, hidden
        End If
    End If
    If rangeName Is Nothing Or boundsChanged Then
        If Not rangeName Is Nothing Then rangeName.Delete
        ws.Names.Add Name:="'" & Replace(ws.Name, "'", "''") & "'!" & TASK_FOLD_RANGE_NAME, _
                     RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & detailColumns.Address, Visible:=False
    End If
    If boundsChanged Then
        If mTaskViewResetSheets Is Nothing Then Set mTaskViewResetSheets = CreateObject("Scripting.Dictionary")
        mTaskViewResetSheets(ws.CodeName) = True
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
                     statusText = STATUS_TODAY Or _
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
