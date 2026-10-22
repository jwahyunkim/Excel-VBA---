Attribute VB_Name = "modGanttHierarchy"
Option Explicit
' 2026-10-22 김좌현 수정 : 행별 들여쓰기·굵기·배경·최소 높이를 모아 적용하고 빈 행 및 루트 경계 유지
' 2026-10-22 김좌현 수정 : 구조가 같은 행 그룹 재생성을 생략하고 초기화된 테두리를 렌더러에 전달
' 2026-10-22 김좌현 수정 : 3단계 웹모드 동기화와 중복 날짜 배경/계층 재생성 제거
' 2026-10-22 김좌현 수정 : 웹 설정/작업시간 3단 열 접기와 기존 접기 설정 이전 및 고정 창 보존
' 2026-10-20 김좌현 수정 : 숨긴 행의 틀 고정 좌표 보정과 화면 변경 중 재진입 방지
' 2026-10-20 김좌현 수정 : 계층 펼치기 후에도 완료 숨김 상태 유지
' 2026-10-20 김좌현 수정 : config 열 접기 영역의 제목/입력 서식 통일

' 2026-10-19 김좌현 수정 : 쉼표로 선택한 여러 열을 접힌 화면에 고정하고 접기 범위 안의 선택 열도 표시

Private mTaskOutlineViewEvents As CGanttOutlineView
Private mTaskOutlineViewStarting As Boolean
Private mTaskViewResetSheets As Object
Private Const TASK_FOLD_RANGE_NAME As String = "_GanttDetailFoldRange"
Private Const TASK_FOLD_SELECTION_NAME As String = "_GanttFoldSelection"

' A single letter retains the existing viewport setting. Lists select exact columns.
Private Function TaskViewColumnSelection(ByVal ws As Worksheet, ByVal lastCol As Long, _
                                         ByRef normalized As String) As Object
    Dim configWs As Worksheet, selected As Object, parts As Variant, part As Variant
    Dim value As String, number As Long, c As Long, hasVisibleSelection As Boolean
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If IsError(configWs.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2) Then Exit Function
    value = CStr(configWs.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2)
    If InStr(value, ",") = 0 Then Exit Function
    Set selected = CreateObject("Scripting.Dictionary")
    parts = Split(value, ",")
    For Each part In parts
        number = TaskFoldColumnNumber(part, lastCol)
        If number = 0 Then Exit Function
        selected(CStr(number)) = True
    Next part
    If selected.Count < 2 Or selected.Count >= lastCol Then Exit Function
    For c = 1 To lastCol
        If selected.Exists(CStr(c)) And Not IsTaskWebDetailColumn(ws, c) And _
           Not IsTaskPrivateColumn(ws, c) Then hasVisibleSelection = True
    Next c
    If Not hasVisibleSelection Then
        number = ws.Range(COL_REQUEST & "1").Column
        If number > lastCol Then number = Application.Min(lastCol, ws.Range(COL_MINOR_CATEGORY & "1").Column)
        selected(CStr(number)) = True
    End If
    normalized = ""
    For c = 1 To lastCol
        If selected.Exists(CStr(c)) Then
            If Len(normalized) > 0 Then normalized = normalized & ","
            normalized = normalized & Replace(ws.Cells(1, c).Address(False, False), "1", "")
        End If
    Next c
    Set TaskViewColumnSelection = selected
End Function

Private Function TaskViewUsesSelection() As Boolean
    Dim value As Variant
    value = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2
    If Not IsError(value) Then TaskViewUsesSelection = (InStr(CStr(value), ",") > 0)
End Function

Private Function TaskSelectionSignature(ByVal ws As Worksheet) As String
    On Error Resume Next
    TaskSelectionSignature = ws.Names(TASK_FOLD_SELECTION_NAME).RefersTo
    On Error GoTo 0
End Function

Private Function TaskSelectionCollapsed(ByVal ws As Worksheet, ByVal selected As Object, _
                                         ByVal lastCol As Long) As Boolean
    Dim c As Long, hasFoldColumn As Boolean
    TaskSelectionCollapsed = True
    For c = 1 To lastCol
        If Not selected.Exists(CStr(c)) And Not IsTaskWebDetailColumn(ws, c) And _
           Not IsTaskPrivateColumn(ws, c) Then
            hasFoldColumn = True
            If Not ws.Columns(c).Hidden Then TaskSelectionCollapsed = False: Exit Function
        End If
    Next c
    If Not hasFoldColumn Then TaskSelectionCollapsed = False
End Function

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
    Dim configWs As Worksheet, viewCol As Long, selected As Object, normalized As String, c As Long
    If firstCol < 1 Or lastCol < firstCol Then Exit Function
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If configWs Is Nothing Then
        viewCol = ws.Range(COL_REQUEST & "1").Column
    Else
        If TaskViewUsesSelection() Then
            Set selected = TaskViewColumnSelection(ws, lastCol, normalized)
            If Not selected Is Nothing Then
                For c = 1 To lastCol
                    If selected.Exists(CStr(c)) And Not IsTaskWebDetailColumn(ws, c) And _
                        Not IsTaskPrivateColumn(ws, c) Then GetTaskViewStartColumn = c: Exit For
                Next c
            End If
            Exit Function
        End If
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
            .Value2 = "고정 열: H 또는 A,B,G (접기 끝 열 이내)" & vbLf & "펼치면 원래 화면으로 복원"
            .Font.Color = RGB(95, 105, 117)
        Else
            .Value2 = "접기 시작 ≤ 끝 · 열 문자를 확인하세요." & vbLf & "고정 열: F 또는 A,B,G (접기 끝 열 이내)"
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
    ws.Range(TASK_FOLD_TITLE_CELL).Value2 = "열 접기 (1 간단 / 2 전체 / 3 웹설정)"
    ws.Range("I13").Value2 = "접기 시작 열"
    ws.Range("I14").Value2 = "접기 끝 열"
    ws.Range("I15").Value2 = "접힌 화면 고정 열"
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_START_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_START_VALUE_CELL).Value2 = COL_MODIFICATION
    End If
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_END_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_END_VALUE_CELL).Value2 = COL_STATUS
    End If
    If Len(Trim$(CStr(ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2))) = 0 Then
        ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = COL_REQUEST
    End If
    With ws.Range("I12:J12")
        .Font.Bold = True
        .Interior.Color = RGB(37, 55, 75)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone
    End With
    ws.Range("I13:J15").Borders.LineStyle = xlContinuous
    ws.Range("I13:J15").Borders.Color = RGB(216, 222, 229)
    ws.Range("I13:J15").Borders.Weight = xlHairline
    ws.Range("I13:J15").Font.Size = 10
    ws.Range("I13:J15").WrapText = False
    With ws.Range(TASK_FOLD_SETTING_VALUE_RANGE)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 249, 228)
        .Font.Color = RGB(36, 72, 111)
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
        If cell.Address = ws.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Address Then
            ' Lists are validated by ApplyTaskDetailColumnSettings before any outline changes.
            cell.Validation.Add Type:=xlValidateInputOnly
            cell.Validation.InputTitle = "접힌 화면 고정 열"
            cell.Validation.InputMessage = "한 열: H / 여러 열: A,B,G,K (접기 끝 열 이내, 시트 순서)"
            cell.Validation.ShowInput = True
            GoTo NextSettingCell
        End If
        cell.Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=formula
        cell.Validation.IgnoreBlank = False
        cell.Validation.InputTitle = "접을 열 범위"
        cell.Validation.InputMessage = "열 문자만 입력하세요. 예: I, W, AA"
        cell.Validation.ErrorTitle = "열 문자 입력 오류"
        cell.Validation.ErrorMessage = "숫자나 셀 주소 대신 열 문자를 입력하세요. 예: I, W"
        cell.Validation.ShowError = True
NextSettingCell:
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
    Dim selected As Object, normalized As String
    Dim previousViewStarting As Boolean
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousViewStarting = mTaskOutlineViewStarting
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mTaskOutlineViewStarting = True
    For Each cell In configWs.Range(TASK_FOLD_SETTING_VALUE_RANGE).Cells
        number = TaskFoldColumnNumber(cell.Value2, configWs.Columns.Count)
        If number > 0 Then cell.Value2 = UCase$(Trim$(CStr(cell.Value2)))
    Next cell
    If Not GetTaskDetailColumnBounds(configWs, firstCol, lastCol) Then
        UpdateTaskFoldHint configWs, False
        GoTo Finished
    End If
    If GetTaskViewStartColumn(configWs, firstCol, lastCol) = 0 Then
        UpdateTaskFoldHint configWs, False
        GoTo Finished
    End If
    If TaskViewUsesSelection() Then
        Set selected = TaskViewColumnSelection(configWs, lastCol, normalized)
        configWs.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = normalized
    End If
    UpdateTaskFoldHint configWs, True
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
    mTaskOutlineViewStarting = previousViewStarting
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
Failed:
    settingError = Err.Number
    settingErrorText = Err.Description
    mTaskOutlineViewStarting = previousViewStarting
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

Private Function CaptureTaskFreezeHiddenRows(ByVal ws As Worksheet, ByVal lastRow As Long) As Range
    Dim r As Long, first As Long, isHidden As Boolean, part As Range, result As Range
    For r = 1 To lastRow + 1
        isHidden = False
        If r <= lastRow Then isHidden = ws.Rows(r).Hidden
        If isHidden Then
            If first = 0 Then first = r
        ElseIf first > 0 Then
            Set part = ws.Rows(CStr(first) & ":" & CStr(r - 1))
            If result Is Nothing Then
                Set result = part
            Else
                Set result = Application.Union(result, part)
            End If
            first = 0
        End If
    Next r
    Set CaptureTaskFreezeHiddenRows = result
End Function

Private Sub RestoreTaskFreezeHiddenRows(ByVal hiddenRows As Range)
    Dim area As Range
    If hiddenRows Is Nothing Then Exit Sub
    For Each area In hiddenRows.Areas
        area.EntireRow.Hidden = True
    Next area
End Sub

Public Sub EnsureTaskRequestFreeze(ByVal ws As Worksheet)
    Dim taskWindow As Window, requestCol As Long, previousRow As Long, c As Long
    Dim firstCol As Long, lastCol As Long, freezeColumns As Long, nextVisibleCol As Long
    Dim wasSaved As Boolean, collapsed As Boolean, resetView As Boolean, hasFoldColumn As Boolean
    Dim detailColumns As Range, wasScreenUpdating As Boolean
    Dim viewError As Long, viewErrorText As String
    Dim selected As Object, normalized As String, selectedColumn As Variant, lastSelected As Long
    Dim freezeRows As Range, hiddenRows As Range, lastTaskRow As Long
    If ActiveWindow Is Nothing Then Exit Sub
    If Not (ActiveWorkbook Is ws.Parent) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) <> "No." Or _
       CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) <> "Level" Then Exit Sub

    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
    Set taskWindow = ActiveWindow
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Sub
    If TaskViewUsesSelection() Then
        Set selected = TaskViewColumnSelection(ws, lastCol, normalized)
        If selected Is Nothing Then Exit Sub
    End If
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
    If selected Is Nothing Then
        collapsed = True
        For c = firstCol To lastCol
            If Not IsTaskWebDetailColumn(ws, c) And Not IsTaskPrivateColumn(ws, c) Then hasFoldColumn = True
            If Not IsTaskWebDetailColumn(ws, c) And Not IsTaskPrivateColumn(ws, c) And _
               Not ws.Columns(c).Hidden Then
                collapsed = False
                Exit For
            End If
        Next c
        If Not hasFoldColumn Then collapsed = False
    Else
        For Each selectedColumn In selected.Keys
            If Not IsTaskWebDetailColumn(ws, CLng(selectedColumn)) And _
               Not IsTaskPrivateColumn(ws, CLng(selectedColumn)) Then _
                lastSelected = Application.Max(lastSelected, CLng(selectedColumn))
        Next selectedColumn
        freezeColumns = lastSelected - requestCol + 1
        collapsed = TaskSelectionCollapsed(ws, selected, lastCol)
    End If
    With taskWindow
        If collapsed Then
            If Not resetView Then
                If .FreezePanes And .SplitRow = HEADER_ROW Then
                    If .Panes(1).ScrollColumn = requestCol Then
                        If selected Is Nothing Then
                            If .SplitColumn = freezeColumns Then Exit Sub
                        ElseIf .SplitColumn > 0 Then
                            Exit Sub
                        End If
                    End If
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
        If collapsed Then
            lastTaskRow = Application.Max(HEADER_ROW, GetLastDataRow(ws))
            Set freezeRows = ws.Rows("1:" & CStr(lastTaskRow))
            Set hiddenRows = CaptureTaskFreezeHiddenRows(ws, lastTaskRow)
            ' Hidden rows also change Excel's split coordinates; retain their exact visibility.
            If Not hiddenRows Is Nothing Then freezeRows.Hidden = False
        End If
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0
        If collapsed Then
            ' Excel counts hidden columns differently when creating a split.
            ' Build the freeze with details visible, then restore the fold.
            If selected Is Nothing Then
                Set detailColumns = ws.Range(ws.Cells(1, firstCol), ws.Cells(1, lastCol)).EntireColumn
            Else
                Set detailColumns = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).EntireColumn
            End If
            If selected Is Nothing Then
                For c = detailColumns.Column To lastCol
                    If ws.Columns(c).Hidden Then ws.Columns(c).Hidden = False
                Next c
            End If
            .ScrollColumn = requestCol
            .ScrollRow = 1
            .SplitColumn = freezeColumns
            .SplitRow = HEADER_ROW
            .FreezePanes = True
            RestoreTaskFreezeHiddenRows hiddenRows
            For c = detailColumns.Column To lastCol
                If selected Is Nothing Then
                    ws.Columns(c).Hidden = True
                Else
                    ws.Columns(c).Hidden = (Not selected.Exists(CStr(c)) Or _
                        IsTaskWebDetailColumn(ws, c) Or IsTaskPrivateColumn(ws, c))
                End If
            Next c
            If selected Is Nothing Then
                nextVisibleCol = requestCol + 1
            Else
                nextVisibleCol = lastCol + 1
            End If
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
    RestoreTaskFreezeHiddenRows hiddenRows
    If Not detailColumns Is Nothing Then
        For c = detailColumns.Column To lastCol
            If selected Is Nothing Then
                ws.Columns(c).Hidden = True
            Else
                ws.Columns(c).Hidden = (Not selected.Exists(CStr(c)) Or _
                        IsTaskWebDetailColumn(ws, c) Or IsTaskPrivateColumn(ws, c))
            End If
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
    If Not Application.EnableEvents Or Not Application.ScreenUpdating Then Exit Sub
    Set ws = ActiveTaskOutlineSheet()
    If Not ws Is Nothing Then
        EnsureTaskRequestFreeze ws
        MaintainCompletedTaskView ws
        RefreshWebModeView ws
    End If
End Sub

Public Function CaptureTaskDetailColumnHidden(ByVal ws As Worksheet) As Variant
    Dim firstCol As Long, lastCol As Long, c As Long
    Dim hidden() As Boolean
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Function
    ' 2026-10-19 김좌현 수정 : 너비 재설정으로 풀리는 선택 열 사이의 숨김 상태도 함께 보존
    If TaskViewUsesSelection() Then firstCol = 1
    firstCol = 1
    lastCol = Application.Max(lastCol, ws.Range(COL_STATUS & "1").Column)
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
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
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

' 2026-10-22 김좌현 수정 : 기존 상세 열은 2단계, 웹 설정과 작업시간은 3단계로 독립 표시
Public Function IsTaskWebDetailColumn(ByVal ws As Worksheet, ByVal columnNumber As Long) As Boolean
    IsTaskWebDetailColumn = _
        (columnNumber >= ws.Range(COL_WEB_ENABLED & "1").Column And _
         columnNumber <= ws.Range(COL_WEB_RESULT & "1").Column) Or _
        columnNumber = ws.Range(COL_WORK_MINUTES & "1").Column
End Function

Private Function IsTaskPrivateColumn(ByVal ws As Worksheet, ByVal columnNumber As Long) As Boolean
    IsTaskPrivateColumn = (columnNumber = ws.Range(COL_WEB_TASK_ID & "1").Column)
End Function

Public Sub EnsureTaskDetailColumnOutline(ByVal ws As Worksheet)
    Dim firstCol As Long, lastCol As Long, c As Long, limit As Long, pass As Long, runStart As Long
    Dim needsGroup As Boolean, wasCollapsed As Boolean, settingsChanged As Boolean
    Dim selected As Object, normalized As String, signature As String, previousSignature As String
    Dim rangeName As Name, oldColumns As Range, detailColumns As Range
    Dim expected() As Long, hidden() As Boolean, oldFirst As Long, oldLast As Long
    If IsWebUtilitySheet(ws) Then Exit Sub
    If Not GetTaskDetailColumnBounds(ws, firstCol, lastCol) Then Exit Sub
    If GetTaskViewStartColumn(ws, firstCol, lastCol) = 0 Then Exit Sub
    previousSignature = TaskSelectionSignature(ws)
    If TaskViewUsesSelection() Then
        Set selected = TaskViewColumnSelection(ws, lastCol, normalized)
        If selected Is Nothing Then Exit Sub
        signature = "=" & Chr$(34) & firstCol & "|" & lastCol & "|" & normalized & Chr$(34)
    End If
    On Error Resume Next
    Set rangeName = ws.Names(TASK_FOLD_RANGE_NAME)
    If Not rangeName Is Nothing Then Set oldColumns = rangeName.RefersToRange
    On Error GoTo 0
    If oldColumns Is Nothing Then Set oldColumns = ws.Range(COL_MODIFICATION & ":" & COL_STATUS)
    oldFirst = oldColumns.Column
    oldLast = oldFirst + oldColumns.Columns.Count - 1
    limit = Application.Max(lastCol, oldLast, ws.Range(COL_STATUS & "1").Column)
    settingsChanged = (oldFirst <> firstCol Or oldLast <> lastCol Or previousSignature <> signature)
    wasCollapsed = True
    If Len(previousSignature) > 0 Then
        wasCollapsed = TaskPreviousSelectionCollapsed(ws, previousSignature)
    Else
        For c = oldFirst To oldLast
            If Not IsTaskWebDetailColumn(ws, c) And Not IsTaskPrivateColumn(ws, c) Then
                If Not ws.Columns(c).Hidden Then wasCollapsed = False
            End If
        Next c
    End If
    ReDim expected(1 To limit)
    ReDim hidden(1 To limit)
    For c = 1 To limit
        expected(c) = 1
        If selected Is Nothing Then
            If c >= firstCol And c <= lastCol Then expected(c) = 2
        ElseIf c <= lastCol Then
            If Not selected.Exists(CStr(c)) Then expected(c) = 2
        End If
        If IsTaskWebDetailColumn(ws, c) Then expected(c) = 3
        If IsTaskPrivateColumn(ws, c) Then expected(c) = 1
        hidden(c) = ws.Columns(c).Hidden
        If ws.Columns(c).OutlineLevel <> expected(c) Then needsGroup = True
    Next c
    If needsGroup Or settingsChanged Then
        ' Full-column Ungroup preserves independent task-row groups.
        UngroupTaskOutline ws.Range(ws.Cells(1, 1), ws.Cells(1, limit)).EntireColumn
        ' Inner groups first. A second group also permits isolated level-3 fields.
        For pass = 3 To 2 Step -1
            runStart = 0
            For c = 1 To limit + 1
                If c <= limit Then
                    If expected(c) >= pass Then
                        If runStart = 0 Then runStart = c
                    ElseIf runStart > 0 Then
                        ws.Range(ws.Cells(1, runStart), ws.Cells(1, c - 1)).EntireColumn.Group
                        runStart = 0
                    End If
                ElseIf runStart > 0 Then
                    ws.Range(ws.Cells(1, runStart), ws.Cells(1, c - 1)).EntireColumn.Group
                End If
            Next c
        Next pass
        For c = 1 To limit
            If settingsChanged And Not IsTaskWebDetailColumn(ws, c) Then
                ws.Columns(c).Hidden = (expected(c) = 2 And wasCollapsed)
            Else
                ws.Columns(c).Hidden = hidden(c)
            End If
        Next c
    End If
    If rangeName Is Nothing Or settingsChanged Then
        If Not rangeName Is Nothing Then rangeName.Delete
        Set detailColumns = ws.Range(ws.Cells(1, firstCol), ws.Cells(1, lastCol)).EntireColumn
        ws.Names.Add Name:="'" & Replace(ws.Name, "'", "''") & "'!" & TASK_FOLD_RANGE_NAME, _
                     RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & detailColumns.Address, Visible:=False
    End If
    If previousSignature <> signature Then
        If Len(previousSignature) > 0 Then ws.Names(TASK_FOLD_SELECTION_NAME).Delete
        If Len(signature) > 0 Then _
            ws.Names.Add Name:="'" & Replace(ws.Name, "'", "''") & "'!" & TASK_FOLD_SELECTION_NAME, _
                         RefersTo:=signature, Visible:=False
    End If
    If settingsChanged Then
        If mTaskViewResetSheets Is Nothing Then Set mTaskViewResetSheets = CreateObject("Scripting.Dictionary")
        mTaskViewResetSheets(ws.CodeName) = True
    End If
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
    If ws.Outline.SummaryColumn <> xlRight Then ws.Outline.SummaryColumn = xlRight
    If Not ws.EnableOutlining Then ws.EnableOutlining = True
    If Not ActiveWindow Is Nothing Then
        If Not ActiveWindow.DisplayOutline Then ActiveWindow.DisplayOutline = True
    End If
End Sub

Private Function TaskPreviousSelectionCollapsed(ByVal ws As Worksheet, ByVal signature As String) As Boolean
    Dim parts As Variant, selected As Object, part As Variant, oldLast As Long
    parts = Split(Replace(Replace(signature, "=", ""), Chr$(34), ""), "|")
    If UBound(parts) <> 2 Then Exit Function
    oldLast = CLng(parts(1))
    Set selected = CreateObject("Scripting.Dictionary")
    For Each part In Split(parts(2), ",")
        selected(CStr(TaskFoldColumnNumber(part, oldLast))) = True
    Next part
    TaskPreviousSelectionCollapsed = TaskSelectionCollapsed(ws, selected, oldLast)
End Function

' 2026-10-22 김좌현 수정 : 기존 설정의 열 문자도 삽입 열 수만큼 한 번 이전하며 G 고정은 유지
Public Sub MigrateTaskWebFoldSettings(ByVal taskWs As Worksheet)
    Dim configWs As Worksheet, marker As Name, cell As Range, signature As String
    Dim parts As Variant, selectedColumns As String, column As Variant
    On Error Resume Next
    Set marker = ThisWorkbook.Names("_GanttWebFoldV6")
    Set configWs = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If marker Is Nothing Then
        If Not configWs Is Nothing Then
            For Each cell In configWs.Range(TASK_FOLD_SETTING_VALUE_RANGE).Cells
                If Not IsError(cell.Value2) Then cell.Value2 = MigrateTaskWebColumnLetters(CStr(cell.Value2), configWs)
            Next cell
        End If
        ThisWorkbook.Names.Add Name:="_GanttWebFoldV6", RefersTo:="=TRUE", Visible:=False
    End If
    signature = TaskSelectionSignature(taskWs)
    If Len(signature) = 0 Then Exit Sub
    parts = Split(Replace(Replace(signature, "=", ""), Chr$(34), ""), "|")
    If UBound(parts) <> 2 Then Exit Sub
    selectedColumns = MigrateTaskWebColumnLetters(CStr(parts(2)), taskWs)
    signature = "=" & Chr$(34) & MigratedTaskWebColumnNumber(CLng(parts(0))) & "|" & _
                MigratedTaskWebColumnNumber(CLng(parts(1))) & "|" & selectedColumns & Chr$(34)
    taskWs.Names(TASK_FOLD_SELECTION_NAME).RefersTo = signature
End Sub

Private Function MigratedTaskWebColumnNumber(ByVal number As Long) As Long
    If number >= 16 Then
        number = number + 11
    ElseIf number >= 8 Then
        number = number + 10
    End If
    MigratedTaskWebColumnNumber = number
End Function

Private Function MigrateTaskWebColumnLetters(ByVal value As String, ByVal ws As Worksheet) As String
    Dim parts As Variant, part As Variant, number As Long, token As String
    parts = Split(value, ",")
    For Each part In parts
        token = UCase$(Trim$(CStr(part)))
        number = TaskFoldColumnNumber(token, ws.Columns.Count)
        If number > 0 Then _
            token = Replace(ws.Cells(1, MigratedTaskWebColumnNumber(number)).Address(False, False), "1", "")
        If Len(MigrateTaskWebColumnLetters) > 0 Then _
            MigrateTaskWebColumnLetters = MigrateTaskWebColumnLetters & ","
        MigrateTaskWebColumnLetters = MigrateTaskWebColumnLetters & token
    Next part
End Function
' Visual hierarchy and native outlines share the same subtree boundaries as
' the task summary calculation. Build them before applying display filters.
Public Sub ApplyTaskHierarchyAppearance(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                        ByVal startCol As Long, ByVal endCol As Long, _
                                        Optional ByVal fitLayout As Boolean = True, Optional ByVal rebuildHierarchy As Boolean = True, _
                                        Optional ByVal bordersAlreadyReset As Boolean = False)
    Dim r As Long, i As Long, taskLevel As Long, fillIndex As Long, contentRunStart As Long
    Dim firstDataCol As Long, hasChildren As Boolean, inRootGroup As Boolean, isWarning As Boolean
    Dim statusText As String, rowRange As Range, taskRange As Range, rootRow As Variant
    Dim indentRanges(1 To 3) As Range, fillRanges(1 To 5) As Range
    Dim heightRanges(1 To 3) As Range, boldTasks As Range, darkRoots As Range
    Dim rowHeightGroup() As Long, fillColors As Variant, rootRows As Collection

    EnsureTaskDetailColumnOutline ws
    If lastRow < DATA_START_ROW Or endCol < startCol Then Exit Sub
    firstDataCol = ws.Range(COL_NO & "1").Column
    ReDim rowHeightGroup(DATA_START_ROW To lastRow)
    Set rootRows = New Collection
    fillColors = Array(RGB(255, 255, 255), RGB(243, 247, 251), RGB(250, 252, 254), _
                       RGB(226, 234, 242), RGB(222, 233, 243))
    With ws.Range(COL_NO & DATA_START_ROW & ":" & COL_STATUS & lastRow).Font
        .Bold = False
        .Color = RGB(40, 50, 62)
    End With

    ' 2026-10-22 김좌현 수정 : 기존 행 판정은 유지하고 동일 속성의 영역만 모아 일괄 서식 적용
    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then
            If contentRunStart > 0 Then
                ApplyHierarchyBottomRun ws, contentRunStart, r - 1, firstDataCol, endCol
                contentRunStart = 0
            End If
            GoTo NextTask
        End If
        If contentRunStart = 0 Then contentRunStart = r
        taskLevel = GetTaskLevel(ws, r)
        hasChildren = HasChildTask(ws, r, lastRow)
        If taskLevel = 1 Then inRootGroup = hasChildren
        statusText = CStr(ws.Cells(r, COL_STATUS).Value)
        isWarning = (statusText = STATUS_HOLD Or statusText = STATUS_DELAY Or _
                     statusText = STATUS_TODAY Or statusText = STATUS_ERROR)
        Set taskRange = ws.Range(COL_REQUEST & r & ":" & COL_MODIFICATION & r)
        AddHierarchyAppearanceRange indentRanges(taskLevel), taskRange
        If taskLevel < 3 And (inRootGroup Or hasChildren) Then _
            AddHierarchyAppearanceRange boldTasks, taskRange

        fillIndex = 1
        If inRootGroup Then
            fillIndex = 2
            If taskLevel = 3 Then fillIndex = 3
        End If
        If taskLevel = 1 And hasChildren Then
            fillIndex = 4
        ElseIf taskLevel = 2 And hasChildren Then
            fillIndex = 5
        End If
        ' Warning rows and status cells retain their existing status colors.
        If Not isWarning Then
            AddHierarchyAppearanceRange fillRanges(fillIndex), _
                ws.Range(COL_NO & r & ":" & COL_ACTUAL_DAYS & r)
            If taskLevel = 1 And hasChildren Then _
                AddHierarchyAppearanceRange darkRoots, ws.Range(COL_NO & r & ":" & COL_MODIFICATION & r)
        End If
        If taskLevel = 1 And hasChildren Then
            rootRows.Add r
            rowHeightGroup(r) = 1
        ElseIf taskLevel = 2 Then
            rowHeightGroup(r) = 2
        Else
            rowHeightGroup(r) = 3
        End If
NextTask:
    Next r
    If contentRunStart > 0 Then _
        ApplyHierarchyBottomRun ws, contentRunStart, lastRow, firstDataCol, endCol

    For i = 1 To 3
        If Not indentRanges(i) Is Nothing Then indentRanges(i).IndentLevel = i - 1
    Next i
    If Not boldTasks Is Nothing Then boldTasks.Font.Bold = True
    For i = 1 To 5
        If Not fillRanges(i) Is Nothing Then fillRanges(i).Interior.Color = CLng(fillColors(i - 1))
    Next i
    If Not darkRoots Is Nothing Then
        With darkRoots
            .Interior.Color = RGB(44, 72, 98)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
    End If
    ' Root edges remain ordered per row so shared boundary precedence is unchanged.
    For Each rootRow In rootRows
        Set rowRange = ws.Range(ws.Cells(CLng(rootRow), firstDataCol), ws.Cells(CLng(rootRow), endCol))
        SetHierarchyRule rowRange, xlEdgeTop, RGB(68, 96, 122), xlMedium
        SetHierarchyRule rowRange, xlEdgeBottom, RGB(126, 150, 173), xlThin
    Next rootRow
    ' Read final row heights after font changes; preserve every height already above its minimum.
    For r = DATA_START_ROW To lastRow
        i = rowHeightGroup(r)
        If i > 0 Then
            Select Case i
                Case 1
                    If ws.Rows(r).RowHeight < 30 Then AddHierarchyAppearanceRange heightRanges(i), ws.Rows(r)
                Case 2
                    If ws.Rows(r).RowHeight < 26 Then AddHierarchyAppearanceRange heightRanges(i), ws.Rows(r)
                Case Else
                    If ws.Rows(r).RowHeight < 23 Then AddHierarchyAppearanceRange heightRanges(i), ws.Rows(r)
            End Select
        End If
    Next r
    If Not heightRanges(1) Is Nothing Then heightRanges(1).RowHeight = 30
    If Not heightRanges(2) Is Nothing Then heightRanges(2).RowHeight = 26
    If Not heightRanges(3) Is Nothing Then heightRanges(3).RowHeight = 23

    ' Fit the final bold text and native indentation, not the old text prefixes.
    If fitLayout Then ws.Range(COL_REQUEST & HEADER_ROW & ":" & COL_MODIFICATION & lastRow).Columns.AutoFit
    If rebuildHierarchy Then RebuildTaskHierarchyOutline ws, lastRow
    RenderHierarchyTaskBars ws, lastRow, startCol, endCol, Nothing, Nothing, bordersAlreadyReset
End Sub

Private Sub AddHierarchyAppearanceRange(ByRef accumulated As Range, ByVal target As Range)
    If accumulated Is Nothing Then
        Set accumulated = target
    Else
        Set accumulated = Application.Union(accumulated, target)
    End If
End Sub

Private Sub ApplyHierarchyBottomRun(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                                    ByVal firstCol As Long, ByVal lastCol As Long)
    Dim target As Range
    Set target = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
    ' Adjacent content rows share internal borders; blank-row gaps are never included.
    If lastRow > firstRow Then SetHierarchyRule target, xlInsideHorizontal, RGB(221, 229, 237), xlThin
    SetHierarchyRule target, xlEdgeBottom, RGB(221, 229, 237), xlThin
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
    If Not ws.EnableOutlining Then ws.EnableOutlining = True
    If Not ActiveWindow Is Nothing Then
        If Not ActiveWindow.DisplayOutline Then ActiveWindow.DisplayOutline = True
    End If
End Sub
