Attribute VB_Name = "modCompletedTaskView"
Option Explicit
' 2026-10-20 김좌현 수정 : 시트별 완료 숨김 토글과 기존 행/계층 보기 상태 보존
' 2026-10-20 김좌현 수정 : 완료 숨김 전 행 상태로 수동 숨김/새 접기를 구분하고 실패 시 보기 메타데이터까지 복원

Public Const COMPLETED_TASK_BUTTON_NAME As String = "btnGanttCompletedToggle"
Private Const FILTER_NAME As String = "_GanttHideCompleted"
Private Const OWNED_PREFIX As String = "_GanttDoneOwned_"
Private Const COLLAPSED_PREFIX As String = "_GanttDoneCollapsed_"
Private Const BASELINE_PREFIX As String = "_GanttDoneBaseline_"
Private mChangingView As Boolean

Public Function IsCompletedTaskSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsCompletedTaskSheet = (CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." And _
                            CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Level" And _
                            CStr(ws.Cells(HEADER_ROW, COL_STATUS).Value2) = "상태")
End Function

Private Function LocalNameKey(ByVal item As Name) As String
    LocalNameKey = Mid$(item.Name, InStrRev(item.Name, "!") + 1)
End Function

Public Function AreCompletedTasksHidden(ByVal ws As Worksheet) As Boolean
    Dim item As Name
    For Each item In ws.Names
        If LocalNameKey(item) = FILTER_NAME Then
            AreCompletedTasksHidden = (item.RefersTo = "=1")
            Exit Function
        End If
    Next item
End Function

Public Function CompletedTaskButtonCaption(ByVal ws As Worksheet) As String
    If AreCompletedTasksHidden(ws) Then
        CompletedTaskButtonCaption = "완료 표시"
    Else
        CompletedTaskButtonCaption = "완료 숨김"
    End If
End Function

Public Sub UpdateCompletedTaskButton(ByVal ws As Worksheet)
    Dim button As Shape, caption As String
    If Not IsCompletedTaskSheet(ws) Then Exit Sub
    On Error Resume Next
    Set button = ws.Shapes(COMPLETED_TASK_BUTTON_NAME)
    On Error GoTo 0
    If button Is Nothing Then Exit Sub
    caption = CompletedTaskButtonCaption(ws)
    If button.TextFrame2.TextRange.Text <> caption Then button.TextFrame2.TextRange.Text = caption
End Sub

Private Sub DeleteStoredNames(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long, key As String
    For i = ws.Names.Count To 1 Step -1
        key = LocalNameKey(ws.Names(i))
        If Left$(key, Len(prefix)) = prefix Then ws.Names(i).Delete
    Next i
End Sub

Private Sub AddLocalName(ByVal ws As Worksheet, ByVal key As String, ByVal formula As String)
    ws.Names.Add Name:="'" & Replace(ws.Name, "'", "''") & "'!" & key, _
                 RefersTo:=formula, Visible:=False
End Sub

Private Function StoredRows(ByVal ws As Worksheet, ByVal prefix As String) As Range
    Dim item As Name, part As Range, result As Range
    For Each item In ws.Names
        If Left$(LocalNameKey(item), Len(prefix)) = prefix Then
            Set part = Nothing
            ' Deleting task rows can remove a stored range completely.
            On Error Resume Next
            Set part = item.RefersToRange
            On Error GoTo 0
            If Not part Is Nothing Then
                If result Is Nothing Then
                    Set result = part
                Else
                    Set result = Application.Union(result, part)
                End If
            End If
        End If
    Next item
    Set StoredRows = result
End Function

Private Function TaskViewLastRow(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim owned As Range, area As Range
    TaskViewLastRow = lastRow
    Set owned = StoredRows(ws, OWNED_PREFIX)
    If owned Is Nothing Then Exit Function
    ' Cleared trailing tasks remain owned until their filter hiding has been removed.
    For Each area In owned.Areas
        TaskViewLastRow = Application.Max(TaskViewLastRow, area.Row + area.Rows.Count - 1)
    Next area
End Function

Private Function TaskViewSubtreeEnd(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal lastRow As Long) As Long
    Dim groupEnd As Long
    groupEnd = GetTaskSubtreeEndRow(ws, rowNumber, lastRow)
    ' Match the outline builder: trailing empty rows are outside the physical group.
    Do While groupEnd > rowNumber
        If HasTaskContent(ws, groupEnd) Then Exit Do
        groupEnd = groupEnd - 1
    Loop
    TaskViewSubtreeEnd = groupEnd
End Function

Private Sub StoreRows(ByVal ws As Worksheet, ByVal prefix As String, ByVal rows As Range)
    Dim area As Range, formula As String, part As String, index As Long
    DeleteStoredNames ws, prefix
    If rows Is Nothing Then Exit Sub
    index = 1
    For Each area In rows.Areas
        part = "'" & Replace(ws.Name, "'", "''") & "'!" & area.Address
        ' Bound each name formula; alternating completed rows can create many areas.
        If Len(formula) + Len(part) > 3000 Then
            AddLocalName ws, prefix & CStr(index), "=" & formula
            index = index + 1
            formula = vbNullString
        End If
        If Len(formula) > 0 Then formula = formula & ","
        formula = formula & part
    Next area
    If Len(formula) > 0 Then AddLocalName ws, prefix & CStr(index), "=" & formula
End Sub

Private Function HasStoredNames(ByVal ws As Worksheet, ByVal prefix As String) As Boolean
    Dim item As Name
    For Each item In ws.Names
        If Left$(LocalNameKey(item), Len(prefix)) = prefix Then
            HasStoredNames = True
            Exit Function
        End If
    Next item
End Function

Private Sub StoreBaselineRows(ByVal ws As Worksheet, ByVal rows As Range)
    StoreRows ws, BASELINE_PREFIX, rows
    ' An empty baseline is different from a workbook saved before baseline tracking existed.
    If rows Is Nothing Then AddLocalName ws, BASELINE_PREFIX & "0", "=0"
End Sub

Private Function ContainsRow(ByVal rows As Range, ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    If rows Is Nothing Then Exit Function
    ContainsRow = Not Application.Intersect(rows, ws.Rows(rowNumber)) Is Nothing
End Function

Private Function RowsFromMask(ByVal ws As Worksheet, ByRef mask() As Boolean, ByVal lastRow As Long) As Range
    Dim r As Long, first As Long, part As Range, result As Range
    For r = DATA_START_ROW To lastRow + 1
        If r <= lastRow Then
            If mask(r) And first = 0 Then first = r
        End If
        If first > 0 Then
            If r > lastRow Then
                Set part = ws.Rows(CStr(first) & ":" & CStr(r - 1))
                first = 0
            ElseIf Not mask(r) Then
                Set part = ws.Rows(CStr(first) & ":" & CStr(r - 1))
                first = 0
            Else
                Set part = Nothing
            End If
            If Not part Is Nothing Then
                If result Is Nothing Then
                    Set result = part
                Else
                    Set result = Application.Union(result, part)
                End If
            End If
        End If
    Next r
    Set RowsFromMask = result
End Function

Public Function CaptureTaskRowView(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim view As Object, owned As Range, savedCollapsed As Range, baseline As Range
    Dim hidden() As Boolean, collapsed() As Boolean, allHidden() As Boolean
    Dim filterOwned() As Boolean, baselineHidden() As Boolean
    Dim r As Long, child As Long, groupEnd As Long, isExpanded As Boolean
    Dim anyOwned As Boolean, newlyHidden As Boolean, allChildrenHidden As Boolean
    Dim filterHidden As Boolean, baselineKnown As Boolean
    lastRow = TaskViewLastRow(ws, lastRow)
    Set view = CreateObject("Scripting.Dictionary")
    view.Add "LastRow", lastRow
    view.Add "Hidden", Nothing
    view.Add "Collapsed", Nothing
    view.Add "AllHidden", Nothing
    Set owned = StoredRows(ws, OWNED_PREFIX)
    Set savedCollapsed = StoredRows(ws, COLLAPSED_PREFIX)
    Set baseline = StoredRows(ws, BASELINE_PREFIX)
    filterHidden = AreCompletedTasksHidden(ws)
    baselineKnown = HasStoredNames(ws, BASELINE_PREFIX)
    view.Add "Owned", owned
    view.Add "StoredCollapsed", savedCollapsed
    view.Add "Baseline", baseline
    view.Add "BaselineKnown", baselineKnown
    view.Add "FilterHidden", filterHidden
    If lastRow >= DATA_START_ROW Then
        ReDim hidden(DATA_START_ROW To lastRow)
        ReDim collapsed(DATA_START_ROW To lastRow)
        ReDim allHidden(DATA_START_ROW To lastRow)
        ReDim filterOwned(DATA_START_ROW To lastRow)
        ReDim baselineHidden(DATA_START_ROW To lastRow)
        For r = DATA_START_ROW To lastRow
            allHidden(r) = ws.Rows(r).Hidden
            filterOwned(r) = ContainsRow(owned, ws, r)
            hidden(r) = allHidden(r) And Not filterOwned(r)
            If baselineKnown Then
                baselineHidden(r) = ContainsRow(baseline, ws, r)
            Else
                ' Old workbooks have no baseline; do not treat pre-existing manual hiding as a new collapse.
                baselineHidden(r) = hidden(r)
            End If
        Next r
        For r = DATA_START_ROW To lastRow
            If HasTaskContent(ws, r) And HasChildTask(ws, r, lastRow) Then
                isExpanded = True
                On Error Resume Next
                isExpanded = ws.Rows(r).ShowDetail
                On Error GoTo 0
                If Not isExpanded Then
                    groupEnd = TaskViewSubtreeEnd(ws, r, lastRow)
                    allChildrenHidden = True
                    anyOwned = False
                    newlyHidden = False
                    For child = r + 1 To groupEnd
                        If Not allHidden(child) Then allChildrenHidden = False
                        If filterOwned(child) Then anyOwned = True
                        If hidden(child) And Not baselineHidden(child) Then newlyHidden = True
                    Next child
                    ' ShowDetail=False alone also describes manual/filter hiding. Require a preserved
                    ' collapse or a change in the hidden rows while the filter was active.
                    collapsed(r) = allChildrenHidden And (Not filterHidden Or Not anyOwned Or _
                        ContainsRow(savedCollapsed, ws, r) Or newlyHidden Or _
                        (filterOwned(r) And Not allHidden(r)))
                End If
            End If
        Next r
        Set view("Hidden") = RowsFromMask(ws, hidden, lastRow)
        Set view("Collapsed") = RowsFromMask(ws, collapsed, lastRow)
        Set view("AllHidden") = RowsFromMask(ws, allHidden, lastRow)
    End If
    Set CaptureTaskRowView = view
End Function

Private Sub SetRowsHidden(ByVal rows As Range, ByVal hidden As Boolean)
    Dim area As Range
    If rows Is Nothing Then Exit Sub
    For Each area In rows.Areas
        area.EntireRow.Hidden = hidden
    Next area
End Sub

Private Sub RestoreCollapsedRows(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal rows As Range)
    Dim area As Range, row As Range, groupEnd As Long
    If rows Is Nothing Then Exit Sub
    For Each area In rows.Areas
        For Each row In area.Rows
            groupEnd = TaskViewSubtreeEnd(ws, row.Row, lastRow)
            ' Repeating ShowDetail=False on an already-hidden inner group can collapse its parent.
            ' Restore the exact descendants instead; ShowDetail reflects their hidden state.
            If groupEnd > row.Row Then ws.Rows(CStr(row.Row + 1) & ":" & CStr(groupEnd)).Hidden = True
        Next row
    Next area
End Sub

Public Sub RestoreTaskRowView(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal view As Object)
    Dim rows As Range
    lastRow = Application.Max(lastRow, CLng(view("LastRow")))
    DeleteStoredNames ws, OWNED_PREFIX
    DeleteStoredNames ws, BASELINE_PREFIX
    If lastRow < DATA_START_ROW Then Exit Sub
    ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(lastRow)).Hidden = False
    Set rows = view("Hidden")
    SetRowsHidden rows, True
    Set rows = view("Collapsed")
    RestoreCollapsedRows ws, lastRow, rows
End Sub

Public Sub RestoreTaskRowViewSnapshot(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal view As Object)
    Dim rows As Range
    If view Is Nothing Then Exit Sub
    lastRow = Application.Max(lastRow, CLng(view("LastRow")))
    If lastRow >= DATA_START_ROW Then
        ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(lastRow)).Hidden = False
        Set rows = view("AllHidden")
        SetRowsHidden rows, True
    End If
    Set rows = view("Owned")
    StoreRows ws, OWNED_PREFIX, rows
    Set rows = view("StoredCollapsed")
    StoreRows ws, COLLAPSED_PREFIX, rows
    DeleteStoredNames ws, BASELINE_PREFIX
    If CBool(view("BaselineKnown")) Then
        Set rows = view("Baseline")
        StoreBaselineRows ws, rows
    End If
    DeleteStoredNames ws, FILTER_NAME
    If CBool(view("FilterHidden")) Then AddLocalName ws, FILTER_NAME, "=1"
    UpdateCompletedTaskButton ws
End Sub

Private Sub ApplyView(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal hideCompleted As Boolean)
    Dim view As Object, owned As Range, collapsed As Range, baseline As Range
    Dim mask() As Boolean, baselineMask() As Boolean, r As Long
    Set view = CaptureTaskRowView(ws, lastRow)
    lastRow = CLng(view("LastRow"))
    RestoreTaskRowView ws, lastRow, view
    DeleteStoredNames ws, FILTER_NAME
    If hideCompleted And lastRow >= DATA_START_ROW Then
        ReDim mask(DATA_START_ROW To lastRow)
        ReDim baselineMask(DATA_START_ROW To lastRow)
        For r = DATA_START_ROW To lastRow
            baselineMask(r) = ws.Rows(r).Hidden
            If HasTaskContent(ws, r) Then
                If CStr(ws.Cells(r, COL_STATUS).Value2) = STATUS_DONE And Not baselineMask(r) Then mask(r) = True
            End If
        Next r
        Set owned = RowsFromMask(ws, mask, lastRow)
        StoreRows ws, OWNED_PREFIX, owned
        Set collapsed = view("Collapsed")
        StoreRows ws, COLLAPSED_PREFIX, collapsed
        Set baseline = RowsFromMask(ws, baselineMask, lastRow)
        StoreBaselineRows ws, baseline
        AddLocalName ws, FILTER_NAME, "=1"
        SetRowsHidden owned, True
    Else
        DeleteStoredNames ws, COLLAPSED_PREFIX
    End If
    UpdateCompletedTaskButton ws
End Sub

Public Sub ApplyCompletedTaskVisibility(ByVal ws As Worksheet, ByVal lastRow As Long)
    If Not IsCompletedTaskSheet(ws) Then Exit Sub
    ApplyView ws, lastRow, AreCompletedTasksHidden(ws)
End Sub

Public Sub ClearCompletedTaskFilter(ByVal ws As Worksheet, ByVal lastRow As Long)
    ApplyView ws, lastRow, False
End Sub

Public Sub ToggleCompletedTaskFilter(ByVal ws As Worksheet)
    Dim lastRow As Long, previousEvents As Boolean, previousScreen As Boolean, wasProtected As Boolean
    Dim before As Object, originalFlag As Boolean
    Dim holidays As Object, workdays As Object, errNumber As Long, errText As String
    If Not IsCompletedTaskSheet(ws) Then Err.Raise vbObjectError + 7600, , "업무 시트에서 실행하세요."
    If mChangingView Then Exit Sub
    lastRow = TaskViewLastRow(ws, GetLastDataRow(ws))
    If lastRow < DATA_START_ROW And Not AreCompletedTasksHidden(ws) Then _
        Err.Raise vbObjectError + 7601, , "업무 데이터가 없습니다."
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
    Set before = CaptureTaskRowView(ws, lastRow)
    originalFlag = AreCompletedTasksHidden(ws)
    On Error GoTo Failed
    mChangingView = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectTaskSheet ws
    EnsureConfigSheet
    LoadHolidaySettings holidays, workdays
    ApplyHierarchySummaryValues ws, lastRow, holidays, workdays
    ApplyView ws, lastRow, Not originalFlag
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    mChangingView = False
    Exit Sub
Failed:
    errNumber = Err.Number
    errText = Err.Description
    On Error Resume Next
    RestoreTaskRowViewSnapshot ws, lastRow, before
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    mChangingView = False
    On Error GoTo 0
    Err.Raise errNumber, "ToggleCompletedTaskFilter", errText
End Sub

Public Sub MaintainCompletedTaskView(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long, owned As Range, needsUpdate As Boolean, wasProtected As Boolean
    Dim before As Object, previousEvents As Boolean, previousScreen As Boolean
    If mChangingView Or Not IsCompletedTaskSheet(ws) Then Exit Sub
    If Not AreCompletedTasksHidden(ws) Then Exit Sub
    lastRow = TaskViewLastRow(ws, GetLastDataRow(ws))
    Set owned = StoredRows(ws, OWNED_PREFIX)
    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then
            If ContainsRow(owned, ws, r) Then needsUpdate = True
        Else
            If CStr(ws.Cells(r, COL_STATUS).Value2) = STATUS_DONE Then
                If Not ws.Rows(r).Hidden Then needsUpdate = True
            ElseIf ContainsRow(owned, ws, r) Then
                needsUpdate = True
            End If
        End If
    Next r
    If Not needsUpdate Then Exit Sub
    Set before = CaptureTaskRowView(ws, lastRow)
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
    On Error GoTo Failed
    mChangingView = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectTaskSheet ws
    ApplyView ws, lastRow, True
Finished:
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    mChangingView = False
    Exit Sub
Failed:
    On Error Resume Next
    RestoreTaskRowViewSnapshot ws, lastRow, before
    GoTo Finished
End Sub
