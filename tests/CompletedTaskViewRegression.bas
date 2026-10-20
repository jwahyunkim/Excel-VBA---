Attribute VB_Name = "CompletedTaskViewRegression"
Option Explicit
Private mPhase As String
' 2026-10-20 김좌현 수정 : 완료 숨김 토글/계층/새로고침/상태 복구/저장 후 유지 검증
' 2026-10-20 김좌현 수정 : 열 접기 순서별 틀 고정 안정성, 숨김 원인 구분 및 새로고침 중간 실패 복구 검증
Private mFreezeResets As Long
Private mNativeUpdates As Long

' The runner adds these callbacks only to the temporary workbook sources.
Public Sub RecordFreezeReset()
    mFreezeResets = mFreezeResets + 1
End Sub

Public Sub RecordNativeUpdate()
    mNativeUpdates = mNativeUpdates + 1
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 7610, "CompletedTaskViewRegression", message
End Sub

Private Function NewTaskSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = sheetName
    SetupDataHeaders ws
    Set NewTaskSheet = ws
End Function

Private Sub CheckStableFold(ByVal ws As Worksheet, ByVal anchor As Long)
    Dim taskWindow As Window, bottomPane As Pane, pass As Long
    Dim scrollColumn As Long, scrollRow As Long, freezeResets As Long
    Dim beforeRows As String, beforeMetadata As String, header As Long
    Set taskWindow = ActiveWindow
    EnsureTaskRequestFreeze ws
    Check taskWindow.FreezePanes, "collapsed columns have no frozen panes"
    Check taskWindow.SplitRow = HEADER_ROW, "hidden data rows changed the frozen header boundary"
    Check taskWindow.Panes(1).ScrollColumn = anchor, "collapsed viewport starts at the wrong column"
    Set bottomPane = taskWindow.Panes(taskWindow.Panes.Count)
    bottomPane.ScrollColumn = 32
    bottomPane.ScrollRow = 12
    scrollColumn = bottomPane.ScrollColumn
    scrollRow = bottomPane.ScrollRow
    beforeRows = HiddenRowsSignature(ws, 20)
    beforeMetadata = TaskViewMetadataSignature(ws)
    ws.Parent.Saved = True
    freezeResets = mFreezeResets
    For pass = 1 To 5
        TaskOutlineViewTick
        DoEvents
        Check taskWindow.FreezePanes And taskWindow.SplitRow = HEADER_ROW, "repeated monitor changed the frozen header boundary"
        Check taskWindow.Panes(taskWindow.Panes.Count).ScrollColumn = scrollColumn, "monitor reset horizontal scrolling"
        Check taskWindow.Panes(taskWindow.Panes.Count).ScrollRow = scrollRow, "monitor reset vertical scrolling"
        Check HiddenRowsSignature(ws, 20) = beforeRows, "monitor changed raw row visibility"
        Check TaskViewMetadataSignature(ws) = beforeMetadata, "monitor changed row-view metadata"
        Check ws.Parent.Saved, "stable monitor dirtied a saved workbook"
    Next pass
    For header = 1 To HEADER_ROW
        Check Not ws.Rows(header).Hidden, "pane maintenance hid a header row"
    Next header
    Check mFreezeResets = freezeResets, "stable monitor repeatedly recreated frozen panes"
    Check ws.Rows(5).Hidden And ws.Rows(6).Hidden And ws.Rows(7).Hidden, "pane maintenance exposed completed rows"
End Sub

Public Function RunCompletedTaskViewOrderRegression() As String
    Dim cfg As Worksheet, ws As Worksheet, r As Long, selectionCase As Long, order As Long
    Dim originalStart As Variant, originalEnd As Variant, originalView As Variant
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean, configCaptured As Boolean
    Dim errorText As String, nativeUpdates As Long
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    StopTaskOutlineViewMonitor
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set cfg = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    originalStart = cfg.Range(TASK_FOLD_START_VALUE_CELL).Value2
    originalEnd = cfg.Range(TASK_FOLD_END_VALUE_CELL).Value2
    originalView = cfg.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2
    configCaptured = True
    cfg.Range(TASK_FOLD_START_VALUE_CELL).Value2 = "I"
    cfg.Range(TASK_FOLD_END_VALUE_CELL).Value2 = "W"
    mNativeUpdates = 0
    For selectionCase = 1 To 2
        cfg.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = IIf(selectionCase = 1, "G", "A,B,G")
        For order = 1 To 2
            mPhase = "fold order variant " & selectionCase & ", order " & order
            Set ws = NewTaskSheet("CompletedFold" & selectionCase & "Order" & order)
            For r = 5 To 20
                SeedTask ws, r, 1, (r <= 7)
            Next r
            ws.Activate
            RefreshGanttSheet False
            ws.Outline.ShowLevels ColumnLevels:=2
            EnsureTaskRequestFreeze ws
            If order = 1 Then
                ToggleCompletedTaskFilter ws
                ws.Outline.ShowLevels ColumnLevels:=1
            Else
                ws.Outline.ShowLevels ColumnLevels:=1
                EnsureTaskRequestFreeze ws
                ToggleCompletedTaskFilter ws
            End If
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            StartTaskOutlineViewMonitor
            CheckStableFold ws, IIf(selectionCase = 1, 7, 1)
            StopTaskOutlineViewMonitor
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            ws.Outline.ShowLevels ColumnLevels:=2
            EnsureTaskRequestFreeze ws
            Check Not ActiveWindow.FreezePanes And ActiveWindow.SplitRow = 0 And ActiveWindow.SplitColumn = 0, "expansion retained collapsed panes"
            Check ws.Rows(5).Hidden And ws.Rows(7).Hidden, "column expansion exposed completed rows"
        Next order
    Next selectionCase
    nativeUpdates = mNativeUpdates
    RunCompletedTaskViewOrderRegression = "PASS: G and A,B,G, both hide/fold orders, header boundary, five stable monitor passes, scroll preservation, native updates=" & nativeUpdates
    GoTo Finished
Failed:
    errorText = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
    RunCompletedTaskViewOrderRegression = errorText
Finished:
    On Error Resume Next
    StopTaskOutlineViewMonitor
    If configCaptured Then
        cfg.Range(TASK_FOLD_START_VALUE_CELL).Value2 = originalStart
        cfg.Range(TASK_FOLD_END_VALUE_CELL).Value2 = originalEnd
        cfg.Range(TASK_FOLD_VIEW_START_VALUE_CELL).Value2 = originalView
    End If
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Function

Public Function RunCompletedTaskVisibilityRegression() As String
    Dim mixed As Worksheet, allDone As Worksheet, nested As Worksheet, nestedAllDone As Worksheet
    Dim cleared As Worksheet, changed As Worksheet, shrink As Worksheet, clearCase As Long
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean
    Dim debugContext As String, subtest As String
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    EnsureConfigSheet
    mPhase = "mixed completed/manual-hidden children"
    Set mixed = NewTaskSheet("CompletedManualChildren")
    SeedTask mixed, 5, 1, False
    SeedTask mixed, 6, 2, True
    SeedTask mixed, 7, 2, False
    mixed.Activate
    RefreshGanttSheet False
    UnprotectTaskSheet mixed
    mixed.Rows(7).Hidden = True
    ApplyCalculatedColumnsProtection mixed, 7
    ToggleCompletedTaskFilter mixed
    Check mixed.Rows(6).Hidden And mixed.Rows(7).Hidden, "mixed hide setup failed"
    ToggleCompletedTaskFilter mixed
    Check Not mixed.Rows(5).Hidden And Not mixed.Rows(6).Hidden, "manual hiding was mistaken for a collapsed group"
    Check mixed.Rows(7).Hidden, "manual child hiding was lost"

    subtest = RunVisibleChildCollapseRegression()
    Check Left$(subtest, 5) = "PASS:", subtest

    mPhase = "unchanged native command while the entire completed group is hidden"
    Set allDone = NewTaskSheet("CompletedNewCollapse")
    SeedTask allDone, 5, 1, False
    SeedTask allDone, 6, 2, True
    SeedTask allDone, 7, 2, True
    SeedTask allDone, 8, 1, False
    allDone.Activate
    RefreshGanttSheet False
    ToggleCompletedTaskFilter allDone
    Check allDone.Rows(6).Hidden And allDone.Rows(7).Hidden, "completed children were not hidden"
    debugContext = " beforeShow=" & CompletedViewDebug(allDone, 8)
    allDone.Outline.ShowLevels RowLevels:=1
    debugContext = debugContext & " afterShow=" & CompletedViewDebug(allDone, 8)
    MaintainCompletedTaskView allDone
    debugContext = debugContext & " afterMaintain=" & CompletedViewDebug(allDone, 8)
    ToggleCompletedTaskFilter allDone
    debugContext = debugContext & " afterOff=" & CompletedViewDebug(allDone, 8)
    ' Excel leaves both visibility and ShowDetail unchanged for this fully hidden top-level group.
    ' This native command is a no-op; disabling the filter must not invent a collapsed group.
    Check Not allDone.Rows(5).Hidden, "unchanged outline command hid its parent after disabling the filter"
    Check Not allDone.Rows(6).Hidden And Not allDone.Rows(7).Hidden, "unchanged outline command invented a collapsed group"
    debugContext = vbNullString

    mPhase = "new collapse reveals an all-completed nested summary"
    Set nestedAllDone = NewTaskSheet("CompletedNestedAllDone")
    SeedTask nestedAllDone, 5, 1, False
    SeedTask nestedAllDone, 6, 2, True
    SeedTask nestedAllDone, 7, 3, True
    SeedTask nestedAllDone, 8, 3, True
    SeedTask nestedAllDone, 9, 2, False
    SeedTask nestedAllDone, 10, 1, False
    nestedAllDone.Activate
    RefreshGanttSheet False
    ToggleCompletedTaskFilter nestedAllDone
    Check nestedAllDone.Rows(6).Hidden And nestedAllDone.Rows(7).Hidden And nestedAllDone.Rows(8).Hidden, "nested all-completed setup failed"
    debugContext = " nestedBeforeShow=" & CompletedViewDebug(nestedAllDone, 10)
    nestedAllDone.Outline.ShowLevels RowLevels:=2
    debugContext = debugContext & " nestedAfterShow=" & CompletedViewDebug(nestedAllDone, 10)
    Check Not nestedAllDone.Rows(6).Hidden, "nested outline command did not reveal its summary"
    MaintainCompletedTaskView nestedAllDone
    debugContext = debugContext & " nestedAfterMaintain=" & CompletedViewDebug(nestedAllDone, 10)
    ToggleCompletedTaskFilter nestedAllDone
    debugContext = debugContext & " nestedAfterOff=" & CompletedViewDebug(nestedAllDone, 10)
    Check Not nestedAllDone.Rows(5).Hidden And Not nestedAllDone.Rows(6).Hidden And Not nestedAllDone.Rows(9).Hidden, "nested all-completed collapse hid its summary or sibling"
    Check nestedAllDone.Rows(7).Hidden And nestedAllDone.Rows(8).Hidden, "new nested all-completed collapse was lost"
    debugContext = vbNullString

    mPhase = "nested pre-existing collapse"
    Set nested = NewTaskSheet("CompletedNestedCollapse")
    SeedTask nested, 5, 1, False
    SeedTask nested, 6, 2, False
    SeedTask nested, 7, 3, True
    SeedTask nested, 8, 3, False
    SeedTask nested, 9, 2, True
    SeedTask nested, 10, 1, False
    nested.Activate
    RefreshGanttSheet False
    UnprotectTaskSheet nested
    nested.Rows(6).ShowDetail = False
    ApplyCalculatedColumnsProtection nested, 10
    ToggleCompletedTaskFilter nested
    ToggleCompletedTaskFilter nested
    Check Not nested.Rows(5).Hidden And Not nested.Rows(6).Hidden And Not nested.Rows(9).Hidden, "nested collapse spread to visible siblings"
    Check nested.Rows(7).Hidden And nested.Rows(8).Hidden, "nested collapse was expanded"

    mPhase = "clearing the last filter-owned task"
    Set cleared = NewTaskSheet("CompletedViewClearedLast")
    SeedTask cleared, 5, 1, True
    cleared.Activate
    RefreshGanttSheet False
    ToggleCompletedTaskFilter cleared
    Check cleared.Rows(5).Hidden, "last completed task was not hidden"
    UnprotectTaskSheet cleared
    cleared.Range("C5:W5").ClearContents
    ApplyCalculatedColumnsProtection cleared, 5
    Check GetLastDataRow(cleared) < DATA_START_ROW, "cleared task still counted as data"
    MaintainCompletedTaskView cleared
    Check Not cleared.Rows(5).Hidden, "clearing the last task left its empty row hidden"

    For clearCase = 1 To 2
        mPhase = "cleared trailing completed row, operation " & clearCase
        Set shrink = NewTaskSheet(IIf(clearCase = 1, "CompletedClearToggle", "CompletedClearRefresh"))
        SeedTask shrink, 5, 1, False
        SeedTask shrink, 6, 1, True
        shrink.Activate
        RefreshGanttSheet False
        ToggleCompletedTaskFilter shrink
        Check shrink.Rows(6).Hidden, "trailing completed task was not hidden"
        UnprotectTaskSheet shrink
        shrink.Range("C6:W6").ClearContents
        ApplyCalculatedColumnsProtection shrink, 6
        Check GetLastDataRow(shrink) = 5, "cleared trailing task did not shrink the data boundary"
        If clearCase = 1 Then
            ToggleCompletedTaskFilter shrink
            Check Not AreCompletedTasksHidden(shrink), "disable after clearing the trailing task left the filter active"
        Else
            RefreshGanttSheet False
            Check AreCompletedTasksHidden(shrink), "refresh after clearing the trailing task disabled the filter"
        End If
        Check Not shrink.Rows(5).Hidden And Not shrink.Rows(6).Hidden, "cleared trailing row retained filter-owned hiding"
    Next clearCase

    mPhase = "inserting and deleting a row while the filter is active"
    Set changed = NewTaskSheet("CompletedViewRowChanges")
    SeedTask changed, 5, 1, False
    SeedTask changed, 6, 2, True
    SeedTask changed, 7, 2, False
    SeedTask changed, 8, 1, False
    changed.Activate
    RefreshGanttSheet False
    UnprotectTaskSheet changed
    changed.Rows(7).Hidden = True
    ApplyCalculatedColumnsProtection changed, 8
    ToggleCompletedTaskFilter changed
    UnprotectTaskSheet changed
    changed.Rows(6).Insert Shift:=xlDown
    SeedTask changed, 6, 2, False
    changed.Rows(6).Hidden = False
    ApplyCalculatedColumnsProtection changed, 9
    MaintainCompletedTaskView changed
    Check Not changed.Rows(6).Hidden And changed.Rows(7).Hidden And changed.Rows(8).Hidden, "inserted row displaced hiding ownership"
    UnprotectTaskSheet changed
    changed.Rows(6).Delete Shift:=xlUp
    ApplyCalculatedColumnsProtection changed, 8
    MaintainCompletedTaskView changed
    Check changed.Rows(6).Hidden And changed.Rows(7).Hidden, "deleted row displaced hiding ownership"
    ToggleCompletedTaskFilter changed
    Check Not changed.Rows(5).Hidden And Not changed.Rows(6).Hidden And changed.Rows(7).Hidden, "row edits confused completed and manual hiding"
    ToggleCompletedTaskFilter changed
    RunCompletedTaskVisibilityRegression = "PASS: manual/completed siblings, native no-op, visible-child and nested all-completed new collapse, cleared rows on maintain/toggle/refresh, row insert/delete ownership"
    GoTo Finished
Failed:
    RunCompletedTaskVisibilityRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description & debugContext
Finished:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Function

Private Function CompletedViewDebug(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim view As Object, key As Variant, rows As Range
    Set view = CaptureTaskRowView(ws, lastRow)
    CompletedViewDebug = "[status5=" & CStr(ws.Cells(5, COL_STATUS).Value2) & ", hidden=" & HiddenRowsSignature(ws, lastRow) & _
        ", showDetail5=" & CStr(ws.Rows(5).ShowDetail) & ", levels5/6=" & ws.Rows(5).OutlineLevel & "/" & ws.Rows(6).OutlineLevel
    For Each key In Array("Collapsed", "StoredCollapsed", "Owned", "Baseline")
        Set rows = view(CStr(key))
        CompletedViewDebug = CompletedViewDebug & ", " & key & "="
        If rows Is Nothing Then
            CompletedViewDebug = CompletedViewDebug & "nothing"
        Else
            CompletedViewDebug = CompletedViewDebug & rows.Address
        End If
    Next key
    CompletedViewDebug = CompletedViewDebug & "]"
End Function

Public Function RunVisibleChildCollapseRegression() As String
    Dim ws As Worksheet, previousEvents As Boolean, previousScreenUpdating As Boolean
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    StopTaskOutlineViewMonitor
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mPhase = "new collapse with a visible unfinished child"
    Set ws = NewTaskSheet("CompletedLiveCollapse")
    SeedTask ws, 5, 1, False
    SeedTask ws, 6, 2, True
    SeedTask ws, 7, 2, False
    SeedTask ws, 8, 1, False
    ws.Activate
    RefreshGanttSheet False
    ToggleCompletedTaskFilter ws
    Check ws.Rows(6).Hidden And Not ws.Rows(7).Hidden, "visible child collapse setup failed"
    ws.Outline.ShowLevels RowLevels:=1
    MaintainCompletedTaskView ws
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(5).Hidden And ws.Rows(6).Hidden And ws.Rows(7).Hidden, "new collapse of a visible unfinished child was lost"
    RunVisibleChildCollapseRegression = "PASS: new collapse while an unfinished child is visible preserves both children after disabling the filter"
    GoTo Finished
Failed:
    RunVisibleChildCollapseRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
Finished:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Function

Private Function HiddenRowsSignature(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        HiddenRowsSignature = HiddenRowsSignature & IIf(ws.Rows(r).Hidden, "1", "0")
    Next r
End Function

Private Function TaskViewMetadataSignature(ByVal ws As Worksheet) As String
    Dim localName As Name
    For Each localName In ws.Names
        TaskViewMetadataSignature = TaskViewMetadataSignature & localName.Name & "=" & localName.RefersTo & vbLf
    Next localName
End Function

Public Function RunRefreshFailureRegression() As String
    Dim ws As Worksheet, beforeRows As String, trappedError As Long, pass As Long
    Dim previousEvents As Boolean, previousScreenUpdating As Boolean
    Dim beforeMetadata As String, trappedDescription As String
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    StopTaskOutlineViewMonitor
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set ws = NewTaskSheet("CompletedViewRefreshInit")
    SeedTask ws, 5, 1, False
    SeedTask ws, 6, 2, True
    SeedTask ws, 7, 2, False
    SeedTask ws, 8, 1, False
    SeedTask ws, 9, 2, True
    SeedTask ws, 10, 2, False
    SeedTask ws, 11, 1, False
    ws.Activate
    RefreshGanttSheet False
    UnprotectTaskSheet ws
    ws.Rows(8).ShowDetail = False
    ws.Rows(11).Hidden = True
    ApplyCalculatedColumnsProtection ws, 11
    ToggleCompletedTaskFilter ws
    beforeRows = HiddenRowsSignature(ws, 11)
    ws.Name = "CompletedViewRefreshFail"
    beforeMetadata = TaskViewMetadataSignature(ws)
    For pass = 1 To 2
        mPhase = "refresh failure after ShowAllTaskRows, protection variant " & pass
        If pass = 2 Then UnprotectTaskSheet ws
        Application.EnableEvents = (pass = 1)
        Application.ScreenUpdating = (pass = 1)
        trappedError = 0
        On Error Resume Next
        RefreshGanttSheet False
        trappedError = Err.Number
        trappedDescription = Err.Description
        Err.Clear
        On Error GoTo Failed
        Check trappedError <> 0, "injected refresh error was swallowed"
        Check InStr(trappedDescription, "Injected after ShowAllTaskRows") > 0, "refresh failed before the injected recovery point: " & trappedDescription
        Check HiddenRowsSignature(ws, 11) = beforeRows, "refresh failure changed row visibility"
        Check TaskViewMetadataSignature(ws) = beforeMetadata, "refresh failure changed row-view metadata"
        Check Not ws.Rows(8).ShowDetail, "refresh failure changed outline state"
        Check AreCompletedTasksHidden(ws), "refresh failure disabled the filter"
        Check ws.ProtectContents = (pass = 1), "refresh failure changed original protection"
        Check Application.EnableEvents = (pass = 1), "refresh failure changed event state"
        Check Application.ScreenUpdating = (pass = 1), "refresh failure changed screen updating state"
    Next pass
    RunRefreshFailureRegression = "PASS: failure after rows were exposed restores hidden rows, outline, filter, protection and application settings"
    GoTo Finished
Failed:
    RunRefreshFailureRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
Finished:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Function

Private Sub SeedTask(ByVal ws As Worksheet, ByVal r As Long, ByVal level As Long, ByVal done As Boolean)
    ws.Cells(r, COL_LEVEL).Value2 = level
    ws.Cells(r, COL_REQUEST).Value2 = "검증 업무 " & r
    ws.Cells(r, COL_PLAN_START).Value = DateSerial(2026, 10, 1)
    ws.Cells(r, COL_PLAN_END).Value = DateSerial(2026, 10, 30)
    ws.Cells(r, COL_ACTUAL_START).Value = DateSerial(2026, 10, 1)
    ws.Cells(r, COL_PROGRESS).Value2 = IIf(done, 1, 0)
    If done Then ws.Cells(r, COL_ACTUAL_END).Value = DateSerial(2026, 10, 5)
End Sub

Public Function RunCompletedTaskRegression() As String
    Dim ws As Worksheet, allDone As Worksheet, undated As Worksheet, cfg As Worksheet
    Dim button As Shape, reference As Shape, r As Long, trappedError As Long
    Dim inputs As Variant, after As Variant, c As Long, originalLimit As Variant
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    EnsureConfigSheet
    Set ws = NewTaskSheet("CompletedViewMixed")
    For r = 5 To 13
        SeedTask ws, r, 1, (r = 6 Or r = 7 Or r = 9 Or r = 11)
    Next r
    ws.Range("C6:C7").Value2 = 2
    ws.Range("C9:C10").Value2 = 2
    ws.Range("S11").Value2 = STATUS_HOLD
    ws.Activate
    mPhase = "RefreshGanttSheet False"
    RefreshGanttSheet False
    Check Not ws.Rows(6).Hidden And Not ws.Rows(9).Hidden, "initial refresh invented collapsed groups"
    UnprotectTaskSheet ws
    CreateVersionButton ws, "btnGanttRefresh", "새로고침", "칸트차트_새로고침", 1, 72
    With ws.Shapes("btnGanttRefresh")
        .Width = 71.25
        .Height = 23.5
        .TextFrame2.TextRange.Font.Name = "Calibri"
        .TextFrame2.TextRange.Font.Size = 11
    End With
    CreateVersionButton ws, COMPLETED_TASK_BUTTON_NAME, CompletedTaskButtonCaption(ws), "칸트차트_완료숨김", 2, 72
    mPhase = "ws.Rows(13).Hidden = True"
    ws.Rows(13).Hidden = True
    ws.Range("M12").Value = DateSerial(2026, 9, 1)
    ws.Range("W9").Value2 = STATUS_NORMAL
    ws.Range("W10").Value2 = STATUS_DONE
    ApplyCalculatedColumnsProtection ws, 13
    inputs = ws.Range("C5:O13").Value2
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check AreCompletedTasksHidden(ws), "filter not enabled"
    Check ws.Rows(5).Hidden And ws.Rows(6).Hidden And ws.Rows(7).Hidden, "completed subtree still visible"
    Check ws.Rows(9).Hidden, "stale normal status not recalculated"
    Check Not ws.Rows(8).Hidden And Not ws.Rows(10).Hidden, "unfinished parent/child hidden"
    Check Not ws.Rows(11).Hidden And Not ws.Rows(12).Hidden, "hold/error hidden"
    Check ws.Rows(13).Hidden, "manual hidden row was expanded"
    Check GetLastDataRow(ws) = 13, "hidden last row not found"
    Check ws.ProtectContents, "protection lost"
    Check Not Application.EnableEvents And Not Application.ScreenUpdating, "application state not restored"
    after = ws.Range("C5:O13").Value2
    For r = 1 To UBound(inputs, 1)
        For c = 1 To UBound(inputs, 2)
            Check CStr(inputs(r, c)) = CStr(after(r, c)), "task input changed"
        Next c
    Next r
    Set button = ws.Shapes(COMPLETED_TASK_BUTTON_NAME)
    Set reference = ws.Shapes("btnGanttRefresh")
    Check button.TextFrame2.TextRange.Text = "완료 표시", "enabled caption"
    Check Abs(button.Width - reference.Width) < 0.1 And Abs(button.Height - reference.Height) < 0.1, "button size differs"
    Check button.AutoShapeType = reference.AutoShapeType, "button shape differs"
    Check button.Fill.ForeColor.RGB = reference.Fill.ForeColor.RGB, "button fill differs"
    Check button.Line.ForeColor.RGB = reference.Line.ForeColor.RGB, "button line differs"
    Check button.TextFrame2.TextRange.Font.Name = reference.TextFrame2.TextRange.Font.Name, "button font differs"
    Check button.TextFrame2.TextRange.Font.Size = reference.TextFrame2.TextRange.Font.Size, "button font size differs"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(5).Hidden And Not ws.Rows(6).Hidden And Not ws.Rows(9).Hidden, "completed rows not restored"
    Check ws.Rows(13).Hidden, "manual hiding lost on disable"
    Check button.TextFrame2.TextRange.Text = "완료 숨김", "disabled caption"
    UnprotectTaskSheet ws
    mPhase = "ws.Rows(8).ShowDetail = False"
    ws.Rows(8).ShowDetail = False
    ws.Range("M12").Value = DateSerial(2026, 10, 30)
    ApplyCalculatedColumnsProtection ws, 13
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    mPhase = "RefreshGanttSheet False"
    RefreshGanttSheet False
    Check AreCompletedTasksHidden(ws) And ws.Rows(5).Hidden, "refresh forgot filter"
    Check ws.Rows(9).Hidden And ws.Rows(10).Hidden, "refresh expanded collapsed group"
    Check ws.Rows(13).Hidden, "refresh expanded manual hiding"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(5).Hidden And Not ws.Rows(6).Hidden, "filter mistaken for outline collapse"
    Check ws.Rows(9).Hidden And ws.Rows(10).Hidden, "disable expanded collapsed group"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    mPhase = "ws.Rows(8).ShowDetail = True"
    ws.Rows(8).ShowDetail = True
    mPhase = "MaintainCompletedTaskView ws"
    MaintainCompletedTaskView ws
    Check ws.Rows(9).Hidden And Not ws.Rows(10).Hidden, "outline expansion exposed completed child"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(9).Hidden And Not ws.Rows(10).Hidden, "expanded group not restored"
    ' Force a calculation failure without opening a password prompt.
    Set cfg = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    originalLimit = cfg.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value2
    cfg.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value2 = 1E+20
    On Error Resume Next
    ToggleCompletedTaskFilter ws
    trappedError = Err.Number
    Err.Clear
    On Error GoTo Failed
    cfg.Range(TASK_MAX_LENGTH_LEVEL1_VALUE_CELL).Value2 = originalLimit
    Check trappedError <> 0, "calculation failure was swallowed"
    Check ws.ProtectContents And Not AreCompletedTasksHidden(ws), "failure changed protection/filter"
    Check ws.Rows(13).Hidden, "failure changed manual visibility"
    Check Not Application.EnableEvents And Not Application.ScreenUpdating, "failure changed application state"
    Set cfg = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    trappedError = 0
    On Error Resume Next
    mPhase = "ToggleCompletedTaskFilter cfg"
    ToggleCompletedTaskFilter cfg
    trappedError = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check trappedError <> 0, "config sheet accepted as task sheet"
    Set undated = NewTaskSheet("CompletedViewUndated")
    undated.Range("C5").Value2 = 1
    undated.Range("H5").Value2 = "날짜 없는 완료 업무"
    undated.Range("P5").Value2 = 1
    mPhase = "ToggleCompletedTaskFilter undated"
    ToggleCompletedTaskFilter undated
    Check undated.Rows(5).Hidden, "undated completed task not hidden"
    mPhase = "ToggleCompletedTaskFilter undated"
    ToggleCompletedTaskFilter undated
    Check Not undated.Rows(5).Hidden, "undated completed task not restored"
    Set allDone = NewTaskSheet("CompletedViewAll")
    For r = 5 To 7
        SeedTask allDone, r, 1, True
    Next r
    allDone.Activate
    mPhase = "RefreshGanttSheet False"
    RefreshGanttSheet False
    UnprotectTaskSheet allDone
    CreateVersionButton allDone, COMPLETED_TASK_BUTTON_NAME, "완료 숨김", "칸트차트_완료숨김", 1, 72
    ApplyCalculatedColumnsProtection allDone, 7
    mPhase = "ToggleCompletedTaskFilter allDone"
    ToggleCompletedTaskFilter allDone
    Check allDone.Rows(5).Hidden And allDone.Rows(7).Hidden, "all-completed hiding failed"
    Check GetLastDataRow(allDone) = 7, "all-hidden data lost"
    mPhase = "RefreshGanttSheet False"
    RefreshGanttSheet False
    Check allDone.Rows(5).Hidden And allDone.Rows(7).Hidden, "all-hidden refresh failed"
    Check Not AreCompletedTasksHidden(ws), "filter leaked between sheets"
    RunCompletedTaskRegression = "PASS: recalculation, hierarchy, manual/outline preservation, toggle/refresh, all-hidden/undated tasks, button design, failure cleanup, per-sheet state"
    Exit Function
Failed:
    RunCompletedTaskRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
End Function

Public Function RunSavedCompletedTaskRegression() As String
    Dim ws As Worksheet, mixed As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("CompletedViewAll")
    Check AreCompletedTasksHidden(ws), "saved filter not restored"
    Check ws.Rows(5).Hidden And ws.Rows(7).Hidden, "saved row visibility not restored"
    Check ws.Shapes(COMPLETED_TASK_BUTTON_NAME).TextFrame2.TextRange.Text = "완료 표시", "saved caption not restored"
    Check Not AreCompletedTasksHidden(ThisWorkbook.Worksheets("CompletedViewMixed")), "saved state leaked between sheets"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(5).Hidden And Not ws.Rows(7).Hidden, "cannot disable after reopen"
    mPhase = "saved mixed visibility after row insertion/deletion"
    Set mixed = ThisWorkbook.Worksheets("CompletedViewRowChanges")
    Check AreCompletedTasksHidden(mixed), "mixed filter was not restored after reopen"
    Check Not mixed.Rows(5).Hidden And mixed.Rows(6).Hidden And mixed.Rows(7).Hidden, "saved mixed visibility was not restored"
    ToggleCompletedTaskFilter mixed
    Check Not mixed.Rows(6).Hidden And mixed.Rows(7).Hidden, "saved baseline lost manual/completed ownership"
    RunSavedCompletedTaskRegression = "PASS: save/reopen, per-sheet state, mixed manual/completed baseline after row edits, caption and restore"
    Exit Function
Failed:
    RunSavedCompletedTaskRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
End Function
