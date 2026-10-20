Attribute VB_Name = "CompletedTaskViewRegression"
Option Explicit
Private mPhase As String
' 2026-10-20 김좌현 수정 : 완료 숨김 토글/계층/새로고침/상태 복구/저장 후 유지 검증

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
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("CompletedViewAll")
    Check AreCompletedTasksHidden(ws), "saved filter not restored"
    Check ws.Rows(5).Hidden And ws.Rows(7).Hidden, "saved row visibility not restored"
    Check ws.Shapes(COMPLETED_TASK_BUTTON_NAME).TextFrame2.TextRange.Text = "완료 표시", "saved caption not restored"
    Check Not AreCompletedTasksHidden(ThisWorkbook.Worksheets("CompletedViewMixed")), "saved state leaked between sheets"
    mPhase = "ToggleCompletedTaskFilter ws"
    ToggleCompletedTaskFilter ws
    Check Not ws.Rows(5).Hidden And Not ws.Rows(7).Hidden, "cannot disable after reopen"
    RunSavedCompletedTaskRegression = "PASS: save/reopen, persisted per-sheet state, caption and restore"
    Exit Function
Failed:
    RunSavedCompletedTaskRegression = "FAIL: " & mPhase & " / " & Err.Source & " / " & Err.Description
End Function
