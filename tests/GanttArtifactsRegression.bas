' 2026-11-11 김좌현 수정 : PPT 삭제 확인·취소, 행 이동·삭제, 반복 가져오기와 잔존물 정리 검증
' 2026-10-29 김좌현 수정 : 통합 새로고침 경로의 기존 결과 보존 검증
' Appended to modVersionMigration only in an isolated workbook.
' 2026-10-29 김좌현 수정 : 첨부 개체의 행 연결·가져오기·반복 새로고침·저장 재열기 검증
Private Sub ArtifactCheck(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 7950, "GanttArtifactsRegression", message
End Sub

Public Function RunActualRowRefreshRegression() As String
    Dim ws As Worksheet, lastRow As Long, r As Long, pass As Long
    Dim before() As Boolean, hiddenBefore As Long, hiddenAfter As Long, changed As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ws.Activate
    lastRow = GetLastDataRow(ws)
    ReDim before(5 To lastRow)
    For r = 5 To lastRow
        before(r) = ws.Rows(r).Hidden
        If before(r) Then hiddenBefore = hiddenBefore + 1
    Next r
    RefreshGanttSheet False
    For r = 5 To lastRow
        If ws.Rows(r).Hidden Then hiddenAfter = hiddenAfter + 1
        If ws.Rows(r).Hidden And Not before(r) Then changed = changed & r & ":" & ws.Cells(r, COL_STATUS).Value2 & ";"
    Next r
    UnprotectTaskSheet ws
    ClearCompletedTaskFilter ws, lastRow
    ShowAllTaskRows ws, lastRow
    For pass = 1 To 3
        RefreshGanttSheet False
        For r = 5 To lastRow
            ArtifactCheck Not ws.Rows(r).Hidden, "filter OFF refresh hid row " & r & " at pass " & pass
        Next r
    Next pass
    RunActualRowRefreshRegression = "PASS: actual row visibility before=" & hiddenBefore & ", after=" & hiddenAfter & _
        ", newly hidden=" & changed & "; filter OFF 3 refreshes all " & (lastRow - 4) & " rows visible"
    Exit Function
Failed:
    RunActualRowRefreshRegression = "FAIL: " & Err.Description & "; newly hidden=" & changed
End Function

Public Function RunActualArtifactRegression(ByVal sourceName As String, ByVal baseline As Boolean) As String
    Dim source As Worksheet, target As Worksheet, copied As Long, failed As Long
    Dim o As OLEObject, result As String, expected As Long
    Dim plan As Object, logs As Long, states As Long
    On Error GoTo Failed
    Set source = Application.Workbooks(sourceName).Worksheets("Sheet1")
    Set target = ThisWorkbook.Worksheets("Sheet1")
    expected = source.OLEObjects.Count
    For Each o In source.OLEObjects
        result = result & o.Name & "@" & o.TopLeftCell.Row & "; "
    Next o
    Set plan = BuildWebMigrationPlan(source.Parent, ThisWorkbook)
    ImportTaskSheet source, target, copied, failed
    EnsureWebHistorySheets
    ApplyWebMigrationPlan plan, ThisWorkbook, logs, states
    PrepareWebTaskSheet target
    result = "source=" & result & "copied=" & copied & ", failed=" & failed & "; target="
    For Each o In target.OLEObjects
        result = result & o.Name & "@" & o.TopLeftCell.Row & "; "
    Next o
    ArtifactCheck copied = expected And failed = 0, result
    If baseline Then
        RunActualArtifactRegression = "PASS: " & result
        Exit Function
    End If
    For Each o In source.OLEObjects
        ArtifactCheck Not target.OLEObjects(o.Name) Is Nothing, "import changed task attachment: " & o.Name
    Next o
    target.Activate
    ArtifactCheck RefreshMigratedSheet(target), "import refresh failed"
    RefreshGanttSheet False
    For Each o In source.OLEObjects
        ArtifactCheck Not target.OLEObjects(o.Name) Is Nothing, "refresh changed task attachment: " & o.Name
    Next o
    RunActualArtifactRegression = "PASS: actual PPT import and refresh: " & result
    Exit Function
Failed:
    RunActualArtifactRegression = "FAIL: " & Err.Description & " / " & result
End Function

Public Function RunGanttArtifactsRegression() As String
    Dim ws As Worksheet, source As Worksheet, sourceBook As Workbook, o As OLEObject
    Dim r As Long, copied As Long, failed As Long, stepName As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "ArtifactFixture"
    SetupDataHeaders ws
    For r = 5 To 12
        ws.Cells(r, COL_LEVEL).Value2 = 1
        ws.Cells(r, COL_REQUEST).Value2 = "Attachment task " & r
        ws.Cells(r, COL_PLAN_START).Value = Date
        ws.Cells(r, COL_PLAN_END).Value = Date + 7
    Next r
    ws.Cells(6, COL_LEVEL).Value2 = 2
    ws.Cells(7, COL_LEVEL).Value2 = 3
    ws.Cells(8, COL_LEVEL).Value2 = 3
    ws.Activate
    stepName = "initial fixture refresh"
    RefreshGanttSheet False
    UnprotectTaskSheet ws
    stepName = "create fixture OLE attachments"
    For r = 7 To 8
        Set o = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Left:=ws.Cells(r, COL_NOTE).Left + 1, _
                                 Top:=ws.Cells(r, COL_NOTE).Top + 1, Width:=1, Height:=1)
        o.Name = "oleGanttPpt_" & r
        o.Placement = xlMoveAndSize
        o.Visible = False
        ws.Cells(r, COL_NOTE).Value2 = ChrW(&H25A0)
    Next r
    stepName = "collapsed attachment refresh"
    ws.Rows(6).ShowDetail = False
    RefreshGanttSheet False
    ArtifactCheck ws.Rows(7).Hidden And ws.Rows(8).Hidden, "collapsed children expanded"
    ArtifactCheck Not ws.Rows(9).Hidden, "unrelated row collapsed"
    ArtifactCheck ws.OLEObjects.Count = 2, "refresh lost attachments"
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_7") Is Nothing, "attachment 7 moved"
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_8") Is Nothing, "attachment 8 moved"
    stepName = "repeated hidden refresh"
    RefreshGanttSheet False
    RefreshGanttSheet False
    UnprotectTaskSheet ws
    ws.Rows(6).ShowDetail = True
    RefreshGanttSheet False
    ArtifactCheck ws.OLEObjects("oleGanttPpt_7").TopLeftCell.Row = 7, "expanded attachment 7 moved"
    ArtifactCheck ws.OLEObjects("oleGanttPpt_8").TopLeftCell.Row = 8, "expanded attachment 8 moved"
    stepName = "row insertion while collapsed"
    UnprotectTaskSheet ws
    ws.Rows(6).ShowDetail = False
    ws.Rows(7).Insert
    ws.Cells(7, COL_LEVEL).Value2 = 3
    ws.Cells(7, COL_REQUEST).Value2 = "Inserted task"
    RefreshGanttSheet False
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_8") Is Nothing, "first attachment did not follow insertion"
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_9") Is Nothing, "second attachment did not follow insertion"
    stepName = "migration of hidden attachments"
    ws.Copy
    Set sourceBook = ActiveWorkbook
    Set source = sourceBook.Worksheets(1)
    ImportTaskSheet source, ws, copied, failed
    ArtifactCheck copied = 2 And failed = 0, "import copied=" & copied & ", failed=" & failed
    ArtifactCheck source.Rows(8).Hidden And source.Rows(9).Hidden, "source hidden state changed"
    ArtifactCheck Not source.OLEObjects("oleGanttPpt_8").Visible, "source attachment visibility changed"
    sourceBook.Close False
    Set sourceBook = Nothing
    ws.Activate
    ArtifactCheck RefreshMigratedSheet(ws), "migration refresh failed"
    RefreshGanttSheet False
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_8") Is Nothing, "import attachment 8 moved"
    ArtifactCheck Not ws.OLEObjects("oleGanttPpt_9") Is Nothing, "import attachment 9 moved"
    RunGanttArtifactsRegression = "PASS: collapsed/repeated unified refresh, row insertion, hidden attachment import and source preservation"
    Exit Function
Failed:
    RunGanttArtifactsRegression = "FAIL: " & stepName & ": " & Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close False
End Function

Public Function RunSavedArtifactRegression() As String
    Dim ws As Worksheet, o As OLEObject, expected As Object, key As Variant
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ArtifactFixture")
    On Error GoTo Failed
    Set expected = CreateObject("Scripting.Dictionary")
    For Each o In ws.OLEObjects
        expected.Add o.Name, True
    Next o
    ws.Activate
    RefreshGanttSheet False
    For Each key In expected.Keys
        ArtifactCheck Not ws.OLEObjects(CStr(key)) Is Nothing, "reopen changed attachment " & CStr(key)
    Next key
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PptDeleteFixture")
    On Error GoTo Failed
    If Not ws Is Nothing Then
        ws.Activate
        RefreshGanttSheet False
        ArtifactCheck ws.OLEObjects.Count = 2, "reopen changed deletion/import object count"
        ArtifactCheck Not ws.OLEObjects("oleGanttPpt_6") Is Nothing, "reopen lost imported PPT"
        ArtifactCheck Not ws.OLEObjects("UnrelatedControl") Is Nothing, "reopen lost unrelated control"
        ArtifactCheck PptAnchorCount(ws) = 1, "reopen left stale anchors"
        ArtifactCheck ws.Cells(6, COL_NOTE).Hyperlinks.Count = 1, "reopen lost PPT link"
        ArtifactCheck ws.Cells(7, COL_NOTE).Value2 = "", "reopen restored deleted icon"
    End If
    RunSavedArtifactRegression = "PASS: saved/reopened attachment names and repeated refresh"
    Exit Function
Failed:
    RunSavedArtifactRegression = "FAIL: " & Err.Description
End Function

Public Function PptTestMessageBox(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal title As String = "") As VbMsgBoxResult
    PptTestLastMessage = prompt
    If (buttons And vbYesNo) = vbYesNo Then
        PptTestMessageBox = PptTestDecision
    ElseIf Left$(prompt, Len("PPT 개체 삽입 완료")) = "PPT 개체 삽입 완료" Or _
           prompt = "해당 행에는 이미 PPT 개체가 삽입되어 있습니다." Then
        PptTestMessageBox = vbOK
    Else
        Err.Raise vbObjectError + 7952, "PptTestMessageBox", prompt
    End If
End Function

Public Sub PptTestTrace(ByVal message As String)
    Dim handle As Integer
    handle = FreeFile
    Open Environ$("TEMP") & "\" & ThisWorkbook.Name & ".trace" For Append As #handle
    Print #handle, message
    Close #handle
End Sub

Public Sub PptTestShowPopup(ByVal menu As Object)
    PptTestPopupCount = PptTestPopupCount + 1
    ArtifactCheck menu.Controls.Count = 1, "unexpected popup control count"
    ArtifactCheck menu.Controls(1).Caption = "첨부 PPT 삭제", "delete menu caption"
    If PptTestExecuteMenu Then menu.Controls(1).Execute
End Sub

Private Sub AddDeleteFixturePpt(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim o As OLEObject
    Set o = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Left:=ws.Cells(targetRow, COL_NOTE).Left + 1, _
                             Top:=ws.Cells(targetRow, COL_NOTE).Top + 1, Width:=1, Height:=1)
    o.Name = "oleGanttPpt_" & targetRow
    o.Visible = False
    o.Placement = xlMoveAndSize
    BindGanttPptObjectRow ws, o, targetRow
    ws.Cells(targetRow, COL_NOTE).Value2 = ChrW(&H25A0)
End Sub

Private Function PptAnchorCount(ByVal ws As Worksheet) As Long
    Dim item As Name, key As String
    For Each item In ws.Names
        key = Mid$(item.Name, InStrRev(item.Name, "!") + 1)
        If Left$(key, Len("_GanttPptRow_")) = "_GanttPptRow_" Then PptAnchorCount = PptAnchorCount + 1
    Next item
End Function

Public Function RunGanttPptDeleteRegression() As String
    Dim ws As Worksheet, source As Worksheet, o As OLEObject, i As Long, n As Long
    Dim copied As Long, failed As Long, stage As String, controlCount As Long, deletedAnchor As String
    Dim previousEvents As Boolean, previousScreen As Boolean
    Dim cancel As Boolean
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "PptDeleteFixture"
    SetupDataHeaders ws
    For i = 5 To 12
        ws.Cells(i, COL_LEVEL).Value2 = 1
        ws.Cells(i, COL_REQUEST).Value2 = "PPT delete task " & i
        ws.Cells(i, COL_PLAN_START).Value2 = Date
        ws.Cells(i, COL_PLAN_END).Value2 = Date + 7
    Next i
    ws.Activate
    AddDeleteFixturePpt ws, 7
    AddDeleteFixturePpt ws, 9
    RefreshGanttSheet False
    ArtifactCheck PptAnchorCount(ws) = 2, "initial anchors"
    stage = "right-click filtering and popup cancel"
    ArtifactCheck IsGanttPptDeleteCell(ws, ws.Cells(7, COL_NOTE)), "PPT cell not recognized"
    ArtifactCheck Not IsGanttPptDeleteCell(ws, ws.Cells(7, COL_REQUEST)), "normal cell intercepted"
    ArtifactCheck Not IsGanttPptDeleteCell(ws, ws.Range(ws.Cells(7, COL_NOTE), ws.Cells(9, COL_NOTE))), "multi-cell intercepted"
    controlCount = Application.CommandBars.Count
    PptTestExecuteMenu = False
    ThisWorkbook.PptTestRightClick ws, ws.Cells(7, COL_REQUEST), cancel
    ArtifactCheck Not cancel And PptTestPopupCount = 0, "normal right-click changed"
    ThisWorkbook.PptTestRightClick ws, ws.Cells(7, COL_NOTE), cancel
    ArtifactCheck cancel And PptTestPopupCount = 1, "PPT right-click event not connected"
    ArtifactCheck ws.OLEObjects.Count = 2 And PptAnchorCount(ws) = 2, "dismissed popup deleted attachment"
    ArtifactCheck Application.CommandBars.Count = controlCount, "popup leaked"
    stage = "confirmation cancel after row insertion"
    UnprotectTaskSheet ws
    ws.Rows(7).Insert
    ws.Cells(7, COL_LEVEL).Value2 = 1
    ws.Cells(7, COL_REQUEST).Value2 = "inserted row without attachment"
    ' Names still end in 7/9; actual attachment rows are now 8/10.
    ws.Cells(5, COL_REQUEST).Select
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    stage = "real PPT insertion with shifted names and duplicate prevention"
    ws.Cells(7, COL_REQUEST).Select
    칸트차트_개체삽입
    ArtifactCheck ws.OLEObjects.Count = 3 And PptAnchorCount(ws) = 3, "new PPT collided with shifted object name"
    ArtifactCheck Left$(PptTestLastMessage, Len("PPT 개체 삽입 완료")) = "PPT 개체 삽입 완료", "PPT insertion did not finish"
    ArtifactCheck DeleteGanttPptForRow(ws, 7), "real PPT deletion failed"
    ws.Cells(8, COL_REQUEST).Select
    칸트차트_개체삽입
    ArtifactCheck ws.OLEObjects.Count = 2 And PptAnchorCount(ws) = 2, "duplicate insertion after row shift"
    ArtifactCheck PptTestLastMessage = "해당 행에는 이미 PPT 개체가 삽입되어 있습니다.", "duplicate insertion was not detected"
    ws.Cells(5, COL_REQUEST).Select
    PptTestExecuteMenu = True
    PptTestDecision = vbNo
    ShowGanttPptDeleteMenu ws, ws.Cells(8, COL_NOTE)
    ArtifactCheck ws.OLEObjects.Count = 2 And PptAnchorCount(ws) = 2, "confirmation cancel deleted attachment"
    stage = "confirmed delete uses context row, preserves protection and application state"
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    PptTestDecision = vbYes
    ShowGanttPptDeleteMenu ws, ws.Cells(8, COL_NOTE)
    ArtifactCheck ws.OLEObjects.Count = 1 And PptAnchorCount(ws) = 1, "delete left object or anchor"
    ArtifactCheck ws.Cells(8, COL_NOTE).Value2 = "" And ws.Cells(8, COL_NOTE).Hyperlinks.Count = 0, "delete left icon/link"
    ArtifactCheck ws.Cells(8, COL_NOTE).Interior.Pattern = xlNone, "delete left red fill"
    ArtifactCheck ws.Cells(8, COL_REQUEST).Value2 = "PPT delete task 7", "delete changed task text"
    ArtifactCheck ws.ProtectContents And Not ws.ProtectDrawingObjects, "protection changed"
    ArtifactCheck Not Application.EnableEvents And Not Application.ScreenUpdating, "application state changed"
    ArtifactCheck GetGanttPptObjectRow(ws, ws.OLEObjects(1)) = 10, "wrong attachment deleted"
    stage = "delete an unattached row with a stale object name elsewhere"
    ArtifactCheck Not DeleteGanttPptForRow(ws, 9), "stale suffix caused wrong deletion"
    ArtifactCheck ws.OLEObjects.Count = 1, "unrelated attachment removed"
    stage = "deleted row cleanup and unrelated object preservation"
    UnprotectTaskSheet ws
    Set o = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Left:=0, Top:=0, Width:=12, Height:=12)
    o.Name = "UnrelatedControl"
    ws.Rows(10).Delete
    RefreshGanttSheet False
    ArtifactCheck ws.OLEObjects.Count = 1 And ws.OLEObjects(1).Name = "UnrelatedControl", "orphan cleanup removed unrelated object or kept PPT"
    ArtifactCheck PptAnchorCount(ws) = 0, "deleted row left anchor"
    stage = "import replacement clears old, temporary and orphan attachments"
    Set source = ThisWorkbook.Worksheets.Add
    source.Name = "PptImportFixture"
    SetupDataHeaders source
    For i = 5 To 9
        source.Cells(i, COL_LEVEL).Value2 = 1
        source.Cells(i, COL_REQUEST).Value2 = "new source " & i
        source.Cells(i, COL_PLAN_START).Value2 = Date
        source.Cells(i, COL_PLAN_END).Value2 = Date + 7
    Next i
    source.Activate
    source.Cells(6, COL_REQUEST).Select
    칸트차트_개체삽입
    ArtifactCheck source.OLEObjects.Count = 1, "real source PPT insertion"
    UnprotectTaskSheet source
    AddDeleteFixturePpt source, 8
    source.Rows(8).Delete
    UnprotectTaskSheet ws
    AddDeleteFixturePpt ws, 7
    AddDeleteFixturePpt ws, 8
    ws.OLEObjects("oleGanttPpt_7").Name = "tmpGanttPptOle_fixture"
    ws.Rows(8).Delete
    ws.Names.Add Name:="_GanttPptRow_unused", RefersTo:="=#REF!", Visible:=False
    Set o = Nothing
    With ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 12, 12)
        .Name = "shpGanttPpt_legacy"
    End With
    stage = "partial import failure removes pasted object and its anchor"
    PptTestFailCopy = True
    copied = 0: failed = 0
    ImportTaskSheet source, ws, copied, failed
    PptTestFailCopy = False
    ArtifactCheck copied = 0 And failed = 1, "injected failure not counted"
    ArtifactCheck ws.OLEObjects.Count = 1 And PptAnchorCount(ws) = 0, "partial paste left object/anchor"
    ArtifactCheck ws.Cells(6, COL_NOTE).Value2 = "", "partial paste left icon"
    For n = 1 To ws.Shapes.Count
        ArtifactCheck ws.Shapes(n).Name <> "shpGanttPpt_legacy", "import left legacy PPT shape"
    Next n
    stage = "repeat real PPT import after failed paste"
    For i = 1 To 2
        copied = 0: failed = 0
        ImportTaskSheet source, ws, copied, failed
        ArtifactCheck copied = 1 And failed = 0, "import counts: " & copied & "/" & failed
        ArtifactCheck ws.OLEObjects.Count = 2, "repeated import left duplicate or removed unrelated control"
        ArtifactCheck PptAnchorCount(ws) = 1, "repeated import left anchors"
        ArtifactCheck Not ws.OLEObjects("oleGanttPpt_6") Is Nothing, "new PPT missing"
        ArtifactCheck ws.Cells(7, COL_NOTE).Value2 = "", "old icon survived replacement"
    Next i
    ArtifactCheck source.OLEObjects.Count = 2, "import modified source attachments"
    ws.Activate
    RefreshGanttSheet False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    RunGanttPptDeleteRegression = "PASS: popup dismiss/cancel/confirm, shifted-row deletion, local cleanup, application/protection state, deleted-row orphans, unrelated controls, repeated import and source preservation"
    Exit Function
Failed:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    RunGanttPptDeleteRegression = "FAIL: " & stage & ": " & Err.Description
End Function
