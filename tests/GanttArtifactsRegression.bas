' Appended to modVersionMigration only in an isolated workbook.
' 2026-10-29 김좌현 수정 : 첨부 개체의 행 연결·가져오기·반복 새로고침·저장 재열기 검증
Private Sub ArtifactCheck(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 7950, "GanttArtifactsRegression", message
End Sub

Public Function RunActualArtifactRegression(ByVal sourceName As String, ByVal baseline As Boolean) As String
    Dim source As Worksheet, target As Worksheet, copied As Long, failed As Long
    Dim o As OLEObject, result As String, expected As Long
    On Error GoTo Failed
    Set source = Application.Workbooks(sourceName).Worksheets("Sheet1")
    Set target = ThisWorkbook.Worksheets("Sheet1")
    expected = source.OLEObjects.Count
    For Each o In source.OLEObjects
        result = result & o.Name & "@" & o.TopLeftCell.Row & "; "
    Next o
    ImportTaskSheet source, target, copied, failed
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
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    UnprotectTaskSheet ws
    ws.Cells.Clear
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
    RefreshGanttSheet False
    UnprotectTaskSheet ws
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
    RefreshGanttOnly False
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
    RunGanttArtifactsRegression = "PASS: collapsed/repeated/full/fast refresh, row insertion, hidden attachment import and source preservation"
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
    Set expected = CreateObject("Scripting.Dictionary")
    For Each o In ws.OLEObjects
        expected.Add o.Name, True
    Next o
    ws.Activate
    RefreshGanttSheet False
    For Each key In expected.Keys
        ArtifactCheck Not ws.OLEObjects(CStr(key)) Is Nothing, "reopen changed attachment " & CStr(key)
    Next key
    RunSavedArtifactRegression = "PASS: saved/reopened attachment names and repeated refresh"
    Exit Function
Failed:
    RunSavedArtifactRegression = "FAIL: " & Err.Description
End Function
