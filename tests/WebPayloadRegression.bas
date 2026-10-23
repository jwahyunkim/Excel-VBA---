' Append to modWebPayload only in a disposable workbook copy.
' 2026-10-22 김좌현 수정 : 등록 기준·본문·필수값·웹업무ID와 읽기 전용 미리보기 회귀 검증
' 2026-10-23 김좌현 수정 : Y 선택 행 한정·미선택 오류 행 제외와 등록 단위 중복 방지 검증
Public Function RunWebPayloadRegression() As String
    Dim ws As Worksheet, configWs As Worksheet, copiedWs As Worksheet, selectionWs As Worksheet, items As Collection, item As Object
    Dim values As Object, first As Object, checked As Long, stage As String
    Dim snapshot As String, body As String, r As Long, sourceId As String
    On Error GoTo Failed
    stage = "seed web options and hierarchy"
    EnsureWebConfigSheet
    Set configWs = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    WebRegressionSeedOption configWs, 5, "동명이인 [user1]", "동명이인", "user1"
    configWs.Cells(10, 5).Value2 = "동명이인 [user2]"
    configWs.Cells(10, 6).Value2 = "동명이인"
    configWs.Cells(10, 7).Value2 = "user2"
    WebRegressionSeedOption configWs, 9, "개발", "개발", "DEV"
    WebRegressionSeedOption configWs, 13, "처리중", "처리중", "IP"
    WebRegressionSeedOption configWs, 17, "생산", "생산", "MES"
    WebRegressionSeedOption configWs, 21, "처리자 [worker]", "처리자", "worker"
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "WebPayloadFixture"
    For r = 5 To 14
        If r <> 9 Then WebRegressionSeedRow ws, r
    Next r
    ws.Cells(5, COL_REQUEST).Value2 = "그룹 A"
    ws.Cells(5, COL_LEVEL).Value2 = 1
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(6, COL_REQUEST).Value2 = "레벨2 분기"
    ws.Cells(6, COL_LEVEL).Value2 = 2
    ws.Cells(7, COL_REQUEST).Value2 = "레벨3 첫 요청"
    ws.Cells(7, COL_LEVEL).Value2 = 3
    ws.Cells(7, COL_MODIFICATION).Value2 = " 첫 수정 " & vbCrLf & vbCrLf & "둘째 수정" & vbLf
    ws.Cells(7, COL_WORK_MINUTES).Value2 = 0
    ws.Cells(8, COL_REQUEST).Value2 = "레벨3 둘째 요청"
    ws.Cells(8, COL_LEVEL).Value2 = 3
    ws.Cells(8, COL_MODIFICATION).Value2 = vbTab & " " & vbCrLf
    ws.Cells(10, COL_REQUEST).Value2 = "레벨3 없는 레벨2"
    ws.Cells(10, COL_LEVEL).Value2 = 2
    ws.Cells(11, COL_REQUEST).Value2 = "그룹 B"
    ws.Cells(11, COL_LEVEL).Value2 = 1
    ws.Cells(11, COL_WEB_LEVEL).Value2 = 1
    ws.Cells(12, COL_REQUEST).Value2 = "B 하위"
    ws.Cells(12, COL_LEVEL).Value2 = 2
    ws.Cells(13, COL_REQUEST).Value2 = "레벨1 단독"
    ws.Cells(13, COL_LEVEL).Value2 = 1
    ws.Cells(13, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(14, COL_REQUEST).Value2 = "선택 안 한 단독"
    ws.Cells(14, COL_LEVEL).Value2 = 1
    ws.Cells(14, COL_WEB_LEVEL).Value2 = 2
    ws.Cells(14, COL_WEB_ENABLED).Value2 = "N"

    stage = "level 3 with level 2 leaf fallback and group isolation"
    snapshot = WebRegressionSnapshot(ws)
    Application.EnableEvents = False
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items.Count), "5", "three A targets, B root and one empty-group error", checked
    WebRegressionEqual CStr(items(1)("Row")), "7", "L2 with L3 is excluded", checked
    WebRegressionEqual CStr(items(2)("Row")), "8", "every L3 is selected", checked
    WebRegressionEqual CStr(items(3)("Row")), "10", "childless L2 fallback", checked
    WebRegressionEqual CStr(items(4)("Row")), "11", "next group owns its own level", checked
    WebRegressionEqual CStr(items(5)("Row")), "13", "root-only group diagnostic", checked
    WebRegressionContains CStr(items(5)("Error")), "해당하는 업무가 없습니다", "root-only group is not silently registered", checked
    For r = 1 To 4
        WebRegressionEqual CStr(items(r)("Error")), "", "valid selected payload " & CStr(r), checked
    Next r
    WebRegressionOnlySelected ws, items, checked
    Set item = items(1)
    Set values = item("Values")
    WebRegressionEqual CStr(item("SourceId")), "fixture-7", "stable source ID", checked
    WebRegressionEqual CStr(item("GroupRow")), "5", "owning level 1 group", checked
    WebRegressionEqual CStr(values("GUBUN")), "PJT", "project performance form", checked
    WebRegressionEqual CStr(values("REQ_TIT")), "[소분류] 레벨3 첫 요청", "title format", checked
    WebRegressionEqual CStr(values("REQ_UID")), "user1", "requester code, not display name", checked
    WebRegressionEqual CStr(values("ACT_GUBUN")), "DEV", "action code", checked
    WebRegressionEqual CStr(values("STS")), "IP", "status code", checked
    WebRegressionEqual CStr(values("PGM_MD")), "MES", "module code", checked
    WebRegressionEqual CStr(values("ACT_UID")), "worker", "handler code", checked
    WebRegressionEqual CStr(values("PGM_NM")), "소분류", "program name is minor classification", checked
    WebRegressionEqual CStr(values("PGM_ID")), "PGM007", "program ID is per row", checked
    WebRegressionEqual CStr(values("ACT_DT_FR")), "2026-10-01", "actual start only", checked
    WebRegressionEqual CStr(values("ACT_DT")), "2026-10-02", "actual finish only", checked
    WebRegressionEqual CStr(values("ACT_TIME")), "0", "explicit zero minutes allowed", checked
    WebRegressionEqual CStr(values.Count), "13", "only expected form fields", checked
    WebRegressionContains CStr(item("Body")), "└ 첫 수정" & vbCrLf & vbCrLf & "  └ 둘째 수정", "multiline modification alignment and blank line", checked
    WebRegressionAssert InStr(CStr(items(2)("Body")), "└") = 0, "blank modification has no marker", checked
    WebRegressionAssert InStr(CStr(items(3)("Body")), "그룹 B") = 0, "subtree stops at next group", checked
    WebRegressionContains CStr(items(4)("Body")), "    - B 하위", "level 1 includes descendant with hierarchy indentation", checked
    WebRegressionEqual WebRegressionSnapshot(ws), snapshot, "preview leaves all task values unchanged", checked
    WebRegressionAssert Not Application.EnableEvents, "preview preserves event setting", checked

    stage = "levels 1 and 2 include their own requests and descendants"
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 1
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items.Count), "3", "multiple selected descendants do not duplicate a level 1 registration unit", checked
    WebRegressionEqual CStr(items(1)("Row")), "5", "level 1 uses root", checked
    body = CStr(items(1)("Body"))
    WebRegressionContains body, "그룹 A", "root request is in body", checked
    WebRegressionContains body, "    - 레벨2 분기", "L2 indentation", checked
    WebRegressionContains body, "        " & ChrW(&HB7) & " 레벨3 첫 요청", "L3 indentation", checked
    WebRegressionAssert InStr(body, "그룹 B") = 0, "level 1 group boundary", checked
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 2
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items(1)("Row")), "6", "level 2 excludes root", checked
    WebRegressionEqual CStr(items(2)("Row")), "10", "level 2 includes leaf L2", checked
    WebRegressionContains CStr(items(1)("Body")), "    " & ChrW(&HB7) & " 레벨3 첫 요청", "level 2 body includes L3", checked
    WebRegressionAssert InStr(CStr(items(1)("Body")), "그룹 A") = 0, "level 2 body excludes ancestor", checked

    stage = "selection is explicit and group criterion comes from root"
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(7, COL_WEB_ENABLED).ClearContents
    ws.Cells(8, COL_WEB_ENABLED).Value2 = "N"
    ws.Cells(10, COL_WEB_LEVEL).Value2 = 1
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items(1)("Row")), "10", "blank and N disabled; child level override ignored", checked
    ws.Cells(10, COL_WEB_ENABLED).Value2 = "N"
    ws.Cells(5, COL_WEB_ENABLED).Value2 = "N"
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items(1)("Row")), "6", "group diagnostic belongs to selected task instead of disabled root", checked
    WebRegressionContains CStr(items(1)("Error")), "해당하는 행의 WEB UPDATE", "selected non-target gives actionable diagnostic", checked
    WebRegressionOnlySelected ws, items, checked
    ws.Cells(5, COL_WEB_ENABLED).ClearContents
    Set items = BuildWebUploadItems(ws)
    WebRegressionEqual CStr(items(1)("Row")), "6", "blank root is not included as a group diagnostic", checked
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(8, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(10, COL_WEB_ENABLED).Value2 = "Y"

    stage = "dates, minutes and option validation never use website defaults"
    ws.Cells(7, COL_ACTUAL_START).ClearContents
    ws.Cells(7, COL_ACTUAL_END).ClearContents
    ws.Cells(7, COL_WORK_MINUTES).ClearContents
    ws.Cells(7, COL_WEB_REQUESTER).Value2 = "동명이인"
    Set items = BuildWebUploadItems(ws)
    Set item = items(1)
    Set values = item("Values")
    WebRegressionContains CStr(item("Error")), "실제 시작일", "missing start rejected", checked
    WebRegressionContains CStr(item("Error")), "실제 종료일", "missing finish rejected", checked
    WebRegressionContains CStr(item("Error")), "작업시간", "missing minutes rejected", checked
    WebRegressionContains CStr(item("Error")), "요청자", "ambiguous name rejected", checked
    WebRegressionEqual CStr(values("ACT_DT_FR")), "", "no planned/today start fallback", checked
    WebRegressionEqual CStr(values("ACT_DT")), "", "no planned/today end fallback", checked
    WebRegressionEqual CStr(values("ACT_TIME")), "", "no 30-minute default", checked
    ws.Cells(7, COL_ACTUAL_START).Value = DateSerial(2026, 10, 3)
    ws.Cells(7, COL_ACTUAL_END).Value = DateSerial(2026, 10, 2)
    ws.Cells(7, COL_WORK_MINUTES).Value2 = 1.5
    ws.Cells(7, COL_WEB_REQUESTER).Value2 = "동명이인 [user2]"
    Set items = BuildWebUploadItems(ws)
    Set item = items(1)
    Set values = item("Values")
    WebRegressionContains CStr(item("Error")), "시작일보다 빠릅니다", "reversed actual dates rejected", checked
    WebRegressionContains CStr(item("Error")), "정수", "fractional minutes rejected", checked
    WebRegressionEqual CStr(values("REQ_UID")), "user2", "same-name requester distinguished by ID", checked
    ws.Cells(7, COL_ACTUAL_START).Value = DateSerial(2026, 10, 1)
    ws.Cells(7, COL_WORK_MINUTES).Value2 = -1
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "작업시간", "negative minutes rejected", checked
    ws.Cells(7, COL_WORK_MINUTES).Value2 = 30

    stage = "duplicate and missing stable identity"
    sourceId = CStr(ws.Cells(8, COL_WEB_TASK_ID).Value2)
    ws.Cells(8, COL_WEB_TASK_ID).Value2 = ws.Cells(7, COL_WEB_TASK_ID).Value2
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "중복", "first duplicate identity also blocked", checked
    WebRegressionContains CStr(items(2)("Error")), "중복", "second duplicate identity blocked", checked
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "N"
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "중복", "selected duplicate is blocked when earlier duplicate is not selected", checked
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(8, COL_WEB_TASK_ID).Value2 = sourceId
    Set copiedWs = ThisWorkbook.Worksheets.Add
    copiedWs.Name = "CopiedWebTaskFixture"
    copiedWs.Cells(HEADER_ROW, COL_REQUEST).Value2 = "요청 내용"
    copiedWs.Cells(5, COL_LEVEL).Value2 = 1
    copiedWs.Cells(5, COL_REQUEST).Value2 = "다른 시트로 복사한 업무"
    copiedWs.Cells(5, COL_WEB_ENABLED).Value2 = "N"
    copiedWs.Cells(5, COL_WEB_TASK_ID).Value2 = ws.Cells(7, COL_WEB_TASK_ID).Value2
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "중복", "copied ID in another task sheet also blocks selected source", checked
    copiedWs.Cells(5, COL_WEB_TASK_ID).ClearContents
    ws.Cells(7, COL_WEB_TASK_ID).ClearContents
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "웹업무ID가 없습니다", "missing identity blocked without mutation", checked
    WebRegressionEqual CStr(ws.Cells(7, COL_WEB_TASK_ID).Value2), "", "model does not assign missing identities", checked

    stage = "invalid hierarchy and register criterion are blocked"
    ws.Cells(7, COL_WEB_TASK_ID).Value2 = "fixture-7"
    ws.Cells(6, COL_REQUEST).ClearContents
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "바로 상위 레벨", "L3 without an L2 parent is blocked", checked
    ws.Cells(6, COL_REQUEST).Value2 = "레벨2 분기"
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 2.5
    Set items = BuildWebUploadItems(ws)
    WebRegressionContains CStr(items(1)("Error")), "1, 2, 3", "fractional group criterion rejected", checked
    WebRegressionEqual CStr(items(1)("Row")), "6", "invalid criterion diagnostic does not include blank group root", checked
    WebRegressionOnlySelected ws, items, checked

    stage = "disabled orphan and invalid groups stay outside the selected queue"
    Set selectionWs = ThisWorkbook.Worksheets.Add
    selectionWs.Name = "WebSelectionFixture"
    For r = 5 To 10
        WebRegressionSeedRow selectionWs, r
        selectionWs.Cells(r, COL_REQUEST).Value2 = "선택 필터 검증 " & CStr(r)
        selectionWs.Cells(r, COL_WEB_TASK_ID).Value2 = "selection-" & CStr(r)
        selectionWs.Cells(r, COL_WEB_ENABLED).Value2 = "N"
    Next r
    selectionWs.Cells(5, COL_LEVEL).Value2 = 2
    selectionWs.Cells(6, COL_LEVEL).Value2 = 3
    selectionWs.Cells(6, COL_WEB_ENABLED).ClearContents
    selectionWs.Cells(7, COL_WEB_LEVEL).Value2 = 2.5
    selectionWs.Cells(8, COL_LEVEL).Value2 = 2
    selectionWs.Cells(9, COL_WEB_LEVEL).Value2 = 3
    selectionWs.Cells(9, COL_WEB_ENABLED).ClearContents
    selectionWs.Cells(10, COL_LEVEL).Value2 = 2
    selectionWs.Cells(10, COL_WEB_ENABLED).ClearContents
    Set items = BuildWebUploadItems(selectionWs)
    WebRegressionEqual CStr(items.Count), "0", "N and blank orphan rows and invalid-basis groups are never queued", checked
    For r = 5 To 10
        If r <> 7 And r <> 9 Then selectionWs.Cells(r, COL_WEB_ENABLED).Value2 = "Y"
    Next r
    Set items = BuildWebUploadItems(selectionWs)
    WebRegressionEqual CStr(items.Count), "4", "selected orphan rows and invalid criterion remain visible as errors", checked
    WebRegressionEqual CStr(items(1)("Row")), "5", "first selected orphan is shown", checked
    WebRegressionContains CStr(items(1)("Error")), "레벨1 그룹이 없습니다", "selected orphan is blocked", checked
    WebRegressionEqual CStr(items(2)("Row")), "6", "second selected orphan is shown", checked
    WebRegressionEqual CStr(items(3)("Row")), "8", "invalid group uses selected child rather than N root", checked
    WebRegressionContains CStr(items(3)("Error")), "1, 2, 3", "selected invalid group is blocked", checked
    WebRegressionEqual CStr(items(4)("Row")), "10", "valid Y leaf uses criterion from blank root", checked
    WebRegressionEqual CStr(items(4)("Error")), "", "disabled group root does not disable selected registration unit", checked
    WebRegressionOnlySelected selectionWs, items, checked
    selectionWs.Cells(10, COL_WEB_ENABLED).Value2 = CVErr(xlErrNA)
    Set items = BuildWebUploadItems(selectionWs)
    WebRegressionEqual CStr(items.Count), "3", "cell error in WEB UPDATE is not Y", checked
    WebRegressionOnlySelected selectionWs, items, checked

    stage = "blank rows with leftover Y do not create queue entries"
    selectionWs.Cells(5, COL_WEB_ENABLED).Value2 = "N"
    selectionWs.Cells(6, COL_WEB_ENABLED).ClearContents
    selectionWs.Cells(8, COL_REQUEST).ClearContents
    Set items = BuildWebUploadItems(selectionWs)
    WebRegressionEqual CStr(items.Count), "0", "leftover Y on a blank task does not create group error", checked

    stage = "blank modification and normalized request formatting"
    WebRegressionEqual CleanWebRequestText("  요청" & vbCrLf & vbTab & "내용  "), "요청 내용", "request whitespace matches TXT rule", checked
    WebRegressionEqual CleanWebModificationText(vbLf & "  " & vbTab & vbLf), "", "whitespace modification is empty", checked
    WebRegressionEqual WebModificationLines("수정" & vbLf & vbLf & "다음", 4), _
        "    └ 수정" & vbCrLf & vbCrLf & "    └ 다음", "blank modification lines survive", checked
    RunWebPayloadRegression = "PASS: " & CStr(checked) & " web payload assertions"
    Exit Function
Failed:
    RunWebPayloadRegression = "FAIL during " & stage & ": " & Err.Description
End Function

Private Sub WebRegressionOnlySelected(ByVal ws As Worksheet, ByVal items As Collection, ByRef checked As Long)
    Dim item As Object
    For Each item In items
        WebRegressionEqual UCase$(Trim$(WebCellText(ws, CLng(item("Row")), COL_WEB_ENABLED))), "Y", _
            "every preview row is explicitly Y at task row " & CStr(item("Row")), checked
    Next item
End Sub

Private Sub WebRegressionSeedOption(ByVal ws As Worksheet, ByVal baseColumn As Long, _
                                    ByVal display As String, ByVal name As String, ByVal code As String)
    ws.Range(ws.Cells(9, baseColumn), ws.Cells(ws.Rows.Count, baseColumn + 2)).ClearContents
    ws.Cells(9, baseColumn).Value2 = display
    ws.Cells(9, baseColumn + 1).Value2 = name
    ws.Cells(9, baseColumn + 2).Value2 = code
End Sub

Private Sub WebRegressionSeedRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    ws.Cells(rowNum, COL_LEVEL).Value2 = 1
    ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2 = "소분류"
    ws.Cells(rowNum, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(rowNum, COL_WEB_LEVEL).Value2 = 1
    ws.Cells(rowNum, COL_WEB_REQUESTER).Value2 = "동명이인 [user1]"
    ws.Cells(rowNum, COL_WEB_ACTION).Value2 = "개발"
    ws.Cells(rowNum, COL_WEB_STATUS).Value2 = "처리중"
    ws.Cells(rowNum, COL_WEB_MODULE).Value2 = "생산"
    ws.Cells(rowNum, COL_WEB_HANDLER).Value2 = "처리자 [worker]"
    ws.Cells(rowNum, COL_WEB_PROGRAM_ID).Value2 = "PGM" & Format$(rowNum, "000")
    ws.Cells(rowNum, COL_WEB_TASK_ID).Value2 = "fixture-" & CStr(rowNum)
    ws.Cells(rowNum, COL_PLAN_START).Value = DateSerial(2026, 9, 1)
    ws.Cells(rowNum, COL_PLAN_END).Value = DateSerial(2026, 9, 2)
    ws.Cells(rowNum, COL_ACTUAL_START).Value = DateSerial(2026, 10, 1)
    ws.Cells(rowNum, COL_ACTUAL_END).Value = DateSerial(2026, 10, 2)
    ws.Range(COL_PLAN_START & rowNum & ":" & COL_ACTUAL_END & rowNum).NumberFormat = "yyyy-mm-dd"
    ws.Cells(rowNum, COL_WORK_MINUTES).Value2 = 45
    ws.Cells(rowNum, COL_WEEKLY_REPORT).Value2 = "untouched status"
End Sub

Private Function WebRegressionSnapshot(ByVal ws As Worksheet) As String
    Dim values As Variant, rowNum As Long, colNum As Long
    values = ws.Range("A1:" & COL_STATUS & "20").Value2
    For rowNum = 1 To UBound(values, 1)
        For colNum = 1 To UBound(values, 2)
            WebRegressionSnapshot = WebRegressionSnapshot & "|" & CStr(values(rowNum, colNum))
        Next colNum
        WebRegressionSnapshot = WebRegressionSnapshot & vbCrLf
    Next rowNum
End Function

Private Sub WebRegressionAssert(ByVal condition As Boolean, ByVal message As String, ByRef checked As Long)
    If Not condition Then Err.Raise vbObjectError + 7891, "WebPayloadRegression", message
    checked = checked + 1
End Sub

Private Sub WebRegressionEqual(ByVal actual As String, ByVal expected As String, _
                                ByVal message As String, ByRef checked As Long)
    WebRegressionAssert (actual = expected), message & ": expected [" & expected & "] actual [" & actual & "]", checked
End Sub

Private Sub WebRegressionContains(ByVal actual As String, ByVal expected As String, _
                                   ByVal message As String, ByRef checked As Long)
    WebRegressionAssert InStr(actual, expected) > 0, message & ": missing [" & expected & "] in [" & actual & "]", checked
End Sub
