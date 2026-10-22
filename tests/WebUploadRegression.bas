' Append to modWebUpload only in a disposable workbook copy. No network calls.
' 2026-10-22 김좌현 수정 : 웹 등록 설정·식별값·선택목록 갱신·미리보기·이력 통합 검증
Public Function RunWebUploadRegression() As String
    Dim ws As Worksheet, config As Worksheet, sortWs As Worksheet, preview As Worksheet
    Dim log As Worksheet, form As Object, item As Object, items As Collection, choices As Collection
    Dim checked As Long, stage As String, firstId As String, secondId As String, newId As String
    Dim snapshot As String, idSnapshot As String, no As String, outcome As String, secret As String
    Dim encrypted As String, r As Long, field As Variant, originalLabel As String, caught As String
    On Error GoTo Failed
    Application.EnableEvents = False
    stage = "configure local option fixture"
    EnsureWebConfigSheet
    Set config = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    Set form = WebUploadRegressionForm(False, False)
    StoreWebOptionLists form
    config.Range("B6").Value2 = 3
    config.Range("B7").Value2 = "N"
    config.Range("B8").Value2 = "요청자 [requester1]"
    config.Range("B9").Value2 = "개발"
    config.Range("B10").Value2 = "처리중"
    config.Range("B11").Value2 = "생산"
    config.Range("B12").Value2 = "처리자 [worker1]"
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "WebUploadFixture"
    SetupDataHeaders ws
    For r = 5 To 9
        WebUploadRegressionSeedRow ws, r
    Next r
    ws.Cells(5, COL_REQUEST).Value2 = "프로젝트 그룹"
    ws.Cells(5, COL_LEVEL).Value2 = 1
    ws.Cells(5, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(6, COL_REQUEST).Value2 = "세부 그룹"
    ws.Cells(6, COL_LEVEL).Value2 = 2
    ws.Cells(7, COL_REQUEST).Value2 = "등록할 레벨3"
    ws.Cells(7, COL_LEVEL).Value2 = 3
    ws.Cells(7, COL_WEB_REQUESTER).Value2 = "요청자 둘 [requester2]"
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(8, COL_REQUEST).Value2 = "레벨3 없는 레벨2"
    ws.Cells(8, COL_LEVEL).Value2 = 2
    ws.Cells(8, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(9, COL_REQUEST).Value2 = "레벨1 단독 오류"
    ws.Cells(9, COL_LEVEL).Value2 = 1
    ws.Cells(9, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(9, COL_WEB_ENABLED).Value2 = "Y"

    stage = "initialize twice preserves business data, choices and stable IDs"
    snapshot = WebUploadRegressionBusinessSnapshot(ws)
    InitializeWebUpload
    WebUploadRegressionEqual WebUploadRegressionBusinessSnapshot(ws), snapshot, "initialization preserves request/modification/dates/minutes/progress", checked
    firstId = GetWebTaskId(ws, 7)
    secondId = GetWebTaskId(ws, 8)
    WebUploadRegressionAssert Len(firstId) = 36 And Len(secondId) = 36, "task IDs are GUIDs", checked
    WebUploadRegressionAssert firstId <> secondId, "new rows receive different IDs", checked
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_ENABLED).Value2), "N", "blank registration selection gets default N", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_ENABLED).Value2), "Y", "explicit registration Y survives", checked
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_REQUESTER).Value2), "요청자 [requester1]", "blank requester gets default", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_REQUESTER).Value2), "요청자 둘 [requester2]", "explicit requester survives", checked
    idSnapshot = WebUploadRegressionIdSnapshot(ws)
    InitializeWebUpload
    WebUploadRegressionEqual WebUploadRegressionBusinessSnapshot(ws), snapshot, "second initialization preserves business data", checked
    WebUploadRegressionEqual WebUploadRegressionIdSnapshot(ws), idSnapshot, "second initialization preserves every task ID", checked
    WebUploadRegressionAssert ws.Columns(COL_WEB_TASK_ID).Hidden, "persistent task ID column stays hidden", checked
    WebUploadRegressionAssert Not Application.EnableEvents, "initialization preserves disabled events", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_REQUESTER).Validation.Type), CStr(xlValidateList), "main requester is a dropdown", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WORK_MINUTES).Validation.Type), CStr(xlValidateWholeNumber), "minutes validation is whole numbers", checked
    WebUploadRegressionAssert ws.Shapes("btnWebPreview").OnAction Like "*PreviewWebTasks", "main preview button is connected", checked

    stage = "child edit propagates criterion across only its own level 1 group"
    ws.Cells(7, COL_WEB_LEVEL).Value2 = 2
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_LEVEL)
    For r = 5 To 8
        WebUploadRegressionEqual CStr(ws.Cells(r, COL_WEB_LEVEL).Value2), "2", "group criterion propagated row " & CStr(r), checked
    Next r
    WebUploadRegressionEqual CStr(ws.Cells(9, COL_WEB_LEVEL).Value2), "3", "next group criterion is isolated", checked
    Application.EnableEvents = True
    ws.Cells(8, COL_WEB_LEVEL).Value2 = 1
    WebUploadRegressionAssert Application.EnableEvents, "criterion event restores enabled events", checked
    WebUploadRegressionEqual CStr(ws.Cells(5, COL_WEB_LEVEL).Value2), "1", "real workbook change event updates root criterion", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_LEVEL).Value2), "1", "real workbook change event updates sibling criterion", checked
    Application.EnableEvents = False
    ws.Cells(6, COL_WEB_LEVEL).Value2 = 3
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_LEVEL)

    stage = "credential persistence is masked and Windows protected"
    secret = "DisposableFixture-" & Replace(firstId, "-", "")
    SetWebCredentials "fixture-login", secret
    encrypted = CStr(config.Range("Z2").Value2)
    WebUploadRegressionEqual CStr(config.Range("B3").Value2), "fixture-login", "credential username stored", checked
    WebUploadRegressionEqual CStr(config.Range("B4").Value2), "********", "password input is masked", checked
    WebUploadRegressionAssert Len(encrypted) > 0 And encrypted <> secret, "password is not stored as plaintext", checked
    WebUploadRegressionEqual WebUnprotectSecret(encrypted), secret, "DPAPI secret roundtrip", checked
    WebUploadRegressionAssert Not WebUploadRegressionContainsCellSecret(secret), "plaintext password is absent from every worksheet cell", checked
    WebUploadRegressionAssert config.Columns("Z").Hidden, "encrypted secret storage is hidden", checked
    WebUploadRegressionAssert Not Application.EnableEvents, "secret setter restores disabled events", checked
    Application.EnableEvents = True
    config.Range("B4").Value2 = secret & "-event"
    WebUploadRegressionAssert Application.EnableEvents, "password change event restores enabled events", checked
    Application.EnableEvents = False
    WebUploadRegressionEqual CStr(config.Range("B4").Value2), "********", "typed password is masked by workbook change event", checked
    WebUploadRegressionEqual WebUnprotectSecret(CStr(config.Range("Z2").Value2)), secret & "-event", "typed password is protected by workbook event", checked
    WebUploadRegressionAssert Not WebUploadRegressionContainsCellSecret(secret), "typed plaintext password is absent from worksheet cells", checked
    SetWebPasswordValue ""
    WebUploadRegressionEqual CStr(config.Range("B4").Value2), "", "cleared password removes mask", checked
    WebUploadRegressionEqual CStr(config.Range("Z2").Value2), "", "cleared password removes protected storage", checked
    secret = vbNullString

    stage = "option refresh keeps selected IDs when labels change"
    Set form = WebUploadRegressionForm(True, False)
    StoreWebOptionLists form
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_REQUESTER).Value2), "요청자 변경 [requester1]", "renamed requester follows existing ID", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_REQUESTER).Value2), "요청자 둘 변경 [requester2]", "second requester follows own ID", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_ACTION).Value2), "개발 변경", "renamed action follows existing ID", checked
    WebUploadRegressionEqual ResolveWebOption("ACT_GUBUN", CStr(ws.Cells(7, COL_WEB_ACTION).Value2)), "DEV", "renamed action retains original code", checked
    WebUploadRegressionEqual CStr(config.Range("B9").Value2), "개발 변경", "default follows same ID", checked
    WebUploadRegressionAssert IsDate(config.Range("B5").Value), "list update timestamp recorded", checked
    WebUploadRegressionAssert InStr(ThisWorkbook.Names("WebRequesters").RefersTo, "$E$10") > 0, "requester dropdown includes both IDs", checked

    stage = "deleted ID stays invalid even if a new code reuses the same label"
    Set form = WebUploadRegressionForm(True, True)
    StoreWebOptionLists form
    WebUploadRegressionEqual ResolveWebOption("ACT_GUBUN", CStr(ws.Cells(7, COL_WEB_ACTION).Value2)), "", "removed selected ID cannot map to replacement by label", checked
    WebUploadRegressionEqual ResolveWebOption("ACT_GUBUN", CStr(config.Range("B9").Value2)), "", "removed default ID cannot map to replacement", checked
    Set items = BuildWebUploadItems(ws)
    WebUploadRegressionContains CStr(items(1)("Error")), "처리구분", "missing code is visible in preview validation", checked
    For r = 5 To 9
        ws.Cells(r, COL_WEB_ACTION).Value2 = "개발 변경"
    Next r
    config.Range("B9").Value2 = "개발 변경"

    stage = "persistent IDs follow rows through sorting and insertion"
    Set sortWs = ThisWorkbook.Worksheets.Add
    sortWs.Name = "WebSortFixture"
    SetupDataHeaders sortWs
    For r = 5 To 7
        WebUploadRegressionSeedRow sortWs, r
        sortWs.Cells(r, COL_LEVEL).Value2 = 1
        sortWs.Cells(r, COL_REQUEST).Value2 = "정렬 " & CStr(r)
    Next r
    PrepareWebTaskSheet sortWs
    firstId = GetWebTaskId(sortWs, 5)
    secondId = GetWebTaskId(sortWs, 7)
    sortWs.Range(COL_NO & "5:" & COL_STATUS & "7").Sort _
        Key1:=sortWs.Range(COL_REQUEST & "5"), Order1:=xlDescending, Header:=xlNo
    WebUploadRegressionEqual GetWebTaskId(sortWs, 7), firstId, "sort carries first ID with its task", checked
    WebUploadRegressionEqual GetWebTaskId(sortWs, 5), secondId, "sort carries last ID with its task", checked
    WebUploadRegressionEqual CStr(sortWs.Cells(7, COL_REQUEST).Value2), "정렬 5", "sort task and ID stay paired", checked
    sortWs.Rows(6).Insert Shift:=xlShiftDown
    WebUploadRegressionEqual GetWebTaskId(sortWs, 8), firstId, "insertion shifts existing ID without reassignment", checked
    WebUploadRegressionEqual CStr(FindWebTaskRow(sortWs, firstId)), "8", "history locates inserted/sorted task by persistent ID", checked
    WebUploadRegressionSeedRow sortWs, 6
    sortWs.Cells(6, COL_REQUEST).Value2 = "새로 삽입한 업무"
    sortWs.Cells(6, COL_LEVEL).Value2 = 1
    PrepareWebTaskSheet sortWs
    newId = GetWebTaskId(sortWs, 6)
    WebUploadRegressionAssert Len(newId) = 36 And newId <> firstId And newId <> secondId, "inserted new task receives new identity", checked
    WebUploadRegressionEqual GetWebTaskId(sortWs, 8), firstId, "prepare preserves shifted original identity", checked

    stage = "actual preview UI shows eligible rows and root-only group error"
    ws.Activate
    PreviewWebTasks
    Set preview = ThisWorkbook.Worksheets(WEB_PREVIEW_SHEET_NAME)
    WebUploadRegressionEqual ActiveSheet.Name, WEB_PREVIEW_SHEET_NAME, "preview sheet activated", checked
    WebUploadRegressionEqual CStr(mPreviewItems.Count), "3", "two eligible tasks and one group diagnostic", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 2).Value2), "7", "preview first source row", checked
    WebUploadRegressionEqual CStr(preview.Cells(6, 2).Value2), "8", "preview childless L2 source row", checked
    WebUploadRegressionEqual CStr(preview.Cells(7, 2).Value2), "9", "preview group error source row", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "등록 가능", "valid task preview status", checked
    WebUploadRegressionEqual CStr(preview.Cells(7, 1).Value2), "확인 필요", "root-only group error preview status", checked
    WebUploadRegressionContains CStr(preview.Cells(7, 13).Value2), "해당하는 업무가 없습니다", "root-only reason is visible", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 3).Value2), "[소분류] 등록할 레벨3", "actual preview title", checked
    WebUploadRegressionContains CStr(preview.Cells(5, 4).Value2), "└ 수정 7", "actual preview body contains modification", checked
    WebUploadRegressionAssert preview.Shapes("btnWebUploadPreview").OnAction Like "*UploadWebTasks", "preview upload button is connected", checked
    snapshot = mPreviewSnapshot
    ws.Cells(7, COL_MODIFICATION).Value2 = "미리보기 후 변경"
    Set items = BuildWebUploadItems(ws)
    WebUploadRegressionAssert PreviewItemsSnapshot(items) <> snapshot, "edited content invalidates prior preview snapshot", checked
    ws.Cells(7, COL_MODIFICATION).Value2 = "수정 7"

    stage = "journal updates source by stable ID and previews block completed/unknown rows"
    Set item = mPreviewItems(1)
    RecordWebOutcome item, "등록완료", "PJT-000123", "fixture only"
    no = vbNullString
    outcome = PreviousWebOutcome(CStr(item("SourceId")), no)
    WebUploadRegressionEqual outcome, "등록완료", "latest success outcome found", checked
    WebUploadRegressionEqual no, "PJT-000123", "registered number found", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_RESULT).Value2), "등록완료 PJT-000123", "source result updated by ID", checked
    Set item = mPreviewItems(2)
    RecordWebOutcome item, "결과 미확인", "PJT-000124", "fixture only"
    PreviewWebTasks
    WebUploadRegressionContains CStr(preview.Cells(5, 13).Value2), "이미 등록", "completed task blocked in preview", checked
    WebUploadRegressionContains CStr(preview.Cells(6, 13).Value2), "웹에서 먼저 확인", "unknown result blocked in preview", checked
    Set item = mPreviewItems(2)
    RecordWebOutcome item, "미등록 확인", "", "checked in fixture"
    PreviewWebTasks
    WebUploadRegressionEqual CStr(preview.Cells(6, 1).Value2), "등록 가능", "explicit unregistered confirmation permits new preview", checked
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    WebUploadRegressionAssert log.Columns("C").Hidden, "journal identity column hidden", checked
    WebUploadRegressionAssert log.Hyperlinks.Count >= 2, "successful/unknown request number links recorded", checked
    WebUploadRegressionAssert Not mWebBusy, "offline tests never enter upload busy state", checked

    stage = "upload orchestration journals before POST, saves results and skips invalid rows"
    WebUploadRegressionResetJournal
    config.Range("B2").Value2 = "http://fixture.invalid/support/"
    SetWebCredentials "mock-login", "Disposable-Mock-Password"
    ws.Activate
    PreviewWebTasks False
    WebUploadMockReset "success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockOpenCount), "1", "one mock session opened for batch", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "2", "fresh form loaded for each valid row", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "2", "two valid rows posted; group error skipped", checked
    WebUploadRegressionEqual CStr(WebUploadMockJournalChecks), "2", "each POST sees persisted matching pre-journal", checked
    WebUploadRegressionEqual CStr(log.Cells(log.Rows.Count, "A").End(xlUp).Row), "5", "two pre-POST and two result journal records", checked
    WebUploadRegressionEqual CStr(log.Range("E2").Value2), "등록 시작", "first journal is start marker", checked
    WebUploadRegressionEqual CStr(log.Range("E3").Value2), "등록완료", "first result journal is success", checked
    WebUploadRegressionEqual CStr(log.Range("E4").Value2), "등록 시작", "second start marker recorded", checked
    WebUploadRegressionEqual CStr(log.Range("E5").Value2), "등록완료", "second result journal is success", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_RESULT).Value2), "등록완료 RQMOCK000001", "first real source result updated", checked
    WebUploadRegressionEqual CStr(ws.Cells(8, COL_WEB_RESULT).Value2), "등록완료 RQMOCK000002", "second real source result updated", checked
    WebUploadRegressionAssert Not mWebBusy And Not Application.EnableEvents, "upload clears busy flag and restores events", checked
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "2", "repeat upload never posts completed rows again", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "2", "repeat upload does not load completed forms", checked

    stage = "stale content and site configuration are rejected before transport"
    WebUploadRegressionResetJournal
    PreviewWebTasks False
    ws.Cells(7, COL_MODIFICATION).Value2 = "stale preview content"
    WebUploadMockReset "success"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description
    Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "미리보기", "stale content raises actionable error", checked
    WebUploadRegressionEqual CStr(WebUploadMockOpenCount), "0", "stale content rejected before login", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "stale content never posts", checked
    ws.Cells(7, COL_MODIFICATION).Value2 = "수정 7"
    PreviewWebTasks False
    config.Range("B2").Value2 = "http://changed.invalid/support/"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description
    Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "미리보기", "changed site invalidates preview", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "changed site never posts stale preview", checked
    config.Range("B2").Value2 = "http://fixture.invalid/support/"

    stage = "unknown result stops the batch and is never retried without confirmation"
    WebUploadRegressionResetJournal
    PreviewWebTasks False
    WebUploadMockReset "unknown"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "1", "unknown result stops before next valid row", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 7), no), "결과 미확인", "unknown result persisted", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 8), no), "", "later valid row not attempted", checked
    ws.Cells(8, COL_WEB_ENABLED).Value2 = "N"
    PreviewWebTasks False
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "1", "unknown task not posted again on repeat", checked
    ws.Cells(8, COL_WEB_ENABLED).Value2 = "Y"

    stage = "known failure allows the next row and only the failed row can be retried"
    WebUploadRegressionResetJournal
    PreviewWebTasks False
    WebUploadMockReset "failed,success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "2", "known failure does not stop later valid row", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 7), no), "등록실패", "known failure journaled", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 8), no), "등록완료", "later row success journaled", checked
    WebUploadMockReset "success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "1", "retry posts only previously failed row", checked
    WebUploadRegressionContains WebUploadMockPostedTitles, "등록할 레벨3", "failed target is retried", checked
    WebUploadRegressionAssert InStr(WebUploadMockPostedTitles, "레벨3 없는 레벨2") = 0, "completed later target skipped on retry", checked

    stage = "exception after POST starts becomes a persisted unknown outcome"
    WebUploadRegressionResetJournal
    PreviewWebTasks False
    WebUploadMockReset "throw"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description
    Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "Mock POST", "transport exception rethrown in noninteractive mode", checked
    WebUploadRegressionEqual CStr(WebUploadMockJournalChecks), "1", "exception occurs only after persisted start marker", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 7), no), "결과 미확인", "exception creates unknown journal marker", checked
    WebUploadRegressionAssert Not mWebBusy And Not Application.EnableEvents, "exception resets busy flag and event setting", checked
    no = vbNullString
    WebUploadRegressionEqual PreviousWebOutcome(GetWebTaskId(ws, 8), no), "", "exception leaves subsequent row unposted", checked
    RunWebUploadRegression = "PASS: " & CStr(checked) & " web upload integration assertions"
    Exit Function
Failed:
    RunWebUploadRegression = "FAIL during " & stage & ": " & Err.Description
End Function

Private Sub WebUploadRegressionResetJournal()
    Dim log As Worksheet, lastRow As Long
    EnsureWebLogSheet
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    log.Hyperlinks.Delete
    lastRow = Application.Max(2, log.Cells(log.Rows.Count, "A").End(xlUp).Row)
    log.Range("A2:J" & lastRow).ClearContents
End Sub

Private Function WebUploadRegressionForm(ByVal renamed As Boolean, ByVal replaced As Boolean) As Object
    Dim form As Object, options As Object, fields As Object, values As Object, field As Variant
    Set form = CreateObject("Scripting.Dictionary")
    Set options = CreateObject("Scripting.Dictionary")
    Set fields = CreateObject("Scripting.Dictionary")
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        Set values = CreateObject("Scripting.Dictionary")
        Select Case CStr(field)
            Case "REQ_UID"
                values.Add "requester1", IIf(renamed, "요청자 변경", "요청자")
                values.Add "requester2", IIf(renamed, "요청자 둘 변경", "요청자 둘")
                fields.Add CStr(field), "requester1"
            Case "ACT_GUBUN"
                values.Add IIf(replaced, "DEV-REPLACED", "DEV"), IIf(renamed, "개발 변경", "개발")
                fields.Add CStr(field), IIf(replaced, "DEV-REPLACED", "DEV")
            Case "STS"
                values.Add "IP", "처리중"
                fields.Add CStr(field), "IP"
            Case "PGM_MD"
                values.Add "MES", "생산"
                fields.Add CStr(field), "MES"
            Case "ACT_UID"
                values.Add "worker1", "처리자"
                fields.Add CStr(field), "worker1"
        End Select
        options.Add CStr(field), values
    Next field
    form.Add "Options", options
    form.Add "Fields", fields
    Set WebUploadRegressionForm = form
End Function

Private Sub WebUploadRegressionSeedRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    ws.Cells(rowNum, COL_MINOR_CATEGORY).Value2 = "소분류"
    ws.Cells(rowNum, COL_MODIFICATION).Value2 = "수정 " & CStr(rowNum)
    ws.Cells(rowNum, COL_OWNER).Value2 = "담당자"
    ws.Cells(rowNum, COL_NOTE).Value2 = "비고"
    ws.Cells(rowNum, COL_PLAN_START).Value = DateSerial(2026, 9, 1)
    ws.Cells(rowNum, COL_PLAN_END).Value = DateSerial(2026, 9, 2)
    ws.Cells(rowNum, COL_ACTUAL_START).Value = DateSerial(2026, 10, 1)
    ws.Cells(rowNum, COL_ACTUAL_END).Value = DateSerial(2026, 10, 2)
    ws.Range(COL_PLAN_START & rowNum & ":" & COL_ACTUAL_END & rowNum).NumberFormat = "yyyy-mm-dd"
    ws.Cells(rowNum, COL_WORK_MINUTES).Value2 = 30
    ws.Cells(rowNum, COL_PROGRESS).Value2 = 0.5
    ws.Cells(rowNum, COL_WEEKLY_REPORT).Value2 = "기존 보고 상태"
End Sub

Private Function WebUploadRegressionBusinessSnapshot(ByVal ws As Worksheet) As String
    Dim r As Long, col As Variant
    For r = 5 To 9
        For Each col In Array(COL_REQUEST, COL_MODIFICATION, COL_OWNER, COL_NOTE, COL_PLAN_START, _
            COL_PLAN_END, COL_ACTUAL_START, COL_ACTUAL_END, COL_WORK_MINUTES, COL_PROGRESS, COL_WEEKLY_REPORT)
            WebUploadRegressionBusinessSnapshot = WebUploadRegressionBusinessSnapshot & "|" & CStr(ws.Cells(r, CStr(col)).Value2)
        Next col
    Next r
End Function

Private Function WebUploadRegressionIdSnapshot(ByVal ws As Worksheet) As String
    Dim r As Long
    For r = 5 To 9
        WebUploadRegressionIdSnapshot = WebUploadRegressionIdSnapshot & "|" & GetWebTaskId(ws, r)
    Next r
End Function

Private Function WebUploadRegressionContainsCellSecret(ByVal secret As String) As Boolean
    Dim ws As Worksheet, found As Range
    For Each ws In ThisWorkbook.Worksheets
        Set found = ws.UsedRange.Find(What:=secret, LookIn:=xlValues, LookAt:=xlPart, SearchFormat:=False)
        If Not found Is Nothing Then
            WebUploadRegressionContainsCellSecret = True
            Exit Function
        End If
    Next ws
End Function

Private Sub WebUploadRegressionAssert(ByVal condition As Boolean, ByVal message As String, ByRef checked As Long)
    If Not condition Then Err.Raise vbObjectError + 7901, "WebUploadRegression", message
    checked = checked + 1
End Sub

Private Sub WebUploadRegressionEqual(ByVal actual As String, ByVal expected As String, _
                                      ByVal message As String, ByRef checked As Long)
    WebUploadRegressionAssert actual = expected, message & ": expected [" & expected & "] actual [" & actual & "]", checked
End Sub

Private Sub WebUploadRegressionContains(ByVal actual As String, ByVal expected As String, _
                                         ByVal message As String, ByRef checked As Long)
    WebUploadRegressionAssert InStr(actual, expected) > 0, message & ": missing [" & expected & "] in [" & actual & "]", checked
End Sub
