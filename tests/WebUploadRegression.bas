' Append to modWebUpload only in a disposable workbook. No real network calls.
' 2026-10-22 김좌현 수정 : 등록단위 동기화·신규/갱신·서버충돌·요청번호 연결 및 미확인 복구 검증
' 2026-10-22 김좌현 수정 : 기존 갱신 기본값과 신규 등록 명시 선택, 연결 소실 및 모드 변조 차단 검증
' 2026-10-22 김좌현 수정 : 이력 삭제 후 새로고침의 선택적 잠금 해제와 저장 후 재복구 방지 검증
Public Function RunWebUploadRegression() As String
    Dim ws As Worksheet, config As Worksheet, preview As Worksheet, log As Worksheet
    Dim form As Object, item As Object, state As Object, server As Object, session As Object, items As Collection
    Dim checked As Long, stage As String, r As Long, firstId As String, idSnapshot As String, snapshot As String
    Dim no As String, secondNo As String, caught As String, originalBody As String, baseline As String, secret As String
    Dim connection As Worksheet, connectionRow As Long, expectedRequestNoExisting As String
    On Error GoTo Failed
    Application.EnableEvents = False
    stage = "local fixture and stable identities"
    EnsureWebConfigSheet
    Set config = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    config.Range("B2").Value2 = "http://fixture.invalid/support/"
    stage = "populate option fixture"
    StoreWebOptionLists WebUploadRegressionForm(False, False)
    config.Range("B6").Value2 = 3: config.Range("B7").Value2 = "N"
    config.Range("B8").Value2 = "요청자 [requester1]": config.Range("B9").Value2 = "개발"
    config.Range("B10").Value2 = "처리중": config.Range("B11").Value2 = "생산": config.Range("B12").Value2 = "처리자 [worker1]"
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "WebUploadFixture"
    stage = "set up fixture headers"
    SetupDataHeaders ws
    For r = 5 To 9
        WebUploadRegressionSeedRow ws, r
    Next r
    ws.Cells(5, COL_REQUEST).Value2 = "프로젝트 그룹": ws.Cells(5, COL_LEVEL).Value2 = 1: ws.Cells(5, COL_WEB_LEVEL).Value2 = 3
    ws.Cells(6, COL_REQUEST).Value2 = "세부 그룹": ws.Cells(6, COL_LEVEL).Value2 = 2
    ws.Cells(7, COL_REQUEST).Value2 = "등록할 레벨3": ws.Cells(7, COL_LEVEL).Value2 = 3: ws.Cells(7, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(8, COL_REQUEST).Value2 = "레벨3 없는 레벨2": ws.Cells(8, COL_LEVEL).Value2 = 2: ws.Cells(8, COL_WEB_ENABLED).Value2 = "Y"
    ws.Cells(9, COL_REQUEST).Value2 = "레벨1 단독 오류": ws.Cells(9, COL_LEVEL).Value2 = 1: ws.Cells(9, COL_WEB_LEVEL).Value2 = 3: ws.Cells(9, COL_WEB_ENABLED).Value2 = "Y"
    snapshot = WebUploadRegressionBusinessSnapshot(ws)
    stage = "first web initialization"
    InitializeWebUpload
    idSnapshot = WebUploadRegressionIdSnapshot(ws)
    stage = "second web initialization"
    InitializeWebUpload
    WebUploadRegressionEqual WebUploadRegressionIdSnapshot(ws), idSnapshot, "initialization preserves IDs", checked
    WebUploadRegressionEqual WebUploadRegressionBusinessSnapshot(ws), snapshot, "initialization preserves business data", checked
    firstId = GetWebTaskId(ws, 7)
    WebUploadRegressionAssert Len(firstId) = 36, "persistent GUID", checked
    WebUploadRegressionAssert ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Visible = xlSheetVeryHidden, "connection state is very hidden", checked
    ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_RESULT).Hidden = True
    ws.Columns(COL_WORK_MINUTES).Hidden = True
    PrepareWebTaskSheet ws
    WebUploadRegressionAssert ws.Columns(COL_WEB_ENABLED).Hidden And ws.Columns(COL_WEB_RESULT).Hidden And ws.Columns(COL_WORK_MINUTES).Hidden, "prepare preserves hidden web mode columns", checked
    ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_RESULT).Hidden = False
    ws.Columns(COL_WORK_MINUTES).Hidden = False
    PrepareWebTaskSheet ws
    WebUploadRegressionAssert Not ws.Columns(COL_WEB_ENABLED).Hidden And Not ws.Columns(COL_WORK_MINUTES).Hidden, "prepare preserves expanded web mode columns", checked
    WebUploadRegressionResetJournal

    stage = "registration unit propagation and merge conflicts"
    ws.Cells(7, COL_WEB_LEVEL).Value2 = 2
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_LEVEL)
    For r = 5 To 8
        WebUploadRegressionEqual CStr(ws.Cells(r, COL_WEB_LEVEL).Value2), "2", "basis propagated within L1", checked
    Next r
    ws.Cells(7, COL_WEB_REQUESTER).Value2 = "요청자 둘 [requester2]"
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_REQUESTER)
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_REQUESTER).Value2), "요청자 둘 [requester2]", "child edit updates L2 root", checked
    WebUploadRegressionEqual CStr(ws.Cells(8, COL_WEB_REQUESTER).Value2), "요청자 [requester1]", "neighbor registration unit unchanged", checked
    On Error Resume Next
    Set form = WebBasisMergePlan(ws, 5, 1)
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "합쳐", "merging conflicting units rejected", checked
    ws.Cells(6, COL_WEB_REQUESTER).Value2 = "요청자 [requester1]"
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_REQUESTER)
    ws.Cells(7, COL_WEB_LEVEL).Value2 = 1
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_LEVEL)
    WebUploadRegressionEqual CStr(ws.Cells(5, COL_WEB_ENABLED).Value2), "Y", "uniform merged unit adopts old representative selection", checked
    ws.Cells(7, COL_WEB_PROGRAM_ID).Value2 = "PGM-ONE"
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_PROGRAM_ID)
    WebUploadRegressionEqual CStr(ws.Cells(5, COL_WEB_PROGRAM_ID).Value2), "PGM-ONE", "L1 unit includes root", checked
    WebUploadRegressionEqual CStr(ws.Cells(8, COL_WEB_PROGRAM_ID).Value2), "PGM-ONE", "L1 unit includes siblings", checked
    ws.Cells(6, COL_WEB_LEVEL).Value2 = 3
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_LEVEL)
    WebUploadRegressionEqual CStr(WebRegistrationRootRow(ws, 6)), "0", "L2 with children is not target in basis3", checked
    WebUploadRegressionEqual CStr(WebRegistrationRootRow(ws, 7)), "7", "L3 is its own target", checked
    WebUploadRegressionEqual CStr(WebRegistrationRootRow(ws, 8)), "8", "childless L2 fallback remains target", checked
    ApplyWebTaskProtection ws
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_ACTION).Font.Color), CStr(RGB(140, 140, 140)), "non-target web field grey", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_ACTION).Font.Color), "0", "active target web field black", checked
    ws.Cells(6, COL_WEB_LEVEL).Value2 = 2
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_LEVEL)
    ApplyWebTaskProtection ws
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_ACTION).Font.Color), "0", "activated target restores black", checked
    ws.Cells(6, COL_WEB_LEVEL).Value2 = 3
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_LEVEL)

    stage = "period freshness without report or minutes effects"
    ws.Cells(7, COL_ACTUAL_START).Value = DateSerial(2026, 10, 10)
    ws.Cells(7, COL_ACTUAL_END).Value = DateSerial(2026, 10, 12)
    RefreshWebTaskPeriods ws
    WebUploadRegressionEqual Format$(ws.Cells(6, COL_ACTUAL_START).Value, "yyyy-mm-dd"), "2026-10-10", "parent actual start refreshed", checked
    WebUploadRegressionEqual Format$(ws.Cells(6, COL_ACTUAL_END).Value, "yyyy-mm-dd"), "2026-10-12", "parent actual end refreshed", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEEKLY_REPORT).Value2), "기존 보고 상태", "web period refresh leaves report state", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WORK_MINUTES).Value2), "30", "minutes not replicated or summed", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_PROGRESS).Value2), "0.5", "web refresh leaves progress", checked

    stage = "protected credentials and fresh preview"
    secret = "Disposable-" & Replace(firstId, "-", "")
    SetWebCredentials "fixture-login", secret
    WebUploadRegressionEqual CStr(config.Range("B4").Value2), "********", "password masked", checked
    WebUploadRegressionEqual WebUnprotectSecret(CStr(config.Range("Z2").Value2)), secret, "DPAPI roundtrip", checked
    WebUploadRegressionAssert Not WebUploadRegressionContainsCellSecret(secret), "no plaintext secret cell", checked
    WebUploadMockReset "success", True
    ws.Activate
    PreviewWebTasks False
    Set preview = ThisWorkbook.Worksheets(WEB_PREVIEW_SHEET_NAME)
    WebUploadRegressionEqual CStr(mPreviewItems.Count), "3", "two targets plus group diagnostic", checked
    WebUploadRegressionEqual CStr(preview.Range("D3").Value2), "기존 갱신", "main preview defaults to existing update", checked
    WebUploadRegressionEqual mPreviewMode, "기존 갱신", "stored preview mode defaults to update", checked
    WebUploadRegressionEqual CStr(preview.Range("D3").Validation.Type), CStr(xlValidateList), "execution mode is a dropdown", checked
    WebUploadRegressionAssert Not preview.Range("D3").Locked, "execution mode remains selectable in protected preview", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "연결 필요", "missing request number is connection-required in update mode", checked
    WebUploadRegressionEqual CStr(preview.Cells(6, 1).Value2), "연결 필요", "every unlinked target is blocked in update mode", checked
    WebUploadRegressionContains CStr(preview.Cells(5, 13).Value2), "요청번호", "missing connection has actionable explanation", checked
    WebUploadRegressionEqual CStr(preview.Cells(7, 1).Value2), "확인 필요", "invalid group displayed", checked
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "default update mode cannot create unlinked requests", checked
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "0", "unlinked update mode cannot update a request", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "default update mode never loads a new registration form", checked

    stage = "execution mode and preview operation cannot silently authorize new registration"
    mPreviewItems(1)("Operation") = "신규"
    mPreviewItems(1)("Error") = vbNullString
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "tampered operation cannot insert in update mode", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "tampered operation cannot load a new form in update mode", checked
    PreviewWebTasks False
    preview.Range("D3").Value2 = "임의 모드"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionAssert Len(caught) > 0, "invalid pasted execution mode is rejected", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "invalid mode cannot insert", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "invalid mode cannot load registration form", checked
    WebUploadRegressionSelectMode "신규 등록"
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "신규", "explicit new mode enables unlinked target", checked
    WebUploadRegressionEqual mPreviewMode, "신규 등록", "selected new mode is captured", checked
    mPreviewItems(1)("Operation") = "갱신"
    mPreviewItems(1)("Error") = vbNullString
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionAssert Len(caught) > 0, "update operation tampering is rejected in new mode", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "tampered new-mode operation cannot insert", checked
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "0", "tampered new-mode operation cannot update", checked
    WebUploadRegressionSelectMode "신규 등록"
    preview.Range("D3").Value2 = "기존 갱신"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "미리보기", "mode change invalidates prepared preview", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "changed mode never posts stale preview", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "changed mode rejected before new form retrieval", checked
    WebUploadRegressionSelectMode "신규 등록"
    ws.Activate
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Range("D3").Value2), "기존 갱신", "opening from main resets prior new mode", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "연결 필요", "main reset restores missing-connection block", checked
    WebUploadRegressionSelectMode "신규 등록"
    snapshot = mPreviewSnapshot
    ws.Cells(7, COL_MODIFICATION).Value2 = "stale"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "미리보기", "stale preview blocked", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "stale preview never posts", checked
    ws.Cells(7, COL_MODIFICATION).Value2 = "수정 7"
    PreviewWebTasks False

    stage = "new registration and stable request binding"
    WebUploadMockReset "success"
    UploadWebTasks False
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "2", "two new requests", checked
    WebUploadRegressionEqual CStr(WebUploadMockJournalChecks), "2", "journal saved before every POST", checked
    Set state = WebTaskState(firstId)
    no = CStr(state("RequestNo"))
    Set state = WebTaskState(GetWebTaskId(ws, 8))
    secondNo = CStr(state("RequestNo"))
    WebUploadRegressionAssert Len(no) > 0 And no <> secondNo, "distinct request bindings", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "기등록", "new mode excludes an already linked request", checked
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "2", "repeat never duplicates", checked
    WebUploadRegressionSelectMode "기존 갱신"
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "변경 없음", "same payload skipped in update mode", checked

    stage = "lost existing request connection never falls back to insert"
    ws.Cells(7, COL_MODIFICATION).Value2 = "연결 소실 재검증"
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "갱신", "existing request prepared for update", checked
    expectedRequestNoExisting = CStr(mPreviewItems(1)("RequestNo"))
    WebUploadRegressionEqual expectedRequestNoExisting, no, "preview retains exact existing request number", checked
    Set connection = ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME)
    connectionRow = 0
    For r = 2 To connection.Cells(connection.Rows.Count, "A").End(xlUp).Row
        If CStr(connection.Cells(r, "A").Value2) = firstId And CStr(connection.Cells(r, "B").Value2) = WebCurrentSite() Then connectionRow = r: Exit For
    Next r
    WebUploadRegressionAssert connectionRow > 0, "existing connection fixture located", checked
    connection.Unprotect
    connection.Cells(connectionRow, "D").ClearContents
    connection.Protect UserInterfaceOnly:=True
    WebUploadMockReset "success"
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionAssert Len(caught) > 0, "lost update binding is rejected", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "lost update binding never inserts", checked
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "0", "lost update binding never updates another request", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "lost update binding does not fetch new form", checked
    mPreviewItems(1)("Operation") = "신규"
    mPreviewItems(1)("Error") = vbNullString
    On Error Resume Next
    UploadWebTasks False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "tampered lost-binding operation cannot insert", checked
    WebUploadRegressionEqual CStr(WebUploadMockLoadCount), "0", "tampered lost-binding operation cannot load new form", checked
    connection.Unprotect
    connection.Cells(connectionRow, "D").Value2 = expectedRequestNoExisting
    connection.Protect UserInterfaceOnly:=True
    ws.Cells(7, COL_MODIFICATION).Value2 = "수정 7"
    PreviewWebTasks False
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), expectedRequestNoExisting, "existing request binding restored without replacement", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "변경 없음", "restored binding is assessed as existing request", checked
    ws.Unprotect
    ApplyWebTaskProtection ws
    WebUploadRegressionAssert ws.Cells(5, COL_WEB_LEVEL).Locked And ws.Cells(8, COL_WEB_LEVEL).Locked, "whole L1 basis locked after partial group registration", checked
    WebUploadRegressionAssert ws.Cells(7, COL_WEB_REQUESTER).Locked, "requester locked after registration", checked
    WebUploadRegressionAssert ws.Cells(7, COL_WEB_RESULT).Locked, "result protected", checked
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "N"
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_ENABLED)
    WebUploadRegressionAssert WebGroupIsLocked(ws, 5), "N selection cannot unlock binding", checked
    ws.Cells(7, COL_WEB_ENABLED).Value2 = "Y"
    HandleWebTaskChange ws, ws.Cells(7, COL_WEB_ENABLED)
    ws.Cells(9, COL_LEVEL).Value2 = 3
    WebUploadRegressionContains WebTaskStructureIssue(ws), "등록 단위", "registered L2 fallback cannot become non-target", checked
    ws.Cells(9, COL_LEVEL).Value2 = 1
    WebUploadRegressionEqual WebTaskStructureIssue(ws), "", "restored hierarchy valid", checked

    stage = "update uses same request and preserves web-only processing text"
    Set session = OpenConfiguredWebSession()
    Set server = WebReadExistingValues(session, no)
    server("ACT_RMK") = "web processing note"
    WebUploadMockSetServer no, server
    ws.Cells(7, COL_MODIFICATION).Value2 = "갱신할 수정 내용"
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "갱신", "edited registered task is update", checked
    WebUploadMockReset "success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "1", "exactly one existing request updated", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "update does not create a request", checked
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), no, "request number retained", checked
    Set server = WebReadExistingValues(session, no)
    WebUploadRegressionEqual CStr(server("ACT_RMK")), "web processing note", "unmanaged web note preserved", checked

    stage = "failed update preserves successful connection and permits safe retry"
    baseline = CStr(state("Baseline"))
    ws.Cells(7, COL_MODIFICATION).Value2 = "갱신 실패 후 재시도"
    PreviewWebTasks False
    WebUploadMockReset "failed"
    UploadWebTasks False
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), no, "failed update retains confirmed request", checked
    WebUploadRegressionEqual CStr(state("Baseline")), baseline, "failed update retains last successful fingerprint", checked
    WebUploadRegressionEqual CStr(state("Outcome")), "갱신실패", "recent attempt separated from binding", checked
    WebUploadRegressionEqual CStr(state("PendingNo")), "", "known update failure clears attempt reservation", checked
    WebUploadMockReset "success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "1", "retry updates existing request", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "0", "failed update never falls back to insert", checked

    stage = "server drift requires explicit verified reconnection"
    Set server = WebReadExistingValues(session, no)
    server("REQ_TIT") = "웹에서 직접 고친 제목"
    WebUploadMockSetServer no, server
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "확인 필요", "server conflict blocks update", checked
    WebUploadRegressionContains CStr(preview.Cells(5, 13).Value2), "웹 내용", "conflict message", checked
    WebUploadMockReset "success"
    UploadWebTasks False
    WebUploadRegressionEqual CStr(WebUploadMockUpdateCount), "0", "conflict never overwrites", checked
    ConnectWebRequest ws, 7, no, session, True
    WebUploadRegressionEqual CStr(preview.Range("A1").Value2), "요청번호 연결 확인: " & no, "full connection review shown", checked
    WebUploadRegressionEqual CStr(preview.Range("B7").Value2), "웹에서 직접 고친 제목", "review contains current web title", checked
    WebUploadRegressionAssert preview.Shapes("btnWebLinkConfirm").Visible, "explicit reviewed connection action", checked
    WebUploadRegressionAssert mPreviewItems Is Nothing, "connection review cannot run stale upload preview", checked
    baseline = mLinkServerFingerprint
    server("REQ_TIT") = "연결 검토 후 웹 변경"
    WebUploadMockSetServer no, server
    On Error Resume Next
    ConnectWebRequest ws, 7, no, session, False, baseline
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "웹 내용이 바뀌", "stale connection review rejected", checked
    ConnectWebRequest ws, 7, no, session, False
    ' The real reviewed-connection action returns to the source sheet.
    ws.Activate
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "갱신", "explicit reconnection accepts reviewed server baseline", checked
    On Error Resume Next
    ConnectWebRequest ws, 8, no, session, False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "다른 업무", "same request cannot bind two units", checked
    config.Range("B2").Value2 = "http://other.invalid/support/"
    PreviewWebTasks False
    WebUploadRegressionContains CStr(preview.Cells(5, 13).Value2), "다른 사이트", "site change cannot reuse existing task binding", checked
    config.Range("B2").Value2 = "http://fixture.invalid/support/"

    stage = "missing requester option retains verified immutable server ID"
    config.Range("E9:G10").ClearContents
    Set items = BuildWebUploadItems(ws)
    WebUploadRegressionEqual CStr(items(1)("Values")("REQ_UID")), "requester1", "removed requester uses confirmed code", checked
    WebUploadRegressionAssert InStr(CStr(items(1)("Error")), "요청자") = 0, "verified requester does not block existing update", checked
    StoreWebOptionLists WebUploadRegressionForm(False, False)

    stage = "partial update is blocked and confirmed binding survives"
    PreviewWebTasks False
    WebUploadMockReset "partial"
    UploadWebTasks False
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), no, "partial update retains request", checked
    WebUploadRegressionEqual CStr(state("Outcome")), "갱신 일부저장", "partial is unresolved", checked
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "확인 필요", "partial does not auto retry", checked

    stage = "unknown new registration reserves candidate and user-confirmed recovery"
    WebUploadRegressionResetJournal
    WebUploadMockReset "unknown", True
    ws.Activate
    PreviewWebTasks False
    WebUploadRegressionEqual CStr(preview.Cells(5, 1).Value2), "연결 필요", "unknown fixture starts with explicit update block", checked
    WebUploadRegressionSelectMode "신규 등록"
    UploadWebTasks False
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("Outcome")), "결과 미확인", "unknown recorded", checked
    WebUploadRegressionAssert Len(CStr(state("PendingNo"))) > 0 And Len(CStr(state("RequestNo"))) = 0, "candidate is not confirmed request", checked
    WebUploadRegressionAssert WebRequestAlreadyLinked(GetWebTaskId(ws, 8), CStr(state("PendingNo"))), "pending request also reserves ownership", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount), "1", "unknown stops batch", checked
    On Error Resume Next
    ConfirmWebAttemptUnregistered ws, 7, False
    caught = Err.Description: Err.Clear
    On Error GoTo Failed
    WebUploadRegressionContains caught, "직접 확인", "no automatic unregistered classification", checked
    ConfirmWebAttemptUnregistered ws, 7, True
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("Outcome")), "미등록 확인", "explicit recovery journal", checked
    WebUploadRegressionEqual CStr(state("PendingNo")), "", "candidate reservation released only after confirmation", checked
    WebUploadRegressionAssert Not WebGroupIsLocked(ws, 5), "unregistered confirmed group unlocks", checked
    stage = "legacy migration preserves confirmed number after start and failure rows"
    WebUploadRegressionResetJournal
    log.Unprotect
    For r = 2 To 4
        log.Cells(r, "A").Value = Now
        log.Cells(r, "B").Value2 = ws.Name
        log.Cells(r, "C").Value2 = firstId
        log.Cells(r, "D").Value2 = "RQLEGACY001"
    Next r
    log.Cells(2, "E").Value2 = "등록 시작"
    log.Cells(3, "E").Value2 = "등록완료"
    log.Cells(4, "E").Value2 = "갱신실패"
    log.Protect UserInterfaceOnly:=True
    MigrateLegacyWebJournal
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), "RQLEGACY001", "legacy confirmed number outranks earlier start", checked
    WebUploadRegressionEqual CStr(state("PendingNo")), "", "legacy confirmed request not demoted to candidate", checked
    WebUploadRegressionEqual CStr(state("Outcome")), "기존이력 확인필요", "legacy still requires verified baseline", checked
    WebUploadRegressionAssert WebGroupIsLocked(ws, 5), "legacy confirmed group locked", checked
    WebUploadRegressionResetJournal
    log.Unprotect
    log.Cells(2, "A").Value = Now
    log.Cells(2, "C").Value2 = firstId
    log.Cells(2, "D").Value2 = "RQOTHERSITE"
    log.Cells(2, "E").Value2 = "등록완료"
    log.Cells(2, "L").Value2 = "http://different.invalid/support/"
    log.Protect UserInterfaceOnly:=True
    MigrateLegacyWebJournal
    Set state = WebTaskState(firstId)
    WebUploadRegressionEqual CStr(state("RequestNo")), "", "site-tagged journal never rebound as current-site legacy", checked
    stage = "registered child ownership cannot change silently"
    WebUploadRegressionResetJournal
    ws.Cells(6, COL_WEB_LEVEL).Value2 = 2
    HandleWebTaskChange ws, ws.Cells(6, COL_WEB_LEVEL)
    Set item = BuildWebItemForRow(ws, 6)
    Set mPreviewSource = ws
    RecordWebOutcome item, "등록완료", "RQOWNERSHIP", "offline fixture"
    ws.Cells(7, COL_MODIFICATION).Value2 = "같은 등록 단위 안의 본문 변경"
    WebUploadRegressionEqual WebTaskStructureIssue(ws), "", "body edits retain registration representative", checked
    ws.Cells(7, COL_LEVEL).Value2 = 2
    WebUploadRegressionContains WebTaskStructureIssue(ws), "다른 등록 단위", "registered child cannot become independent registration unit", checked
    ws.Cells(7, COL_LEVEL).Value2 = 3
    WebUploadRegressionEqual WebTaskStructureIssue(ws), "", "restored child ownership accepted", checked
    stage = "deleted journal entries release only unsupported web connections on refresh"
    WebUploadRegressionHistoryRefresh ws, config, checked
    WebUploadRegressionAssert Not mWebBusy And Not Application.EnableEvents, "busy and events restored", checked
    RunWebUploadRegression = "PASS: " & CStr(checked) & " web upload integration assertions"
    Exit Function
Failed:
    RunWebUploadRegression = "FAIL during " & stage & ": " & Err.Source & " / " & Err.Description
End Function

Private Sub WebUploadRegressionHistoryRefresh(ByVal ws As Worksheet, ByVal config As Worksheet, ByRef checked As Long)
    Dim log As Worksheet, connection As Worksheet, state As Object, item As Object
    Dim selectedId As String, neighborId As String, currentSite As String, oldSite As String
    Dim businessSnapshot As String, identitySnapshot As String, selectionSnapshot As String, connectionSnapshot As String
    Dim neighborSnapshot As String, oldSiteSnapshot As String, legacySnapshot As String
    Dim r As Long, legacyRow As Long, oldStateRow As Long, logLastRow As Long, previousAlerts As Boolean
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    Set connection = ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME)
    currentSite = WebCurrentSite()
    oldSite = "http://history-old.invalid/support/"
    selectedId = GetWebTaskId(ws, 6)
    neighborId = GetWebTaskId(ws, 8)
    ws.Range(COL_WEB_ENABLED & "6:" & COL_WEB_ENABLED & "8").Value2 = "Y"
    Set mPreviewSource = ws
    Set item = BuildWebItemForRow(ws, 8)
    RecordWebOutcome item, "등록완료", "RQHISTORYNEIGHBOR", "unrelated history fixture"
    Set item = BuildWebItemForRow(ws, 6)
    config.Range("B2").Value2 = oldSite
    RecordWebOutcome item, "등록완료", "RQHISTORYOLD", "valid old-site history fixture"
    config.Range("B2").Value2 = currentSite
    RecordWebOutcome item, "갱신실패", "RQOWNERSHIP", "latest history fixture"
    ws.Cells(9, COL_WEB_LEVEL).Value2 = 1
    Set item = BuildWebItemForRow(ws, 9)
    RecordWebOutcome item, "등록실패", "RQHISTORYFAILED", "failed registration retains its journal without a state request number"
    connectionSnapshot = WebUploadRegressionRangeSnapshot(connection.Range("A2:O" & connection.Cells(connection.Rows.Count, "A").End(xlUp).Row))
    WebUploadMockReset "success"
    RefreshWebHistoryState ws
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(connection.Range("A2:O" & connection.Cells(connection.Rows.Count, "A").End(xlUp).Row)), connectionSnapshot, "valid current and former-site history preserve all connection values", checked
    WebUploadRegressionContains CStr(ws.Cells(9, COL_WEB_RESULT).Value2), "등록실패", "failed registration display survives while its candidate-number journal remains", checked
    WebUploadRegressionClearHistory log, GetWebTaskId(ws, 9), currentSite
    RefreshWebHistoryState ws
    connectionSnapshot = WebUploadRegressionRangeSnapshot(connection.Range("A2:O" & connection.Cells(connection.Rows.Count, "A").End(xlUp).Row))
    log.Unprotect
    log.Rows(log.Cells(log.Rows.Count, "A").End(xlUp).Row).Delete
    log.Protect UserInterfaceOnly:=True
    RefreshWebHistoryState ws
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(connection.Range("A2:O" & connection.Cells(connection.Rows.Count, "A").End(xlUp).Row)), connectionSnapshot, "deleting only latest attempt retains earlier evidence for same request", checked
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    WebUploadRegressionAssert ws.Cells(6, COL_WEB_REQUESTER).Locked And ws.Cells(8, COL_WEB_REQUESTER).Locked, "supported requests remain locked before journal removal", checked
    businessSnapshot = WebUploadRegressionBusinessSnapshot(ws)
    identitySnapshot = WebUploadRegressionIdSnapshot(ws)
    selectionSnapshot = WebUploadRegressionRangeSnapshot(ws.Range(COL_WEB_ENABLED & "6:" & COL_WEB_PROGRAM_ID & "8"))
    Set state = WebTaskState(neighborId)
    neighborSnapshot = WebUploadRegressionStateSnapshot(state)
    For r = 2 To connection.Cells(connection.Rows.Count, "A").End(xlUp).Row
        If CStr(connection.Cells(r, "A").Value2) = selectedId And CStr(connection.Cells(r, "B").Value2) = oldSite Then oldStateRow = r: Exit For
    Next r
    WebUploadRegressionAssert oldStateRow > 0, "former-site connection fixture exists", checked
    oldSiteSnapshot = WebUploadRegressionRangeSnapshot(connection.Range("A" & oldStateRow & ":O" & oldStateRow))
    WebUploadRegressionClearHistory log, selectedId, currentSite
    Set mPreviewItems = New Collection
    mPreviewItems.Add item
    Set mPreviewSource = ws
    mPreviewSnapshot = "stale-deleted-history-preview"
    mPreviewMode = "기존 갱신"
    Set mLinkSource = ws
    mLinkSourceId = selectedId
    mLinkRequestNo = "RQOWNERSHIP"
    RefreshWebHistoryState ws
    Set state = WebTaskState(selectedId)
    WebUploadRegressionEqual CStr(state("RequestNo")) & CStr(state("PendingNo")), "", "deleting all matching history clears current-site binding", checked
    WebUploadRegressionEqual CStr(state("Outcome")), "이력 삭제", "deleted history leaves an explicit tombstone", checked
    WebUploadRegressionEqual WebUploadRegressionBusinessSnapshot(ws), businessSnapshot, "history reconciliation preserves body dates minutes and report data", checked
    WebUploadRegressionEqual WebUploadRegressionIdSnapshot(ws), identitySnapshot, "history reconciliation preserves stable task IDs", checked
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(ws.Range(COL_WEB_ENABLED & "6:" & COL_WEB_PROGRAM_ID & "8")), selectionSnapshot, "history reconciliation preserves Y selection basis and web values", checked
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_WEB_RESULT).Value2) & CStr(ws.Cells(7, COL_WEB_RESULT).Value2), "", "unsupported root and child result displays are cleared", checked
    WebUploadRegressionAssert Not ws.Cells(6, COL_WEB_REQUESTER).Locked And Not ws.Cells(7, COL_WEB_REQUESTER).Locked, "deleted current-site history releases root and inherited requester", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_REQUESTER).Font.Color), CStr(ws.Cells(7, COL_REQUEST).Font.Color), "released requester is no longer grey", checked
    Set state = WebTaskState(neighborId)
    WebUploadRegressionEqual WebUploadRegressionStateSnapshot(state), neighborSnapshot, "other registration unit keeps its request and baseline", checked
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(connection.Range("A" & oldStateRow & ":O" & oldStateRow)), oldSiteSnapshot, "valid former-site evidence keeps its separate binding", checked
    WebUploadRegressionAssert ws.Cells(8, COL_WEB_REQUESTER).Locked And WebGroupIsLocked(ws, 5), "remaining requests keep their requester and group locks", checked
    WebUploadRegressionEqual WebTaskStructureIssue(ws), "", "history tombstone does not invalidate a supported other-site connection", checked
    WebUploadRegressionAssert ws.ProtectContents And Not Application.EnableEvents, "history refresh restores protected sheet and event state", checked
    WebUploadRegressionAssert mPreviewItems Is Nothing And Len(mPreviewSnapshot) = 0 And Len(mPreviewMode) = 0, "history deletion invalidates prepared upload", checked
    WebUploadRegressionAssert mLinkSource Is Nothing And Len(mLinkSourceId) = 0, "history deletion invalidates pending connection review", checked
    WebUploadRegressionAssert Not ThisWorkbook.Saved, "history reconciliation does not auto-save the workbook", checked

    log.Unprotect
    legacyRow = log.Cells(log.Rows.Count, "A").End(xlUp).Row + 1
    log.Cells(legacyRow, "A").Value = Now
    log.Cells(legacyRow, "B").Value2 = ws.Name
    log.Cells(legacyRow, "C").Value2 = selectedId
    log.Cells(legacyRow, "D").Value2 = "RQHISTORYLEGACY"
    log.Cells(legacyRow, "E").Value2 = "등록완료"
    log.Protect UserInterfaceOnly:=True
    legacySnapshot = WebUploadRegressionRangeSnapshot(log.Range("A" & legacyRow & ":M" & legacyRow))
    MigrateLegacyWebJournal
    Set state = WebTaskState(selectedId)
    WebUploadRegressionEqual CStr(state("RequestNo")), "", "same-site legacy migration does not restore deleted history", checked
    config.Range("B2").Value2 = "http://history-future.invalid/support/"
    MigrateLegacyWebJournal
    Set state = WebTaskState(selectedId)
    WebUploadRegressionEqual CStr(state("RequestNo")) & CStr(state("PendingNo")), "", "site changes cannot revive obsolete legacy request", checked
    config.Range("B2").Value2 = currentSite
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(log.Range("A" & legacyRow & ":M" & legacyRow)), legacySnapshot, "history reconciliation does not rewrite remaining journal rows", checked

    WebUploadRegressionClearHistory log, neighborId, currentSite
    EnsureConfigSheet
    ws.Activate
    RefreshGanttOnly False
    Set state = WebTaskState(neighborId)
    WebUploadRegressionEqual CStr(state("Outcome")), "이력 삭제", "Gantt-only refresh reconciles removed task history", checked
    WebUploadRegressionEqual CStr(state("RequestNo")), "", "Gantt-only refresh clears unsupported neighbor binding", checked
    WebUploadRegressionAssert Not ws.Cells(8, COL_WEB_REQUESTER).Locked, "Gantt-only refresh removes requester lock", checked
    WebUploadRegressionEqual CStr(ws.Cells(8, COL_WEB_ENABLED).Value2), "Y", "Gantt-only refresh preserves upload selection", checked
    WebUploadRegressionAssert WebTaskHasOtherSiteConnection(selectedId), "Gantt-only refresh retains supported former-site binding", checked
    WebUploadRegressionAssert Not Application.EnableEvents, "Gantt-only refresh restores disabled events", checked

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    log.Delete
    Application.DisplayAlerts = previousAlerts
    Set log = Nothing
    RefreshGanttSheet False
    For r = 2 To connection.Cells(connection.Rows.Count, "A").End(xlUp).Row
        WebUploadRegressionEqual CStr(connection.Cells(r, "D").Value2) & CStr(connection.Cells(r, "E").Value2), "", "all-refresh clears every unsupported site after journal sheet deletion", checked
        WebUploadRegressionEqual CStr(connection.Cells(r, "I").Value2), "이력 삭제", "deleted journal sheet retains inactive state identities", checked
    Next r
    WebUploadRegressionAssert Not WebGroupIsLocked(ws, 5) And Not ws.Cells(5, COL_WEB_LEVEL).Locked, "all-refresh unlocks basis when every history record is removed", checked
    WebUploadRegressionAssert Not ws.Cells(6, COL_WEB_REQUESTER).Locked And Not ws.Cells(8, COL_WEB_REQUESTER).Locked, "all-refresh releases all unsupported requester locks", checked
    WebUploadRegressionEqual WebUploadRegressionRangeSnapshot(ws.Range(COL_WEB_ENABLED & "6:" & COL_WEB_PROGRAM_ID & "8")), selectionSnapshot, "both refresh paths preserve all web selections", checked

    ws.Unprotect
    ws.Range(COL_WEB_LEVEL & "5:" & COL_WEB_LEVEL & "8").Value2 = 3
    Set item = BuildWebItemForRow(ws, 7)
    Set mPreviewSource = ws
    RecordWebOutcome item, "등록완료", "RQHISTORYCHILD", "new child registration after deleting old parent history"
    RefreshWebHistoryState ws
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_WEB_RESULT).Value2), "등록완료 RQHISTORYCHILD", "old parent tombstone does not clear a new child unit result", checked
    Set state = WebTaskState(GetWebTaskId(ws, 7))
    WebUploadRegressionEqual CStr(state("RequestNo")), "RQHISTORYCHILD", "old parent tombstone preserves the new child connection", checked
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    WebUploadRegressionClearHistory log, GetWebTaskId(ws, 7), currentSite
    RefreshWebHistoryState ws
    ws.Range(COL_WEB_LEVEL & "5:" & COL_WEB_LEVEL & "8").Value2 = 2
    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)

    Set item = BuildWebItemForRow(ws, 8)
    Set mPreviewSource = ws
    RecordWebOutcome item, "등록 시작", "RQHISTORYUNKNOWN", "unresolved history fixture"
    RecordWebOutcome item, "결과 미확인", "RQHISTORYUNKNOWN", "unresolved history fixture"
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    log.Unprotect
    logLastRow = log.Cells(log.Rows.Count, "A").End(xlUp).Row
    log.Range("A2:M" & Application.Max(2, logLastRow)).ClearContents
    log.Protect UserInterfaceOnly:=True
    PrepareWebTaskSheet ws
    Set state = WebTaskState(neighborId)
    WebUploadRegressionEqual CStr(state("PendingNo")), "", "setup refresh releases candidate whose entire journal was cleared", checked
    WebUploadRegressionAssert Not WebStateUnresolved(state) And Not WebGroupIsLocked(ws, 5), "clearing unresolved history releases its group", checked
    WebUploadRegressionEqual CStr(ws.Cells(8, COL_WEB_ENABLED).Value2), "Y", "unresolved history reset retains upload selection", checked
    WebUploadRegressionAssert ws.ProtectContents And Not ws.Cells(8, COL_WEB_REQUESTER).Locked, "setup refresh reapplies protection with requester editable", checked
    WebUploadRegressionEqual CStr(WebUploadMockOpenCount + WebUploadMockReadCount + WebUploadMockLoadCount), "0", "history and both Gantt refresh paths perform no web reads", checked
    WebUploadRegressionEqual CStr(WebUploadMockPostCount + WebUploadMockUpdateCount), "0", "history refresh performs no web writes", checked
    WebUploadRegressionEqual WebTaskStructureIssue(ws), "", "remaining hierarchy remains valid after history removal", checked
End Sub

Public Function RunSavedWebHistoryRegression() As String
    Dim ws As Worksheet, state As Object, r As Variant, checked As Long
    Dim businessSnapshot As String, identitySnapshot As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("WebUploadFixture")
    businessSnapshot = WebUploadRegressionBusinessSnapshot(ws)
    identitySnapshot = WebUploadRegressionIdSnapshot(ws)
    InitializeWebUpload
    WebUploadRegressionEqual WebCurrentSite(), "http://fixture.invalid/support/", "saved history fixture retains original site", checked
    WebUploadRegressionEqual WebUploadRegressionBusinessSnapshot(ws), businessSnapshot, "reopening after history deletion preserves business data", checked
    WebUploadRegressionEqual WebUploadRegressionIdSnapshot(ws), identitySnapshot, "reopening preserves persistent task identities", checked
    WebUploadRegressionEqual CStr(ws.Cells(6, COL_REQUEST).Value2), "세부 그룹", "original group body survives save and reopen", checked
    WebUploadRegressionEqual CStr(ws.Cells(7, COL_MODIFICATION).Value2), "같은 등록 단위 안의 본문 변경", "child body survives save and reopen", checked
    For Each r In Array(6, 8)
        Set state = WebTaskState(GetWebTaskId(ws, CLng(r)))
        WebUploadRegressionEqual CStr(state("Outcome")), "이력 삭제", "deleted-history tombstone survives reopening", checked
        WebUploadRegressionEqual CStr(state("RequestNo")) & CStr(state("PendingNo")), "", "reopening does not revive confirmed or candidate request", checked
        WebUploadRegressionAssert Not WebTaskHasOtherSiteConnection(GetWebTaskId(ws, CLng(r))), "reopening does not revive old-site state", checked
        WebUploadRegressionEqual CStr(ws.Cells(CLng(r), COL_WEB_ENABLED).Value2), "Y", "reopening preserves explicit web selection", checked
        WebUploadRegressionEqual CStr(ws.Cells(CLng(r), COL_WEB_RESULT).Value2), "", "unsupported result remains empty after reopening", checked
        WebUploadRegressionEqual CStr(ws.Cells(CLng(r), COL_WORK_MINUTES).Value2), "30", "work minutes survive history deletion and reopening", checked
        WebUploadRegressionAssert Not ws.Cells(CLng(r), COL_WEB_REQUESTER).Locked, "requester remains editable after reopening", checked
        WebUploadRegressionEqual CStr(ws.Cells(CLng(r), COL_WEB_REQUESTER).Font.Color), CStr(ws.Cells(CLng(r), COL_REQUEST).Font.Color), "requester stays in normal text color after reopening", checked
    Next r
    WebUploadRegressionAssert Not ws.Cells(7, COL_WEB_REQUESTER).Locked, "inherited child requester stays editable", checked
    WebUploadRegressionAssert Not WebGroupIsLocked(ws, 5) And Not ws.Cells(5, COL_WEB_LEVEL).Locked, "group basis remains unlocked after reopening", checked
    WebUploadRegressionAssert ws.ProtectContents And Not Application.EnableEvents, "reopening preserves protection and disabled events", checked
    RunSavedWebHistoryRegression = "PASS: " & CStr(checked) & " saved web history assertions"
    Exit Function
Failed:
    RunSavedWebHistoryRegression = "FAIL during saved web history: " & Err.Source & " / " & Err.Description
End Function

Private Sub WebUploadRegressionClearHistory(ByVal log As Worksheet, ByVal sourceId As String, ByVal site As String)
    Dim r As Long
    log.Unprotect
    For r = 2 To log.Cells(log.Rows.Count, "A").End(xlUp).Row
        If CStr(log.Cells(r, "C").Value2) = sourceId And CStr(log.Cells(r, "L").Value2) = site Then
            log.Range("A" & r & ":M" & r).ClearContents
        End If
    Next r
    log.Protect UserInterfaceOnly:=True
End Sub

Private Function WebUploadRegressionRangeSnapshot(ByVal cells As Range) As String
    Dim cell As Range, value As String
    For Each cell In cells.Cells
        value = CStr(cell.Value2)
        WebUploadRegressionRangeSnapshot = WebUploadRegressionRangeSnapshot & CStr(Len(value)) & ":" & value
    Next cell
End Function

Private Function WebUploadRegressionStateSnapshot(ByVal state As Object) As String
    Dim field As Variant, value As String
    For Each field In Array("SourceId", "Site", "Account", "RequestNo", "PendingNo", "GroupId", "Basis", _
                            "Operation", "Outcome", "AttemptedAt", "ConfirmedAt", "Baseline", "Members", "PendingSnapshot", "RequesterCode")
        value = CStr(state(CStr(field)))
        WebUploadRegressionStateSnapshot = WebUploadRegressionStateSnapshot & CStr(Len(value)) & ":" & value
    Next field
End Function

Private Sub WebUploadRegressionSelectMode(ByVal executionMode As String)
    Dim preview As Worksheet
    Set preview = ThisWorkbook.Worksheets(WEB_PREVIEW_SHEET_NAME)
    preview.Activate
    preview.Range("D3").Value2 = executionMode
    PreviewWebTasks False
End Sub

Private Sub WebUploadRegressionResetJournal()
    Dim log As Worksheet, state As Worksheet, ws As Worksheet
    EnsureWebLogSheet
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    log.Unprotect
    log.Hyperlinks.Delete
    log.Range("A2:M" & Application.Max(2, log.Cells(log.Rows.Count, "A").End(xlUp).Row)).ClearContents
    log.Protect UserInterfaceOnly:=True
    EnsureWebStateSheet
    Set state = ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME)
    state.Unprotect
    state.Range("A2:O" & Application.Max(2, state.Cells(state.Rows.Count, "A").End(xlUp).Row)).ClearContents
    state.Protect UserInterfaceOnly:=True
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
