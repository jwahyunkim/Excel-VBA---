' Append to modVersionMigration in a disposable workbook. The file picker and confirmation UI are bypassed.
' 2026-10-23 김좌현 수정 : 데이터 가져오기의 현행·구형 열 매핑, 웹 이력·연결 이관 및 충돌 사전 차단 검증
Private Const MIG_MATCH As String = "MigrationMatch"
Private Const MIG_KEEP As String = "MigrationKeep"
Private Const MIG_LEGACY As String = "MigrationLegacy"
Private Const MIG_SITE As String = "http://migration.invalid/support/"
Private Const MIG_OTHER_SITE As String = "http://migration-other.invalid/support/"
Private Const MIG_GROUP As String = "11111111-1111-4111-8111-111111111111"
Private Const MIG_TASK As String = "22222222-2222-4222-8222-222222222222"
Private Const MIG_CHILD As String = "33333333-3333-4333-8333-333333333333"
Private Const MIG_PENDING As String = "44444444-4444-4444-8444-444444444444"
Private Const MIG_TOMB As String = "55555555-5555-4555-8555-555555555555"
Private Const MIG_KEEP_ID As String = "66666666-6666-4666-8666-666666666666"
Private Const MIG_OLD_ID As String = "77777777-7777-4777-8777-777777777777"
Private Const MIG_DELETED_ID As String = "88888888-8888-4888-8888-888888888888"
Private Const MIG_OLD_DELETED As String = "99999999-9999-4999-8999-999999999999"

Public Function RunVersionMigrationRegression() As String
    Dim source As Workbook, legacy As Workbook, target As Worksheet, keep As Worksheet, old As Worksheet
    Dim sourceWs As Worksheet, sourceLog As Worksheet, sourceState As Worksheet, config As Worksheet
    Dim plan As Object, checked As Long, stage As String, copied As Long, failed As Long, logs As Long, states As Long
    Dim sourceBefore As String, configBefore As String, keepBefore As String, snapshot As String, caught As String
    Dim r As Long, headers As Variant, newId As String, sourcePath As String, preview As Worksheet
    On Error GoTo Failed
    Application.EnableEvents = False
    stage = "initialize disposable target workbook"
    InitializeWebUpload
    EnsureWebHistorySheets
    Set config = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    config.Range("B2").Value2 = MIG_SITE
    config.Range("B3").Value2 = "migration-target-account"
    config.Range("B6").Value2 = 2
    config.Range("B7").Value2 = "N"
    configBefore = MigrationCells(config.Range("B2:S12")) & MigrationCells(config.Range("Z2"))
    stage = "prepare isolated current-layout target sheet"
    Set target = ThisWorkbook.Worksheets.Add
    target.Name = MIG_MATCH
    SetupDataHeaders target
    MigrationSeedCurrent target, False
    target.Cells(6, COL_WEB_TASK_ID).Value2 = MIG_OLD_ID
    target.Cells(10, COL_LEVEL).Value2 = 2
    target.Cells(10, COL_REQUEST).Value2 = "old target-only row"
    target.Cells(10, COL_WEB_TASK_ID).Value2 = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    target.Range("AB6").Formula = "=AA6"
    target.Range("R6").Interior.Color = RGB(210, 225, 240)
    Set keep = ThisWorkbook.Worksheets.Add
    keep.Name = MIG_KEEP
    SetupDataHeaders keep
    MigrationSeedBusiness keep, 5, 1, "untouched target task"
    keep.Cells(5, COL_WEB_LEVEL).Value2 = 1
    keep.Cells(5, COL_WEB_ENABLED).Value2 = "Y"
    keep.Cells(5, COL_WEB_TASK_ID).Value2 = MIG_KEEP_ID
    MigrationLog ThisWorkbook, MIG_KEEP, MIG_KEEP_ID, "RQ2001", "등록완료", MIG_SITE
    MigrationState ThisWorkbook, MIG_KEEP_ID, "RQ2001", "", "등록완료", MIG_SITE, MIG_KEEP_ID, 1, "keep-baseline"
    MigrationLog ThisWorkbook, MIG_MATCH, MIG_OLD_ID, "RQ3001", "등록완료", MIG_SITE
    MigrationState ThisWorkbook, MIG_OLD_ID, "RQ3001", "", "등록완료", MIG_SITE, MIG_GROUP, 2, "old-baseline"
    MigrationLog ThisWorkbook, MIG_MATCH, MIG_OLD_DELETED, "RQ3002", "등록완료", MIG_SITE
    keepBefore = MigrationCells(keep.Range("B4:AE5")) & MigrationIdHistory(MIG_KEEP_ID)
    stage = "prepare external source workbook current-layout headers"
    Set source = Application.Workbooks.Add(xlWBATWorksheet)
    Set sourceWs = source.Worksheets(1)
    sourceWs.Name = MIG_MATCH
    sourceWs.Range("B4:AH4").Value2 = target.Range("B4:AH4").Value2
    MigrationSeedCurrent sourceWs, True
    MigrationLog source, MIG_MATCH, MIG_TASK, "RQ1001", "등록완료", MIG_SITE
    MigrationLog source, MIG_MATCH, MIG_TASK, "RQ1002", "등록완료", MIG_OTHER_SITE
    MigrationLog source, MIG_MATCH, MIG_PENDING, "RQ1003", "결과 미확인", MIG_SITE
    MigrationLog source, MIG_MATCH, MIG_DELETED_ID, "RQ1004", "등록완료", MIG_SITE
    MigrationState source, MIG_TASK, "RQ1001", "", "등록완료", MIG_SITE, MIG_GROUP, 2, "confirmed-baseline"
    MigrationState source, MIG_TASK, "RQ1002", "", "등록완료", MIG_OTHER_SITE, MIG_GROUP, 2, "other-site-baseline"
    MigrationState source, MIG_PENDING, "", "RQ1003", "결과 미확인", MIG_SITE, MIG_GROUP, 2, ""
    MigrationState source, MIG_TOMB, "", "", "이력 삭제", MIG_SITE, "", 0, ""
    Set old = source.Worksheets.Add
    old.Name = "MigrationSkip"
    old.Range("B4:AH4").Value2 = target.Range("B4:AH4").Value2
    MigrationSeedBusiness old, 5, 1, "unmatched source task"
    old.Cells(5, COL_WEB_TASK_ID).Value2 = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb"
    MigrationLog source, old.Name, CStr(old.Cells(5, COL_WEB_TASK_ID).Value2), "RQ9001", "등록완료", MIG_SITE
    Set old = source.Worksheets.Add
    old.Name = WEB_CONFIG_SHEET_NAME
    old.Range("B2").Value2 = "http://never-import-source-settings.invalid/"
    old.Range("B3").Value2 = "source-account-must-not-replace-target"
    old.Range("B4").Value2 = "source-password-must-not-copy"
    Set sourceLog = source.Worksheets(WEB_LOG_SHEET_NAME)
    Set sourceState = source.Worksheets(WEB_STATE_SHEET_NAME)
    sourceLog.Cells(2, "L").ClearContents
    sourceLog.Hyperlinks.Add Anchor:=sourceLog.Cells(2, "D"), Address:=MIG_SITE & "C0/c0200_02.aspx?REQ_NO=RQ1001", TextToDisplay:="RQ1001"
    sourceLog.Protect UserInterfaceOnly:=True
    sourceState.Visible = xlSheetVeryHidden
    sourceState.Protect UserInterfaceOnly:=True
    stage = "save external source workbook fixture"
    sourcePath = ThisWorkbook.Path & "\migration-source.xlsx"
    source.SaveAs Filename:=sourcePath, FileFormat:=xlOpenXMLWorkbook
    sourceBefore = MigrationSourceSnapshot(source)
    ThisWorkbook.Activate
    Set preview = MigrationUtility(ThisWorkbook, WEB_PREVIEW_SHEET_NAME)
    preview.Unprotect
    preview.Range("A5").Value2 = "stale pre-migration upload row"

    stage = "current layout maps actual business and web columns"
    Set plan = BuildWebMigrationPlan(source, ThisWorkbook)
    ImportOutputReportConfig source, ThisWorkbook
    ImportTaskSheet sourceWs, target, copied, failed
    MigrationEqual MigrationCells(target.Range("C5:AA9")), MigrationCells(sourceWs.Range("C5:AA9")), "current task fields, web inputs, identities and minutes copied by real columns", checked
    MigrationEqual MigrationCells(target.Range("AC5:AE9")), MigrationCells(sourceWs.Range("AC5:AE9")), "manual fields copied separately from calculated AB", checked
    MigrationEqual CStr(target.Range("AB6").Formula), "=AA6", "target calculated formula preserved before refresh", checked
    MigrationEqual CStr(target.Range("R6").Interior.Color), CStr(RGB(210, 225, 240)), "target input formatting preserved during value import", checked
    MigrationEqual CStr(Application.CountA(target.Range("H10:AE10"))), "0", "old trailing target input and task identity cleared", checked
    EnsureWebHistorySheets
    ApplyWebMigrationPlan plan, ThisWorkbook, logs, states
    ResetWebPreviewAfterMigration
    MigrationEqual CStr(preview.Range("A5").Value2), "", "migration clears stale upload preview rows", checked
    MigrationAssert logs > 0 And states > 0, "history plan reports copied journal and connection records", checked
    MigrationPrepareImported source
    MigrationCheckImported checked

    MigrationEqual MigrationCells(keep.Range("B4:AE5")) & MigrationIdHistory(MIG_KEEP_ID), keepBefore, "unmatched target sheet and all its web history remain intact", checked
    MigrationAssert configBefore = MigrationCells(config.Range("B2:S12")) & MigrationCells(config.Range("Z2")), "target site account credentials and option lists remain unchanged", checked
    MigrationEqual MigrationSourceSnapshot(source), sourceBefore, "source workbook values history identities and protection are unchanged", checked
    MigrationAssert source.Saved, "read-only migration does not dirty source workbook", checked
    MigrationEqual CStr(WebUploadMockPostCount) & CStr(WebUploadMockUpdateCount), "00", "data import never submits a real or mock web request", checked

    stage = "repeat import replaces the same authoritative history without duplicates"
    snapshot = MigrationIdHistory(MIG_TASK) & MigrationIdHistory(MIG_PENDING) & MigrationIdHistory(MIG_TOMB) & MigrationIdHistory(MIG_DELETED_ID)
    MigrationExecute source
    MigrationEqual MigrationIdHistory(MIG_TASK) & MigrationIdHistory(MIG_PENDING) & MigrationIdHistory(MIG_TOMB) & MigrationIdHistory(MIG_DELETED_ID), snapshot, "repeat import keeps exact web history and state without duplicate rows", checked
    MigrationCheckImported checked

    stage = "legacy blank-site journal uses exact request connection before changed source settings"
    sourceLog.Unprotect
    sourceLog.Range("D2").Hyperlinks.Delete
    sourceLog.Protect UserInterfaceOnly:=True
    MigrationExecute source
    MigrationEqual CStr(WebTaskState(MIG_TASK)("RequestNo")), "RQ1001", "blank-site history uses the original request connection site", checked
    MigrationEqual CStr(WebTaskState(MIG_TASK)("Baseline")), "confirmed-baseline", "changed source site setting does not erase existing confirmed baseline", checked
    MigrationEqual CStr(sourceLog.Cells(2, "L").Value2), "", "site inference does not rewrite source journal", checked
    stage = "unidentified legacy journal site is rejected before mutation"
    sourceLog.Unprotect
    sourceLog.Cells(5, "L").ClearContents
    sourceLog.Protect UserInterfaceOnly:=True
    source.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").ClearContents
    MigrationRejectPlan source, checked, "legacy history site cannot be identified"
    sourceLog.Unprotect
    sourceLog.Cells(5, "L").Value2 = MIG_SITE
    sourceLog.Protect UserInterfaceOnly:=True
    source.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2 = "http://never-import-source-settings.invalid/"
    stage = "planning rejects duplicate imported identities before mutation"
    sourceWs.Cells(7, COL_WEB_TASK_ID).Value2 = MIG_TASK
    MigrationRejectPlan source, checked, "duplicate source task ID"
    sourceWs.Cells(7, COL_WEB_TASK_ID).Value2 = MIG_CHILD
    stage = "planning rejects collision with an unmatched target task"
    keep.Cells(5, COL_WEB_TASK_ID).Value2 = MIG_TASK
    MigrationRejectPlan source, checked, "unmatched target task ID collision"
    keep.Cells(5, COL_WEB_TASK_ID).Value2 = MIG_KEEP_ID
    stage = "planning rejects one site request owned by different tasks"
    r = MigrationFindState(ThisWorkbook, MIG_KEEP_ID, MIG_SITE)
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Unprotect
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Cells(r, "D").Value2 = "RQ1001"
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Protect UserInterfaceOnly:=True
    MigrationRejectPlan source, checked, "same site request belongs to two task identities"
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Unprotect
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Cells(r, "D").Value2 = "RQ2001"
    ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Protect UserInterfaceOnly:=True

    stage = "web-free legacy layout clears unrelated target bindings and maps split fields"
    Set target = ThisWorkbook.Worksheets.Add
    target.Name = MIG_LEGACY
    SetupDataHeaders target
    MigrationSeedBusiness target, 5, 1, "old legacy target"
    target.Cells(5, COL_WEB_TASK_ID).Value2 = "cccccccc-cccc-4ccc-8ccc-cccccccccccc"
    target.Cells(5, COL_WEB_ENABLED).Value2 = "Y"
    target.Cells(5, COL_WEB_LEVEL).Value2 = 1
    target.Cells(5, COL_WORK_MINUTES).Value2 = 123
    MigrationLog ThisWorkbook, MIG_LEGACY, CStr(target.Cells(5, COL_WEB_TASK_ID).Value2), "RQ4001", "등록완료", MIG_SITE
    MigrationState ThisWorkbook, CStr(target.Cells(5, COL_WEB_TASK_ID).Value2), "RQ4001", "", "등록완료", MIG_SITE, CStr(target.Cells(5, COL_WEB_TASK_ID).Value2), 1, "obsolete-legacy-target"
    Set legacy = Application.Workbooks.Add(xlWBATWorksheet)
    Set old = legacy.Worksheets(1)
    old.Name = MIG_LEGACY
    headers = Array("No.", "Level", "타입", "대분류", "중분류", "소분류", "요청 내용", "수정 내용", "담당", "비고", "계획 시작일", "계획 종료일", "실제 시작일", "실제 종료일", "진행률", "정상 진행률", "진행률 수동", "상태 수동", "보고 상태")
    For r = 0 To UBound(headers)
        old.Cells(HEADER_ROW, r + 2).Value2 = headers(r)
    Next r
    old.Range("C5").Value2 = 1
    old.Range("D5:G5").Value2 = "legacy category"
    old.Range("H5").Value2 = "legacy request"
    old.Range("I5").Value2 = "legacy modification"
    old.Range("J5").Value2 = "legacy owner"
    old.Range("K5").Value2 = "legacy note"
    old.Range("L5").Value2 = CDbl(DateSerial(2026, 10, 1))
    old.Range("M5").Value2 = CDbl(DateSerial(2026, 10, 2))
    old.Range("N5").Value2 = CDbl(DateSerial(2026, 10, 3))
    old.Range("O5").Value2 = CDbl(DateSerial(2026, 10, 4))
    old.Range("P5").Value2 = 0.4
    old.Range("R5").Value2 = 0.7
    old.Range("S5").Value2 = "legacy manual state"
    old.Range("T5").Value2 = "legacy report state"
    snapshot = MigrationCells(old.Range("B4:T5"))
    ThisWorkbook.Activate
    Set plan = BuildWebMigrationPlan(legacy, ThisWorkbook)
    ImportOutputReportConfig legacy, ThisWorkbook
    ImportTaskSheet old, target, copied, failed
    MigrationEqual CStr(target.Cells(5, COL_PROGRESS).Value2), "0.4", "legacy raw progress maps past inserted minutes before refresh", checked
    MigrationEqual CStr(target.Cells(5, COL_WEEKLY_REPORT).Value2), "legacy report state", "legacy raw report state maps before normal refresh", checked
    EnsureWebHistorySheets
    ApplyWebMigrationPlan plan, ThisWorkbook, logs, states
    ResetWebPreviewAfterMigration
    MigrationPrepareImported legacy
    MigrationEqual CStr(target.Cells(5, COL_REQUEST).Value2), "legacy request", "legacy request mapped to current R column", checked
    MigrationEqual CStr(target.Cells(5, COL_MODIFICATION).Value2), "legacy modification", "legacy modification mapped separately", checked
    MigrationEqual CStr(target.Cells(5, COL_OWNER).Value2), "legacy owner", "legacy seven-column details preserve owner", checked
    MigrationEqual CStr(target.Cells(5, COL_PROGRESS).Value2), "1", "normal refresh marks an actually ended leaf as complete", checked
    MigrationEqual CStr(target.Cells(5, COL_WORK_MINUTES).Value2), "", "legacy file without minutes does not inherit old target minutes", checked
    MigrationEqual CStr(target.Cells(5, COL_MANUAL_PROGRESS).Value2), "0.7", "legacy manual progress preserved", checked
    MigrationEqual CStr(target.Cells(5, COL_WEEKLY_REPORT).Value2), REPORT_STATUS_COMPLETED, "normal refresh recalculates report state from actual end date", checked
    newId = GetWebTaskId(target, 5)
    MigrationAssert Len(newId) = 36 And newId <> "cccccccc-cccc-4ccc-8ccc-cccccccccccc", "legacy task receives its own fresh persistent identity", checked
    MigrationEqual CStr(target.Cells(5, COL_WEB_ENABLED).Value2), "N", "legacy task starts with safe target selection default", checked
    MigrationEqual MigrationIdHistory("cccccccc-cccc-4ccc-8ccc-cccccccccccc"), "", "legacy replacement removes old target web connection and history", checked
    MigrationEqual MigrationCells(old.Range("B4:T5")), snapshot, "legacy source remains unchanged", checked
    MigrationEqual MigrationCells(keep.Range("B4:AE5")) & MigrationIdHistory(MIG_KEEP_ID), keepBefore, "legacy import still preserves unmatched target sheet and history", checked
    MigrationEqual CStr(WebUploadMockPostCount) & CStr(WebUploadMockUpdateCount), "00", "migration regression performs no web writes", checked
    source.Close SaveChanges:=False
    legacy.Close SaveChanges:=False
    ThisWorkbook.Worksheets(MIG_MATCH).Activate
    RunVersionMigrationRegression = "PASS: " & CStr(checked) & " version migration assertions"
    Exit Function
Failed:
    RunVersionMigrationRegression = "FAIL during " & stage & ": " & Err.Source & " / " & Err.Description
    On Error Resume Next
    If Not source Is Nothing Then source.Close SaveChanges:=False
    If Not legacy Is Nothing Then legacy.Close SaveChanges:=False
End Function

Public Function RunSavedVersionMigrationRegression() As String
    Dim checked As Long, state As Object, ws As Worksheet
    On Error GoTo Failed
    Application.EnableEvents = False
    InitializeWebUpload
    MigrationCheckImported checked
    Set ws = ThisWorkbook.Worksheets(MIG_LEGACY)
    MigrationEqual CStr(ws.Cells(5, COL_REQUEST).Value2), "legacy request", "legacy imported request survives save and reopen", checked
    MigrationAssert Len(GetWebTaskId(ws, 5)) = 36, "generated legacy identity survives reopen", checked
    MigrationEqual CStr(ws.Cells(5, COL_WEB_ENABLED).Value2), "N", "legacy default N survives reopen", checked
    MigrationEqual CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2), MIG_SITE, "source web settings do not replace target site after reopen", checked
    MigrationAssert ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).Visible = xlSheetVeryHidden, "connection state stays very hidden after reopen", checked
    MigrationAssert ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME).ProtectContents, "connection state stays protected after reopen", checked
    RunSavedVersionMigrationRegression = "PASS: " & CStr(checked) & " saved version migration assertions"
    Exit Function
Failed:
    RunSavedVersionMigrationRegression = "FAIL during saved version migration: " & Err.Source & " / " & Err.Description
End Function

Private Sub MigrationExecute(ByVal source As Workbook)
    Dim plan As Object, ws As Worksheet, copied As Long, failed As Long, logs As Long, states As Long
    Set plan = BuildWebMigrationPlan(source, ThisWorkbook)
    ImportOutputReportConfig source, ThisWorkbook
    For Each ws In source.Worksheets
        If IsMigrationTaskPair(ws, ThisWorkbook) Then ImportTaskSheet ws, ThisWorkbook.Worksheets(ws.Name), copied, failed
    Next ws
    EnsureWebHistorySheets
    ApplyWebMigrationPlan plan, ThisWorkbook, logs, states
    ResetWebPreviewAfterMigration
    MigrationPrepareImported source
End Sub

Private Sub MigrationPrepareImported(ByVal source As Workbook)
    Dim ws As Worksheet, target As Worksheet, ignored As Boolean
    For Each ws In source.Worksheets
        If IsMigrationTaskPair(ws, ThisWorkbook) Then
            Set target = ThisWorkbook.Worksheets(ws.Name)
            PrepareWebTaskSheet target
            ignored = RefreshMigratedSheet(target)
        End If
    Next ws
End Sub

Private Sub MigrationCheckImported(ByRef checked As Long)
    Dim ws As Worksheet, state As Object
    Set ws = ThisWorkbook.Worksheets(MIG_MATCH)
    MigrationEqual GetWebTaskId(ws, 6), MIG_TASK, "imported representative keeps original persistent identity", checked
    MigrationEqual GetWebTaskId(ws, 7), MIG_CHILD, "imported child keeps original persistent identity", checked
    MigrationEqual CStr(ws.Cells(6, COL_REQUEST).Value2), "source representative", "modern request survives preparation and refresh", checked
    MigrationEqual CStr(ws.Cells(7, COL_MODIFICATION).Value2), "source modification 7", "modern child modification survives preparation and refresh", checked
    MigrationEqual CStr(ws.Cells(6, COL_WORK_MINUTES).Value2), "37", "modern parent work minutes preserve correct child sum", checked
    Set state = WebTaskState(MIG_TASK)
    MigrationEqual CStr(state("RequestNo")), "RQ1001", "confirmed request is available for update instead of new registration", checked
    MigrationEqual CStr(state("Baseline")), "confirmed-baseline", "confirmed web baseline preserved", checked
    MigrationAssert InStr(CStr(state("Members")), MIG_CHILD) > 0, "registered unit child membership preserved", checked
    Set state = WebTaskState(MIG_PENDING)
    MigrationEqual CStr(state("PendingNo")), "RQ1003", "uncertain candidate reservation preserved", checked
    MigrationAssert WebStateUnresolved(state), "uncertain import remains blocked from automatic duplicate registration", checked
    Set state = WebTaskState(MIG_TOMB)
    MigrationEqual CStr(state("Outcome")), "이력 삭제", "deleted-history tombstone preserved", checked
    MigrationEqual CStr(state("RequestNo")) & CStr(state("PendingNo")), "", "tombstone does not revive an old request", checked
    MigrationEqual CStr(MigrationCountLog(MIG_TASK, MIG_SITE)), "1", "one current-site registration history imported", checked
    MigrationEqual CStr(MigrationCountLog(MIG_TASK, MIG_OTHER_SITE)), "1", "separate other-site registration history imported", checked
    MigrationEqual CStr(MigrationCountLog(MIG_DELETED_ID, MIG_SITE)), "1", "deleted source task history retained by matching sheet name", checked
    MigrationEqual MigrationIdHistory(MIG_OLD_ID) & MigrationIdHistory(MIG_OLD_DELETED), "", "replaced target identities and deleted-task history removed", checked
    MigrationEqual CStr(MigrationCountLog("bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb", MIG_SITE)), "0", "unmatched source task history is excluded", checked
    MigrationAssert ws.ProtectContents And ws.Cells(6, COL_WEB_REQUESTER).Locked, "imported confirmed connection restores requester protection", checked
End Sub

Private Sub MigrationRejectPlan(ByVal source As Workbook, ByRef checked As Long, ByVal label As String)
    Dim before As String, sourceBefore As String, plan As Object, caught As String
    before = MigrationTargetSnapshot()
    sourceBefore = MigrationSourceSnapshot(source)
    On Error Resume Next
    Set plan = BuildWebMigrationPlan(source, ThisWorkbook)
    caught = Err.Description: Err.Clear
    On Error GoTo 0
    MigrationAssert Len(caught) > 0, label & " is rejected", checked
    MigrationEqual MigrationTargetSnapshot(), before, label & " leaves all target task and web values unchanged", checked
    MigrationEqual MigrationSourceSnapshot(source), sourceBefore, label & " does not mutate source", checked
End Sub

Private Sub MigrationSeedCurrent(ByVal ws As Worksheet, ByVal source As Boolean)
    Dim r As Long, prefix As String
    prefix = IIf(source, "source", "target")
    For r = 5 To 9
        MigrationSeedBusiness ws, r, IIf(r = 5, 1, IIf(r = 7, 3, 2)), prefix & " task " & r
        ws.Cells(r, COL_WEB_LEVEL).Value2 = 2
        ws.Cells(r, COL_WEB_ENABLED).Value2 = IIf(r = 5, "N", "Y")
        ws.Cells(r, COL_WEB_REQUESTER).Value2 = "요청자 [requester1]"
        ws.Cells(r, COL_WEB_ACTION).Value2 = "개발"
        ws.Cells(r, COL_WEB_STATUS).Value2 = "처리중"
        ws.Cells(r, COL_WEB_MODULE).Value2 = "생산"
        ws.Cells(r, COL_WEB_HANDLER).Value2 = "처리자 [worker1]"
        ws.Cells(r, COL_WEB_PROGRAM_ID).Value2 = "migration-program"
    Next r
    ws.Cells(5, COL_WEB_TASK_ID).Value2 = MIG_GROUP
    ws.Cells(6, COL_WEB_TASK_ID).Value2 = MIG_TASK
    ws.Cells(7, COL_WEB_TASK_ID).Value2 = MIG_CHILD
    ws.Cells(8, COL_WEB_TASK_ID).Value2 = MIG_PENDING
    ws.Cells(9, COL_WEB_TASK_ID).Value2 = MIG_TOMB
    ws.Cells(6, COL_REQUEST).Value2 = prefix & " representative"
    ws.Cells(5, COL_WORK_MINUTES).Value2 = 111
    ws.Cells(6, COL_WORK_MINUTES).Value2 = 37
End Sub

Private Sub MigrationSeedBusiness(ByVal ws As Worksheet, ByVal row As Long, ByVal level As Long, ByVal request As String)
    ws.Cells(row, COL_LEVEL).Value2 = level
    ws.Range("D" & row & ":G" & row).Value2 = "migration category"
    ws.Cells(row, COL_REQUEST).Value2 = request
    ws.Cells(row, COL_MODIFICATION).Value2 = "source modification " & row
    ws.Cells(row, COL_OWNER).Value2 = "migration owner"
    ws.Cells(row, COL_NOTE).Value2 = "migration note"
    ws.Cells(row, COL_PLAN_START).Value2 = CDbl(DateSerial(2026, 10, 1))
    ws.Cells(row, COL_PLAN_END).Value2 = CDbl(DateSerial(2026, 10, 2))
    ws.Cells(row, COL_ACTUAL_START).Value2 = CDbl(DateSerial(2026, 10, 3))
    ws.Cells(row, COL_ACTUAL_END).Value2 = CDbl(DateSerial(2026, 10, 4))
    ws.Cells(row, COL_WORK_MINUTES).Value2 = 37
    ws.Cells(row, COL_PROGRESS).Value2 = 0.4
    ws.Cells(row, COL_MANUAL_PROGRESS).Value2 = 0.7
    ws.Cells(row, COL_WEEKLY_REPORT).Value2 = "migration report state"
End Sub

Private Function MigrationUtility(ByVal book As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set MigrationUtility = book.Worksheets(name)
    On Error GoTo 0
    If MigrationUtility Is Nothing Then
        Set MigrationUtility = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        MigrationUtility.Name = name
        MigrationUtility.Range("A1:O1").Value2 = "migration header"
    End If
End Function

Private Sub MigrationLog(ByVal book As Workbook, ByVal sheet As String, ByVal id As String, ByVal request As String, ByVal outcome As String, ByVal site As String)
    Dim ws As Worksheet, row As Long
    Set ws = MigrationUtility(book, WEB_LOG_SHEET_NAME)
    ws.Unprotect
    row = Application.Max(2, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1)
    ws.Cells(row, "A").Value2 = CDbl(DateSerial(2026, 10, 20))
    ws.Cells(row, "B").Value2 = sheet
    ws.Cells(row, "C").Value2 = id
    ws.Cells(row, "D").Value2 = request
    ws.Cells(row, "E").Value2 = outcome
    ws.Cells(row, "F").Value2 = "migration fixture journal"
    ws.Cells(row, "G").Value2 = "migration title"
    ws.Cells(row, "J").Value2 = 37
    ws.Cells(row, "K").Value2 = "등록"
    ws.Cells(row, "L").Value2 = site
    ws.Cells(row, "M").Value2 = "original-account"
End Sub

Private Sub MigrationState(ByVal book As Workbook, ByVal id As String, ByVal request As String, ByVal pending As String, ByVal outcome As String, ByVal site As String, ByVal group As String, ByVal basis As Long, ByVal baseline As String)
    Dim ws As Worksheet, row As Long
    Set ws = MigrationUtility(book, WEB_STATE_SHEET_NAME)
    ws.Unprotect
    row = Application.Max(2, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1)
    ws.Range("A" & row & ":O" & row).NumberFormat = "@"
    ws.Cells(row, "A").Value2 = id
    ws.Cells(row, "B").Value2 = site
    ws.Cells(row, "C").Value2 = "original-account"
    ws.Cells(row, "D").Value2 = request
    ws.Cells(row, "E").Value2 = pending
    ws.Cells(row, "F").Value2 = group
    If basis > 0 Then ws.Cells(row, "G").Value2 = CStr(basis)
    ws.Cells(row, "H").Value2 = IIf(outcome = "이력 삭제", "이력 삭제", "등록")
    ws.Cells(row, "I").Value2 = outcome
    ws.Cells(row, "J").Value2 = "2026-10-20 12:00:00"
    ws.Cells(row, "L").Value2 = baseline
    If Len(request) > 0 Then
        ws.Cells(row, "K").Value2 = "2026-10-20 12:00:00"
        ws.Cells(row, "O").Value2 = "requester1"
    End If
    ws.Cells(row, "M").Value2 = "|" & id & IIf(id = MIG_TASK, "|" & MIG_CHILD, "")
    ws.Protect UserInterfaceOnly:=True
End Sub

Private Function MigrationFindState(ByVal book As Workbook, ByVal id As String, ByVal site As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = book.Worksheets(WEB_STATE_SHEET_NAME)
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "A").Value2) = id And CStr(ws.Cells(r, "B").Value2) = site Then MigrationFindState = r: Exit Function
    Next r
End Function

Private Function MigrationCountLog(ByVal id As String, ByVal site As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "C").Value2) = id And CStr(ws.Cells(r, "L").Value2) = site Then MigrationCountLog = MigrationCountLog + 1
    Next r
End Function

Private Function MigrationIdHistory(ByVal id As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "C").Value2) = id Then MigrationIdHistory = MigrationIdHistory & MigrationCells(ws.Range("A" & r & ":M" & r))
    Next r
    Set ws = ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME)
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "A").Value2) = id Then MigrationIdHistory = MigrationIdHistory & MigrationCells(ws.Range("A" & r & ":O" & r))
    Next r
End Function

Private Function MigrationSourceSnapshot(ByVal book As Workbook) As String
    Dim ws As Worksheet, link As Hyperlink
    For Each ws In book.Worksheets
        MigrationSourceSnapshot = MigrationSourceSnapshot & ws.Name & CStr(ws.Visible) & CStr(ws.ProtectContents) & MigrationUsedCells(ws)
        For Each link In ws.Hyperlinks
            MigrationSourceSnapshot = MigrationSourceSnapshot & link.Range.Address & link.Address & link.SubAddress
        Next link
    Next ws
End Function

Private Function MigrationTargetSnapshot() As String
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = MIG_MATCH Or ws.Name = MIG_KEEP Or ws.Name = MIG_LEGACY Or ws.Name = WEB_LOG_SHEET_NAME Or ws.Name = WEB_STATE_SHEET_NAME Or ws.Name = WEB_CONFIG_SHEET_NAME Then _
            MigrationTargetSnapshot = MigrationTargetSnapshot & ws.Name & CStr(ws.Visible) & CStr(ws.ProtectContents) & MigrationUsedCells(ws)
    Next ws
End Function

Private Function MigrationUsedCells(ByVal ws As Worksheet) As String
    Dim lastRow As Range, lastColumn As Range
    Set lastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastRow Is Nothing Then Exit Function
    Set lastColumn = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    MigrationUsedCells = MigrationCells(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow.Row, lastColumn.Column)))
End Function

Private Function MigrationCells(ByVal range As Range) As String
    Dim values As Variant, r As Long, c As Long, value As String
    values = range.Formula
    If range.Cells.CountLarge = 1 Then
        If IsError(values) Then value = "#ERROR" Else value = CStr(values)
        MigrationCells = Len(value) & ":" & value
        Exit Function
    End If
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If IsError(values(r, c)) Then value = "#ERROR" Else value = CStr(values(r, c))
            MigrationCells = MigrationCells & Len(value) & ":" & value
        Next c
    Next r
End Function

Private Sub MigrationAssert(ByVal ok As Boolean, ByVal label As String, ByRef checked As Long)
    If Not ok Then Err.Raise vbObjectError + 7991, "VersionMigrationRegression", label
    checked = checked + 1
End Sub

Private Sub MigrationEqual(ByVal actual As String, ByVal expected As String, ByVal label As String, ByRef checked As Long)
    MigrationAssert actual = expected, label, checked
End Sub
