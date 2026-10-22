Attribute VB_Name = "modWebState"
Option Explicit
' 2026-10-22 김좌현 수정 : 요청번호 연결과 최근 시도를 분리하고 등록 단위 편집 및 구조 잠금 추가

Public Const WEB_STATE_SHEET_NAME As String = "웹등록_연결"
Private Const WEB_STATE_FIRST_ROW As Long = 2
Private mWebBasisCache As Object

Public Function WebManagedFields() As Variant
    WebManagedFields = Array("GUBUN", "REQ_UID", "REQ_TIT", "REQ_RMK", "STS", "PGM_MD", "ACT_GUBUN", "ACT_UID", "PGM_NM", "PGM_ID", "ACT_DT_FR", "ACT_DT", "ACT_TIME")
End Function

Public Function WebCurrentSite() As String
    On Error Resume Next
    WebCurrentSite = LCase$(Trim$(CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2)))
    On Error GoTo 0
    If Len(WebCurrentSite) > 0 Then
        If Right$(WebCurrentSite, 1) <> "/" Then WebCurrentSite = WebCurrentSite & "/"
    End If
End Function

Private Function WebStateSheet(Optional ByVal create As Boolean = False) As Worksheet
    Dim previous As Object, headers As Variant, c As Long
    On Error Resume Next
    Set WebStateSheet = ThisWorkbook.Worksheets(WEB_STATE_SHEET_NAME)
    On Error GoTo 0
    If Not WebStateSheet Is Nothing Or Not create Then Exit Function
    Set previous = ThisWorkbook.ActiveSheet
    Set WebStateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    With WebStateSheet
        .Name = WEB_STATE_SHEET_NAME
        headers = Array("업무 ID", "사이트", "실행 계정", "확정 요청번호", "확인중 요청번호", "레벨1 그룹 ID", "등록 기준", "최근 작업", "최근 결과", "최근 시각", "마지막 확인 시각", "마지막 확인값 지문", "등록단위 업무 ID", "전송 예정값 지문", "확인된 요청자 ID")
        For c = 0 To UBound(headers)
            .Cells(1, c + 1).Value2 = headers(c)
        Next c
        .Columns("A:O").NumberFormat = "@"
        .Rows(1).Font.Bold = True
        previous.Activate
        .Visible = xlSheetVeryHidden
        .Protect UserInterfaceOnly:=True
    End With
End Function

Public Sub EnsureWebStateSheet()
    Dim ws As Worksheet
    Set ws = WebStateSheet(True)
    ws.Visible = xlSheetVeryHidden
End Sub

Public Sub MigrateLegacyWebJournal()
    Dim log As Worksheet, stateWs As Worksheet, ws As Worksheet, r As Long, target As Long, rowNum As Long, groupRow As Long
    Dim sourceId As String, outcome As String, requestNo As String, seen As Object, j As Long
    On Error Resume Next
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    On Error GoTo 0
    If log Is Nothing Then Exit Sub
    Set stateWs = WebStateSheet(True)
    Set seen = CreateObject("Scripting.Dictionary")
    For r = log.Cells(log.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        sourceId = CStr(log.Cells(r, "C").Value2)
        outcome = CStr(log.Cells(r, "E").Value2)
        requestNo = CStr(log.Cells(r, "D").Value2)
        If Len(sourceId) = 0 Then GoTo NextLegacy
        ' New journal rows already carry an explicit site and are represented by
        ' their site-scoped connection record. They must never be rebound to the
        ' currently selected site while migrating older workbooks.
        If Len(CStr(log.Cells(r, "L").Value2)) > 0 Then GoTo NextLegacy
        If seen.Exists(sourceId) Then GoTo NextLegacy
        seen(sourceId) = True
        target = WebStateRow(sourceId)
        If target > 0 Then GoTo NextLegacy
        For j = log.Cells(log.Rows.Count, "A").End(xlUp).Row To 2 Step -1
            If CStr(log.Cells(j, "C").Value2) = sourceId And CStr(log.Cells(j, "E").Value2) = "등록완료" And Len(CStr(log.Cells(j, "L").Value2)) = 0 Then
                outcome = "등록완료"
                requestNo = CStr(log.Cells(j, "D").Value2)
                Exit For
            End If
        Next j
        If outcome <> "등록완료" And outcome <> "등록 시작" And outcome <> "결과 미확인" Then GoTo NextLegacy
        For Each ws In ThisWorkbook.Worksheets
            If CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) = "요청 내용" Then
                rowNum = WebFindTaskRow(ws, sourceId)
                If rowNum > 0 Then Exit For
            End If
        Next ws
        If rowNum = 0 Then GoTo NextLegacy
        target = WebStateRow(sourceId, True)
        groupRow = WebGroupRootRow(ws, rowNum)
        stateWs.Unprotect
        stateWs.Cells(target, "C").Value2 = CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2)
        If outcome = "등록완료" Then stateWs.Cells(target, "D").Value2 = requestNo Else stateWs.Cells(target, "E").Value2 = requestNo
        If groupRow > 0 Then
            stateWs.Cells(target, "F").Value2 = GetWebTaskId(ws, groupRow)
            stateWs.Cells(target, "G").Value2 = CStr(ws.Cells(groupRow, COL_WEB_LEVEL).Value2)
        End If
        stateWs.Cells(target, "H").Value2 = "이전 이력"
        stateWs.Cells(target, "I").Value2 = "기존이력 확인필요"
        stateWs.Protect UserInterfaceOnly:=True
NextLegacy:
    Next r
End Sub

Public Sub RememberWebGroupBases(ByVal ws As Worksheet)
    Dim r As Long, raw As Variant, key As String
    If mWebBasisCache Is Nothing Then Set mWebBasisCache = CreateObject("Scripting.Dictionary")
    For r = DATA_START_ROW To GetLastDataRow(ws)
        If HasTaskContent(ws, r) And GetTaskLevel(ws, r) = 1 Then
            raw = ws.Cells(r, COL_WEB_LEVEL).Value2
            If Not IsError(raw) Then
                If IsNumeric(raw) Then
                    key = GetWebTaskId(ws, r)
                    If Len(key) > 0 Then mWebBasisCache(key) = CLng(raw)
                End If
            End If
        End If
    Next r
End Sub

Public Function WebBasisMergePlan(ByVal ws As Worksheet, ByVal groupRow As Long, ByVal newBasis As Long) As Object
    Dim result As Object, fields As Variant, oldBasis As Long, groupEnd As Long, r As Long, child As Long
    Dim field As Variant, newEnd As Long, consensus As String, hasValue As Boolean, key As String
    Set result = CreateObject("Scripting.Dictionary")
    fields = Array(COL_WEB_ENABLED, COL_WEB_REQUESTER, COL_WEB_ACTION, COL_WEB_STATUS, COL_WEB_MODULE, COL_WEB_HANDLER, COL_WEB_PROGRAM_ID)
    key = GetWebTaskId(ws, groupRow)
    If Not mWebBasisCache Is Nothing Then
        If mWebBasisCache.Exists(key) Then oldBasis = CLng(mWebBasisCache(key))
    End If
    If oldBasis = 0 Then oldBasis = newBasis
    groupEnd = GetTaskSubtreeEndRow(ws, groupRow, GetLastDataRow(ws))
    For r = groupRow To groupEnd
        If IsWebUploadRow(ws, r, newBasis, groupEnd) Then
            newEnd = GetTaskSubtreeEndRow(ws, r, groupEnd)
            For Each field In fields
                hasValue = False: consensus = vbNullString
                For child = r To newEnd
                    If IsWebUploadRow(ws, child, oldBasis, groupEnd) Then
                        If Not hasValue Then
                            consensus = CStr(ws.Cells(child, CStr(field)).Value2)
                            hasValue = True
                        ElseIf consensus <> CStr(ws.Cells(child, CStr(field)).Value2) Then
                            Err.Raise vbObjectError + 2610, , "등록 기준을 바꾸면 서로 다른 웹 설정이 한 건으로 합쳐집니다. 먼저 기존 등록 단위의 " & _
                                CStr(ws.Cells(HEADER_ROW, CStr(field)).Value2) & " 값을 통일하세요."
                        End If
                    End If
                Next child
                If hasValue Then result(CStr(r) & ":" & CStr(ws.Range(CStr(field) & "1").Column)) = consensus
            Next field
        End If
    Next r
    Set WebBasisMergePlan = result
End Function

Private Function WebStateRow(ByVal sourceId As String, Optional ByVal create As Boolean = False) As Long
    Dim ws As Worksheet, r As Long
    Set ws = WebStateSheet(create)
    If ws Is Nothing Then Exit Function
    For r = WEB_STATE_FIRST_ROW To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "A").Value2) = sourceId And CStr(ws.Cells(r, "B").Value2) = WebCurrentSite() Then
            WebStateRow = r
            Exit Function
        End If
    Next r
    If create Then
        WebStateRow = Application.Max(WEB_STATE_FIRST_ROW, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1)
        ws.Unprotect
        ws.Cells(WebStateRow, "A").Value2 = sourceId
        ws.Cells(WebStateRow, "B").Value2 = WebCurrentSite()
        ws.Protect UserInterfaceOnly:=True
    End If
End Function

Public Function WebTaskState(ByVal sourceId As String) As Object
    Dim result As Object, ws As Worksheet, r As Long, fields As Variant, c As Long
    Set result = CreateObject("Scripting.Dictionary")
    fields = Array("SourceId", "Site", "Account", "RequestNo", "PendingNo", "GroupId", "Basis", "Operation", "Outcome", "AttemptedAt", "ConfirmedAt", "Baseline", "Members", "PendingSnapshot", "RequesterCode")
    Set ws = WebStateSheet()
    r = WebStateRow(sourceId)
    For c = 0 To UBound(fields)
        If r = 0 Then
            result.Add CStr(fields(c)), vbNullString
        Else
            result.Add CStr(fields(c)), CStr(ws.Cells(r, c + 1).Value2)
        End If
    Next c
    Set WebTaskState = result
End Function

' Two modular hashes plus length avoid storing long report bodies in a single cell.
' Each managed field is compared independently, including an explicit empty value.
Private Function WebValueFingerprint(ByVal value As String) As String
    Dim a As Double, b As Double, i As Long, code As Long
    value = Replace(Replace(value, vbCrLf, vbLf), vbCr, vbLf)
    a = 7: b = 11
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code < 0 Then code = code + 65536
        a = a * 131# + code
        a = a - Fix(a / 2147483629#) * 2147483629#
        b = b * 137# + code
        b = b - Fix(b / 2147483587#) * 2147483587#
    Next i
    WebValueFingerprint = CStr(Len(value)) & ":" & Format$(a, "0") & ":" & Format$(b, "0")
End Function

Public Function WebValuesFingerprint(ByVal values As Object) As String
    Dim field As Variant, value As String
    For Each field In WebManagedFields()
        If Not values.Exists(CStr(field)) Then
            WebValuesFingerprint = WebValuesFingerprint & CStr(field) & "=?|"
        Else
            value = CStr(values(CStr(field)))
            WebValuesFingerprint = WebValuesFingerprint & CStr(field) & "=" & WebValueFingerprint(value) & "|"
        End If
    Next field
End Function

Public Function WebGroupRootRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    Dim r As Long
    For r = rowNum To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) And GetTaskLevel(ws, r) = 1 Then
            WebGroupRootRow = r
            Exit Function
        End If
    Next r
End Function

Public Function WebRegistrationRootRow(ByVal ws As Worksheet, ByVal rowNum As Long, Optional ByVal lastDataRow As Long = 0) As Long
    Dim groupRow As Long, groupEnd As Long, basis As Long, r As Long, raw As Variant
    groupRow = WebGroupRootRow(ws, rowNum)
    If groupRow = 0 Then Exit Function
    raw = ws.Cells(groupRow, COL_WEB_LEVEL).Value2
    If IsError(raw) Then Exit Function
    If Not IsNumeric(raw) Then Exit Function
    If CDbl(raw) < 1 Or CDbl(raw) > 3 Or CDbl(raw) <> Fix(CDbl(raw)) Then Exit Function
    basis = CLng(raw)
    If lastDataRow = 0 Then lastDataRow = GetLastDataRow(ws)
    groupEnd = GetTaskSubtreeEndRow(ws, groupRow, lastDataRow)
    For r = rowNum To groupRow Step -1
        If IsWebUploadRow(ws, r, basis, groupEnd) Then
            If GetTaskSubtreeEndRow(ws, r, groupEnd) >= rowNum Then
                WebRegistrationRootRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Public Function WebStateUnresolved(ByVal state As Object) As Boolean
    Select Case CStr(state("Outcome"))
        Case "등록 시작", "갱신 시작", "결과 미확인", "갱신 결과 미확인", "갱신 일부저장", "기존이력 확인필요"
            WebStateUnresolved = True
    End Select
End Function

Private Function WebHasAnyConnection(ByVal sourceId As String) As Boolean
    Dim ws As Worksheet, r As Long, outcome As String
    Set ws = WebStateSheet()
    If ws Is Nothing Then Exit Function
    For r = WEB_STATE_FIRST_ROW To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "A").Value2) = sourceId Then
            outcome = CStr(ws.Cells(r, "I").Value2)
            If Len(CStr(ws.Cells(r, "D").Value2)) > 0 Or outcome = "등록 시작" Or outcome = "갱신 시작" Or _
               outcome = "결과 미확인" Or outcome = "갱신 결과 미확인" Or outcome = "갱신 일부저장" Or outcome = "기존이력 확인필요" Then
                WebHasAnyConnection = True
                Exit Function
            End If
        End If
    Next r
End Function

Public Function WebTaskHasOtherSiteConnection(ByVal sourceId As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = WebStateSheet()
    If ws Is Nothing Then Exit Function
    For r = WEB_STATE_FIRST_ROW To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "A").Value2) = sourceId And CStr(ws.Cells(r, "B").Value2) <> WebCurrentSite() Then
            If Len(CStr(ws.Cells(r, "D").Value2)) > 0 Or Len(CStr(ws.Cells(r, "E").Value2)) > 0 Then WebTaskHasOtherSiteConnection = True: Exit Function
        End If
    Next r
End Function

Public Function WebGroupIsLocked(ByVal ws As Worksheet, ByVal groupRow As Long) As Boolean
    Dim r As Long, groupEnd As Long
    groupEnd = GetTaskSubtreeEndRow(ws, groupRow, GetLastDataRow(ws))
    For r = groupRow To groupEnd
        If WebHasAnyConnection(GetWebTaskId(ws, r)) Then WebGroupIsLocked = True: Exit Function
    Next r
End Function

Public Sub ApplyWebTaskProtection(ByVal ws As Worksheet)
    Dim r As Long, root As Long, lastRow As Long, state As Object, lockedGroup As Boolean
    Dim stateWs As Worksheet, rows As Variant, i As Long, connections As Object, requesters As Object
    Dim sourceId As String, groupEnd As Long
    If ws Is Nothing Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) <> "요청 내용" Then Exit Sub
    lastRow = GetLastDataRow(ws)
    ws.Columns(COL_WEB_TASK_ID).Locked = True
    ws.Range(COL_WEB_RESULT & DATA_START_ROW & ":" & COL_WEB_RESULT & ws.Rows.Count).Locked = True
    Set connections = CreateObject("Scripting.Dictionary")
    Set requesters = CreateObject("Scripting.Dictionary")
    Set stateWs = WebStateSheet()
    If Not stateWs Is Nothing Then
        If stateWs.Cells(stateWs.Rows.Count, "A").End(xlUp).Row >= WEB_STATE_FIRST_ROW Then
            rows = stateWs.Range("A2:I" & stateWs.Cells(stateWs.Rows.Count, "A").End(xlUp).Row).Value2
            For i = 1 To UBound(rows, 1)
                If Len(CStr(rows(i, 4))) > 0 Or CStr(rows(i, 9)) = "등록 시작" Or CStr(rows(i, 9)) = "갱신 시작" Or _
                   CStr(rows(i, 9)) = "결과 미확인" Or CStr(rows(i, 9)) = "갱신 결과 미확인" Or CStr(rows(i, 9)) = "갱신 일부저장" Or CStr(rows(i, 9)) = "기존이력 확인필요" Then
                    connections(CStr(rows(i, 1))) = True
                    If CStr(rows(i, 2)) = WebCurrentSite() Then requesters(CStr(rows(i, 1))) = True
                End If
            Next i
        End If
    End If
    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) = 1 Then
                lockedGroup = False
                groupEnd = GetTaskSubtreeEndRow(ws, r, lastRow)
                For i = r To groupEnd
                    If connections.Exists(GetWebTaskId(ws, i)) Then lockedGroup = True: Exit For
                Next i
            End If
            ws.Cells(r, COL_WEB_LEVEL).Locked = lockedGroup
            root = WebRegistrationRootRow(ws, r, lastRow)
            If root = 0 Then
                ws.Range(COL_WEB_ENABLED & r & ":" & COL_WEB_ENABLED & r).Locked = True
                ws.Range(COL_WEB_REQUESTER & r & ":" & COL_WEB_PROGRAM_ID & r).Locked = True
                Union(ws.Cells(r, COL_WEB_ENABLED), ws.Range(COL_WEB_REQUESTER & r & ":" & COL_WEB_PROGRAM_ID & r)).Font.Color = RGB(140, 140, 140)
            Else
                ws.Cells(r, COL_WEB_ENABLED).Locked = False
                ws.Range(COL_WEB_REQUESTER & r & ":" & COL_WEB_PROGRAM_ID & r).Locked = False
                Union(ws.Cells(r, COL_WEB_ENABLED), ws.Range(COL_WEB_REQUESTER & r & ":" & COL_WEB_PROGRAM_ID & r)).Font.Color = ws.Cells(r, COL_REQUEST).Font.Color
                If requesters.Exists(GetWebTaskId(ws, root)) Then
                    ws.Cells(r, COL_WEB_REQUESTER).Locked = True
                    ws.Cells(r, COL_WEB_REQUESTER).Font.Color = RGB(140, 140, 140)
                End If
            End If
        End If
    Next r
End Sub

Public Function WebUnitMembers(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim r As Long, lastRow As Long
    lastRow = GetTaskSubtreeEndRow(ws, rowNum, GetLastDataRow(ws))
    For r = rowNum To lastRow
        If HasTaskContent(ws, r) Then WebUnitMembers = WebUnitMembers & "|" & GetWebTaskId(ws, r)
    Next r
End Function

Public Sub WebRecordState(ByVal ws As Worksheet, ByVal item As Object, ByVal operation As String, _
                          ByVal outcome As String, ByVal requestNo As String, Optional ByVal confirmedValues As Object = Nothing)
    Dim stateWs As Worksheet, r As Long, groupRow As Long
    r = WebStateRow(CStr(item("SourceId")), True)
    Set stateWs = WebStateSheet()
    groupRow = CLng(item("GroupRow"))
    stateWs.Unprotect
    stateWs.Cells(r, "C").Value2 = CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2)
    If groupRow > 0 Then
        stateWs.Cells(r, "F").Value2 = GetWebTaskId(ws, groupRow)
        stateWs.Cells(r, "G").Value2 = CStr(ws.Cells(groupRow, COL_WEB_LEVEL).Value2)
    End If
    stateWs.Cells(r, "H").Value2 = operation
    stateWs.Cells(r, "I").Value2 = outcome
    stateWs.Cells(r, "J").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    stateWs.Cells(r, "N").Value2 = WebValuesFingerprint(item("Values"))
    If Len(requestNo) > 0 Then stateWs.Cells(r, "E").Value2 = requestNo
    If outcome = "등록완료" Or outcome = "갱신완료" Or outcome = "연결완료" Then
        stateWs.Cells(r, "D").Value2 = requestNo
        stateWs.Cells(r, "E").ClearContents
        stateWs.Cells(r, "K").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        If confirmedValues Is Nothing Then Set confirmedValues = item("Values")
        stateWs.Cells(r, "L").Value2 = WebValuesFingerprint(confirmedValues)
        stateWs.Cells(r, "M").Value2 = WebUnitMembers(ws, CLng(item("Row")))
        If confirmedValues.Exists("REQ_UID") Then stateWs.Cells(r, "O").Value2 = CStr(confirmedValues("REQ_UID"))
    ElseIf outcome = "등록실패" Or outcome = "갱신실패" Or outcome = "미등록 확인" Then
        stateWs.Cells(r, "E").ClearContents
    End If
    stateWs.Protect UserInterfaceOnly:=True
End Sub

Public Function WebRequestAlreadyLinked(ByVal sourceId As String, ByVal requestNo As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = WebStateSheet()
    If ws Is Nothing Then Exit Function
    For r = WEB_STATE_FIRST_ROW To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(r, "B").Value2) = WebCurrentSite() And CStr(ws.Cells(r, "A").Value2) <> sourceId Then
            If CStr(ws.Cells(r, "D").Value2) = requestNo Or CStr(ws.Cells(r, "E").Value2) = requestNo Then WebRequestAlreadyLinked = True: Exit Function
        End If
    Next r
End Function

' A previously registered representative must remain a representative of the same
' level-1 group and basis. In particular, adding a level-3 child to a level-2
' fallback cannot silently replace an existing web request with new requests.
Public Function WebTaskStructureIssue(ByVal ws As Worksheet) As String
    Dim stateWs As Worksheet, r As Long, taskRow As Long, groupRow As Long, sourceId As String
    Dim state As Object, taskWs As Worksheet, foundAnywhere As Boolean
    Dim rowMap As Object, entry As Object, member As Variant, memberId As String, owner As Long, memberWs As Worksheet
    Dim mapRow As Long, mapLast As Long
    Set stateWs = WebStateSheet()
    If stateWs Is Nothing Then Exit Function
    If stateWs.Cells(stateWs.Rows.Count, "A").End(xlUp).Row < WEB_STATE_FIRST_ROW Then Exit Function
    Set rowMap = CreateObject("Scripting.Dictionary")
    For Each taskWs In ThisWorkbook.Worksheets
        If CStr(taskWs.Cells(HEADER_ROW, COL_REQUEST).Value2) = "요청 내용" Then
            mapLast = GetLastDataRow(taskWs)
            For mapRow = DATA_START_ROW To mapLast
                sourceId = GetWebTaskId(taskWs, mapRow)
                If Len(sourceId) > 0 And HasTaskContent(taskWs, mapRow) Then
                    Set entry = CreateObject("Scripting.Dictionary")
                    entry.Add "Sheet", taskWs
                    entry.Add "Row", mapRow
                    entry.Add "Owner", WebRegistrationRootRow(taskWs, mapRow, mapLast)
                    Set rowMap(sourceId) = entry
                End If
            Next mapRow
        End If
    Next taskWs
    For r = WEB_STATE_FIRST_ROW To stateWs.Cells(stateWs.Rows.Count, "A").End(xlUp).Row
        sourceId = CStr(stateWs.Cells(r, "A").Value2)
        If Not WebHasAnyConnection(sourceId) Then GoTo NextState
        taskRow = 0: foundAnywhere = rowMap.Exists(sourceId)
        If foundAnywhere Then
            Set entry = rowMap(sourceId)
            Set taskWs = entry("Sheet")
            taskRow = CLng(entry("Row"))
        End If
        If Not foundAnywhere Then
            WebTaskStructureIssue = "등록 이력이 있는 업무 행이 삭제되었거나 식별값이 사라졌습니다. 요청번호 연결을 확인하세요."
            Exit Function
        End If
        If Not taskWs Is ws Then GoTo NextState
        groupRow = WebGroupRootRow(ws, taskRow)
        If groupRow = 0 Then
            WebTaskStructureIssue = "등록된 업무의 레벨1 그룹이 없습니다: " & CStr(taskRow) & "행"
        ElseIf GetWebTaskId(ws, groupRow) <> CStr(stateWs.Cells(r, "F").Value2) Or _
               CStr(ws.Cells(groupRow, COL_WEB_LEVEL).Value2) <> CStr(stateWs.Cells(r, "G").Value2) Or _
               WebRegistrationRootRow(ws, taskRow) <> taskRow Then
            WebTaskStructureIssue = "등록된 업무의 그룹·레벨·등록 단위가 바뀌었습니다: " & CStr(taskRow) & "행. 기존 구조를 복원한 뒤 갱신하세요."
        End If
        If Len(WebTaskStructureIssue) > 0 Then Exit Function
        For Each member In Split(CStr(stateWs.Cells(r, "M").Value2), "|")
            memberId = CStr(member)
            If Len(memberId) > 0 And memberId <> sourceId And rowMap.Exists(memberId) Then
                Set entry = rowMap(memberId)
                Set memberWs = entry("Sheet")
                If Not memberWs Is taskWs Or CLng(entry("Owner")) <> taskRow Then
                    WebTaskStructureIssue = "등록된 업무의 하위 행이 다른 등록 단위로 이동했습니다. 기존 구조를 복원한 뒤 확인하세요."
                    Exit Function
                End If
            End If
        Next member
NextState:
    Next r
End Function

Public Function WebFindTaskRow(ByVal ws As Worksheet, ByVal sourceId As String) As Long
    Dim r As Long
    For r = DATA_START_ROW To GetLastDataRow(ws)
        If GetWebTaskId(ws, r) = sourceId Then WebFindTaskRow = r: Exit Function
    Next r
End Function

Public Sub RefreshWebTaskPeriods(ByVal ws As Worksheet)
    Dim r As Long, child As Long, lastRow As Long, subtreeEnd As Long, minStart As Date, maxEnd As Date
    Dim hasStart As Boolean, hasEnd As Boolean, allEnded As Boolean, raw As Variant
    lastRow = GetLastDataRow(ws)
    For r = lastRow To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) And HasChildTask(ws, r, lastRow) Then
            hasStart = False: hasEnd = False: allEnded = True
            subtreeEnd = GetTaskSubtreeEndRow(ws, r, lastRow)
            child = r + 1
            Do While child <= subtreeEnd
                If HasTaskContent(ws, child) Then
                    raw = ws.Cells(child, COL_ACTUAL_START).Value
                    If Not IsError(raw) Then
                        If IsDate(raw) Then
                            If Not hasStart Then minStart = CDate(raw): hasStart = True
                            If CDate(raw) < minStart Then minStart = CDate(raw)
                        End If
                    End If
                    raw = ws.Cells(child, COL_ACTUAL_END).Value
                    If Not IsError(raw) Then
                        If IsDate(raw) Then
                            If Not hasEnd Then maxEnd = CDate(raw): hasEnd = True
                            If CDate(raw) > maxEnd Then maxEnd = CDate(raw)
                        Else
                            allEnded = False
                        End If
                    Else
                        allEnded = False
                    End If
                    child = GetTaskSubtreeEndRow(ws, child, lastRow) + 1
                Else
                    child = child + 1
                End If
            Loop
            If hasStart Then ws.Cells(r, COL_ACTUAL_START).Value = minStart Else ws.Cells(r, COL_ACTUAL_START).ClearContents
            If allEnded And hasEnd Then ws.Cells(r, COL_ACTUAL_END).Value = maxEnd Else ws.Cells(r, COL_ACTUAL_END).ClearContents
        End If
    Next r
End Sub
