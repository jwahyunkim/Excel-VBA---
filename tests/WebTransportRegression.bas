Attribute VB_Name = "WebTransportRegression"
Option Explicit
' 2026-10-22 김좌현 수정 : UTF-8·폼·암호화 및 로컬 웹 세션/등록 결과 확인 회귀 검증
' 2026-10-22 김좌현 수정 : 요청번호 갱신·충돌·부분 성공·웹 전용 본문 보존 회귀 검증

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1941, "WebTransportRegression", message
End Sub

Private Sub Trace(ByVal path As String, ByVal message As String)
    Dim handle As Integer
    If Len(path) = 0 Then Exit Sub
    handle = FreeFile
    Open path For Append As #handle
    Print #handle, message
    Close #handle
End Sub

Private Sub CheckExistingUpdates(ByVal session As Object, ByVal tracePath As String)
    Dim form As Object, values As Object, result As Object, actual As Object, key As Variant
    Dim invalidError As Long, title As String, body As String
    Trace tracePath, "existing update / two distinct saves"
    Set form = WebLoadExisting(session, "RQ9101")
    Check form("RequestNo") = "RQ9101" And form("Values")("REQ_UID") = "u1", "exact existing identity"
    Check Not form("Fields").Exists("REQ_UID"), "requester is read-only span"
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In form("Values").Keys
        values.Add CStr(key), form("Values")(key)
    Next key
    title = "[소분류] 갱신 ' + & %"
    body = "<p>첫 줄 &amp; 50% + ! '()</p>" & vbCrLf & "둘째 줄"
    values("REQ_TIT") = title
    values("REQ_RMK") = body
    values("STS") = "5"
    values("ACT_TIME") = "90"
    Set result = WebPostUpdate(session, form, values)
    Check result("Success") And result("Code") = "verified", "both tabs read back: " & result("Message")
    Check result("Values")("REQ_TIT") = title And result("Values")("REQ_RMK") = body, "HTML/entity/percent/body preserved"
    Check result("Values")("STS") = "5" And result("Values")("ACT_TIME") = "90", "processing values updated"
    Set actual = WebReadExistingValues(session, "RQ9101")
    Check actual("REQ_UID") = "u1" And actual("ACT_TIME") = "90", "read API managed values"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "failed", "consumed update form not retried"

    Trace tracePath, "unchanged / forbidden requester / forbidden field / stale choice"
    Set form = WebLoadExisting(session, "RQ9101")
    Set result = WebPostUpdate(session, form, actual)
    Check result("Success"), "unchanged values need no POST"
    Set form = WebLoadExisting(session, "RQ9101")
    Set values = CreateObject("Scripting.Dictionary")
    values.Add "REQ_UID", "u2"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "failed" And result("Code") = "requester_locked", "requester change rejected before POST"
    values.RemoveAll
    values.Add "ACT_RMK", "must not overwrite web-only processing body"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "failed" And result("Code") = "validation", "web-only field is not managed"
    values.RemoveAll
    values.Add "STS", "removed-status"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "failed", "stale processing choice rejected"

    Trace tracePath, "remote edit conflict"
    Set form = WebLoadExisting(session, "RQ9104")
    values.RemoveAll
    values.Add "REQ_TIT", "local-title"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "failed" And result("Code") = "conflict", "remote change prevents overwrite"
    Check result("Values")("REQ_TIT") = "remote-title", "conflict includes current remote values"

    Trace tracePath, "partial save / lost second response"
    Set form = WebLoadExisting(session, "RQ9102")
    values.RemoveAll
    values.Add "REQ_TIT", "partial-title"
    values.Add "REQ_RMK", "부분 저장 내용 % + &"
    values.Add "ACT_TIME", "120"
    Set result = WebPostUpdate(session, form, values)
    Check result("Outcome") = "partial" And Not result("Success"), "request saved, processing not saved"
    Check result("RequestNo") = "RQ9102" And result("Values")("REQ_TIT") = "partial-title", "partial preserves exact request and confirmed values"
    Check result("Values")("ACT_TIME") = "30", "partial reports original processing value"

    Trace tracePath, "processing persisted despite lost response"
    Set form = WebLoadExisting(session, "RQ9103")
    values.RemoveAll
    values.Add "ACT_TIME", "60"
    Set result = WebPostUpdate(session, form, values)
    Check result("Success") And result("Values")("ACT_TIME") = "60", "GET confirms lost processing POST without retry"

    Trace tracePath, "invalid identity / non-project refused"
    On Error Resume Next
    Set form = WebLoadExisting(session, "RQ9101&other=1")
    invalidError = Err.Number
    Err.Clear
    On Error GoTo 0
    Check invalidError <> 0, "invalid request number rejected"
    invalidError = 0
    On Error Resume Next
    Set form = WebLoadExisting(session, "RQ9199")
    invalidError = Err.Number
    Err.Clear
    On Error GoTo 0
    Check invalidError <> 0, "non-project detail refused"
End Sub

Public Function RunWebTransportRegression(Optional ByVal fixtureBaseUrl As String = "", Optional ByVal tracePath As String = "") As String
    Dim fields As Object, parsed As Object, session As Object, form As Object, values As Object, result As Object
    Dim html As String, encrypted As String, again As String, invalidError As Long
    On Error GoTo Failed
    Trace tracePath, "encoding"
    Check WebFormEncode("") = "", "empty encoding"
    Check WebFormEncode("a b+c&d=e%") = "a+b%2Bc%26d%3De%25", "form escaping"
    Check WebFormEncode("한글") = "%ED%95%9C%EA%B8%80", "UTF-8 encoding"
    Check WebFormEncode(vbCrLf) = "%0D%0A", "newlines"
    Set fields = CreateObject("Scripting.Dictionary")
    fields.Add "Login1$UserName", "a+b&c"
    fields.Add "ACT_TIME", ""
    Check WebFormEncode(fields) = "Login1%24UserName=a%2Bb%26c&ACT_TIME=", "form keys / explicit blank time"

    Trace tracePath, "parser"
    html = "<form method='post' action='./c0210_in.aspx'><input type='hidden' name='__VIEWSTATE' value='v+1&amp;x'>" & _
        "<input name='REQ_TIT' id='REQ_TIT' value='제목 &quot;A&quot; &amp; B'><textarea name='REQ_RMK' id='REQ_RMK'>첫 줄" & vbLf & "둘째 &amp; 셋째</textarea>" & _
        "<span id='outer'><span id='REQ_NO'>RQ123</span></span><select name='REQ_UID' id='REQ_UID'><option value='u1'>동명</option><option selected value='u2'>동명</option></select>" & _
        "<select name='GUBUN'><option value=''>선택</option><option selected value='PJT'>프로젝트</option><option disabled value='BAD'>사용 안 함</option></select>" & _
        "<input type='image' name='btnSave'><input disabled name='disabledField' value='x'><input name='ACT_TIME' value='30'></form>"
    Set parsed = WebParseRegistrationHtml(html, "http://127.0.0.1/support/C0/c0210_in.aspx")
    Check parsed("Fields")("__VIEWSTATE") = "v+1&x", "hidden state HTML decoding"
    Check parsed("Fields")("REQ_TIT") = "제목 ""A"" & B", "input HTML decoding"
    Check parsed("Fields")("REQ_UID") = "u2", "selected code"
    Check parsed("Options")("REQ_UID").Count = 2, "duplicate labels retain distinct IDs"
    Check parsed("Options")("GUBUN").Exists("") And Not parsed("Options")("GUBUN").Exists("BAD"), "blank and disabled options"
    Check Not parsed("Fields").Exists("disabledField") And Not parsed("Fields").Exists("btnSave"), "successful form controls"
    Check parsed("Buttons").Exists("btnSave"), "image button metadata"
    Check parsed("Action") = "http://127.0.0.1/support/C0/c0210_in.aspx", "relative form action"
    Check parsed("RequestNo") = "RQ123", "preallocated request number parsed only as metadata"
    Check Replace(parsed("Fields")("REQ_RMK"), vbCrLf, vbLf) = "첫 줄" & vbLf & "둘째 & 셋째", "textarea decoding"
    Set parsed = WebParseRegistrationHtml("<form><textarea name='blank'>" & vbCrLf & "</textarea><textarea name='leading'>" & vbCrLf & vbCrLf & "text</textarea></form>")
    Check parsed("Fields")("blank") = "", "ASP.NET empty textarea initial CRLF"
    Check parsed("Fields")("leading") = vbCrLf & "text", "intentional leading newline survives"
    On Error Resume Next
    Set parsed = WebParseRegistrationHtml("<form action='http://other.invalid/save'></form>", "http://127.0.0.1/support/")
    invalidError = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check invalidError <> 0, "cross-origin form rejected"

    Trace tracePath, "DPAPI"
    Check WebProtectSecret("") = "" And WebUnprotectSecret("") = "", "empty protected setting"
    encrypted = WebProtectSecret("fixture-only 한글 +& secret")
    again = WebProtectSecret("fixture-only 한글 +& secret")
    Check Left$(encrypted, 9) = "dpapi:v1:", "DPAPI format"
    Check InStr(encrypted, "fixture-only") = 0 And encrypted <> again, "DPAPI randomized ciphertext"
    Check WebUnprotectSecret(encrypted) = "fixture-only 한글 +& secret", "DPAPI current-user roundtrip"
    On Error Resume Next
    again = WebUnprotectSecret("unprotected-text")
    invalidError = Err.Number
    Err.Clear
    On Error GoTo Failed
    Check invalidError <> 0, "plaintext rejected by credential reader"

    If Len(fixtureBaseUrl) > 0 Then
        Trace tracePath, "login"
        Set session = WebOpenSession(fixtureBaseUrl, "fixture-user", "fixture-password")
        Check session("Authenticated") And session("Cookies").Count >= 2, "login / 302 / cookies"
        Check Not session.Exists("Password"), "password absent from session"
        Trace tracePath, "registration GET"
        Set form = WebLoadRegistration(session)
        Set values = CreateObject("Scripting.Dictionary")
        values.Add "GUBUN", "PJT"
        values.Add "REQ_UID", "u1"
        values.Add "REQ_TIT", "[소분류] 제목 + & 검증"
        values.Add "REQ_RMK", "첫째 줄" & vbCrLf & "둘째 + & = %"
        values.Add "STS", "1"
        values.Add "PGM_NM", "소분류"
        values.Add "PGM_ID", "P+A&B"
        values.Add "ACT_TIME", ""
        Trace tracePath, "registration POST + readback"
        Set result = WebPostRegistration(session, form, values)
        Check result("Success") And result("Outcome") = "success", "persisted detail readback: " & result("Message")
        Check Left$(result("RequestNo"), 2) = "RQ", "persisted ID"
        Trace tracePath, "stale option"
        Set form = WebLoadRegistration(session)
        values("REQ_UID") = "removed-user"
        Set result = WebPostRegistration(session, form, values)
        Check Not result("Success") And result("Outcome") = "failed", "stale options rejected before POST"
        Check result("RequestNo") = form("RequestNo"), "prevalidation retains journal candidate"
        values("REQ_UID") = "u1"
        values("REQ_TIT") = "fixture-connection-drop"
        Trace tracePath, "uncertain POST"
        Set result = WebPostRegistration(session, form, values)
        Check Not result("Success") And result("Outcome") = "unknown", "lost POST response is unknown, never retried"
        Check result("RequestNo") = form("RequestNo"), "uncertain outcome retains journal candidate"
        Set result = WebPostRegistration(session, form, values)
        Check result("Outcome") = "failed", "consumed form cannot be posted again"
        CheckExistingUpdates session, tracePath
    End If
    Trace tracePath, "passed"
    RunWebTransportRegression = "PASS: encoding, HTML forms/options, DPAPI" & IIf(Len(fixtureBaseUrl) > 0, ", loopback registration/update/conflict/partial/readback/no retry", "")
    Exit Function
Failed:
    RunWebTransportRegression = "FAIL: " & CStr(Err.Number) & " / " & Err.Description
    Trace tracePath, RunWebTransportRegression
End Function
