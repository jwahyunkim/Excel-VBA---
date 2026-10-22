Attribute VB_Name = "modWebTransport"
Option Explicit
' 2026-10-22 김좌현 수정 : 웹 세션·폼 전송·저장 후 상세 조회 확인과 Windows 사용자별 비밀번호 암호화

Private Type WebDataBlob
    cbData As Long
#If VBA7 Then
    pbData As LongPtr
#Else
    pbData As Long
#End If
End Type

#If VBA7 Then
Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" (ByRef dataIn As WebDataBlob, ByVal description As LongPtr, ByVal entropy As LongPtr, ByVal reserved As LongPtr, ByVal prompt As LongPtr, ByVal flags As Long, ByRef dataOut As WebDataBlob) As Long
Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" (ByRef dataIn As WebDataBlob, ByVal description As LongPtr, ByVal entropy As LongPtr, ByVal reserved As LongPtr, ByVal prompt As LongPtr, ByVal flags As Long, ByRef dataOut As WebDataBlob) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Sub WebCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)
#Else
Private Declare Function CryptProtectData Lib "crypt32.dll" (ByRef dataIn As WebDataBlob, ByVal description As Long, ByVal entropy As Long, ByVal reserved As Long, ByVal prompt As Long, ByVal flags As Long, ByRef dataOut As WebDataBlob) As Long
Private Declare Function CryptUnprotectData Lib "crypt32.dll" (ByRef dataIn As WebDataBlob, ByVal description As Long, ByVal entropy As Long, ByVal reserved As Long, ByVal prompt As Long, ByVal flags As Long, ByRef dataOut As WebDataBlob) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Sub WebCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Private Const WEB_ERROR As Long = vbObjectError + 1940
Private Const WEB_SECRET_PREFIX As String = "dpapi:v1:"

Private Function WebDictionary() As Object
    Set WebDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function WebString(ByVal value As Variant) As String
    If Not IsNull(value) And Not IsEmpty(value) Then WebString = CStr(value)
End Function

Private Function WebUtf8Bytes(ByVal value As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3 ' ADODB.Stream writes a UTF-8 BOM which is not form data.
    WebUtf8Bytes = stream.Read
    stream.Close
End Function

Private Function WebEncodeValue(ByVal value As String) As String
    Dim bytes As Variant, i As Long, b As Long, result As String
    If Len(value) = 0 Then Exit Function
    bytes = WebUtf8Bytes(value)
    For i = LBound(bytes) To UBound(bytes)
        b = bytes(i)
        Select Case b
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & Chr$(b)
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Right$("0" & Hex$(b), 2)
        End Select
    Next i
    WebEncodeValue = result
End Function

' A string encodes one value. A Dictionary encodes the complete form body.
Public Function WebFormEncode(ByVal value As Variant) As String
    Dim key As Variant, result As String
    If IsObject(value) Then
        For Each key In value.Keys
            If Len(result) > 0 Then result = result & "&"
            result = result & WebEncodeValue(CStr(key)) & "=" & WebEncodeValue(WebString(value(key)))
        Next key
        WebFormEncode = result
    Else
        WebFormEncode = WebEncodeValue(WebString(value))
    End If
End Function

Private Function WebMatch(ByVal value As String, ByVal pattern As String) As String
    Dim expression As Object, matches As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = pattern
    expression.IgnoreCase = True
    expression.MultiLine = True
    Set matches = expression.Execute(value)
    If matches.Count > 0 Then WebMatch = matches(0).SubMatches(0)
End Function

Private Function WebOrigin(ByVal url As String) As String
    WebOrigin = LCase$(WebMatch(url, "^(https?://[^/\?#]+)"))
    If Len(WebOrigin) = 0 Or InStr(WebOrigin, "@") > 0 Then _
        Err.Raise WEB_ERROR, "modWebTransport", "웹 주소가 올바르지 않습니다."
End Function

Private Function WebResolveUrl(ByVal currentUrl As String, ByVal reference As String) As String
    Dim origin As String, path As String, suffix As String, parts As Variant
    Dim stack As Collection, part As Variant, i As Long, at As Long
    origin = WebOrigin(currentUrl)
    reference = Trim$(reference)
    If InStr(reference, "#") > 0 Then reference = Left$(reference, InStr(reference, "#") - 1)
    If Len(reference) = 0 Then
        WebResolveUrl = Split(currentUrl, "#")(0)
        Exit Function
    End If
    If LCase$(Left$(reference, 7)) = "http://" Or LCase$(Left$(reference, 8)) = "https://" Then
        If WebOrigin(reference) <> origin Then Err.Raise WEB_ERROR, "modWebTransport", "다른 사이트로 이동하는 웹 응답을 중단했습니다."
        WebResolveUrl = reference
        Exit Function
    End If
    If Left$(reference, 2) = "//" Or InStr(reference, ":") > 0 Then _
        Err.Raise WEB_ERROR, "modWebTransport", "지원하지 않는 웹 이동 주소입니다."
    path = Mid$(currentUrl, Len(WebOrigin(currentUrl)) + 1)
    path = Split(Split(path, "#")(0), "?")(0)
    If Len(path) = 0 Then path = "/"
    If Left$(reference, 1) = "?" Then
        WebResolveUrl = origin & path & reference
        Exit Function
    End If
    If Left$(reference, 1) = "/" Then
        path = reference
    Else
        path = Left$(path, InStrRev(path, "/")) & reference
    End If
    at = InStr(path, "?")
    If at > 0 Then
        suffix = Mid$(path, at)
        path = Left$(path, at - 1)
    End If
    Set stack = New Collection
    parts = Split(path, "/")
    For Each part In parts
        If part = ".." Then
            If stack.Count > 0 Then stack.Remove stack.Count
        ElseIf part <> "." And Len(part) > 0 Then
            stack.Add CStr(part)
        End If
    Next part
    path = "/"
    For i = 1 To stack.Count
        If i > 1 Then path = path & "/"
        path = path & stack(i)
    Next i
    If Right$(reference, 1) = "/" And Right$(path, 1) <> "/" Then path = path & "/"
    WebResolveUrl = origin & path & suffix
End Function

Private Function WebMatches(ByVal value As String, ByVal pattern As String) As Object
    Dim expression As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = pattern
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = True
    Set WebMatches = expression.Execute(value)
End Function

Private Function WebHtmlDecode(ByVal value As String) As String
    Dim matches As Object, item As Object, decoded As String, token As String
    Dim code As Long, i As Long
    Set matches = WebMatches(value, "&(#x[0-9a-f]+|#[0-9]+|amp|quot|apos|lt|gt|nbsp);")
    For i = matches.Count - 1 To 0 Step -1
        Set item = matches(i)
        token = LCase$(item.SubMatches(0))
        Select Case token
            Case "amp": decoded = "&"
            Case "quot": decoded = """"
            Case "apos": decoded = "'"
            Case "lt": decoded = "<"
            Case "gt": decoded = ">"
            Case "nbsp": decoded = ChrW$(160)
            Case Else
                If Left$(token, 2) = "#x" Then
                    code = CLng("&H" & Mid$(token, 3))
                Else
                    code = CLng(Mid$(token, 2))
                End If
                If code < 0 Or code > &H10FFFF Then Err.Raise WEB_ERROR, , "웹 문자 인코딩을 확인하세요."
                If code > &HFFFF& Then
                    code = code - &H10000
                    decoded = ChrW$(&HD800& + (code \ 1024)) & ChrW$(&HDC00& + (code Mod 1024))
                Else
                    decoded = ChrW$(code)
                End If
        End Select
        value = Left$(value, item.FirstIndex) & decoded & Mid$(value, item.FirstIndex + item.length + 1)
    Next i
    WebHtmlDecode = value
End Function

Private Function WebAttributes(ByVal tag As String) As Object
    Dim attributes As Object, item As Object, name As String, value As String
    Set attributes = WebDictionary()
    attributes.CompareMode = 1
    For Each item In WebMatches(tag, "(?:^|\s)([a-z_:][-a-z0-9_:.]*)(?:\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s=<>`]+)))?")
        name = CStr(item.SubMatches(0))
        value = CStr(item.SubMatches(1))
        If Len(value) = 0 Then value = CStr(item.SubMatches(2))
        If Len(value) = 0 Then value = CStr(item.SubMatches(3))
        attributes(name) = WebHtmlDecode(value)
    Next item
    Set WebAttributes = attributes
End Function

Private Function WebAttribute(ByVal attributes As Object, ByVal name As String) As String
    If attributes.Exists(name) Then WebAttribute = CStr(attributes(name))
End Function

Private Function WebPlainText(ByVal html As String) As String
    Dim expression As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "<[^>]*>"
    WebPlainText = WebHtmlDecode(expression.Replace(html, ""))
End Function

Private Function WebTextareaText(ByVal html As String) As String
    ' HTML textarea parsing ignores exactly one initial markup newline.
    ' ASP.NET renders that CRLF even when the stored text is empty.
    If Left$(html, 2) = vbCrLf Then
        html = Mid$(html, 3)
    ElseIf Left$(html, 1) = vbCr Or Left$(html, 1) = vbLf Then
        html = Mid$(html, 2)
    End If
    WebTextareaText = WebHtmlDecode(html)
End Function

Public Function WebParseRegistrationHtml(ByVal html As String, Optional ByVal url As String = "") As Object
    Dim result As Object, fields As Object, options As Object, buttons As Object, display As Object
    Dim forms As Object, element As Object, optionItem As Object, choices As Object, attributes As Object
    Dim optionAttributes As Object, expression As Object, name As String, kind As String
    Dim action As String, id As String, body As String, value As String, firstValue As String, selected As Boolean, spanId As Variant
    Set result = WebDictionary()
    Set fields = WebDictionary()
    Set options = WebDictionary()
    Set buttons = WebDictionary()
    Set display = WebDictionary()
    ' Tokenize known WebForms controls as text. Never instantiate MSHTML,
    ' navigate a browser, load an image, or execute a page script/event.
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.IgnoreCase = True
    expression.Pattern = "<(script|style|iframe|object|embed|applet)\b[^>]*>[\s\S]*?</\1\s*>"
    body = expression.Replace(html, "")
    Set forms = WebMatches(body, "<form\b([^>]*)>([\s\S]*?)</form\s*>")
    If forms.Count > 0 Then
        Set attributes = WebAttributes(CStr(forms(0).SubMatches(0)))
        action = WebAttribute(attributes, "action")
        body = CStr(forms(0).SubMatches(1))
        For Each element In WebMatches(body, "<input\b((?:[^""'>]|""[^""]*""|'[^']*')*)>")
            Set attributes = WebAttributes(CStr(element.SubMatches(0)))
            name = WebAttribute(attributes, "name")
            kind = LCase$(WebAttribute(attributes, "type"))
            value = WebAttribute(attributes, "value")
            id = WebAttribute(attributes, "id")
            If Len(id) > 0 Then display(id) = value
            If Len(name) > 0 And Not attributes.Exists("disabled") Then
                Select Case kind
                    Case "submit", "image", "button"
                        buttons(name) = value
                    Case "file", "reset"
                    Case "checkbox", "radio"
                        If attributes.Exists("checked") Then fields(name) = value
                    Case Else
                        fields(name) = value
                End Select
            End If
        Next element
        For Each element In WebMatches(body, "<textarea\b([^>]*)>([\s\S]*?)</textarea\s*>")
            Set attributes = WebAttributes(CStr(element.SubMatches(0)))
            name = WebAttribute(attributes, "name")
            id = WebAttribute(attributes, "id")
            value = WebTextareaText(CStr(element.SubMatches(1)))
            If Len(name) > 0 And Not attributes.Exists("disabled") Then fields(name) = value
            If Len(id) > 0 Then display(id) = value
        Next element
        For Each element In WebMatches(body, "<select\b([^>]*)>([\s\S]*?)</select\s*>")
            Set attributes = WebAttributes(CStr(element.SubMatches(0)))
            name = WebAttribute(attributes, "name")
            id = WebAttribute(attributes, "id")
            If Len(name) > 0 And Not attributes.Exists("disabled") Then
                Set choices = WebDictionary()
                value = ""
                firstValue = ""
                selected = False
                For Each optionItem In WebMatches(CStr(element.SubMatches(1)), "<option\b([^>]*)>([\s\S]*?)</option\s*>")
                    Set optionAttributes = WebAttributes(CStr(optionItem.SubMatches(0)))
                    If Not optionAttributes.Exists("disabled") Then
                        If choices.Count = 0 Then firstValue = WebAttribute(optionAttributes, "value")
                        choices(WebAttribute(optionAttributes, "value")) = Trim$(WebPlainText(CStr(optionItem.SubMatches(1))))
                        If optionAttributes.Exists("selected") Then
                            value = WebAttribute(optionAttributes, "value")
                            selected = True
                        End If
                    End If
                Next optionItem
                If Not selected Then value = firstValue
                fields(name) = value
                If Len(id) > 0 Then display(id) = value
                options.Add name, choices
            End If
        Next element
    End If
    For Each spanId In Array("REQ_NO", "REQ_UID")
        For Each element In WebMatches(body, "<span\b[^>]*\bid\s*=\s*[""']" & CStr(spanId) & "[""'][^>]*>([\s\S]*?)</span\s*>")
            display(CStr(spanId)) = Trim$(WebPlainText(CStr(element.SubMatches(0))))
        Next element
    Next spanId
    result.Add "Url", url
    If Len(url) > 0 Then action = WebResolveUrl(url, action)
    result.Add "Action", action
    result.Add "Html", html
    result.Add "Fields", fields
    result.Add "Options", options
    result.Add "Buttons", buttons
    result.Add "DisplayValues", display
    result.Add "RequestNo", WebAttribute(display, "REQ_NO")
    Set WebParseRegistrationHtml = result
End Function

Private Sub WebCaptureCookies(ByVal session As Object, ByVal headers As String, ByVal url As String)
    Dim line As Variant, parts As Variant, pair As String, name As String, value As String
    Dim at As Long, path As String, cookiePath As String, cookieKey As String, part As Variant
    Dim expired As Boolean, secure As Boolean, cookie As Object
    path = Mid$(url, Len(WebOrigin(url)) + 1)
    path = Split(path, "?")(0)
    If Len(path) = 0 Then path = "/"
    For Each line In Split(Replace(headers, vbCr, ""), vbLf)
        If LCase$(Left$(CStr(line), 11)) = "set-cookie:" Then
            parts = Split(Trim$(Mid$(CStr(line), 12)), ";")
            pair = Trim$(parts(0))
            at = InStr(pair, "=")
            If at > 1 Then
                name = Left$(pair, at - 1)
                value = Mid$(pair, at + 1)
                cookiePath = Left$(path, InStrRev(path, "/"))
                expired = False
                secure = False
                For Each part In parts
                    part = Trim$(CStr(part))
                    If LCase$(Left$(part, 5)) = "path=" Then cookiePath = Mid$(part, 6)
                    If LCase$(part) = "max-age=0" Then expired = True
                    If LCase$(part) = "secure" Then secure = True
                Next part
                If Len(cookiePath) = 0 Then cookiePath = "/"
                cookieKey = cookiePath & vbTab & name
                If expired Or Len(value) = 0 Then
                    If session("Cookies").Exists(cookieKey) Then session("Cookies").Remove cookieKey
                Else
                    Set cookie = WebDictionary()
                    cookie.Add "Name", name
                    cookie.Add "Value", value
                    cookie.Add "Path", cookiePath
                    cookie.Add "Secure", secure
                    If session("Cookies").Exists(cookieKey) Then session("Cookies").Remove cookieKey
                    session("Cookies").Add cookieKey, cookie
                End If
            End If
        End If
    Next line
End Sub

Private Function WebCookieHeader(ByVal session As Object, ByVal url As String) As String
    Dim key As Variant, cookie As Object, path As String, result As String, allowed As Boolean
    path = Mid$(url, Len(WebOrigin(url)) + 1)
    path = Split(path, "?")(0)
    For Each key In session("Cookies").Keys
        Set cookie = session("Cookies")(key)
        allowed = (Left$(path, Len(cookie("Path"))) = cookie("Path"))
        If cookie("Secure") And LCase$(Left$(url, 8)) <> "https://" Then allowed = False
        If allowed Then
            If Len(result) > 0 Then result = result & "; "
            result = result & cookie("Name") & "=" & cookie("Value")
        End If
    Next key
    WebCookieHeader = result
End Function

Private Function WebResponseText(ByVal request As Object) As String
    Dim stream As Object, charset As String, contentType As String, bytes As Variant, byteCount As Long
    bytes = request.ResponseBody
    On Error Resume Next
    byteCount = UBound(bytes) - LBound(bytes) + 1
    On Error GoTo 0
    If byteCount <= 0 Then Exit Function
    On Error Resume Next
    contentType = request.GetResponseHeader("Content-Type")
    On Error GoTo 0
    charset = WebMatch(contentType, "charset\s*=\s*[""']?([^;""'\s]+)")
    If Len(charset) = 0 Then charset = "utf-8"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = charset
    WebResponseText = stream.ReadText
    stream.Close
End Function

Private Function WebHttpRequest(ByVal session As Object, ByVal method As String, ByVal url As String, _
                                Optional ByVal body As String = "", Optional ByVal follow As Boolean = True) As Object
    Dim request As Object, response As Object, hop As Long, status As Long
    Dim location As String, cookie As String, headers As String
    For hop = 0 To 8
        If WebOrigin(url) <> WebOrigin(CStr(session("BaseUrl"))) Then _
            Err.Raise WEB_ERROR, "modWebTransport", "설정된 웹사이트 밖으로 전송할 수 없습니다."
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts 15000, 15000, CLng(session("TimeoutMs")), CLng(session("TimeoutMs"))
        request.Open method, url, False
        request.Option(6) = False ' Inspect redirects and cookies explicitly.
        request.SetRequestHeader "User-Agent", "REO-Excel-WebUpload/1.0"
        request.SetRequestHeader "Accept", "text/html,application/xhtml+xml"
        cookie = WebCookieHeader(session, url)
        If Len(cookie) > 0 Then request.SetRequestHeader "Cookie", cookie
        If method = "POST" Then
            request.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
            request.SetRequestHeader "Referer", CStr(session("LastUrl"))
            request.Send body
        Else
            request.Send
        End If
        status = request.Status
        headers = request.GetAllResponseHeaders
        WebCaptureCookies session, headers, url
        location = ""
        On Error Resume Next
        location = request.GetResponseHeader("Location")
        On Error GoTo 0
        session("LastUrl") = url
        If follow And (status = 301 Or status = 302 Or status = 303) And Len(location) > 0 Then
            url = WebResolveUrl(url, location)
            method = "GET"
            body = ""
        ElseIf follow And (status = 307 Or status = 308) And method = "GET" And Len(location) > 0 Then
            url = WebResolveUrl(url, location)
        Else
            Set response = WebDictionary()
            response.Add "Status", status
            response.Add "Html", WebResponseText(request)
            response.Add "Url", url
            response.Add "Location", location
            Set WebHttpRequest = response
            Exit Function
        End If
    Next hop
    Err.Raise WEB_ERROR, "modWebTransport", "웹사이트 이동 횟수가 제한을 초과했습니다."
End Function

Public Function WebOpenSession(ByVal baseUrl As String, ByVal userId As String, ByVal password As String) As Object
    Dim session As Object, response As Object, form As Object, fields As Object
    Dim cookies As Object, errorNumber As Long
    On Error GoTo Failed
    baseUrl = Trim$(baseUrl)
    If Right$(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    If Len(Trim$(userId)) = 0 Or Len(password) = 0 Then Err.Raise WEB_ERROR
    Set session = WebDictionary()
    Set cookies = WebDictionary()
    session.Add "BaseUrl", baseUrl
    session.Add "Cookies", cookies
    session.Add "TimeoutMs", 30000
    session.Add "LastUrl", baseUrl
    session.Add "UserId", userId
    session.Add "Authenticated", False
    Set response = WebHttpRequest(session, "GET", baseUrl)
    Set form = WebParseRegistrationHtml(CStr(response("Html")), CStr(response("Url")))
    Set fields = form("Fields")
    If Not fields.Exists("Login1$UserName") Or Not fields.Exists("Login1$Password") Then Err.Raise WEB_ERROR
    fields("Login1$UserName") = userId
    fields("Login1$Password") = password
    fields("Login1$LoginButton") = "로그인"
    If form("Buttons").Exists("Login1$LoginButton") Then fields("Login1$LoginButton") = form("Buttons")("Login1$LoginButton")
    Set response = WebHttpRequest(session, "POST", CStr(form("Action")), WebFormEncode(fields))
    fields("Login1$Password") = ""
    password = ""
    ' /support/ can show login.aspx even after authentication; use Default.aspx.
    Set response = WebHttpRequest(session, "GET", WebResolveUrl(baseUrl, "Default.aspx"))
    Set form = WebParseRegistrationHtml(CStr(response("Html")), CStr(response("Url")))
    If form("Fields").Exists("Login1$Password") Or CLng(response("Status")) <> 200 Then Err.Raise WEB_ERROR
    session("Authenticated") = True
    Set WebOpenSession = session
    Exit Function
Failed:
    password = ""
    If Not fields Is Nothing Then
        If fields.Exists("Login1$Password") Then fields("Login1$Password") = ""
    End If
    Err.Raise WEB_ERROR, "modWebTransport", "웹 로그인에 실패했습니다. 주소·계정·네트워크를 확인하세요."
End Function

Public Function WebLoadRegistration(ByVal session As Object) As Object
    Dim response As Object, form As Object
    Set response = WebHttpRequest(session, "GET", WebResolveUrl(CStr(session("BaseUrl")), "C0/c0210_in.aspx"))
    If CLng(response("Status")) <> 200 Then Err.Raise WEB_ERROR, "modWebTransport", "프로젝트 실적 등록 화면을 읽지 못했습니다."
    Set form = WebParseRegistrationHtml(CStr(response("Html")), CStr(response("Url")))
    If Not form("Fields").Exists("REQ_TIT") Or Not form("Fields").Exists("REQ_RMK") Or Not form("Options").Exists("GUBUN") Then _
        Err.Raise WEB_ERROR, "modWebTransport", "로그인 세션 또는 프로젝트 실적 등록 화면을 확인하세요."
    form.Add "LoadedAt", Now
    form.Add "Consumed", False
    Set WebLoadRegistration = form
End Function

Private Function WebNormalizeText(ByVal value As String) As String
    WebNormalizeText = Replace(Replace(value, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function WebReadbackMatches(ByVal form As Object, ByVal values As Object, ByVal requestNo As String) As Boolean
    Dim key As Variant, actual As Object
    Set actual = form("DisplayValues")
    If CStr(form("RequestNo")) <> requestNo Then Exit Function
    If Not actual.Exists("REQ_TIT") Or Not actual.Exists("REQ_RMK") Then Exit Function
    If WebNormalizeText(CStr(actual("REQ_TIT"))) <> WebNormalizeText(CStr(values("REQ_TIT"))) Then Exit Function
    If WebNormalizeText(CStr(actual("REQ_RMK"))) <> WebNormalizeText(CStr(values("REQ_RMK"))) Then Exit Function
    For Each key In Array("REQ_UID", "GUBUN", "STS", "PGM_MD", "PGM_NM", "PGM_ID", "ACT_GUBUN", "ACT_UID", "ACT_DT_FR", "ACT_DT", "ACT_TIME")
        If values.Exists(CStr(key)) Then
            If Not actual.Exists(CStr(key)) Then Exit Function
            If CStr(actual(CStr(key))) <> CStr(values(CStr(key))) Then Exit Function
        End If
    Next key
    WebReadbackMatches = True
End Function

Private Function WebOutcome(ByVal outcome As String, ByVal message As String, Optional ByVal requestNo As String = "") As Object
    Dim result As Object
    Set result = WebDictionary()
    result.Add "Success", (outcome = "success")
    result.Add "Outcome", outcome
    result.Add "RequestNo", requestNo
    result.Add "Message", message
    Set WebOutcome = result
End Function

Public Function WebPostRegistration(ByVal session As Object, ByVal form As Object, ByVal values As Object) As Object
    Dim fresh As Object, fields As Object, options As Object, response As Object, readback As Object
    Dim key As Variant, candidates As Object, requestNo As Variant, posted As Boolean, message As String, candidateNo As String
    On Error GoTo Failed
    ' The caller journals this exact form's candidate number before POST.
    ' Fetching another registration here would invalidate that journal entry.
    Set fresh = form
    If fresh Is Nothing Then Err.Raise WEB_ERROR, , "등록 화면을 먼저 불러오세요."
    candidateNo = CStr(fresh("RequestNo"))
    If Not fresh.Exists("Consumed") Then Err.Raise WEB_ERROR, , "등록 화면을 먼저 불러오세요."
    If CBool(fresh("Consumed")) Then Err.Raise WEB_ERROR, , "이미 전송한 화면입니다. 결과를 확인한 뒤 등록 화면을 다시 불러오세요."
    If DateDiff("n", CDate(fresh("LoadedAt")), Now) > 10 Then Err.Raise WEB_ERROR, , "등록 화면이 오래되었습니다. 다시 불러오세요."
    Set fields = WebDictionary()
    For Each key In fresh("Fields").Keys
        fields.Add CStr(key), fresh("Fields")(key)
    Next key
    Set options = fresh("Options")
    If Not values.Exists("REQ_TIT") Or Not values.Exists("REQ_RMK") Then Err.Raise WEB_ERROR, , "제목과 내용을 입력하세요."
    If Len(Trim$(CStr(values("REQ_TIT")))) = 0 Or Len(Trim$(CStr(values("REQ_RMK")))) = 0 Then _
        Err.Raise WEB_ERROR, , "제목과 내용을 입력하세요."
    If Len(CStr(values("REQ_TIT"))) > 300 Then Err.Raise WEB_ERROR, , "제목은 300자 이내로 입력하세요."
    For Each key In values.Keys
        If Not fields.Exists(CStr(key)) Then Err.Raise WEB_ERROR, , "현재 등록 화면에 없는 입력 항목입니다: " & CStr(key)
        If Left$(CStr(key), 2) = "__" Then Err.Raise WEB_ERROR, , "웹 상태값을 수정할 수 없습니다."
        If options.Exists(CStr(key)) Then
            If Not options(CStr(key)).Exists(CStr(values(key))) Then _
                Err.Raise WEB_ERROR, , "웹 선택 목록이 변경되었습니다. 목록을 새로고침하세요: " & CStr(key)
        End If
        fields(CStr(key)) = CStr(values(key))
    Next key
    If fields.Exists("PGM_NM") Then
        If Len(CStr(fields("PGM_NM"))) > 100 Then Err.Raise WEB_ERROR, , "프로그램명은 100자 이내로 입력하세요."
    End If
    If fields.Exists("PGM_ID") Then
        If Len(CStr(fields("PGM_ID"))) > 100 Then Err.Raise WEB_ERROR, , "프로그램ID는 100자 이내로 입력하세요."
    End If
    fields("btnSave.x") = "1"
    fields("btnSave.y") = "1"
    session("LastUrl") = fresh("Url")
    Set candidates = WebDictionary()
    If Len(CStr(fresh("RequestNo"))) > 0 Then candidates(CStr(fresh("RequestNo"))) = True
    fresh("Consumed") = True
    posted = True
    Set response = WebHttpRequest(session, "POST", CStr(fresh("Action")), WebFormEncode(fields))
    Set readback = WebParseRegistrationHtml(CStr(response("Html")), CStr(response("Url")))
    If Len(CStr(readback("RequestNo"))) > 0 Then candidates(CStr(readback("RequestNo"))) = True
    For Each requestNo In candidates.Keys
        If Len(WebMatch(CStr(requestNo), "^(RQ[0-9]+)$")) > 0 Then
            Set response = WebHttpRequest(session, "GET", WebResolveUrl(CStr(session("BaseUrl")), "C0/c0200_02.aspx?REQ_NO=" & WebEncodeValue(CStr(requestNo))))
            If CLng(response("Status")) = 200 Then
                Set readback = WebParseRegistrationHtml(CStr(response("Html")), CStr(response("Url")))
                If WebReadbackMatches(readback, values, CStr(requestNo)) Then
                    Set WebPostRegistration = WebOutcome("success", "웹 상세 화면에서 등록 내용을 확인했습니다.", CStr(requestNo))
                    Exit Function
                End If
            End If
        End If
    Next requestNo
    Set WebPostRegistration = WebOutcome("unknown", "등록 결과를 확인하지 못했습니다. 웹 요청내역에서 확인한 뒤 다시 처리하세요.", candidateNo)
    Exit Function
Failed:
    If posted Then
        Set WebPostRegistration = WebOutcome("unknown", "전송 후 등록 결과를 확인하지 못했습니다. 자동으로 재전송하지 않습니다.", candidateNo)
    Else
        message = Err.Description
        If Len(message) = 0 Then message = "등록 전 입력값 또는 웹 연결을 확인하세요."
        Set WebPostRegistration = WebOutcome("failed", message, candidateNo)
    End If
End Function

Private Function WebBase64Encode(ByVal bytes As Variant) As String
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("value")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    WebBase64Encode = Replace(Replace(node.Text, vbCr, ""), vbLf, "")
End Function

Private Function WebBase64Decode(ByVal value As String) As Variant
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("value")
    node.DataType = "bin.base64"
    node.Text = value
    WebBase64Decode = node.nodeTypedValue
End Function

Public Function WebProtectSecret(ByVal value As String) As String
    Dim bytes() As Byte, protected() As Byte, dataIn As WebDataBlob, dataOut As WebDataBlob
    On Error GoTo Failed
    If Len(value) = 0 Then Exit Function
    bytes = WebUtf8Bytes(value)
    dataIn.cbData = UBound(bytes) - LBound(bytes) + 1
    dataIn.pbData = VarPtr(bytes(LBound(bytes)))
    If CryptProtectData(dataIn, 0, 0, 0, 0, 1, dataOut) = 0 Then Err.Raise WEB_ERROR
    ReDim protected(0 To dataOut.cbData - 1)
    WebCopyMemory protected(0), ByVal dataOut.pbData, dataOut.cbData
    LocalFree dataOut.pbData
    dataOut.pbData = 0
    WebProtectSecret = WEB_SECRET_PREFIX & WebBase64Encode(protected)
    value = ""
    Exit Function
Failed:
    If dataOut.pbData <> 0 Then LocalFree dataOut.pbData
    value = ""
    Err.Raise WEB_ERROR, "modWebTransport", "비밀번호를 현재 Windows 사용자용으로 암호화하지 못했습니다."
End Function

Public Function WebUnprotectSecret(ByVal value As String) As String
    Dim bytes() As Byte, clear() As Byte, dataIn As WebDataBlob, dataOut As WebDataBlob, stream As Object
    On Error GoTo Failed
    If Len(value) = 0 Then Exit Function
    If Left$(value, Len(WEB_SECRET_PREFIX)) <> WEB_SECRET_PREFIX Then Err.Raise WEB_ERROR
    bytes = WebBase64Decode(Mid$(value, Len(WEB_SECRET_PREFIX) + 1))
    dataIn.cbData = UBound(bytes) - LBound(bytes) + 1
    dataIn.pbData = VarPtr(bytes(LBound(bytes)))
    If CryptUnprotectData(dataIn, 0, 0, 0, 0, 1, dataOut) = 0 Then Err.Raise WEB_ERROR
    ReDim clear(0 To dataOut.cbData - 1)
    WebCopyMemory clear(0), ByVal dataOut.pbData, dataOut.cbData
    LocalFree dataOut.pbData
    dataOut.pbData = 0
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write clear
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    WebUnprotectSecret = stream.ReadText
    stream.Close
    Exit Function
Failed:
    If dataOut.pbData <> 0 Then LocalFree dataOut.pbData
    Err.Raise WEB_ERROR, "modWebTransport", "저장된 비밀번호를 읽지 못했습니다. 이 Windows 사용자로 비밀번호를 다시 설정하세요."
End Function
