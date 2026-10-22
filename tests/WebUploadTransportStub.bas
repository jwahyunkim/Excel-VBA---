Attribute VB_Name = "WebUploadTransportStub"
Option Explicit
' 2026-10-22 김좌현 수정 : 실제 HTTP 없이 신규·갱신·충돌·부분저장·이력 보존 검증
Public WebUploadMockOpenCount As Long
Public WebUploadMockLoadCount As Long
Public WebUploadMockPostCount As Long
Public WebUploadMockUpdateCount As Long
Public WebUploadMockReadCount As Long
Public WebUploadMockJournalChecks As Long
Public WebUploadMockPostedTitles As String
Private mMockOutcomes As Variant
Private mMockSequence As Long
Private mMockRequests As Object

Public Sub WebUploadMockReset(Optional ByVal outcomes As String = "success", Optional ByVal clearServer As Boolean = False)
    WebUploadMockOpenCount = 0: WebUploadMockLoadCount = 0: WebUploadMockPostCount = 0
    WebUploadMockUpdateCount = 0: WebUploadMockReadCount = 0: WebUploadMockJournalChecks = 0
    WebUploadMockPostedTitles = ""
    mMockOutcomes = Split(outcomes, ",")
    If mMockRequests Is Nothing Or clearServer Then Set mMockRequests = CreateObject("Scripting.Dictionary")
End Sub

Private Function MockClone(ByVal values As Object) As Object
    Dim result As Object, key As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In values.Keys
        result.Add CStr(key), values(key)
    Next key
    Set MockClone = result
End Function

Public Sub WebUploadMockSetServer(ByVal requestNo As String, ByVal values As Object)
    If mMockRequests Is Nothing Then Set mMockRequests = CreateObject("Scripting.Dictionary")
    Set mMockRequests(requestNo) = MockClone(values)
End Sub

Public Function WebOpenSession(ByVal baseUrl As String, ByVal userId As String, ByVal password As String) As Object
    Dim session As Object
    If Len(baseUrl) = 0 Or Len(userId) = 0 Or Len(password) = 0 Then Err.Raise vbObjectError + 7990, , "Missing mock credentials."
    WebUploadMockOpenCount = WebUploadMockOpenCount + 1
    Set session = CreateObject("Scripting.Dictionary")
    session.Add "BaseUrl", baseUrl
    session.Add "Authenticated", True
    Set WebOpenSession = session
End Function

Public Function WebLoadRegistration(ByVal session As Object) As Object
    Dim form As Object
    WebUploadMockLoadCount = WebUploadMockLoadCount + 1
    mMockSequence = mMockSequence + 1
    Set form = CreateObject("Scripting.Dictionary")
    form.Add "RequestNo", "RQMOCK" & Format$(mMockSequence, "000000")
    Set WebLoadRegistration = form
End Function

Public Function WebReadExistingValues(ByVal session As Object, ByVal requestNo As String) As Object
    WebUploadMockReadCount = WebUploadMockReadCount + 1
    If mMockRequests Is Nothing Then Err.Raise vbObjectError + 7998, , "Mock request not found."
    If Not mMockRequests.Exists(requestNo) Then Err.Raise vbObjectError + 7998, , "Mock request not found."
    Set WebReadExistingValues = MockClone(mMockRequests(requestNo))
End Function

Public Function WebLoadExisting(ByVal session As Object, ByVal requestNo As String) As Object
    Dim form As Object
    Set form = CreateObject("Scripting.Dictionary")
    form.Add "RequestNo", requestNo
    form.Add "Values", WebReadExistingValues(session, requestNo)
    Set WebLoadExisting = form
End Function

Private Function MockPost(ByVal form As Object, ByVal values As Object, ByVal updateExisting As Boolean) As Object
    Dim log As Worksheet, rowNum As Long, result As Object, outcome As String, index As Long, stored As Object
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    rowNum = log.Cells(log.Rows.Count, "A").End(xlUp).Row
    If CStr(log.Cells(rowNum, "E").Value2) <> IIf(updateExisting, "갱신 시작", "등록 시작") Then Err.Raise vbObjectError + 7991, , "POST before start journal."
    If CStr(log.Cells(rowNum, "D").Value2) <> CStr(form("RequestNo")) Then Err.Raise vbObjectError + 7992, , "Candidate differs from journal."
    If Not ThisWorkbook.Saved Then Err.Raise vbObjectError + 7993, , "Journal was not persisted before POST."
    If values.Count <> 13 Or CStr(values("GUBUN")) <> "PJT" Then Err.Raise vbObjectError + 7995, , "Unexpected form payload."
    WebUploadMockJournalChecks = WebUploadMockJournalChecks + 1
    If updateExisting Then WebUploadMockUpdateCount = WebUploadMockUpdateCount + 1 Else WebUploadMockPostCount = WebUploadMockPostCount + 1
    WebUploadMockPostedTitles = WebUploadMockPostedTitles & "|" & CStr(values("REQ_TIT"))
    index = WebUploadMockPostCount + WebUploadMockUpdateCount - 1
    If index > UBound(mMockOutcomes) Then index = UBound(mMockOutcomes)
    outcome = CStr(mMockOutcomes(index))
    If outcome = "throw" Then Err.Raise vbObjectError + 7996, , "Mock POST response interrupted."
    If outcome = "success" Then
        Set stored = MockClone(values)
        If updateExisting Then
            If mMockRequests(CStr(form("RequestNo"))).Exists("ACT_RMK") Then stored("ACT_RMK") = mMockRequests(CStr(form("RequestNo")))("ACT_RMK")
        End If
        WebUploadMockSetServer CStr(form("RequestNo")), stored
    ElseIf outcome = "partial" And updateExisting Then
        Set stored = MockClone(mMockRequests(CStr(form("RequestNo"))))
        stored("REQ_RMK") = values("REQ_RMK")
        WebUploadMockSetServer CStr(form("RequestNo")), stored
    End If
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Outcome", outcome
    result.Add "RequestNo", CStr(form("RequestNo"))
    result.Add "Message", "mock " & outcome
    result.Add "Success", (outcome = "success")
    If Not stored Is Nothing Then result.Add "Values", stored
    Set MockPost = result
End Function

Public Function WebPostRegistration(ByVal session As Object, ByVal form As Object, ByVal values As Object) As Object
    Set WebPostRegistration = MockPost(form, values, False)
End Function

Public Function WebPostUpdate(ByVal session As Object, ByVal form As Object, ByVal values As Object) As Object
    Set WebPostUpdate = MockPost(form, values, True)
End Function
