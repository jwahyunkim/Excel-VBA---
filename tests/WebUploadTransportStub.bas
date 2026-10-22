Attribute VB_Name = "WebUploadTransportStub"
Option Explicit
' Used only by test-web-upload.ps1 in a disposable workbook. Never imported in releases.
' 2026-10-22 김좌현 수정 : 실제 HTTP 없이 등록 오케스트레이션·전송 전 이력 저장 검증
Public WebUploadMockOpenCount As Long
Public WebUploadMockLoadCount As Long
Public WebUploadMockPostCount As Long
Public WebUploadMockJournalChecks As Long
Public WebUploadMockPostedTitles As String
Private mMockOutcomes As Variant

Public Sub WebUploadMockReset(Optional ByVal outcomes As String = "success")
    WebUploadMockOpenCount = 0
    WebUploadMockLoadCount = 0
    WebUploadMockPostCount = 0
    WebUploadMockJournalChecks = 0
    WebUploadMockPostedTitles = ""
    mMockOutcomes = Split(outcomes, ",")
End Sub

Public Function WebOpenSession(ByVal baseUrl As String, ByVal userId As String, ByVal password As String) As Object
    Dim session As Object
    If Len(baseUrl) = 0 Or Len(userId) = 0 Or Len(password) = 0 Then _
        Err.Raise vbObjectError + 7990, "WebUploadTransportStub", "Missing mock session credentials."
    WebUploadMockOpenCount = WebUploadMockOpenCount + 1
    Set session = CreateObject("Scripting.Dictionary")
    session.Add "BaseUrl", baseUrl
    session.Add "Authenticated", True
    Set WebOpenSession = session
End Function

Public Function WebLoadRegistration(ByVal session As Object) As Object
    Dim form As Object
    WebUploadMockLoadCount = WebUploadMockLoadCount + 1
    Set form = CreateObject("Scripting.Dictionary")
    form.Add "RequestNo", "RQMOCK" & Format$(WebUploadMockLoadCount, "000000")
    Set WebLoadRegistration = form
End Function

Public Function WebPostRegistration(ByVal session As Object, ByVal form As Object, ByVal values As Object) As Object
    Dim log As Worksheet, rowNum As Long, result As Object, outcome As String, index As Long
    WebUploadMockPostCount = WebUploadMockPostCount + 1
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    rowNum = log.Cells(log.Rows.Count, "A").End(xlUp).Row
    If CStr(log.Cells(rowNum, "E").Value2) <> "등록 시작" Then _
        Err.Raise vbObjectError + 7991, "WebUploadTransportStub", "POST started before the journal entry."
    If CStr(log.Cells(rowNum, "D").Value2) <> CStr(form("RequestNo")) Then _
        Err.Raise vbObjectError + 7992, "WebUploadTransportStub", "Journal candidate does not match the posted form."
    If Not ThisWorkbook.Saved Then _
        Err.Raise vbObjectError + 7993, "WebUploadTransportStub", "Journal was not persisted before POST."
    If Len(ThisWorkbook.Path) = 0 Or Len(Dir$(ThisWorkbook.FullName)) = 0 Then _
        Err.Raise vbObjectError + 7994, "WebUploadTransportStub", "Persisted workbook is missing."
    If values.Count <> 13 Or CStr(values("GUBUN")) <> "PJT" Then _
        Err.Raise vbObjectError + 7995, "WebUploadTransportStub", "Unexpected project form payload."
    WebUploadMockJournalChecks = WebUploadMockJournalChecks + 1
    WebUploadMockPostedTitles = WebUploadMockPostedTitles & "|" & CStr(values("REQ_TIT"))
    index = WebUploadMockPostCount - 1
    If index > UBound(mMockOutcomes) Then index = UBound(mMockOutcomes)
    outcome = CStr(mMockOutcomes(index))
    If outcome = "throw" Then Err.Raise vbObjectError + 7996, "WebUploadTransportStub", "Mock POST response interrupted."
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Outcome", outcome
    result.Add "RequestNo", CStr(form("RequestNo"))
    result.Add "Message", "mock " & outcome
    result.Add "Success", (outcome = "success")
    Set WebPostRegistration = result
End Function
