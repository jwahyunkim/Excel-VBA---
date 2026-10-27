Attribute VB_Name = "modWebUpload"
Option Explicit
' 2026-10-27 김좌현 수정 : 웹 관련 입력·업무 구조 변경에만 검사하고 일반 문구 입력의 전체 순회 생략
' 2026-10-27 김좌현 수정 : 등록 기준 입력 오류는 편집을 막지 않고 미리보기·모두 새로고침에서 행별 안내
' 2026-10-27 김좌현 수정 : 기존 웹 연결과 설정이 유지되는 빈 행 삽입을 등록 기준 직접 편집과 구분
' 2026-10-26 김좌현 수정 : 웹 연결 구조와 보호값을 보존하는 행 삭제의 등록 기준 변경 오탐 방지
' 2026-10-23 김좌현 수정 : 업무가 한 행인 이관 시트의 단일 셀 초기화 비교와 재열기 오류 방지
' 2026-10-23 김좌현 수정 : 데이터 이관용 웹 이력 시트 준비와 이전 미리보기·연결 확인 초기화
' 2026-10-23 김좌현 수정 : 미리보기 실행 종료 시 결과와 관계없이 원본 WEB UPDATE 전체 N 초기화
' 2026-10-23 김좌현 수정 : WEB UPDATE 선택 행의 신규·갱신 자동 판정과 완료 목록 정리 및 미리보기 재검증
' 2026-10-23 김좌현 수정 : 업무 초기화와 재열기에서 하위 작업시간의 상위 합계 갱신
' 2026-10-23 김좌현 수정 : 행 복사 시 웹 선택목록 이름 충돌과 원본 파일 외부 연결 방지
' 2026-10-22 김좌현 수정 : 프로젝트 실적 웹 설정, 선택목록, 미리보기와 등록 이력 추가
' 2026-10-22 김좌현 수정 : 등록 단위 동기화, 요청번호 연결과 충돌 확인을 포함한 신규/갱신 실행
' 2026-10-22 김좌현 수정 : 기존 갱신을 기본으로 하고 요청번호 없는 업무는 연결 필요로 차단
' 2026-10-22 김좌현 수정 : 삭제한 등록 이력을 새로고침에 반영하고 요청자 잠금 해제
' 2026-10-22 김좌현 수정 : 웹 초기화의 실제 값이 같을 때만 저장 상태를 복원

Private mPreviewItems As Collection
Private mPreviewSource As Worksheet
Private mPreviewSnapshot As String
Private mPreviewDecisionSnapshot As String
Private mWebBusy As Boolean
Private mConfigTaskSheet As Worksheet
Private mConfigTaskRow As Long
Private mLinkSourceId As String
Private mLinkRequestNo As String
Private mLinkServerFingerprint As String
Private mLinkPayloadFingerprint As String
Private mLinkSite As String
Private mLinkAccount As String
Private mLinkSource As Worksheet
Private Const WEB_OPTION_FIRST_ROW As Long = 9
Private Type WEB_TASK_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As WEB_TASK_GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As WEB_TASK_GUID, ByVal buffer As LongPtr, ByVal count As Long) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32" (ByRef value As WEB_TASK_GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef value As WEB_TASK_GUID, ByVal buffer As Long, ByVal count As Long) As Long
#End If

Public Function IsWebUtilitySheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    Select Case ws.Name
        Case WEB_CONFIG_SHEET_NAME, WEB_PREVIEW_SHEET_NAME, WEB_LOG_SHEET_NAME, WEB_STATE_SHEET_NAME
            IsWebUtilitySheet = True
    End Select
End Function

Private Function WebSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set WebSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If WebSheet Is Nothing Then
        Set WebSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WebSheet.Name = sheetName
    End If
End Function

Private Function WebOptionColumn(ByVal fieldName As String) As Long
    Select Case fieldName
        Case "REQ_UID": WebOptionColumn = 5
        Case "ACT_GUBUN": WebOptionColumn = 9
        Case "STS": WebOptionColumn = 13
        Case "PGM_MD": WebOptionColumn = 17
        Case "ACT_UID": WebOptionColumn = 21
        Case Else: Err.Raise vbObjectError + 2600, , "알 수 없는 웹 선택 항목입니다."
    End Select
End Function

Private Function WebOptionName(ByVal fieldName As String) As String
    Select Case fieldName
        Case "REQ_UID": WebOptionName = "WebRequesters"
        Case "ACT_GUBUN": WebOptionName = "WebActions"
        Case "STS": WebOptionName = "WebStates"
        Case "PGM_MD": WebOptionName = "WebModules"
        Case "ACT_UID": WebOptionName = "WebHandlers"
    End Select
End Function

Private Function WebOptionTitle(ByVal fieldName As String) As String
    Select Case fieldName
        Case "REQ_UID": WebOptionTitle = "요청자"
        Case "ACT_GUBUN": WebOptionTitle = "처리구분"
        Case "STS": WebOptionTitle = "진행상태"
        Case "PGM_MD": WebOptionTitle = "모듈/업무"
        Case "ACT_UID": WebOptionTitle = "처리자"
    End Select
End Function

Private Function WebTaskOptionColumn(ByVal fieldName As String) As String
    Select Case fieldName
        Case "REQ_UID": WebTaskOptionColumn = COL_WEB_REQUESTER
        Case "ACT_GUBUN": WebTaskOptionColumn = COL_WEB_ACTION
        Case "STS": WebTaskOptionColumn = COL_WEB_STATUS
        Case "PGM_MD": WebTaskOptionColumn = COL_WEB_MODULE
        Case "ACT_UID": WebTaskOptionColumn = COL_WEB_HANDLER
    End Select
End Function

Private Function WebDefaultCell(ByVal fieldName As String) As String
    Select Case fieldName
        Case "REQ_UID": WebDefaultCell = "B8"
        Case "ACT_GUBUN": WebDefaultCell = "B9"
        Case "STS": WebDefaultCell = "B10"
        Case "PGM_MD": WebDefaultCell = "B11"
        Case "ACT_UID": WebDefaultCell = "B12"
    End Select
End Function

Public Sub EnsureWebConfigSheet()
    Dim ws As Worksheet, field As Variant, c As Long, lastOptionRow As Long, hintRow As Long
    Set ws = WebSheet(WEB_CONFIG_SHEET_NAME)
    ws.Range("A1").Value2 = "프로젝트 실적 웹 등록 설정"
    ws.Range("A2").Value2 = "사이트 주소"
    ws.Range("A3").Value2 = "로그인 아이디"
    ws.Range("A4").Value2 = "비밀번호 입력"
    ws.Range("A5").Value2 = "선택목록 갱신 시각"
    ws.Range("A6").Value2 = "새 업무의 등록 기준"
    ws.Range("A7").Value2 = "새 업무의 WEB UPDATE 선택"
    ws.Range("A8").Value2 = "기본 요청자"
    ws.Range("A9").Value2 = "기본 처리구분"
    ws.Range("A10").Value2 = "기본 진행상태"
    ws.Range("A11").Value2 = "기본 모듈/업무"
    ws.Range("A12").Value2 = "기본 처리자"
    If Len(CStr(ws.Range("B2").Value2)) = 0 Then ws.Range("B2").Value2 = "http://www.thinkgood.co.kr/support/"
    If Len(CStr(ws.Range("B6").Value2)) = 0 Then ws.Range("B6").Value2 = 1
    If Len(CStr(ws.Range("B7").Value2)) = 0 Then ws.Range("B7").Value2 = "N"
    ws.Range("A14").Value2 = "비밀번호는 B4에 입력하면 이 Windows 사용자용으로 보호 저장됩니다."
    ws.Range("A15").Value2 = "기본값은 빈 행에만 적용합니다. 각 업무의 선택값은 메인 시트에서 바꿀 수 있습니다."
    ws.Range("A16").Value2 = "등록 기준은 레벨1 그룹 전체에 적용합니다. 대상 행의 WEB UPDATE를 Y로 선택하면 신규·갱신을 자동 판단합니다."
    ws.Range("A17").Value2 = "작업시간은 분 단위입니다. 실제 날짜와 작업시간을 비워 둔 행은 등록하지 않습니다."
    ws.Range("A18").Value2 = "기존 요청번호 연결은 메인 시트의 등록 단위 행을 선택한 뒤 연결 버튼을 누르세요. 미확인 건은 자동 재전송하지 않습니다."
    For hintRow = 14 To 18
        With ws.Range("A" & hintRow & ":B" & hintRow)
            If Not .MergeCells Then .Merge
            .WrapText = True
            .Font.Size = 10
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(hintRow).RowHeight = 32
    Next hintRow
    ws.Rows(13).RowHeight = 24
    ws.Range("Z1").Value2 = "Windows 사용자 보호 비밀번호"
    ws.Columns("Z").Hidden = True
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 48
    ws.Columns("C:D").ColumnWidth = 3
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Font.Size = 14
    ws.Range("A2:A12").Interior.Color = RGB(229, 239, 249)
    ws.Range("B2:B12").Interior.Color = RGB(255, 252, 229)
    ws.Range("B8:B12").NumberFormat = "@"
    With ws.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="1,2,3"
    End With
    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Y,N"
    End With
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        c = WebOptionColumn(CStr(field))
        ws.Cells(7, c).Value2 = WebOptionTitle(CStr(field))
        ws.Cells(8, c).Value2 = "선택 표시"
        ws.Cells(8, c + 1).Value2 = "웹 이름"
        ws.Cells(8, c + 2).Value2 = "웹 값"
        ws.Range(ws.Cells(7, c), ws.Cells(8, c + 2)).Interior.Color = RGB(229, 239, 249)
        ws.Range(ws.Cells(7, c), ws.Cells(8, c + 2)).Font.Bold = True
        ws.Columns(c).ColumnWidth = 28
        ws.Columns(c + 1).ColumnWidth = 24
        ws.Columns(c + 2).ColumnWidth = 18
        lastOptionRow = Application.Max(WEB_OPTION_FIRST_ROW, ws.Cells(ws.Rows.Count, c).End(xlUp).Row)
        ThisWorkbook.Names.Add Name:=WebOptionName(CStr(field)), _
            RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(WEB_OPTION_FIRST_ROW, c), ws.Cells(lastOptionRow, c)).Address
        With ws.Range(WebDefaultCell(CStr(field))).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""" & WebOptionName(CStr(field)) & """)"
            .IgnoreBlank = True
        End With
    Next field
    AddWebButton ws, "btnWebLogin", "로그인 확인", "TestWebLogin", ws.Range("A13").Left, ws.Range("A13").Top, 105
    AddWebButton ws, "btnWebOptions", "선택목록 갱신", "RefreshWebOptions", ws.Range("B13").Left, ws.Range("B13").Top, 120
    AddWebButton ws, "btnWebLinkRequest", "선택행 요청번호 연결", "LinkSelectedWebRequest", ws.Range("B19").Left, ws.Range("B19").Top, 160
    AddWebButton ws, "btnWebResolveAttempt", "선택행 전송결과 확인", "ResolveSelectedWebAttempt", ws.Range("B20").Left, ws.Range("B20").Top, 160
End Sub

Public Sub SetWebCredentials(ByVal userId As String, ByVal password As String)
    EnsureWebConfigSheet
    ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2 = userId
    SetWebPasswordValue password
End Sub

Public Sub SetWebPasswordValue(ByVal password As String)
    Dim ws As Worksheet, protectedValue As String, eventsEnabled As Boolean
    eventsEnabled = Application.EnableEvents
    Set ws = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    On Error GoTo Failed
    If Len(password) > 0 Then protectedValue = WebProtectSecret(password)
    Application.EnableEvents = False
    ws.Range("Z2").Value2 = protectedValue
    If Len(protectedValue) = 0 Then
        ws.Range("B4").ClearContents
    Else
        ws.Range("B4").Value2 = "********"
    End If
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    Application.EnableEvents = False
    ws.Range("B4").ClearContents
    Application.EnableEvents = eventsEnabled
    Err.Raise vbObjectError + 2601, , "비밀번호를 보호 저장하지 못했습니다. 다시 입력해 주세요."
End Sub

Private Function OpenConfiguredWebSession() As Object
    Dim ws As Worksheet, password As String
    EnsureWebConfigSheet
    Set ws = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    If Len(Trim$(CStr(ws.Range("B3").Value2))) = 0 Then Err.Raise vbObjectError + 2602, , "config_web에서 로그인 아이디를 입력하세요."
    If Len(CStr(ws.Range("Z2").Value2)) = 0 Then Err.Raise vbObjectError + 2602, , "config_web에서 비밀번호를 입력하세요."
    password = WebUnprotectSecret(CStr(ws.Range("Z2").Value2))
    Set OpenConfiguredWebSession = WebOpenSession(Trim$(CStr(ws.Range("B2").Value2)), Trim$(CStr(ws.Range("B3").Value2)), password)
    password = vbNullString
End Function

Public Sub TestWebLogin()
    Dim session As Object
    On Error GoTo Failed
    Set session = OpenConfiguredWebSession()
    MsgBox "로그인 연결을 확인했습니다.", vbInformation, "웹 로그인"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "웹 로그인"
End Sub

Public Function ResolveWebOption(ByVal fieldName As String, ByVal displayValue As String) As String
    Dim ws As Worksheet, c As Long, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    c = WebOptionColumn(fieldName)
    For r = WEB_OPTION_FIRST_ROW To ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If StrComp(CStr(ws.Cells(r, c).Value2), displayValue, vbBinaryCompare) = 0 Then
            ResolveWebOption = CStr(ws.Cells(r, c + 2).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function WebOptionLabel(ByVal fieldName As String, ByVal code As String) As String
    Dim ws As Worksheet, c As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    c = WebOptionColumn(fieldName)
    For r = WEB_OPTION_FIRST_ROW To ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If StrComp(CStr(ws.Cells(r, c + 2).Value2), code, vbBinaryCompare) = 0 Then
            WebOptionLabel = CStr(ws.Cells(r, c).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function CaptureWebChoices() As Collection
    Dim result As New Collection, ws As Worksheet, field As Variant, r As Long, entry As Object, code As String
    For Each ws In ThisWorkbook.Worksheets
        If IsWebTaskSheet(ws) Then
            For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
                For r = DATA_START_ROW To GetLastDataRow(ws)
                    code = ResolveWebOption(CStr(field), CStr(ws.Cells(r, WebTaskOptionColumn(CStr(field))).Value2))
                    If Len(code) > 0 Then
                        Set entry = CreateObject("Scripting.Dictionary")
                        entry.Add "Sheet", ws.Name
                        entry.Add "Row", r
                        entry.Add "Field", CStr(field)
                        entry.Add "Code", code
                        result.Add entry
                    End If
                Next r
            Next field
        End If
    Next ws
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        code = ResolveWebOption(CStr(field), CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range(WebDefaultCell(CStr(field))).Value2))
        If Len(code) > 0 Then
            Set entry = CreateObject("Scripting.Dictionary")
            entry.Add "Sheet", WEB_CONFIG_SHEET_NAME
            entry.Add "Row", 0
            entry.Add "Field", CStr(field)
            entry.Add "Code", code
            result.Add entry
        End If
    Next field
    Set CaptureWebChoices = result
End Function

Public Sub StoreWebOptionLists(ByVal form As Object)
    Dim ws As Worksheet, options As Object, choices As Collection, entry As Object
    Dim field As Variant, code As Variant, values As Object, c As Long, r As Long, label As String, populated As Boolean
    Set ws = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    Set options = form("Options")
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        If Not options.Exists(CStr(field)) Then Err.Raise vbObjectError + 2603, , "웹 선택목록을 가져오지 못했습니다: " & WebOptionTitle(CStr(field))
        Set values = options(CStr(field))
        If values.Count = 0 Then Err.Raise vbObjectError + 2603, , "웹 선택목록이 비어 있습니다: " & WebOptionTitle(CStr(field))
        populated = False
        For Each code In values.Keys
            If Len(CStr(code)) > 0 Then populated = True
        Next code
        If Not populated Then Err.Raise vbObjectError + 2603, , "웹 선택목록이 비어 있습니다: " & WebOptionTitle(CStr(field))
    Next field
    Set choices = CaptureWebChoices()
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        c = WebOptionColumn(CStr(field))
        ws.Range(ws.Cells(WEB_OPTION_FIRST_ROW, c), ws.Cells(Application.Max(WEB_OPTION_FIRST_ROW, ws.Cells(ws.Rows.Count, c).End(xlUp).Row), c + 2)).ClearContents
        Set values = options(CStr(field))
        ws.Range(ws.Cells(WEB_OPTION_FIRST_ROW, c), ws.Cells(WEB_OPTION_FIRST_ROW + values.Count, c + 2)).NumberFormat = "@"
        r = WEB_OPTION_FIRST_ROW
        For Each code In values.Keys
            If Len(CStr(code)) > 0 Then
                label = CStr(values(code))
                If CStr(field) = "REQ_UID" Or CStr(field) = "ACT_UID" Then label = label & " [" & CStr(code) & "]"
                ws.Cells(r, c).Value2 = label
                ws.Cells(r, c + 1).Value2 = CStr(values(code))
                ws.Cells(r, c + 2).NumberFormat = "@"
                ws.Cells(r, c + 2).Value2 = CStr(code)
                r = r + 1
            End If
        Next code
    Next field
    For Each entry In choices
        label = WebOptionLabel(CStr(entry("Field")), CStr(entry("Code")))
        If Len(label) = 0 Then label = "[웹 목록에서 제거됨: " & CStr(entry("Code")) & "]"
        If CLng(entry("Row")) = 0 Then
            ws.Range(WebDefaultCell(CStr(entry("Field")))).Value2 = label
        Else
            ThisWorkbook.Worksheets(CStr(entry("Sheet"))).Cells(CLng(entry("Row")), WebTaskOptionColumn(CStr(entry("Field")))).Value2 = label
        End If
    Next entry
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        If Len(CStr(ws.Range(WebDefaultCell(CStr(field))).Value2)) = 0 Then
            If form("Fields").Exists(CStr(field)) Then
                ws.Range(WebDefaultCell(CStr(field))).Value2 = WebOptionLabel(CStr(field), CStr(form("Fields")(CStr(field))))
            End If
        End If
    Next field
    ws.Range("B5").Value = Now
    ws.Range("B5").NumberFormat = "yyyy-mm-dd hh:mm"
    EnsureWebConfigSheet
End Sub

Public Sub RefreshWebOptions(Optional ByVal showCompletionMessage As Boolean = True)
    Dim session As Object, form As Object, ws As Worksheet, eventsEnabled As Boolean
    Dim errorNumber As Long, errorText As String
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    Set session = OpenConfiguredWebSession()
    Set form = WebLoadRegistration(session)
    Application.EnableEvents = False
    StoreWebOptionLists form
    For Each ws In ThisWorkbook.Worksheets
        If IsWebTaskSheet(ws) Then PrepareWebTaskSheet ws
    Next ws
    Application.EnableEvents = eventsEnabled
    If showCompletionMessage Then MsgBox "웹 선택목록을 갱신했습니다. 기존 업무의 선택은 웹 ID를 기준으로 유지했습니다.", vbInformation, "웹 선택목록"
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Application.EnableEvents = eventsEnabled
    If showCompletionMessage Then
        MsgBox errorText, vbExclamation, "웹 선택목록"
    Else
        Err.Raise errorNumber, "RefreshWebOptions", errorText
    End If
End Sub

Private Function IsWebTaskSheet(ByVal ws As Worksheet) As Boolean
    If IsWebUtilitySheet(ws) Then Exit Function
    IsWebTaskSheet = (CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) = "요청 내용")
End Function

Public Function GetWebTaskId(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    GetWebTaskId = Trim$(CStr(ws.Cells(rowNum, COL_WEB_TASK_ID).Value2))
End Function

Private Function NewWebTaskId() As String
    Dim value As WEB_TASK_GUID, rawGuid As String
    rawGuid = String$(39, vbNullChar)
    If CoCreateGuid(value) <> 0 Then
        Err.Raise vbObjectError + 2604, , "업무 식별값을 만들지 못했습니다."
    End If
    If StringFromGUID2(value, StrPtr(rawGuid), 39) = 0 Then Err.Raise vbObjectError + 2604, , "업무 식별값을 만들지 못했습니다."
    NewWebTaskId = Mid$(rawGuid, 2, 36)
End Function

Public Sub PrepareWebTaskSheet(ByVal ws As Worksheet)
    Dim r As Long, lastRow As Long, rootRow As Long, levelValue As Variant, field As Variant
    Dim config As Worksheet, wasProtected As Boolean, eventsEnabled As Boolean, seen As Object, taskId As String
    Dim errNumber As Long, errText As String, stepName As String, hiddenColumns As Object, columnKey As Variant, c As Long
    Dim isNewTask As Boolean, taskBasis As Variant
    If ws Is Nothing Or IsWebUtilitySheet(ws) Then Exit Sub
    Set hiddenColumns = CreateObject("Scripting.Dictionary")
    For c = ws.Range(COL_WEB_ENABLED & "1").Column To ws.Range(COL_WEB_RESULT & "1").Column
        hiddenColumns(CStr(c)) = CBool(ws.Columns(c).Hidden)
    Next c
    hiddenColumns(CStr(ws.Range(COL_WORK_MINUTES & "1").Column)) = CBool(ws.Columns(COL_WORK_MINUTES).Hidden)
    stepName = "웹 설정"
    EnsureWebConfigSheet
    Set config = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    eventsEnabled = Application.EnableEvents
    wasProtected = ws.ProtectContents
    On Error GoTo Failed
    Application.EnableEvents = False
    If wasProtected Then UnprotectTaskSheet ws
    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = GetLastDataRow(ws)
    ws.Range(COL_WEB_REQUESTER & DATA_START_ROW & ":" & COL_WEB_PROGRAM_ID & Application.Max(DATA_START_ROW, lastRow)).NumberFormat = "@"
    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            taskId = GetWebTaskId(ws, r)
            isNewTask = (Len(taskId) = 0)
            If isNewTask Then
                taskId = NewWebTaskId()
                ws.Cells(r, COL_WEB_TASK_ID).NumberFormat = "@"
                ws.Cells(r, COL_WEB_TASK_ID).Value2 = taskId
            End If
            If seen.Exists(taskId) Then
                ws.Cells(r, COL_WEB_RESULT).Value2 = "업무 식별값 중복: 복사한 행을 확인하세요."
            Else
                seen.Add taskId, r
            End If
            If GetTaskLevel(ws, r) = 1 Then
                rootRow = r
                levelValue = ws.Cells(r, COL_WEB_LEVEL).Value2
                If Not IsError(levelValue) Then
                    If Len(CStr(levelValue)) = 0 And isNewTask Then levelValue = config.Range("B6").Value2
                End If
            End If
            ' Keep existing invalid values visible until the requested validation action.
            If rootRow > 0 And isNewTask And WebBasisValueIsValid(levelValue) Then
                taskBasis = ws.Cells(r, COL_WEB_LEVEL).Value2
                If Not IsError(taskBasis) Then
                    If Len(Trim$(CStr(taskBasis))) = 0 Then ws.Cells(r, COL_WEB_LEVEL).Value2 = levelValue
                End If
            End If
            If Len(CStr(ws.Cells(r, COL_WEB_ENABLED).Value2)) = 0 Then ws.Cells(r, COL_WEB_ENABLED).Value2 = config.Range("B7").Value2
            For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
                If Len(CStr(ws.Cells(r, WebTaskOptionColumn(CStr(field))).Value2)) = 0 Then
                    ws.Cells(r, WebTaskOptionColumn(CStr(field))).Value2 = config.Range(WebDefaultCell(CStr(field))).Value2
                End If
            Next field
        End If
    Next r
    stepName = "작업시간 합계"
    UpdateTaskWorkMinutes ws, lastRow
    stepName = "등록 이력 동기화"
    RefreshWebHistoryState ws
    stepName = "등록 단위 동기화"
    SynchronizeWebUnits ws
    RememberWebGroupBases ws
    stepName = "입력 목록"
    ApplyWebTaskValidation ws
    ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_PROGRAM_ID).ColumnWidth = 14
    ws.Columns(COL_WEB_REQUESTER).ColumnWidth = 25
    ws.Columns(COL_WEB_HANDLER).ColumnWidth = 25
    ws.Columns(COL_WEB_RESULT).ColumnWidth = 28
    ws.Columns(COL_WORK_MINUTES).ColumnWidth = 13
    For Each columnKey In hiddenColumns.Keys
        ws.Columns(CLng(columnKey)).Hidden = CBool(hiddenColumns(columnKey))
    Next columnKey
    stepName = "버튼"
    EnsureWebTaskButtons ws
    ws.Columns(COL_WEB_TASK_ID).Hidden = True
    stepName = "시트 보호"
    If wasProtected Then ApplyCalculatedColumnsProtection ws, Application.Max(DATA_START_ROW, lastRow)
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errNumber = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not hiddenColumns Is Nothing Then
        For Each columnKey In hiddenColumns.Keys
            ws.Columns(CLng(columnKey)).Hidden = CBool(hiddenColumns(columnKey))
        Next columnKey
    End If
    If wasProtected Then ApplyCalculatedColumnsProtection ws, Application.Max(DATA_START_ROW, lastRow)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errNumber, "PrepareWebTaskSheet", stepName & ": " & errText
End Sub

Public Sub RefreshWebHistoryState(ByVal ws As Worksheet)
    Dim eventsEnabled As Boolean, wasProtected As Boolean, changed As Boolean
    Dim errorNumber As Long, errorText As String
    If Not IsWebTaskSheet(ws) Then Exit Sub
    eventsEnabled = Application.EnableEvents
    wasProtected = ws.ProtectContents
    On Error GoTo Failed
    Application.EnableEvents = False
    If wasProtected Then UnprotectTaskSheet ws
    changed = ReconcileDeletedWebHistory(ws)
    If changed Then
        Set mPreviewItems = Nothing
        mPreviewSnapshot = vbNullString
        mPreviewDecisionSnapshot = vbNullString
        Set mLinkSource = Nothing
        mLinkSourceId = vbNullString
        mLinkRequestNo = vbNullString
        mLinkServerFingerprint = vbNullString
        mLinkPayloadFingerprint = vbNullString
        mLinkSite = vbNullString
        mLinkAccount = vbNullString
    End If
    If wasProtected Then
        ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    ElseIf changed Then
        ApplyWebTaskProtection ws
    End If
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "RefreshWebHistoryState", errorText
End Sub

Public Sub ApplyWebTaskValidation(ByVal ws As Worksheet)
    Dim field As Variant, lastRow As Long, rng As Range
    If ws Is Nothing Or IsWebUtilitySheet(ws) Then Exit Sub
    EnsureWebConfigSheet
    lastRow = Application.Max(5000, GetLastDataRow(ws) + 1000)
    For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
        Set rng = ws.Range(WebTaskOptionColumn(CStr(field)) & DATA_START_ROW & ":" & WebTaskOptionColumn(CStr(field)) & lastRow)
        rng.NumberFormat = "@"
        With rng.Validation
            .Delete
            ' Keep copied dropdowns tied to the receiving workbook's existing lists.
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""" & WebOptionName(CStr(field)) & """)"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ErrorTitle = "웹 선택값 확인"
            .ErrorMessage = "config_web의 웹 선택목록에서 선택하세요."
        End With
    Next field
    With ws.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_ENABLED & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Y,N"
        .IgnoreBlank = True
    End With
    With ws.Range(COL_WEB_LEVEL & DATA_START_ROW & ":" & COL_WEB_LEVEL & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="1,2,3"
        .IgnoreBlank = True
        .InputTitle = "그룹 등록 기준"
        .InputMessage = "선택한 값은 같은 레벨1 그룹 전체에 적용됩니다."
        .ShowError = False
    End With
    With ws.Range(COL_WORK_MINUTES & DATA_START_ROW & ":" & COL_WORK_MINUTES & lastRow).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="2147483647"
        .IgnoreBlank = True
        .InputTitle = "작업시간(분)"
        .InputMessage = "실제 작업시간을 분 단위 정수로 입력하세요."
    End With
    ws.Range(COL_WORK_MINUTES & DATA_START_ROW & ":" & COL_WORK_MINUTES & lastRow).NumberFormat = "0"
    ws.Columns(COL_WEB_TASK_ID).Locked = True
    ws.Columns(COL_WEB_TASK_ID).Hidden = True
End Sub

Public Sub HandleWebTaskChange(ByVal ws As Worksheet, ByVal target As Range)
    Dim changed As Range, cell As Range, groups As Object, rootRow As Long, r As Long, lastRow As Long, key As Variant
    Dim units As Object, unitKey As String, columnNum As Long, structureIssue As String, state As Object
    Dim basisChange As Boolean, unitChange As Boolean, wasProtected As Boolean
    Dim mergePlans As Object, mergePlan As Object, mergeKey As Variant, errorNumber As Long, errorText As String
    Dim deferBasisValidation As Boolean
    Dim structureChanged As Boolean, watched As Range, cacheChanges As Range, changedUnit As Range
    If ws Is Nothing Or target Is Nothing Then Exit Sub
    If Not IsWebTaskSheet(ws) Then Exit Sub
    structureChanged = WebTaskContentStructureChanged(ws, target)
    Set watched = Union(ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_LEVEL & ws.Rows.Count), _
                        ws.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_PROGRAM_ID & ws.Rows.Count), _
                        ws.Range(COL_WEB_TASK_ID & DATA_START_ROW & ":" & COL_MODIFICATION & ws.Rows.Count))
    If Intersect(target, watched) Is Nothing Then
        UpdateWebRowCacheForChange ws, target
        Exit Sub
    End If
    If Not structureChanged Then
        If Intersect(target, ws.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_TASK_ID & ws.Rows.Count)) Is Nothing Then
            UpdateWebRowCacheForChange ws, target
            Exit Sub
        End If
    End If
    If structureChanged Then structureIssue = WebTaskStructureIssue(ws)
    If Len(structureIssue) > 0 Then
        If Not Intersect(target, Union(ws.Columns(COL_LEVEL), ws.Columns(COL_REQUEST), ws.Columns(COL_MODIFICATION), ws.Columns(COL_WEB_LEVEL))) Is Nothing Then
            Application.Undo
            Err.Raise vbObjectError + 2605, , structureIssue
        End If
    End If
    If WebRowDeletionKeepsSettings(ws, target) Then
        RememberWebGroupBases ws
        Exit Sub
    End If
    If WebRowInsertionKeepsSettings(ws, target) Then
        RememberWebGroupBases ws
        Exit Sub
    End If
    If Not structureChanged Then lastRow = WebCachedLastDataRow(ws)
    If lastRow = 0 Or structureChanged Then lastRow = GetLastDataRow(ws)
    Set changed = Intersect(target, ws.Range(COL_WEB_REQUESTER & DATA_START_ROW & ":" & COL_WEB_REQUESTER & Application.Max(DATA_START_ROW, lastRow)))
    If Not changed Is Nothing Then
        For Each cell In changed.Cells
            rootRow = WebRegistrationRootRow(ws, cell.Row, lastRow)
            If rootRow > 0 Then
                Set state = WebTaskState(GetWebTaskId(ws, rootRow))
                If Len(CStr(state("RequestNo"))) > 0 Or WebStateUnresolved(state) Then
                    Application.Undo
                    Err.Raise vbObjectError + 2605, , "등록된 요청자는 웹 상세 화면에서 변경할 수 없어 엑셀에서도 잠깁니다."
                End If
            End If
        Next cell
    End If
    Set changed = Intersect(target, ws.Range(COL_WEB_LEVEL & DATA_START_ROW & ":" & COL_WEB_LEVEL & ws.Rows.Count))
    Set groups = CreateObject("Scripting.Dictionary")
    Set units = CreateObject("Scripting.Dictionary")
    If Not changed Is Nothing Then
        Set changed = Intersect(changed, ws.Rows(DATA_START_ROW & ":" & Application.Max(DATA_START_ROW, lastRow)))
        If Not changed Is Nothing Then
            For Each cell In changed.Cells
                rootRow = WebGroupRootRow(ws, cell.Row)
                If rootRow > 0 Then
                    If WebGroupIsLocked(ws, rootRow) Then
                        Application.Undo
                        Err.Raise vbObjectError + 2605, , "등록 이력 또는 미확인 전송이 있는 레벨1 그룹의 등록 기준은 변경할 수 없습니다."
                    End If
                    If Not WebBasisValueIsValid(cell.Value2) Then
                        deferBasisValidation = True
                    ElseIf groups.Exists(CStr(rootRow)) Then
                        If groups(CStr(rootRow)) <> CStr(cell.Value2) Then
                            deferBasisValidation = True
                        End If
                    Else
                        groups.Add CStr(rootRow), CStr(cell.Value2)
                    End If
                End If
            Next cell
        End If
    End If
    If deferBasisValidation Then
        RememberWebGroupBases ws
        Exit Sub
    End If
    Set changed = Intersect(target, Union(ws.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_ENABLED & Application.Max(DATA_START_ROW, lastRow)), _
                                        ws.Range(COL_WEB_REQUESTER & DATA_START_ROW & ":" & COL_WEB_PROGRAM_ID & Application.Max(DATA_START_ROW, lastRow))))
    If Not changed Is Nothing Then
        For Each cell In changed.Cells
            rootRow = WebRegistrationRootRow(ws, cell.Row, lastRow)
            If rootRow = 0 Then
                r = WebGroupRootRow(ws, cell.Row)
                If r > 0 Then
                    If Not WebBasisValueIsValid(ws.Cells(r, COL_WEB_LEVEL).Value2) Then
                        UpdateWebRowCacheForChange ws, target
                        Exit Sub
                    End If
                End If
                If Len(CStr(cell.Value2)) > 0 Then
                    Application.Undo
                    Err.Raise vbObjectError + 2605, , "이 상위 행은 현재 등록 기준의 등록 대상이 아닙니다. 실제 등록 단위에서 설정하세요."
                End If
            Else
                unitKey = CStr(rootRow) & ":" & CStr(cell.Column)
                If units.Exists(unitKey) Then
                    If CStr(units(unitKey)) <> CStr(cell.Value2) Then
                        Application.Undo
                        Err.Raise vbObjectError + 2605, , "같은 등록 단위에 서로 다른 웹 값을 동시에 붙여넣을 수 없습니다."
                    End If
                Else
                    units.Add unitKey, cell.Value2
                End If
            End If
        Next cell
    End If
    If groups.Count = 0 And units.Count = 0 Then
        If structureChanged Then
            RememberWebGroupBases ws
        Else
            UpdateWebRowCacheForChange ws, target
        End If
        Exit Sub
    End If
    Set mergePlans = CreateObject("Scripting.Dictionary")
    On Error GoTo InvalidMerge
    For Each key In groups.Keys
        Set mergePlan = WebBasisMergePlan(ws, CLng(key), CLng(groups(key)))
        For Each mergeKey In mergePlan.Keys
            mergePlans(CStr(mergeKey)) = mergePlan(mergeKey)
        Next mergeKey
    Next key
    On Error GoTo 0
    wasProtected = ws.ProtectContents And (groups.Count > 0 Or Not ws.ProtectionMode)
    If wasProtected Then UnprotectTaskSheet ws
    For Each key In groups.Keys
        rootRow = CLng(key)
        ws.Range(COL_WEB_LEVEL & rootRow & ":" & COL_WEB_LEVEL & GetTaskSubtreeEndRow(ws, rootRow, lastRow)).Value2 = groups(key)
    Next key
    For Each key In mergePlans.Keys
        ws.Cells(CLng(Split(CStr(key), ":")(0)), CLng(Split(CStr(key), ":")(1))).Value2 = mergePlans(key)
    Next key
    For Each key In units.Keys
        rootRow = CLng(Split(CStr(key), ":")(0))
        columnNum = CLng(Split(CStr(key), ":")(1))
        Set changedUnit = ws.Range(ws.Cells(rootRow, columnNum), ws.Cells(GetTaskSubtreeEndRow(ws, rootRow, lastRow), columnNum))
        changedUnit.Value2 = units(key)
        If cacheChanges Is Nothing Then
            Set cacheChanges = changedUnit
        Else
            Set cacheChanges = Union(cacheChanges, changedUnit)
        End If
    Next key
    If groups.Count > 0 Then SynchronizeWebUnits ws
    If groups.Count > 0 Or structureChanged Then
        RememberWebGroupBases ws
    ElseIf cacheChanges Is Nothing Then
        UpdateWebRowCacheForChange ws, target
    Else
        UpdateWebRowCacheForChange ws, Union(target, cacheChanges)
    End If
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Exit Sub
InvalidMerge:
    errorNumber = Err.Number: errorText = Err.Description
    Application.Undo
    Err.Raise errorNumber, "HandleWebTaskChange", errorText
End Sub

Public Sub SynchronizeWebUnits(ByVal ws As Worksheet)
    Dim r As Long, root As Long, lastRow As Long, c As Variant, child As Long
    If Len(WebTaskBasisIssues(ws)) > 0 Then Exit Sub
    lastRow = GetLastDataRow(ws)
    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            root = WebRegistrationRootRow(ws, r, lastRow)
            If root > 0 And root <> r Then
                For Each c In Array(COL_WEB_ENABLED, COL_WEB_REQUESTER, COL_WEB_ACTION, COL_WEB_STATUS, COL_WEB_MODULE, COL_WEB_HANDLER, COL_WEB_PROGRAM_ID)
                    ws.Cells(r, CStr(c)).Value2 = ws.Cells(root, CStr(c)).Value2
                Next c
                ws.Cells(r, COL_WEB_RESULT).Value2 = ws.Cells(root, COL_WEB_RESULT).Value2
            ElseIf root = 0 And HasChildTask(ws, r, lastRow) Then
                ' A former representative can keep a locked Y after changing the group basis.
                For child = r + 1 To GetTaskSubtreeEndRow(ws, r, lastRow)
                    If HasTaskContent(ws, child) Then
                        If WebRegistrationRootRow(ws, child, lastRow) > 0 Then
                            ws.Cells(r, COL_WEB_ENABLED).Value2 = "N"
                            Exit For
                        End If
                    End If
                Next child
            End If
        End If
    Next r
End Sub

Private Function WebInitializationRangeSnapshot(ByVal cells As Range) As String
    Dim values As Variant, parts() As String, r As Long, c As Long, i As Long, value As String
    values = cells.Formula
    If cells.Cells.CountLarge = 1 Then
        If IsError(values) Then value = cells.Text Else value = CStr(values)
        WebInitializationRangeSnapshot = cells.Address & "|" & CStr(VarType(values)) & ":" & CStr(Len(value)) & ":" & value
        Exit Function
    End If
    ReDim parts(1 To UBound(values, 1) * UBound(values, 2))
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If IsError(values(r, c)) Then
                value = cells.Cells(r, c).Text
            Else
                value = CStr(values(r, c))
            End If
            i = i + 1
            parts(i) = CStr(VarType(values(r, c))) & ":" & CStr(Len(value)) & ":" & value
        Next c
    Next r
    WebInitializationRangeSnapshot = cells.Address & "|" & Join(parts, vbNullString)
End Function

Private Function WebInitializationSnapshot() As String
    Dim ws As Worksheet, lastRow As Long, c As Long, lastColumn As Long, value As String
    For Each ws In ThisWorkbook.Worksheets
        value = vbNullString
        lastColumn = 0
        Select Case ws.Name
            Case WEB_CONFIG_SHEET_NAME: lastColumn = 26
            Case WEB_STATE_SHEET_NAME: lastColumn = 15
            Case WEB_LOG_SHEET_NAME: lastColumn = 13
        End Select
        If lastColumn > 0 Then
            lastRow = 1
            For c = 1 To lastColumn
                lastRow = Application.Max(lastRow, ws.Cells(ws.Rows.Count, c).End(xlUp).Row)
            Next c
            value = WebInitializationRangeSnapshot(ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)))
        ElseIf IsWebTaskSheet(ws) Then
            lastRow = Application.Max(DATA_START_ROW, GetLastDataRow(ws))
            value = WebInitializationRangeSnapshot(ws.Range("B" & HEADER_ROW & ":AE" & HEADER_ROW)) & _
                    WebInitializationRangeSnapshot(ws.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_TASK_ID & lastRow)) & _
                    WebInitializationRangeSnapshot(ws.Range(COL_WORK_MINUTES & DATA_START_ROW & ":" & COL_WORK_MINUTES & lastRow))
        End If
        If Len(value) > 0 Then WebInitializationSnapshot = WebInitializationSnapshot & _
            CStr(Len(ws.Name)) & ":" & ws.Name & CStr(Len(value)) & ":" & value
    Next ws
End Function

Public Sub InitializeWebUpload()
    Dim ws As Worksheet, eventsEnabled As Boolean, previousSheet As Object, errorNumber As Long, errorText As String
    Dim stepName As String, wasSaved As Boolean, initialValues As String
    wasSaved = ThisWorkbook.Saved
    eventsEnabled = Application.EnableEvents
    Set previousSheet = ThisWorkbook.ActiveSheet
    On Error GoTo Failed
    If wasSaved Then initialValues = WebInitializationSnapshot()
    Application.EnableEvents = False
    stepName = "웹 설정"
    EnsureWebConfigSheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsWebUtilitySheet(ws) And ws.Name <> CONFIG_SHEET_NAME And ws.Name <> OUTPUT_CONFIG_SHEET_NAME Then
            If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." Then
                EnsureTaskSheetLayout ws
                stepName = ws.Name & " 업무 설정"
                PrepareWebTaskSheet ws
            End If
        End If
    Next ws
    stepName = "웹 이력"
    EnsureWebLogSheet
    stepName = "웹 연결"
    EnsureWebStateSheet
    stepName = "이전 이력 이관"
    MigrateLegacyWebJournal
    previousSheet.Activate
    Application.EnableEvents = eventsEnabled
    If wasSaved Then
        If initialValues = WebInitializationSnapshot() Then ThisWorkbook.Saved = True
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    previousSheet.Activate
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "InitializeWebUpload", stepName & ": " & errorText
End Sub

Private Sub AddWebButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal caption As String, _
                         ByVal action As String, ByVal left As Double, ByVal top As Double, ByVal width As Double)
    Dim button As Shape
    On Error Resume Next
    Set button = ws.Shapes(buttonName)
    On Error GoTo 0
    If button Is Nothing Then Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, width, 18)
    With button
        .Name = buttonName
        .Visible = msoTrue
        .Left = left: .Top = top: .Width = width: .Height = 18
        .Placement = xlFreeFloating
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
        .Fill.ForeColor.RGB = RGB(229, 239, 249)
        .Line.ForeColor.RGB = RGB(139, 164, 185)
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Name = "맑은 고딕"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 45, 65)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.MarginLeft = 2: .TextFrame2.MarginRight = 2
    End With
End Sub

Public Sub EnsureWebTaskButtons(ByVal ws As Worksheet)
    If Not IsWebTaskSheet(ws) Then Exit Sub
    If ws.Rows(3).RowHeight < 20 Then ws.Rows(3).RowHeight = 20
    AddWebButton ws, "btnWebPreview", "웹등록 미리보기", "PreviewWebTasks", ws.Range("B3").Left, ws.Range("B3").Top, 112
    On Error Resume Next
    ws.Shapes("btnWebUpload").Delete
    On Error GoTo 0
    AddWebButton ws, "btnWebConfig", "웹 설정", "OpenWebConfig", ws.Range("B3").Left + 118, ws.Range("B3").Top, 80
    RefreshWebModeView ws, True
End Sub

Public Sub OpenWebConfig()
    If IsWebTaskSheet(ActiveSheet) Then
        Set mConfigTaskSheet = ActiveSheet
        mConfigTaskRow = ActiveCell.Row
    End If
    EnsureWebConfigSheet
    ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Activate
End Sub

Public Sub EnsureWebHistorySheets()
    EnsureWebLogSheet
    EnsureWebStateSheet
End Sub

Public Sub ResetWebPreviewAfterMigration()
    Dim preview As Worksheet, buttonName As Variant
    Set mPreviewItems = Nothing
    Set mPreviewSource = Nothing
    mPreviewSnapshot = vbNullString
    mPreviewDecisionSnapshot = vbNullString
    Set mConfigTaskSheet = Nothing
    mConfigTaskRow = 0
    Set mLinkSource = Nothing
    mLinkSourceId = vbNullString
    mLinkRequestNo = vbNullString
    mLinkServerFingerprint = vbNullString
    mLinkPayloadFingerprint = vbNullString
    mLinkSite = vbNullString
    mLinkAccount = vbNullString
    On Error Resume Next
    Set preview = ThisWorkbook.Worksheets(WEB_PREVIEW_SHEET_NAME)
    On Error GoTo 0
    If preview Is Nothing Then Exit Sub
    preview.Unprotect
    If preview.AutoFilterMode Then preview.AutoFilterMode = False
    preview.Cells.Clear
    preview.Cells.Locked = True
    preview.Range("A1").Value2 = "프로젝트 실적 미리보기"
    preview.Range("A2").Value2 = "데이터를 가져왔습니다. 메인 시트에서 웹등록 미리보기를 다시 만드세요."
    For Each buttonName In Array("btnWebUploadPreview", "btnWebPreviewRefresh", "btnWebLinkConfirm", "btnWebLinkCancel")
        On Error Resume Next
        preview.Shapes(CStr(buttonName)).Visible = msoFalse
        On Error GoTo 0
    Next buttonName
    preview.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Private Sub EnsureWebLogSheet()
    Dim ws As Worksheet, headers As Variant, i As Long
    Set ws = WebSheet(WEB_LOG_SHEET_NAME)
    ws.Unprotect
    headers = Array("기록 시각", "업무 시트", "업무 ID", "웹 요청번호", "상태", "결과", "제목", "실제 시작일", "실제 종료일", "작업시간(분)", "작업", "사이트", "실행 계정")
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value2 = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(229, 239, 249)
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").Hidden = True
    ws.Columns("D:E").ColumnWidth = 20
    ws.Columns("F:G").ColumnWidth = 45
    ws.Columns("F:G").NumberFormat = "@"
    ws.Columns("H:J").ColumnWidth = 15
    ws.Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Private Function PreviousWebOutcome(ByVal sourceId As String, ByRef requestNo As String) As String
    Dim ws As Worksheet, r As Long, state As Object
    Set state = WebTaskState(sourceId)
    If Len(CStr(state("Outcome"))) > 0 Then
        PreviousWebOutcome = CStr(state("Outcome"))
        requestNo = CStr(state("RequestNo"))
        If Len(requestNo) = 0 Then requestNo = CStr(state("PendingNo"))
        Exit Function
    End If
    EnsureWebLogSheet
    Set ws = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    For r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If CStr(ws.Cells(r, "C").Value2) = sourceId Then
            PreviousWebOutcome = CStr(ws.Cells(r, "E").Value2)
            requestNo = CStr(ws.Cells(r, "D").Value2)
            Exit Function
        End If
    Next r
End Function

Private Function NewWebTaskIssue(ByVal sourceId As String, ByVal state As Object) As String
    Dim log As Worksheet, r As Long, outcome As String, latestOutcome As String, site As String
    If WebTaskHasOtherSiteConnection(sourceId) Then
        NewWebTaskIssue = "다른 사이트의 등록 이력이 있습니다. 사이트와 요청번호 연결을 확인하세요."
        Exit Function
    End If
    If Len(CStr(state("RequestNo"))) > 0 Or Len(CStr(state("PendingNo"))) > 0 Or _
       Len(CStr(state("Baseline"))) > 0 Or Len(CStr(state("ConfirmedAt"))) > 0 Or _
       Len(CStr(state("RequesterCode"))) > 0 Then
        NewWebTaskIssue = "기존 전송 또는 등록 이력이 있습니다. 요청번호 연결과 전송결과를 먼저 확인하세요."
        Exit Function
    End If
    outcome = CStr(state("Outcome"))
    If outcome <> vbNullString And outcome <> "등록실패" And outcome <> "미등록 확인" Then
        NewWebTaskIssue = "기존 이력의 요청번호가 연결되지 않았습니다. 웹 설정에서 요청번호 연결을 확인하세요."
        Exit Function
    End If
    On Error Resume Next
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    On Error GoTo 0
    If log Is Nothing Then Exit Function
    For r = log.Cells(log.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If CStr(log.Cells(r, "C").Value2) = sourceId Then
            outcome = CStr(log.Cells(r, "E").Value2)
            site = LCase$(Trim$(CStr(log.Cells(r, "L").Value2)))
            If Len(site) > 0 And Right$(site, 1) <> "/" Then site = site & "/"
            If Len(site) > 0 And site <> WebCurrentSite() Then
                If outcome <> "등록실패" And outcome <> "미등록 확인" Then
                    NewWebTaskIssue = "다른 사이트의 전송 이력이 있습니다. 사이트와 요청번호 연결을 확인하세요."
                    Exit Function
                End If
            Else
                If Len(latestOutcome) = 0 Then latestOutcome = outcome
                If outcome = "등록완료" Or outcome = "갱신완료" Or outcome = "연결완료" Then
                    NewWebTaskIssue = "등록 이력은 있으나 요청번호 연결이 없습니다. 기존 업무를 연결한 뒤 갱신하세요."
                    Exit Function
                End If
            End If
        End If
    Next r
    If latestOutcome <> vbNullString And latestOutcome <> "등록실패" And latestOutcome <> "미등록 확인" Then _
        NewWebTaskIssue = "앞선 전송 이력이 있습니다. 요청번호 연결과 전송결과를 먼저 확인하세요."
End Function

Private Sub AssessWebItem(ByVal item As Object, ByRef session As Object)
    Dim state As Object, server As Object, requestNo As String, currentFingerprint As String
    item("Operation") = "확인 필요"
    item("RequestNo") = vbNullString
    item("ServerSnapshot") = vbNullString
    If Len(CStr(item("Error"))) > 0 Then Exit Sub
    Set state = WebTaskState(CStr(item("SourceId")))
    requestNo = CStr(state("RequestNo"))
    If Len(requestNo) = 0 Then
        item("Error") = NewWebTaskIssue(CStr(item("SourceId")), state)
        If Len(CStr(item("Error"))) = 0 Then
            item("Operation") = "신규"
        Else
            item("RequestNo") = CStr(state("PendingNo"))
        End If
        Exit Sub
    End If
    item("RequestNo") = requestNo
    If WebStateUnresolved(state) Or Len(CStr(state("PendingNo"))) > 0 Then
        item("Error") = "앞선 전송 결과를 웹에서 먼저 확인하세요: " & requestNo & ". 실제 요청번호를 연결하여 확인값을 복구할 수 있습니다."
        Exit Sub
    End If
    If Len(CStr(state("Baseline"))) = 0 Then
        item("Error") = "기존 이력의 마지막 확인값이 없습니다. 요청번호 연결로 웹 내용을 확인하세요."
        Exit Sub
    End If
    On Error GoTo ReadFailed
    If session Is Nothing Then Set session = OpenConfiguredWebSession()
    Set server = WebReadExistingValues(session, requestNo)
    currentFingerprint = WebValuesFingerprint(server)
    item("ServerSnapshot") = currentFingerprint
    If currentFingerprint <> CStr(state("Baseline")) Then
        item("Error") = "웹 내용이 마지막 확인 이후 바뀌었습니다. 요청번호 연결에서 현재 웹 내용을 확인한 뒤 기준값을 갱신하세요."
    ElseIf CStr(server("REQ_UID")) <> CStr(item("Values")("REQ_UID")) Then
        item("Error") = "이미 등록된 요청자는 변경할 수 없습니다. 기존 요청자를 유지하세요."
    ElseIf currentFingerprint = WebValuesFingerprint(item("Values")) Then
        item("Operation") = "변경 없음"
    Else
        item("Operation") = "갱신"
    End If
    Exit Sub
ReadFailed:
    item("Error") = "요청번호 " & requestNo & " 조회 확인이 필요합니다: " & Err.Description
End Sub

Private Function PreviewItemsSnapshot(ByVal items As Collection) As String
    Dim item As Object, field As Variant, output As String, value As String
    output = CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2) & vbNullChar & _
             CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2) & vbNullChar
    For Each item In items
        value = CStr(item("SourceId")) & "|" & CStr(item("Row")) & "|" & _
                CStr(item("GroupRow")) & "|" & CStr(item("Error"))
        output = output & Len(value) & ":" & value
        For Each field In Array("GUBUN", "REQ_UID", "REQ_TIT", "REQ_RMK", "STS", "PGM_MD", "ACT_GUBUN", "ACT_UID", "PGM_NM", "PGM_ID", "ACT_DT_FR", "ACT_DT", "ACT_TIME")
            If item("Values").Exists(CStr(field)) Then value = CStr(item("Values")(CStr(field))) Else value = vbNullString
            output = output & Len(value) & ":" & value
        Next field
    Next item
    PreviewItemsSnapshot = output
End Function

Private Function PreviewDecisionSnapshot(ByVal items As Collection) As String
    Dim item As Object, field As Variant, value As String
    PreviewDecisionSnapshot = PreviewItemsSnapshot(items)
    For Each item In items
        For Each field In Array("Operation", "RequestNo", "ServerSnapshot")
            value = CStr(item(CStr(field)))
            PreviewDecisionSnapshot = PreviewDecisionSnapshot & Len(value) & ":" & value
        Next field
    Next item
End Function

Public Sub PreviewWebTasks(Optional ByVal showCompletionMessage As Boolean = True)
    Dim ws As Worksheet, preview As Worksheet, item As Object, r As Long, i As Long, headers As Variant
    Dim requestNo As String, outcome As String, issue As String, periodFields As Variant, errorNumber As Long, errorText As String
    Dim session As Object, structureIssue As String, eventsEnabled As Boolean, wasProtected As Boolean
    Dim selectedItems As Collection
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    If mWebBusy Then Exit Sub
    If ActiveSheet.Name = WEB_PREVIEW_SHEET_NAME Then
        If mPreviewSource Is Nothing Then Err.Raise vbObjectError + 2606, , "메인 업무 시트에서 미리보기를 다시 만드세요."
        Set ws = mPreviewSource
    Else
        Set ws = ActiveSheet
    End If
    If Not IsWebTaskSheet(ws) Then Err.Raise vbObjectError + 2606, , "메인 업무 시트에서 실행하세요."
    AssertWebTaskBases ws
    structureIssue = WebTaskStructureIssue(ws)
    PrepareWebTaskSheet ws
    eventsEnabled = Application.EnableEvents
    wasProtected = ws.ProtectContents
    Application.EnableEvents = False
    If wasProtected Then UnprotectTaskSheet ws
    RefreshWebTaskPeriods ws
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    Set mPreviewSource = ws
    Set selectedItems = BuildWebUploadItems(ws)
    Set mPreviewItems = New Collection
    For Each item In selectedItems
        If Len(structureIssue) > 0 Then item("Error") = structureIssue
        AssessWebItem item, session
        If CStr(item("Operation")) = "변경 없음" And Len(CStr(item("Error"))) = 0 Then
            CompleteWebUpdateSelection ws, item
        Else
            mPreviewItems.Add item
        End If
    Next item
    mPreviewSnapshot = PreviewItemsSnapshot(BuildWebUploadItems(ws))
    mPreviewDecisionSnapshot = PreviewDecisionSnapshot(mPreviewItems)
    Set preview = WebSheet(WEB_PREVIEW_SHEET_NAME)
    preview.Unprotect
    On Error Resume Next
    preview.Shapes("btnWebLinkConfirm").Visible = msoFalse
    preview.Shapes("btnWebLinkCancel").Visible = msoFalse
    On Error GoTo Failed
    If preview.AutoFilterMode Then preview.AutoFilterMode = False
    preview.Cells.Clear
    preview.Cells.Locked = True
    preview.Cells.Font.Name = "맑은 고딕"
    preview.Cells.Font.Size = 10
    preview.Columns("E:I").NumberFormat = "@"
    preview.Rows(3).RowHeight = 21
    preview.Range("A1").Value2 = "프로젝트 실적 미리보기"
    preview.Range("A2").Value2 = "원본: " & ws.Name & " / WEB UPDATE=Y인 업무의 신규·갱신을 이력으로 자동 판단합니다. 실행 후 원본 시트의 WEB UPDATE는 모두 N으로 초기화됩니다."
    headers = Array("실행 구분", "원본 행", "제목", "요청/처리내용", "요청자", "처리구분", "진행상태", "모듈/업무", "처리자", "시작일", "종료일", "시간(분)", "확인 사항", "요청번호")
    periodFields = Array("ACT_DT_FR", "ACT_DT", "ACT_TIME")
    For i = 0 To UBound(headers)
        preview.Cells(4, i + 1).Value2 = headers(i)
    Next i
    r = 5
    For Each item In mPreviewItems
        issue = CStr(item("Error"))
        preview.Cells(r, 1).Value2 = item("Operation")
        preview.Cells(r, 14).NumberFormat = "@"
        preview.Cells(r, 14).Value2 = item("RequestNo")
        preview.Cells(r, 2).Value2 = item("Row")
        preview.Cells(r, 3).Value2 = Left$(CStr(item("Title")), 32767)
        preview.Cells(r, 4).Value2 = Left$(CStr(item("Body")), 32767)
        If CLng(item("Row")) >= DATA_START_ROW Then
            preview.Cells(r, 5).Value2 = ws.Cells(CLng(item("Row")), COL_WEB_REQUESTER).Value2
            preview.Cells(r, 6).Value2 = ws.Cells(CLng(item("Row")), COL_WEB_ACTION).Value2
            preview.Cells(r, 7).Value2 = ws.Cells(CLng(item("Row")), COL_WEB_STATUS).Value2
            preview.Cells(r, 8).Value2 = ws.Cells(CLng(item("Row")), COL_WEB_MODULE).Value2
            preview.Cells(r, 9).Value2 = ws.Cells(CLng(item("Row")), COL_WEB_HANDLER).Value2
        End If
        For i = 0 To 2
            If item("Values").Exists(CStr(periodFields(i))) Then preview.Cells(r, 10 + i).Value2 = item("Values")(CStr(periodFields(i)))
        Next i
        preview.Cells(r, 13).Value2 = issue
        If Len(issue) > 0 Then preview.Range(preview.Cells(r, 1), preview.Cells(r, 14)).Interior.Color = RGB(255, 239, 224)
        r = r + 1
    Next item
    If mPreviewItems.Count = 0 Then preview.Range("A5").Value2 = "처리할 업무가 없습니다. 메인 시트에서 대상 행의 WEB UPDATE를 Y로 선택하세요."
    preview.Rows(4).Font.Bold = True
    preview.Rows(4).Interior.Color = RGB(229, 239, 249)
    preview.Columns("A").ColumnWidth = 14
    preview.Columns("B").ColumnWidth = 9
    preview.Columns("C").ColumnWidth = 46
    preview.Columns("D").ColumnWidth = 70
    preview.Columns("E:I").ColumnWidth = 22
    preview.Columns("E:I").NumberFormat = "@"
    preview.Columns("J:L").ColumnWidth = 13
    preview.Columns("M").ColumnWidth = 55
    preview.Columns("N").ColumnWidth = 20
    preview.Range("C5:M" & Application.Max(5, r - 1)).WrapText = True
    preview.Range("A5:M" & Application.Max(5, r - 1)).VerticalAlignment = xlTop
    preview.Rows("5:" & Application.Max(5, r - 1)).AutoFit
    AddWebButton preview, "btnWebUploadPreview", "등록·갱신 실행", "UploadWebTasks", preview.Range("A3").Left, preview.Range("A3").Top, 105
    AddWebButton preview, "btnWebPreviewRefresh", "미리보기 갱신", "PreviewWebTasks", preview.Range("C3").Left, preview.Range("C3").Top, 112
    preview.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    preview.Activate
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Set mPreviewItems = Nothing
    mPreviewSnapshot = vbNullString
    mPreviewDecisionSnapshot = vbNullString
    On Error Resume Next
    If Not ws Is Nothing And wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    If showCompletionMessage Then
        MsgBox errorText, vbExclamation, "웹등록 미리보기"
    Else
        Err.Raise errorNumber, "PreviewWebTasks", errorText
    End If
End Sub

Private Sub SaveWebJournal()
    Dim eventsEnabled As Boolean, errorNumber As Long, errorText As String
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = True
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then Err.Raise vbObjectError + 2608, , "등록 이력을 파일에 저장하지 못해 업로드를 중단했습니다."
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Application.EnableEvents = eventsEnabled
    Err.Raise errorNumber, "SaveWebJournal", errorText
End Sub

Private Function FindWebTaskRow(ByVal ws As Worksheet, ByVal sourceId As String) As Long
    Dim r As Long
    For r = DATA_START_ROW To GetLastDataRow(ws)
        If GetWebTaskId(ws, r) = sourceId Then
            FindWebTaskRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub CompleteWebUpdateSelection(ByVal ws As Worksheet, ByVal item As Object)
    Dim rowNum As Long, r As Long, lastRow As Long, wasProtected As Boolean, eventsEnabled As Boolean
    Dim errorNumber As Long, errorText As String
    rowNum = FindWebTaskRow(ws, CStr(item("SourceId")))
    If rowNum = 0 Or rowNum <> CLng(item("Row")) Then _
        Err.Raise vbObjectError + 2607, , "업무 위치가 바뀌었습니다. 미리보기를 다시 만드세요."
    lastRow = GetLastDataRow(ws)
    wasProtected = ws.ProtectContents
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    If wasProtected Then UnprotectTaskSheet ws
    For r = rowNum To GetTaskSubtreeEndRow(ws, rowNum, lastRow)
        If HasTaskContent(ws, r) Then ws.Cells(r, COL_WEB_ENABLED).Value2 = "N"
    Next r
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "CompleteWebUpdateSelection", errorText
End Sub

Private Sub ResetWebUpdateSelections(ByVal ws As Worksheet)
    Dim lastRow As Long, lastFlagRow As Long, wasProtected As Boolean, eventsEnabled As Boolean
    Dim errorNumber As Long, errorText As String
    lastRow = GetLastDataRow(ws)
    lastFlagRow = Application.Max(lastRow, ws.Cells(ws.Rows.Count, COL_WEB_ENABLED).End(xlUp).Row)
    wasProtected = ws.ProtectContents
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    If wasProtected Then UnprotectTaskSheet ws
    If lastFlagRow >= DATA_START_ROW Then _
        ws.Range(ws.Cells(DATA_START_ROW, COL_WEB_ENABLED), ws.Cells(lastFlagRow, COL_WEB_ENABLED)).Value2 = "N"
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, lastRow
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "ResetWebUpdateSelections", errorText
End Sub

Private Sub RecordWebOutcome(ByVal item As Object, ByVal outcome As String, ByVal requestNo As String, ByVal message As String, _
                             Optional ByVal operation As String = "등록", Optional ByVal confirmedValues As Object = Nothing)
    Dim log As Worksheet, r As Long, sourceRow As Long, wasProtected As Boolean
    EnsureWebLogSheet
    Set log = ThisWorkbook.Worksheets(WEB_LOG_SHEET_NAME)
    r = log.Cells(log.Rows.Count, "A").End(xlUp).Row + 1
    log.Cells(r, "A").Value = Now
    log.Cells(r, "B").Value2 = mPreviewSource.Name
    log.Cells(r, "C").Value2 = item("SourceId")
    log.Cells(r, "D").Value2 = requestNo
    log.Cells(r, "E").Value2 = outcome
    log.Cells(r, "F").Value2 = Left$(message, 32767)
    log.Cells(r, "G").Value2 = item("Title")
    log.Cells(r, "H").Value2 = item("Values")("ACT_DT_FR")
    log.Cells(r, "I").Value2 = item("Values")("ACT_DT")
    log.Cells(r, "J").Value2 = item("Values")("ACT_TIME")
    log.Cells(r, "K").Value2 = operation
    log.Cells(r, "L").Value2 = WebCurrentSite()
    log.Cells(r, "M").Value2 = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2
    If Len(requestNo) > 0 Then
        log.Hyperlinks.Add Anchor:=log.Cells(r, "D"), Address:=CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2) & "C0/c0200_02.aspx?REQ_NO=" & requestNo, TextToDisplay:=requestNo
    End If
    WebRecordState mPreviewSource, item, operation, outcome, requestNo, confirmedValues
    sourceRow = FindWebTaskRow(mPreviewSource, CStr(item("SourceId")))
    If sourceRow > 0 Then
        wasProtected = mPreviewSource.ProtectContents
        If wasProtected Then UnprotectTaskSheet mPreviewSource
        mPreviewSource.Range(COL_WEB_RESULT & sourceRow & ":" & COL_WEB_RESULT & GetTaskSubtreeEndRow(mPreviewSource, sourceRow, GetLastDataRow(mPreviewSource))).Value2 = outcome & IIf(Len(requestNo) > 0, " " & requestNo, "")
        If wasProtected Then ApplyCalculatedColumnsProtection mPreviewSource, GetLastDataRow(mPreviewSource)
    End If
End Sub

Public Sub UploadWebTasks(Optional ByVal showCompletionMessage As Boolean = True)
    Dim currentItems As Collection, item As Object, session As Object, form As Object, result As Object
    Dim requestNo As String, previous As String, count As Long, skipped As Long, outcome As String
    Dim eventsEnabled As Boolean, currentStarted As Boolean, errorText As String, errorNumber As Long
    Dim operation As String, state As Object, confirmedValues As Object, structureIssue As String, sourceProtected As Boolean, eligible As Long
    Dim newIssue As String, executionAccepted As Boolean, selectionsReset As Boolean
    If mWebBusy Then Exit Sub
    If ActiveSheet.Name <> WEB_PREVIEW_SHEET_NAME Then
        PreviewWebTasks showCompletionMessage
        Exit Sub
    End If
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    If mPreviewItems Is Nothing Or mPreviewSource Is Nothing Then Err.Raise vbObjectError + 2607, , "메인 시트에서 미리보기를 다시 만드세요."
    If PreviewDecisionSnapshot(mPreviewItems) <> mPreviewDecisionSnapshot Then _
        Err.Raise vbObjectError + 2607, , "미리보기의 실행 정보가 바뀌었습니다. 미리보기를 갱신하세요."
    AssertWebTaskBases mPreviewSource
    structureIssue = WebTaskStructureIssue(mPreviewSource)
    If Len(structureIssue) > 0 Then Err.Raise vbObjectError + 2607, , structureIssue
    Application.EnableEvents = False
    sourceProtected = mPreviewSource.ProtectContents
    If sourceProtected Then UnprotectTaskSheet mPreviewSource
    RefreshWebTaskPeriods mPreviewSource
    If sourceProtected Then ApplyCalculatedColumnsProtection mPreviewSource, GetLastDataRow(mPreviewSource)
    Application.EnableEvents = eventsEnabled
    Set currentItems = BuildWebUploadItems(mPreviewSource)
    If PreviewItemsSnapshot(currentItems) <> mPreviewSnapshot Then Err.Raise vbObjectError + 2607, , "미리보기 후 업무 내용이 바뀌었습니다. 미리보기를 갱신하세요."
    For Each item In currentItems
        AssessWebItem item, session
    Next item
    If PreviewDecisionSnapshot(currentItems) <> mPreviewDecisionSnapshot Then _
        Err.Raise vbObjectError + 2607, , "이력 또는 웹 내용이 미리보기 이후 바뀌었습니다. 미리보기를 갱신하세요."
    Set mPreviewItems = currentItems
    For Each item In mPreviewItems
        operation = CStr(item("Operation"))
        If Len(CStr(item("Error"))) = 0 And (operation = "신규" Or operation = "갱신") Then eligible = eligible + 1
    Next item
    If ThisWorkbook.ReadOnly Or Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 2607, , "등록 이력을 보관할 수 있도록 쓰기 가능한 파일로 먼저 저장하세요."
    executionAccepted = True
    mWebBusy = True
    Application.EnableEvents = False
    ' 확정한 실행 목록은 유지하고 선택값을 먼저 저장해 전송 중단 시 미처리 행도 초기화한다.
    ResetWebUpdateSelections mPreviewSource
    selectionsReset = True
    SaveWebJournal
    If eligible = 0 Then GoTo Completed
    If session Is Nothing Then Set session = OpenConfiguredWebSession()
    For Each item In mPreviewItems
        operation = CStr(item("Operation"))
        If Len(CStr(item("Error"))) > 0 Or (operation <> "신규" And operation <> "갱신") Then
            skipped = skipped + 1
        Else
            Set state = WebTaskState(CStr(item("SourceId")))
            requestNo = CStr(state("RequestNo"))
            If WebStateUnresolved(state) Or Len(CStr(state("PendingNo"))) > 0 Then _
                Err.Raise vbObjectError + 2607, , "앞선 전송 결과를 먼저 확인하세요."
            If operation = "신규" Then
                newIssue = NewWebTaskIssue(CStr(item("SourceId")), state)
                If Len(newIssue) > 0 Then Err.Raise vbObjectError + 2607, , newIssue
                Set form = WebLoadRegistration(session)
                requestNo = CStr(form("RequestNo"))
                RecordWebOutcome item, "등록 시작", requestNo, "저장 결과 확인 중"
            Else
                If Len(requestNo) = 0 Or requestNo <> CStr(item("RequestNo")) Then _
                    Err.Raise vbObjectError + 2607, , "미리보기 이후 요청번호 연결이 바뀌었습니다. 미리보기를 갱신하세요."
                Set form = WebLoadExisting(session, requestNo)
                If WebValuesFingerprint(form("Values")) <> CStr(item("ServerSnapshot")) Or _
                   WebValuesFingerprint(form("Values")) <> CStr(state("Baseline")) Then _
                    Err.Raise vbObjectError + 2607, , "미리보기 이후 웹 내용이 바뀌었습니다. 다시 확인하세요: " & requestNo
                RecordWebOutcome item, "갱신 시작", requestNo, "기존 요청번호로 저장 결과 확인 중", "갱신"
            End If
                Application.StatusBar = "프로젝트 실적 " & operation & ": " & CStr(item("Title"))
                SaveWebJournal
                currentStarted = True
                If operation = "신규" Then
                    Set result = WebPostRegistration(session, form, item("Values"))
                Else
                    Set result = WebPostUpdate(session, form, item("Values"))
                End If
                Select Case CStr(result("Outcome"))
                    Case "success": outcome = IIf(operation = "신규", "등록완료", "갱신완료")
                    Case "failed": outcome = IIf(operation = "신규", "등록실패", "갱신실패")
                    Case "partial": outcome = "갱신 일부저장"
                    Case Else: outcome = IIf(operation = "신규", "결과 미확인", "갱신 결과 미확인")
                End Select
                Set confirmedValues = Nothing
                If result.Exists("Values") Then Set confirmedValues = result("Values")
                RecordWebOutcome item, outcome, CStr(result("RequestNo")), CStr(result("Message")), IIf(operation = "신규", "등록", "갱신"), confirmedValues
                SaveWebJournal
                currentStarted = False
                If outcome = "등록완료" Or outcome = "갱신완료" Then count = count + 1
                If outcome = "결과 미확인" Or outcome = "갱신 결과 미확인" Or outcome = "갱신 일부저장" Then Exit For
        End If
    Next item
Completed:
    mWebBusy = False
    Application.EnableEvents = eventsEnabled
    Application.StatusBar = False
    PreviewWebTasks showCompletionMessage
    executionAccepted = False
    If showCompletionMessage Then
        If eligible = 0 Then
            MsgBox "실행 가능한 업무가 없어 전송하지 않았습니다. WEB UPDATE는 모두 N으로 초기화했습니다." & vbCrLf & "확인 사항을 해결한 뒤 대상 업무를 다시 Y로 선택하세요.", vbInformation, "웹등록"
        Else
            MsgBox "등록/갱신 완료 " & count & "건 / 제외 " & skipped & "건" & vbCrLf & "WEB UPDATE는 모두 N으로 초기화했습니다. 자세한 결과는 웹등록_이력 시트에서 확인하세요.", vbInformation, "웹등록 결과"
        End If
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If sourceProtected And Not mPreviewSource Is Nothing Then ApplyCalculatedColumnsProtection mPreviewSource, GetLastDataRow(mPreviewSource)
    If currentStarted Then
        RecordWebOutcome item, IIf(operation = "신규", "결과 미확인", "갱신 결과 미확인"), requestNo, "전송 후 결과를 확인하지 못했습니다. 웹에서 먼저 확인하세요.", IIf(operation = "신규", "등록", "갱신")
        SaveWebJournal
    End If
    If executionAccepted Then
        Set mPreviewItems = Nothing
        mPreviewSnapshot = vbNullString
        mPreviewDecisionSnapshot = vbNullString
    End If
    mWebBusy = False
    Application.EnableEvents = eventsEnabled
    Application.StatusBar = False
    If selectionsReset Then PreviewWebTasks False
    On Error GoTo 0
    If showCompletionMessage Then
        MsgBox errorText, vbExclamation, "웹등록"
    Else
        Err.Raise errorNumber, "UploadWebTasks", errorText
    End If
End Sub

Private Function SelectedWebLinkItem() As Object
    Dim ws As Worksheet, rowNum As Long
    If IsWebTaskSheet(ActiveSheet) Then
        Set ws = ActiveSheet
        rowNum = ActiveCell.Row
    ElseIf Not mConfigTaskSheet Is Nothing Then
        Set ws = mConfigTaskSheet
        rowNum = mConfigTaskRow
    Else
        Err.Raise vbObjectError + 2609, , "메인 시트에서 업무 행을 선택하고 웹 설정을 여세요."
    End If
    PrepareWebTaskSheet ws
    If rowNum < DATA_START_ROW Then Err.Raise vbObjectError + 2609, , "연결할 업무 행을 먼저 선택하세요."
    Set mPreviewSource = ws
    Set SelectedWebLinkItem = BuildWebItemForRow(ws, rowNum)
End Function

Public Sub LinkSelectedWebRequest()
    Dim item As Object, requestNo As Variant, session As Object, server As Object, state As Object
    Dim field As Variant, summary As String
    On Error GoTo Failed
    Set item = SelectedWebLinkItem()
    Set state = WebTaskState(CStr(item("SourceId")))
    requestNo = Application.InputBox("선택한 업무: " & CStr(item("Title")) & vbCrLf & "연결할 웹 요청번호를 입력하세요.", "기존 요청번호 연결", CStr(state("RequestNo")), Type:=2)
    If VarType(requestNo) = vbBoolean Then Exit Sub
    If Len(Trim$(CStr(requestNo))) = 0 Then Exit Sub
    Set session = OpenConfiguredWebSession()
    ConnectWebRequest mPreviewSource, CLng(item("Row")), Trim$(CStr(requestNo)), session, True
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "요청번호 연결"
End Sub

' The interactive wrapper always displays the fetched record before binding or
' accepting its current web values as the conflict baseline.
Public Sub ConnectWebRequest(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal requestNo As String, _
                             ByVal session As Object, Optional ByVal confirmWithUser As Boolean = True, _
                             Optional ByVal approvedServerFingerprint As String = "")
    Dim item As Object, state As Object, server As Object, detail As String, field As Variant, wasProtected As Boolean
    Dim eventsEnabled As Boolean, requesterLabel As String
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    Set item = BuildWebItemForRow(ws, rowNum)
    If Len(CStr(item("SourceId"))) = 0 Then Err.Raise vbObjectError + 2609, , "업무 식별값이 없습니다."
    If InStr(CStr(item("Error")), "중복") > 0 Then Err.Raise vbObjectError + 2609, , CStr(item("Error"))
    If WebRequestAlreadyLinked(CStr(item("SourceId")), requestNo) Then Err.Raise vbObjectError + 2609, , "이 요청번호는 다른 업무에 이미 연결되어 있습니다."
    Set state = WebTaskState(CStr(item("SourceId")))
    If Len(CStr(state("RequestNo"))) > 0 And CStr(state("RequestNo")) <> requestNo Then _
        Err.Raise vbObjectError + 2609, , "이미 확정된 요청번호를 다른 요청번호로 교체할 수 없습니다."
    Set server = WebReadExistingValues(session, requestNo)
    For Each field In WebManagedFields()
        If Not server.Exists(CStr(field)) Then Err.Raise vbObjectError + 2609, , "요청번호 상세값이 부족합니다: " & CStr(field)
    Next field
    If CStr(server("GUBUN")) <> "PJT" Then Err.Raise vbObjectError + 2609, , "프로젝트 실적 요청번호만 연결할 수 있습니다."
    If confirmWithUser Then
        ShowWebConnectionReview ws, item, requestNo, server
        Exit Sub
    End If
    If Len(approvedServerFingerprint) > 0 Then
        If WebValuesFingerprint(server) <> approvedServerFingerprint Then Err.Raise vbObjectError + 2609, , "확인 화면을 만든 뒤 웹 내용이 바뀌었습니다. 요청번호 연결을 다시 검토하세요."
    End If
    Set mPreviewSource = ws
    Application.EnableEvents = False
    wasProtected = ws.ProtectContents
    If wasProtected Then UnprotectTaskSheet ws
    requesterLabel = WebOptionLabel("REQ_UID", CStr(server("REQ_UID")))
    If Len(requesterLabel) = 0 Then requesterLabel = "[웹 요청자: " & CStr(server("REQ_UID")) & "]"
    ws.Range(COL_WEB_REQUESTER & CLng(item("Row")) & ":" & COL_WEB_REQUESTER & GetTaskSubtreeEndRow(ws, CLng(item("Row")), GetLastDataRow(ws))).Value2 = requesterLabel
    RememberWebGroupBases ws
    RecordWebOutcome item, "연결완료", requestNo, "요청번호 상세 확인 및 현재 웹 비교값 저장", "연결", server
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    SaveWebJournal
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "ConnectWebRequest", errorText
End Sub

Private Sub ShowWebConnectionReview(ByVal ws As Worksheet, ByVal item As Object, ByVal requestNo As String, ByVal server As Object)
    Dim review As Worksheet, field As Variant, r As Long
    Set mLinkSource = ws
    mLinkSourceId = CStr(item("SourceId"))
    mLinkRequestNo = requestNo
    mLinkServerFingerprint = WebValuesFingerprint(server)
    mLinkPayloadFingerprint = WebValuesFingerprint(item("Values"))
    mLinkSite = WebCurrentSite()
    mLinkAccount = CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2)
    Set mPreviewItems = Nothing
    mPreviewSnapshot = vbNullString
    mPreviewDecisionSnapshot = vbNullString
    Set review = WebSheet(WEB_PREVIEW_SHEET_NAME)
    review.Unprotect
    If review.AutoFilterMode Then review.AutoFilterMode = False
    review.Cells.Clear
    review.Cells.Locked = True
    On Error Resume Next
    review.Shapes("btnWebUploadPreview").Visible = msoFalse
    review.Shapes("btnWebPreviewRefresh").Visible = msoFalse
    On Error GoTo 0
    review.Range("A1").Value2 = "요청번호 연결 확인: " & requestNo
    review.Range("A2").Value2 = "웹 실적과 엑셀 업무를 비교하세요. 연결 확정은 현재 웹값을 비교 기준으로 저장하며, 이후 웹등록 미리보기에서 엑셀 변경값을 갱신합니다."
    review.Range("A4").Value2 = "항목"
    review.Range("B4").Value2 = "현재 웹값"
    review.Range("C4").Value2 = "현재 엑셀값"
    review.Columns("A:C").NumberFormat = "@"
    r = 5
    For Each field In WebManagedFields()
        review.Cells(r, 1).Value2 = WebManagedFieldTitle(CStr(field))
        review.Cells(r, 2).Value2 = Left$(CStr(server(CStr(field))), 32767)
        review.Cells(r, 3).Value2 = Left$(CStr(item("Values")(CStr(field))), 32767)
        If CStr(server(CStr(field))) <> CStr(item("Values")(CStr(field))) Then review.Range("B" & r & ":C" & r).Interior.Color = RGB(255, 246, 178)
        r = r + 1
    Next field
    review.Cells(r + 1, 1).Value2 = "웹 상세 열기"
    review.Hyperlinks.Add review.Cells(r + 1, 2), WebCurrentSite() & "C0/c0200_02.aspx?REQ_NO=" & requestNo, TextToDisplay:=requestNo
    review.Cells(r + 2, 1).Value2 = "요청자는 웹값으로 고정됩니다. 긴 본문은 셀 표시 한도 32,767자까지이며 웹 상세에서도 확인할 수 있습니다."
    review.Columns("A").ColumnWidth = 22
    review.Columns("B:C").ColumnWidth = 70
    review.Range("A4:C4").Font.Bold = True
    review.Range("A4:C4").Interior.Color = RGB(229, 239, 249)
    review.Range("B5:C" & r - 1).WrapText = True
    review.Range("A5:C" & r - 1).VerticalAlignment = xlTop
    review.Rows("5:" & r - 1).AutoFit
    AddWebButton review, "btnWebLinkConfirm", "연결 확정", "ConfirmReviewedWebConnection", review.Range("A3").Left, review.Range("A3").Top, 105
    AddWebButton review, "btnWebLinkCancel", "취소", "CancelWebConnectionReview", review.Range("B3").Left, review.Range("B3").Top, 80
    review.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    review.Activate
End Sub

Private Function WebManagedFieldTitle(ByVal field As String) As String
    Select Case field
        Case "GUBUN": WebManagedFieldTitle = "요청구분"
        Case "REQ_UID": WebManagedFieldTitle = "요청자 ID"
        Case "REQ_TIT": WebManagedFieldTitle = "제목"
        Case "REQ_RMK": WebManagedFieldTitle = "요청 내용"
        Case "STS": WebManagedFieldTitle = "진행상태"
        Case "PGM_MD": WebManagedFieldTitle = "모듈/업무"
        Case "ACT_GUBUN": WebManagedFieldTitle = "처리구분"
        Case "ACT_UID": WebManagedFieldTitle = "처리자 ID"
        Case "PGM_NM": WebManagedFieldTitle = "프로그램명"
        Case "PGM_ID": WebManagedFieldTitle = "프로그램ID"
        Case "ACT_DT_FR": WebManagedFieldTitle = "실제 시작일"
        Case "ACT_DT": WebManagedFieldTitle = "실제 종료일"
        Case "ACT_TIME": WebManagedFieldTitle = "작업시간(분)"
    End Select
End Function

Public Sub ConfirmReviewedWebConnection()
    Dim rowNum As Long, item As Object, session As Object
    On Error GoTo Failed
    If mLinkSource Is Nothing Or Len(mLinkSourceId) = 0 Then Err.Raise vbObjectError + 2609, , "요청번호 연결 확인 화면을 다시 만드세요."
    If WebCurrentSite() <> mLinkSite Or CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2) <> mLinkAccount Then _
        Err.Raise vbObjectError + 2609, , "연결 검토 이후 사이트나 계정이 바뀌었습니다. 다시 검토하세요."
    rowNum = WebFindTaskRow(mLinkSource, mLinkSourceId)
    If rowNum = 0 Then Err.Raise vbObjectError + 2609, , "연결할 업무가 없어졌습니다."
    Set item = BuildWebItemForRow(mLinkSource, rowNum)
    If CStr(item("SourceId")) <> mLinkSourceId Or WebValuesFingerprint(item("Values")) <> mLinkPayloadFingerprint Then _
        Err.Raise vbObjectError + 2609, , "연결 검토 이후 엑셀 업무가 바뀌었습니다. 다시 검토하세요."
    Set session = OpenConfiguredWebSession()
    ConnectWebRequest mLinkSource, rowNum, mLinkRequestNo, session, False, mLinkServerFingerprint
    mLinkSource.Activate
    mLinkSourceId = vbNullString
    MsgBox "요청번호 연결과 웹 비교값을 저장했습니다. 웹등록 미리보기에서 변경 내용을 확인하세요.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "연결 확정"
End Sub

Public Sub CancelWebConnectionReview()
    mLinkSourceId = vbNullString
    If Not mLinkSource Is Nothing Then mLinkSource.Activate
End Sub

Public Sub ResolveSelectedWebAttempt()
    Dim item As Object, state As Object, session As Object, server As Object, requestNo As String, readError As String
    On Error GoTo Failed
    Set item = SelectedWebLinkItem()
    Set state = WebTaskState(CStr(item("SourceId")))
    If Not WebStateUnresolved(state) Then Err.Raise vbObjectError + 2611, , "미확인 전송이 없는 업무입니다."
    requestNo = CStr(state("RequestNo"))
    If Len(requestNo) = 0 Then requestNo = CStr(state("PendingNo"))
    Set session = OpenConfiguredWebSession()
    On Error Resume Next
    Set server = WebReadExistingValues(session, requestNo)
    readError = Err.Description
    Err.Clear
    On Error GoTo Failed
    If Not server Is Nothing Then
        ConnectWebRequest mPreviewSource, CLng(item("Row")), requestNo, session, True
    ElseIf Len(CStr(state("RequestNo"))) > 0 Then
        Err.Raise vbObjectError + 2611, , "확정 요청번호의 조회에 실패했습니다. 기존 연결을 유지합니다: " & readError
    Else
        If MsgBox("후보 요청번호 " & requestNo & "를 조회하지 못했습니다." & vbCrLf & readError & vbCrLf & vbCrLf & _
                  "조회 실패만으로 미등록을 판단할 수 없습니다. 웹의 프로젝트 실적 목록에서 직접 검색하여 이 업무가 등록되지 않은 것을 확인했습니까?" & vbCrLf & _
                  "[예]를 누르면 미등록 확인 이력을 저장하고 신규 등록을 다시 허용합니다.", _
                  vbYesNo + vbQuestion + vbDefaultButton2, "웹 미등록 직접 확인") = vbYes Then
            ConfirmWebAttemptUnregistered mPreviewSource, CLng(item("Row")), True
        End If
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "전송결과 확인"
End Sub

Public Sub ConfirmWebAttemptUnregistered(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal confirmedByUser As Boolean)
    Dim item As Object, state As Object
    If Not confirmedByUser Then Err.Raise vbObjectError + 2611, , "웹 목록에서 미등록을 직접 확인해야 합니다."
    Set item = BuildWebItemForRow(ws, rowNum)
    Set state = WebTaskState(CStr(item("SourceId")))
    If Len(CStr(state("RequestNo"))) > 0 Then Err.Raise vbObjectError + 2611, , "확정 요청번호 연결은 미등록으로 해제할 수 없습니다."
    If Not WebStateUnresolved(state) Then Err.Raise vbObjectError + 2611, , "미확인 전송이 없는 업무입니다."
    Set mPreviewSource = ws
    RecordWebOutcome item, "미등록 확인", "", "사용자가 웹 목록을 직접 확인하여 미등록으로 확정", "확인"
    SaveWebJournal
End Sub
