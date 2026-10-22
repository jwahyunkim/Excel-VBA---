Attribute VB_Name = "modWebUpload"
Option Explicit
' 2026-10-22 김좌현 수정 : 프로젝트 실적 웹 설정, 선택목록, 미리보기와 등록 이력 추가

Private mPreviewItems As Collection
Private mPreviewSource As Worksheet
Private mPreviewSnapshot As String
Private mWebBusy As Boolean
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
        Case WEB_CONFIG_SHEET_NAME, WEB_PREVIEW_SHEET_NAME, WEB_LOG_SHEET_NAME
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
    ws.Range("A7").Value2 = "새 업무의 웹등록 선택"
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
    ws.Range("A16").Value2 = "등록 기준은 레벨1 그룹 전체에 적용합니다. 등록할 대상 행의 웹등록을 Y로 선택하세요."
    ws.Range("A17").Value2 = "작업시간은 분 단위입니다. 실제 날짜와 작업시간을 비워 둔 행은 등록하지 않습니다."
    ws.Range("A18").Value2 = "결과 미확인은 웹에서 먼저 확인한 뒤 이력 상태를 '미등록 확인'으로 바꾸면 다시 등록할 수 있습니다."
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
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & WebOptionName(CStr(field))
            .IgnoreBlank = True
        End With
    Next field
    AddWebButton ws, "btnWebLogin", "로그인 확인", "TestWebLogin", ws.Range("A13").Left, ws.Range("A13").Top, 105
    AddWebButton ws, "btnWebOptions", "선택목록 갱신", "RefreshWebOptions", ws.Range("B13").Left, ws.Range("B13").Top, 120
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
    Dim errNumber As Long, errText As String
    If ws Is Nothing Or IsWebUtilitySheet(ws) Then Exit Sub
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
            If Len(taskId) = 0 Then
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
                If Len(CStr(levelValue)) = 0 Then levelValue = config.Range("B6").Value2
            End If
            If rootRow > 0 Then ws.Cells(r, COL_WEB_LEVEL).Value2 = levelValue
            If Len(CStr(ws.Cells(r, COL_WEB_ENABLED).Value2)) = 0 Then ws.Cells(r, COL_WEB_ENABLED).Value2 = config.Range("B7").Value2
            For Each field In Array("REQ_UID", "ACT_GUBUN", "STS", "PGM_MD", "ACT_UID")
                If Len(CStr(ws.Cells(r, WebTaskOptionColumn(CStr(field))).Value2)) = 0 Then
                    ws.Cells(r, WebTaskOptionColumn(CStr(field))).Value2 = config.Range(WebDefaultCell(CStr(field))).Value2
                End If
            Next field
        End If
    Next r
    ApplyWebTaskValidation ws
    ws.Columns(COL_WEB_ENABLED & ":" & COL_WEB_PROGRAM_ID).ColumnWidth = 14
    ws.Columns(COL_WEB_REQUESTER).ColumnWidth = 25
    ws.Columns(COL_WEB_HANDLER).ColumnWidth = 25
    ws.Columns(COL_WEB_RESULT).ColumnWidth = 28
    ws.Columns(COL_WORK_MINUTES).ColumnWidth = 13
    EnsureWebTaskButtons ws
    ws.Columns(COL_WEB_TASK_ID).Hidden = True
    If wasProtected Then ApplyCalculatedColumnsProtection ws, Application.Max(DATA_START_ROW, lastRow)
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errNumber = Err.Number: errText = Err.Description
    On Error Resume Next
    If wasProtected Then ApplyCalculatedColumnsProtection ws, Application.Max(DATA_START_ROW, lastRow)
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errNumber, "PrepareWebTaskSheet", errText
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
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & WebOptionName(CStr(field))
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
    If Not IsWebTaskSheet(ws) Then Exit Sub
    Set changed = Intersect(target, ws.Range(COL_WEB_LEVEL & DATA_START_ROW & ":" & COL_WEB_LEVEL & ws.Rows.Count))
    If changed Is Nothing Then Exit Sub
    Set groups = CreateObject("Scripting.Dictionary")
    lastRow = GetLastDataRow(ws)
    Set changed = Intersect(changed, ws.Rows(DATA_START_ROW & ":" & Application.Max(DATA_START_ROW, lastRow)))
    If changed Is Nothing Then Exit Sub
    For Each cell In changed.Cells
        rootRow = 0
        For r = cell.Row To DATA_START_ROW Step -1
            If HasTaskContent(ws, r) And GetTaskLevel(ws, r) = 1 Then
                rootRow = r
                Exit For
            End If
        Next r
        If rootRow > 0 Then
            If CStr(cell.Value2) <> "1" And CStr(cell.Value2) <> "2" And CStr(cell.Value2) <> "3" Then
                Application.Undo
                Err.Raise vbObjectError + 2605, , "등록 기준은 1, 2, 3 중에서 선택하세요."
            End If
            If groups.Exists(CStr(rootRow)) Then
                If groups(CStr(rootRow)) <> CStr(cell.Value2) Then
                    Application.Undo
                    Err.Raise vbObjectError + 2605, , "한 레벨1 그룹에는 같은 등록 기준을 적용해야 합니다."
                End If
            Else
                groups.Add CStr(rootRow), CStr(cell.Value2)
            End If
        End If
    Next cell
    For Each key In groups.Keys
        rootRow = CLng(key)
        ws.Range(COL_WEB_LEVEL & rootRow & ":" & COL_WEB_LEVEL & GetTaskSubtreeEndRow(ws, rootRow, lastRow)).Value2 = groups(key)
    Next key
End Sub

Public Sub InitializeWebUpload()
    Dim ws As Worksheet, eventsEnabled As Boolean, previousSheet As Object, errorNumber As Long, errorText As String
    eventsEnabled = Application.EnableEvents
    Set previousSheet = ThisWorkbook.ActiveSheet
    On Error GoTo Failed
    Application.EnableEvents = False
    EnsureWebConfigSheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsWebUtilitySheet(ws) And ws.Name <> CONFIG_SHEET_NAME And ws.Name <> OUTPUT_CONFIG_SHEET_NAME Then
            If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." Then
                EnsureTaskSheetLayout ws
                PrepareWebTaskSheet ws
            End If
        End If
    Next ws
    EnsureWebLogSheet
    previousSheet.Activate
    Application.EnableEvents = eventsEnabled
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    previousSheet.Activate
    Application.EnableEvents = eventsEnabled
    On Error GoTo 0
    Err.Raise errorNumber, "InitializeWebUpload", errorText
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
    AddWebButton ws, "btnWebUpload", "웹등록 실행", "UploadWebTasks", ws.Range("B3").Left + 118, ws.Range("B3").Top, 96
    AddWebButton ws, "btnWebConfig", "웹 설정", "OpenWebConfig", ws.Range("B3").Left + 220, ws.Range("B3").Top, 80
End Sub

Public Sub OpenWebConfig()
    EnsureWebConfigSheet
    ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Activate
End Sub

Private Sub EnsureWebLogSheet()
    Dim ws As Worksheet, headers As Variant, i As Long
    Set ws = WebSheet(WEB_LOG_SHEET_NAME)
    headers = Array("기록 시각", "업무 시트", "업무 ID", "웹 등록번호", "상태", "결과", "제목", "실제 시작일", "실제 종료일", "작업시간(분)")
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
End Sub

Private Function PreviousWebOutcome(ByVal sourceId As String, ByRef requestNo As String) As String
    Dim ws As Worksheet, r As Long
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

Private Function PreviewItemsSnapshot(ByVal items As Collection) As String
    Dim item As Object, field As Variant, output As String, value As String
    output = CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2) & vbNullChar & _
             CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B3").Value2) & vbNullChar
    For Each item In items
        value = CStr(item("SourceId")) & "|" & CStr(item("Error"))
        output = output & Len(value) & ":" & value
        For Each field In Array("GUBUN", "REQ_UID", "REQ_TIT", "REQ_RMK", "STS", "PGM_MD", "ACT_GUBUN", "ACT_UID", "PGM_NM", "PGM_ID", "ACT_DT_FR", "ACT_DT", "ACT_TIME")
            If item("Values").Exists(CStr(field)) Then value = CStr(item("Values")(CStr(field))) Else value = vbNullString
            output = output & Len(value) & ":" & value
        Next field
    Next item
    PreviewItemsSnapshot = output
End Function

Public Sub PreviewWebTasks(Optional ByVal showCompletionMessage As Boolean = True)
    Dim ws As Worksheet, preview As Worksheet, item As Object, r As Long, i As Long, headers As Variant
    Dim requestNo As String, outcome As String, issue As String, periodFields As Variant, errorNumber As Long, errorText As String
    On Error GoTo Failed
    If mWebBusy Then Exit Sub
    If ActiveSheet.Name = WEB_PREVIEW_SHEET_NAME Then
        If mPreviewSource Is Nothing Then Err.Raise vbObjectError + 2606, , "메인 업무 시트에서 미리보기를 다시 만드세요."
        Set ws = mPreviewSource
    Else
        Set ws = ActiveSheet
    End If
    If Not IsWebTaskSheet(ws) Then Err.Raise vbObjectError + 2606, , "메인 업무 시트에서 실행하세요."
    PrepareWebTaskSheet ws
    Set mPreviewSource = ws
    Set mPreviewItems = BuildWebUploadItems(ws)
    mPreviewSnapshot = PreviewItemsSnapshot(mPreviewItems)
    Set preview = WebSheet(WEB_PREVIEW_SHEET_NAME)
    preview.Unprotect
    If preview.AutoFilterMode Then preview.AutoFilterMode = False
    preview.Cells.Clear
    preview.Cells.Font.Name = "맑은 고딕"
    preview.Cells.Font.Size = 10
    preview.Columns("E:I").NumberFormat = "@"
    preview.Rows(3).RowHeight = 21
    preview.Range("A1").Value2 = "프로젝트 실적 등록 미리보기"
    preview.Range("A2").Value2 = "원본: " & ws.Name & " / 확인 후 웹등록 실행 버튼을 누르세요. 수정이 필요하면 메인 시트에서 고친 뒤 다시 미리보기하세요."
    headers = Array("상태", "원본 행", "제목", "요청/처리내용", "요청자", "처리구분", "진행상태", "모듈/업무", "처리자", "시작일", "종료일", "시간(분)", "확인 사항")
    periodFields = Array("ACT_DT_FR", "ACT_DT", "ACT_TIME")
    For i = 0 To UBound(headers)
        preview.Cells(4, i + 1).Value2 = headers(i)
    Next i
    r = 5
    For Each item In mPreviewItems
        issue = CStr(item("Error"))
        requestNo = vbNullString
        outcome = PreviousWebOutcome(CStr(item("SourceId")), requestNo)
        Select Case outcome
            Case "등록완료": issue = "이미 등록한 업무입니다: " & requestNo
            Case "결과 미확인", "등록 시작": issue = "앞선 등록 결과를 웹에서 먼저 확인하세요: " & requestNo
        End Select
        preview.Cells(r, 1).Value2 = IIf(Len(issue) = 0, "등록 가능", "확인 필요")
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
        If Len(issue) > 0 Then preview.Range(preview.Cells(r, 1), preview.Cells(r, 13)).Interior.Color = RGB(255, 239, 224)
        r = r + 1
    Next item
    If mPreviewItems.Count = 0 Then preview.Range("A5").Value2 = "등록할 행이 없습니다. 메인 시트에서 등록 대상 행의 웹등록을 Y로 선택하세요."
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
    preview.Range("C5:M" & Application.Max(5, r - 1)).WrapText = True
    preview.Range("A5:M" & Application.Max(5, r - 1)).VerticalAlignment = xlTop
    preview.Rows("5:" & Application.Max(5, r - 1)).AutoFit
    AddWebButton preview, "btnWebUploadPreview", "웹등록 실행", "UploadWebTasks", preview.Range("A3").Left, preview.Range("A3").Top, 105
    AddWebButton preview, "btnWebPreviewRefresh", "미리보기 갱신", "PreviewWebTasks", preview.Range("C3").Left, preview.Range("C3").Top, 112
    preview.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    preview.Activate
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
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

Private Sub RecordWebOutcome(ByVal item As Object, ByVal outcome As String, ByVal requestNo As String, ByVal message As String)
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
    If Len(requestNo) > 0 Then
        log.Hyperlinks.Add Anchor:=log.Cells(r, "D"), Address:=CStr(ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME).Range("B2").Value2) & "C0/c0200_02.aspx?REQ_NO=" & requestNo, TextToDisplay:=requestNo
    End If
    sourceRow = FindWebTaskRow(mPreviewSource, CStr(item("SourceId")))
    If sourceRow > 0 Then
        wasProtected = mPreviewSource.ProtectContents
        If wasProtected Then UnprotectTaskSheet mPreviewSource
        mPreviewSource.Cells(sourceRow, COL_WEB_RESULT).Value2 = outcome & IIf(Len(requestNo) > 0, " " & requestNo, "")
        If wasProtected Then ApplyCalculatedColumnsProtection mPreviewSource, GetLastDataRow(mPreviewSource)
    End If
End Sub

Public Sub UploadWebTasks(Optional ByVal showCompletionMessage As Boolean = True)
    Dim currentItems As Collection, item As Object, session As Object, form As Object, result As Object
    Dim requestNo As String, previous As String, count As Long, skipped As Long, outcome As String
    Dim eventsEnabled As Boolean, currentStarted As Boolean, errorText As String, errorNumber As Long
    If mWebBusy Then Exit Sub
    If ActiveSheet.Name <> WEB_PREVIEW_SHEET_NAME Then
        PreviewWebTasks showCompletionMessage
        Exit Sub
    End If
    eventsEnabled = Application.EnableEvents
    On Error GoTo Failed
    If mPreviewItems Is Nothing Or mPreviewSource Is Nothing Then Err.Raise vbObjectError + 2607, , "메인 시트에서 미리보기를 다시 만드세요."
    Set currentItems = BuildWebUploadItems(mPreviewSource)
    If PreviewItemsSnapshot(currentItems) <> mPreviewSnapshot Then Err.Raise vbObjectError + 2607, , "미리보기 후 업무 내용이 바뀌었습니다. 미리보기를 갱신하세요."
    If mPreviewItems.Count = 0 Then Exit Sub
    If ThisWorkbook.ReadOnly Or Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 2607, , "등록 이력을 보관할 수 있도록 쓰기 가능한 파일로 먼저 저장하세요."
    Set session = OpenConfiguredWebSession()
    mWebBusy = True
    Application.EnableEvents = False
    For Each item In mPreviewItems
        If Len(CStr(item("Error"))) > 0 Then
            skipped = skipped + 1
        Else
            requestNo = vbNullString
            previous = PreviousWebOutcome(CStr(item("SourceId")), requestNo)
            If previous = "등록완료" Or previous = "결과 미확인" Or previous = "등록 시작" Then
                skipped = skipped + 1
            Else
                Application.StatusBar = "프로젝트 실적 등록: " & CStr(item("Title"))
                Set form = WebLoadRegistration(session)
                requestNo = CStr(form("RequestNo"))
                RecordWebOutcome item, "등록 시작", requestNo, "저장 결과 확인 중"
                SaveWebJournal
                currentStarted = True
                Set result = WebPostRegistration(session, form, item("Values"))
                Select Case CStr(result("Outcome"))
                    Case "success": outcome = "등록완료"
                    Case "failed": outcome = "등록실패"
                    Case Else: outcome = "결과 미확인"
                End Select
                RecordWebOutcome item, outcome, CStr(result("RequestNo")), CStr(result("Message"))
                SaveWebJournal
                currentStarted = False
                If outcome = "등록완료" Then count = count + 1
                If outcome = "결과 미확인" Then Exit For
            End If
        End If
    Next item
    mWebBusy = False
    Application.EnableEvents = eventsEnabled
    Application.StatusBar = False
    PreviewWebTasks showCompletionMessage
    If showCompletionMessage Then MsgBox "등록 완료 " & count & "건 / 제외 " & skipped & "건" & vbCrLf & "자세한 결과는 웹등록_이력 시트에서 확인하세요.", vbInformation, "웹등록 결과"
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If currentStarted Then
        RecordWebOutcome item, "결과 미확인", requestNo, "등록 후 결과를 확인하지 못했습니다. 웹에서 등록 여부를 먼저 확인하세요."
        SaveWebJournal
    End If
    mWebBusy = False
    Application.EnableEvents = eventsEnabled
    Application.StatusBar = False
    On Error GoTo 0
    If showCompletionMessage Then
        MsgBox errorText, vbExclamation, "웹등록"
    Else
        Err.Raise errorNumber, "UploadWebTasks", errorText
    End If
End Sub
