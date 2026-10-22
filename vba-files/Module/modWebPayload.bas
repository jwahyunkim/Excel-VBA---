Attribute VB_Name = "modWebPayload"
Option Explicit
' 2026-10-22 김좌현 수정 : 레벨1 그룹 기준 프로젝트 실적 등록 대상 및 TXT 형식 본문 생성
' 2026-10-22 김좌현 수정 : 요청번호 연결용 단일 등록 단위 검증 공개

' Reading a preview never changes task dates, progress, hierarchy or report status.
' The configuration module prepares persistent task IDs before calling this module.
Public Function BuildWebUploadItems(ByVal ws As Worksheet) As Collection
    Dim items As New Collection, identityCounts As Object
    Dim item As Object
    Dim lastRow As Long, r As Long, groupRow As Long, groupEnd As Long
    Dim groupLevel As Long, candidateCount As Long, selectedCount As Long
    Dim hasSelection As Boolean

    If ws Is Nothing Then Err.Raise vbObjectError + 7801, "BuildWebUploadItems", "업무 시트를 지정하세요."
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = OUTPUT_CONFIG_SHEET_NAME Or _
       ws.Name = WEB_CONFIG_SHEET_NAME Or ws.Name = WEB_PREVIEW_SHEET_NAME Or _
       ws.Name = WEB_LOG_SHEET_NAME Or ws.Name = WEB_STATE_SHEET_NAME Then _
        Err.Raise vbObjectError + 7801, "BuildWebUploadItems", "업무 시트에서 실행하세요."
    Set identityCounts = CountWebTaskIds(ws)
    lastRow = GetLastDataRow(ws)
    r = DATA_START_ROW
    Do While r <= lastRow
        If Not WebHasTaskContent(ws, r) Then
            r = r + 1
        ElseIf GetTaskLevel(ws, r) <> 1 Then
            If WebRowEnabled(ws, r) Then
                Set item = NewWebUploadItem(ws, r, 0, lastRow, identityCounts)
                AppendWebItemError item, "상위 레벨1 그룹이 없습니다."
                items.Add item
            End If
            r = r + 1
        Else
            groupRow = r
            groupEnd = GetTaskSubtreeEndRow(ws, groupRow, lastRow)
            groupLevel = WebGroupLevel(ws, groupRow)
            hasSelection = False
            candidateCount = 0
            selectedCount = 0
            For r = groupRow To groupEnd
                If WebRowEnabled(ws, r) Then hasSelection = True
                If IsWebUploadRow(ws, r, groupLevel, groupEnd) Then
                    candidateCount = candidateCount + 1
                    If WebRowEnabled(ws, r) Then
                        Set item = NewWebUploadItem(ws, r, groupRow, lastRow, identityCounts)
                        If WebExactTaskLevel(ws, groupRow) <> 1 Then _
                            AppendWebItemError item, "그룹의 레벨1 값이 올바르지 않습니다."
                        items.Add item
                        selectedCount = selectedCount + 1
                    End If
                End If
            Next r
            If hasSelection And selectedCount = 0 Then
                Set item = NewWebUploadItem(ws, groupRow, groupRow, lastRow, identityCounts)
                If groupLevel = 0 Then
                    AppendWebItemError item, "등록 기준 레벨은 1, 2, 3 중에서 선택하세요."
                ElseIf candidateCount = 0 Then
                    AppendWebItemError item, "이 그룹에는 등록 기준 레벨 " & CStr(groupLevel) & "에 해당하는 업무가 없습니다."
                Else
                    AppendWebItemError item, "등록 기준에 해당하는 행의 업로드 대상을 Y로 선택하세요."
                End If
                items.Add item
            End If
            r = groupEnd + 1
        End If
    Loop
    Set BuildWebUploadItems = items
End Function

Public Function BuildWebItemForRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Object
    Dim groupRow As Long, root As Long, counts As Object
    groupRow = WebGroupRootRow(ws, rowNum)
    root = WebRegistrationRootRow(ws, rowNum)
    If groupRow = 0 Or root = 0 Then Err.Raise vbObjectError + 7802, , "실제 등록 단위에 포함되는 업무 행을 선택하세요."
    Set counts = CountWebTaskIds(ws)
    Set BuildWebItemForRow = NewWebUploadItem(ws, root, groupRow, GetLastDataRow(ws), counts)
End Function

' Level 3 includes level 2 leaves; level 1 never becomes a fallback target.
' This predicate describes structure only; selection is checked by the caller.
Public Function IsWebUploadRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                               ByVal groupLevel As Long, ByVal lastRow As Long) As Boolean
    Dim taskLevel As Long, childRow As Long, subtreeEnd As Long
    If rowNum < DATA_START_ROW Or rowNum > lastRow Then Exit Function
    If Not WebHasTaskContent(ws, rowNum) Then Exit Function
    taskLevel = WebExactTaskLevel(ws, rowNum)
    Select Case groupLevel
        Case 1: IsWebUploadRow = (taskLevel = 1)
        Case 2: IsWebUploadRow = (taskLevel = 2)
        Case 3
            If taskLevel = 3 Then
                IsWebUploadRow = True
            ElseIf taskLevel = 2 Then
                subtreeEnd = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
                For childRow = rowNum + 1 To subtreeEnd
                    If WebHasTaskContent(ws, childRow) Then
                        If GetTaskLevel(ws, childRow) = 3 Then Exit Function
                    End If
                Next childRow
                IsWebUploadRow = True
            End If
    End Select
End Function

' Same whitespace and modification-line rules as modTextReport. The upload
' always includes both fields, independently of the period-report display flags.
Public Function BuildWebTaskBody(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                  ByVal lastRow As Long) As String
    Dim r As Long, subtreeEnd As Long, rootLevel As Long, taskLevel As Long
    Dim requestText As String, modificationText As String, line As String, bullet As String
    Dim indent As Long, requestStart As Long
    rootLevel = GetTaskLevel(ws, rowNum)
    subtreeEnd = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    For r = rowNum To subtreeEnd
        If WebHasTaskContent(ws, r) Then
            taskLevel = GetTaskLevel(ws, r)
            indent = (taskLevel - rootLevel) * 4
            If indent < 0 Then indent = 0
            Select Case taskLevel
                Case 1: bullet = ChrW(&H2022)
                Case 2: bullet = "-"
                Case Else: bullet = ChrW(&HB7)
            End Select
            requestText = CleanWebRequestText(WebCellText(ws, r, COL_REQUEST))
            modificationText = CleanWebModificationText(WebCellText(ws, r, COL_MODIFICATION))
            If Len(requestText) = 0 Then requestText = "(요청 내용 없음)"
            line = Space$(indent) & bullet & " " & requestText
            If Len(BuildWebTaskBody) > 0 Then BuildWebTaskBody = BuildWebTaskBody & vbCrLf
            BuildWebTaskBody = BuildWebTaskBody & line
            If Len(modificationText) > 0 Then
                requestStart = indent + Len(bullet) + 1
                BuildWebTaskBody = BuildWebTaskBody & vbCrLf & _
                    WebModificationLines(modificationText, requestStart)
            End If
        End If
    Next r
End Function

Private Function NewWebUploadItem(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                   ByVal groupRow As Long, ByVal lastRow As Long, _
                                   ByVal identityCounts As Object) As Object
    Dim item As Object, values As Object, sourceId As String, title As String
    Dim body As String, minor As String, requestText As String, r As Long, subtreeEnd As Long
    Set item = CreateObject("Scripting.Dictionary")
    Set values = CreateObject("Scripting.Dictionary")
    item.Add "Row", rowNum
    item.Add "GroupRow", groupRow
    sourceId = GetWebTaskId(ws, rowNum)
    item.Add "SourceId", sourceId
    item.Add "Error", ""
    minor = CleanWebRequestText(WebCellText(ws, rowNum, COL_MINOR_CATEGORY))
    requestText = CleanWebRequestText(WebCellText(ws, rowNum, COL_REQUEST))
    title = "[" & minor & "] " & requestText
    body = BuildWebTaskBody(ws, rowNum, lastRow)
    item.Add "Title", title
    item.Add "Body", body
    values.Add "GUBUN", "PJT"
    values.Add "REQ_TIT", title
    values.Add "REQ_RMK", body
    values.Add "PGM_NM", minor
    values.Add "PGM_ID", Trim$(WebCellText(ws, rowNum, COL_WEB_PROGRAM_ID))
    AddWebOptionValue ws, rowNum, COL_WEB_REQUESTER, "REQ_UID", "요청자", values, item
    AddWebOptionValue ws, rowNum, COL_WEB_ACTION, "ACT_GUBUN", "처리구분", values, item
    AddWebOptionValue ws, rowNum, COL_WEB_STATUS, "STS", "진행상태", values, item
    AddWebOptionValue ws, rowNum, COL_WEB_MODULE, "PGM_MD", "모듈/업무", values, item
    AddWebOptionValue ws, rowNum, COL_WEB_HANDLER, "ACT_UID", "처리자", values, item
    AddWebDateValues ws, rowNum, values, item
    AddWebMinutesValue ws, rowNum, values, item
    item.Add "Values", values
    If Len(sourceId) = 0 Then AppendWebItemError item, "웹업무ID가 없습니다. 웹 등록 설정을 초기화하세요."
    If Len(sourceId) > 0 Then
        If identityCounts.Exists(sourceId) Then
            If CLng(identityCounts(sourceId)) > 1 Then _
                AppendWebItemError item, "웹업무ID가 다른 업무 행과 중복됩니다. 복사한 행을 확인하세요."
        End If
    End If
    If Len(minor) = 0 Then AppendWebItemError item, "소분류를 입력하세요."
    If Len(requestText) = 0 Then AppendWebItemError item, "제목에 사용할 요청 내용을 입력하세요."
    If Len(title) > 300 Then AppendWebItemError item, "등록 제목은 300자 이내로 입력하세요."
    If Len(minor) > 100 Then AppendWebItemError item, "프로그램명(소분류)은 100자 이내로 입력하세요."
    If Len(CStr(values("PGM_ID"))) > 100 Then AppendWebItemError item, "프로그램ID는 100자 이내로 입력하세요."
    If Len(body) = 0 Then AppendWebItemError item, "등록 본문이 비어 있습니다."
    subtreeEnd = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    For r = rowNum To subtreeEnd
        If WebHasTaskContent(ws, r) Then
            If WebExactTaskLevel(ws, r) = 0 Then _
                AppendWebItemError item, CStr(r) & "행의 레벨은 1~3의 정수여야 합니다."
            If Not WebHasExpectedParent(ws, r) Then _
                AppendWebItemError item, CStr(r) & "행의 바로 상위 레벨 업무가 없습니다."
            If IsError(ws.Cells(r, COL_REQUEST).Value2) Or _
               IsError(ws.Cells(r, COL_MODIFICATION).Value2) Then _
                AppendWebItemError item, CStr(r) & "행의 요청 내용 또는 수정 내용에 셀 오류가 있습니다."
        End If
    Next r
    Set NewWebUploadItem = item
End Function

Private Function CountWebTaskIds(ByVal source As Worksheet) As Object
    Dim counts As Object, ws As Worksheet, rowNum As Long, sourceId As String, isTaskSheet As Boolean
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    For Each ws In source.Parent.Worksheets
        isTaskSheet = (ws Is source)
        If Not isTaskSheet Then
            If ws.Name <> CONFIG_SHEET_NAME And ws.Name <> OUTPUT_CONFIG_SHEET_NAME And _
               ws.Name <> WEB_CONFIG_SHEET_NAME And ws.Name <> WEB_PREVIEW_SHEET_NAME And _
               ws.Name <> WEB_LOG_SHEET_NAME Then _
                isTaskSheet = (WebCellText(ws, HEADER_ROW, COL_REQUEST) = "요청 내용")
        End If
        If isTaskSheet Then
            For rowNum = DATA_START_ROW To GetLastDataRow(ws)
                If WebHasTaskContent(ws, rowNum) Then
                    sourceId = GetWebTaskId(ws, rowNum)
                    If Len(sourceId) > 0 Then
                        If counts.Exists(sourceId) Then
                            counts(sourceId) = CLng(counts(sourceId)) + 1
                        Else
                            counts.Add sourceId, 1
                        End If
                    End If
                End If
            Next rowNum
        End If
    Next ws
    Set CountWebTaskIds = counts
End Function

Private Sub AddWebOptionValue(ByVal ws As Worksheet, ByVal rowNum As Long, _
                              ByVal columnName As String, ByVal fieldName As String, _
                              ByVal label As String, ByVal values As Object, ByVal item As Object)
    Dim display As String, code As String, state As Object
    display = Trim$(WebCellText(ws, rowNum, columnName))
    If Len(display) > 0 Then code = ResolveWebOption(fieldName, display)
    If fieldName = "REQ_UID" And Len(code) = 0 Then
        Set state = WebTaskState(GetWebTaskId(ws, rowNum))
        If Len(CStr(state("RequestNo"))) > 0 Then code = CStr(state("RequesterCode"))
    End If
    values.Add fieldName, code
    If Len(code) = 0 Then AppendWebItemError item, label & "를 최신 웹 선택목록에서 선택하세요."
End Sub

Private Sub AddWebDateValues(ByVal ws As Worksheet, ByVal rowNum As Long, _
                             ByVal values As Object, ByVal item As Object)
    Dim startValue As Variant, endValue As Variant, hasStart As Boolean, hasEnd As Boolean
    startValue = ws.Cells(rowNum, COL_ACTUAL_START).Value
    endValue = ws.Cells(rowNum, COL_ACTUAL_END).Value
    If Not IsError(startValue) Then
        If VarType(startValue) <> vbBoolean Then hasStart = IsDate(startValue)
    End If
    If Not IsError(endValue) Then
        If VarType(endValue) <> vbBoolean Then hasEnd = IsDate(endValue)
    End If
    values.Add "ACT_DT_FR", ""
    values.Add "ACT_DT", ""
    If hasStart Then values("ACT_DT_FR") = Format$(CDate(startValue), "yyyy-mm-dd")
    If hasEnd Then values("ACT_DT") = Format$(CDate(endValue), "yyyy-mm-dd")
    If Not hasStart Then AppendWebItemError item, "실제 시작일을 날짜로 입력하세요."
    If Not hasEnd Then AppendWebItemError item, "실제 종료일을 날짜로 입력하세요."
    If hasStart And hasEnd Then
        If DateValue(CDate(startValue)) > DateValue(CDate(endValue)) Then _
            AppendWebItemError item, "실제 종료일이 실제 시작일보다 빠릅니다."
    End If
End Sub

Private Sub AddWebMinutesValue(ByVal ws As Worksheet, ByVal rowNum As Long, _
                               ByVal values As Object, ByVal item As Object)
    Dim raw As Variant, minutes As Double, valid As Boolean
    raw = ws.Cells(rowNum, COL_WORK_MINUTES).Value2
    values.Add "ACT_TIME", ""
    If Not IsError(raw) Then
        If VarType(raw) <> vbBoolean And Not IsEmpty(raw) Then
            If Len(Trim$(CStr(raw))) > 0 And IsNumeric(raw) Then
                minutes = CDbl(raw)
                valid = (minutes >= 0 And minutes <= 2147483647#)
                If valid Then valid = (minutes = Fix(minutes))
            End If
        End If
    End If
    If valid Then
        values("ACT_TIME") = Format$(minutes, "0")
    Else
        AppendWebItemError item, "작업시간(분)은 0 이상의 정수로 입력하세요."
    End If
End Sub

Private Function WebGroupLevel(ByVal ws As Worksheet, ByVal groupRow As Long) As Long
    Dim raw As Variant
    raw = ws.Cells(groupRow, COL_WEB_LEVEL).Value2
    If IsError(raw) Or IsEmpty(raw) Then Exit Function
    If VarType(raw) = vbBoolean Then Exit Function
    If Not IsNumeric(raw) Then Exit Function
    If CDbl(raw) < 1 Or CDbl(raw) > 3 Then Exit Function
    If CDbl(raw) <> Fix(CDbl(raw)) Then Exit Function
    WebGroupLevel = CLng(raw)
End Function

Private Function WebExactTaskLevel(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    Dim raw As Variant
    raw = ws.Cells(rowNum, COL_LEVEL).Value2
    If IsError(raw) Then Exit Function
    If Len(Trim$(CStr(raw))) = 0 Then
        WebExactTaskLevel = 1
        Exit Function
    End If
    If VarType(raw) = vbBoolean Or Not IsNumeric(raw) Then Exit Function
    If CDbl(raw) < 1 Or CDbl(raw) > 3 Then Exit Function
    If CDbl(raw) <> Fix(CDbl(raw)) Then Exit Function
    WebExactTaskLevel = CLng(raw)
End Function

Private Function WebHasExpectedParent(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim taskLevel As Long, r As Long, parentLevel As Long
    taskLevel = WebExactTaskLevel(ws, rowNum)
    If taskLevel <= 1 Then
        WebHasExpectedParent = True
        Exit Function
    End If
    For r = rowNum - 1 To DATA_START_ROW Step -1
        If WebHasTaskContent(ws, r) Then
            parentLevel = WebExactTaskLevel(ws, r)
            If parentLevel < taskLevel Then
                WebHasExpectedParent = (parentLevel = taskLevel - 1)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function WebRowEnabled(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    WebRowEnabled = (UCase$(Trim$(WebCellText(ws, rowNum, COL_WEB_ENABLED))) = "Y")
End Function

Private Function WebHasTaskContent(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    WebHasTaskContent = (Len(Trim$(WebCellText(ws, rowNum, COL_REQUEST))) > 0 Or _
                         Len(Trim$(WebCellText(ws, rowNum, COL_MODIFICATION))) > 0 Or _
                         IsError(ws.Cells(rowNum, COL_REQUEST).Value2) Or _
                         IsError(ws.Cells(rowNum, COL_MODIFICATION).Value2))
End Function

Private Function WebCellText(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal columnName As String) As String
    Dim value As Variant
    value = ws.Cells(rowNum, columnName).Value2
    If Not IsError(value) Then WebCellText = CStr(value)
End Function

Private Sub AppendWebItemError(ByVal item As Object, ByVal message As String)
    If Len(CStr(item("Error"))) > 0 Then item("Error") = CStr(item("Error")) & vbCrLf
    item("Error") = CStr(item("Error")) & message
End Sub

Private Function CleanWebRequestText(ByVal text As String) As String
    text = Replace$(text, vbCr, " ")
    text = Replace$(text, vbLf, " ")
    text = Replace$(text, vbTab, " ")
    text = Replace$(text, ChrW(11), " ")
    Do While InStr(text, "  ") > 0
        text = Replace$(text, "  ", " ")
    Loop
    CleanWebRequestText = Trim$(text)
End Function

Private Function CleanWebModificationText(ByVal text As String) As String
    Dim lines As Variant, i As Long, firstLine As Long, lastLine As Long
    text = Replace$(text, vbCrLf, vbLf)
    text = Replace$(text, vbCr, vbLf)
    text = Replace$(text, ChrW(11), vbLf)
    If Len(text) = 0 Then Exit Function
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        lines(i) = Trim$(Replace$(CStr(lines(i)), vbTab, " "))
    Next i
    firstLine = LBound(lines)
    lastLine = UBound(lines)
    Do While firstLine <= lastLine
        If Len(CStr(lines(firstLine))) > 0 Then Exit Do
        firstLine = firstLine + 1
    Loop
    Do While lastLine >= firstLine
        If Len(CStr(lines(lastLine))) > 0 Then Exit Do
        lastLine = lastLine - 1
    Loop
    For i = firstLine To lastLine
        If i > firstLine Then CleanWebModificationText = CleanWebModificationText & vbLf
        CleanWebModificationText = CleanWebModificationText & CStr(lines(i))
    Next i
End Function

Private Function WebModificationLines(ByVal text As String, ByVal requestStart As Long) As String
    Dim lines As Variant, i As Long
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then WebModificationLines = WebModificationLines & vbCrLf
        If Len(CStr(lines(i))) > 0 Then _
            WebModificationLines = WebModificationLines & Space$(requestStart) & ChrW(&H2514) & " " & CStr(lines(i))
    Next i
End Function
