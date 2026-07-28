Attribute VB_Name = "modDevProgressReport"
Option Explicit

Private Const HISTORY_HEADER_ROW As Long = 1
Private Const HISTORY_DATA_START_ROW As Long = 2
Private Const REPORT_SCOPE_PERSONAL As String = "PERSONAL"
Private Const REPORT_SCOPE_TEAM As String = "TEAM"
Private Const REPORT_SCOPE_MODULE As String = "MODULE"
Private Const REPORT_UNASSIGNED_MODULE As String = "모듈 미지정"
Private Const REPORT_UNASSIGNED_OWNER As String = "담당 미지정"

Public Sub 개발진행보고_텍스트생성()
    팀개발보고_텍스트생성
End Sub

Public Sub 개인개발보고_텍스트생성()
    Dim selectedOwner As String

    selectedOwner = PromptReportFilterValue(ActiveSheet, COL_OWNER, "담당자")
    If Len(selectedOwner) = 0 Then Exit Sub

    Call GenerateScopedDevProgressReport(True, REPORT_SCOPE_PERSONAL, selectedOwner)
End Sub

Public Sub 팀개발보고_텍스트생성()
    Call GenerateScopedDevProgressReport(True, REPORT_SCOPE_TEAM, "")
End Sub

Public Sub 모듈개발보고_텍스트생성()
    Dim selectedModule As String

    selectedModule = PromptReportFilterValue(ActiveSheet, COL_MODULE, "모듈")
    If Len(selectedModule) = 0 Then Exit Sub

    Call GenerateScopedDevProgressReport(True, REPORT_SCOPE_MODULE, selectedModule)
End Sub

Public Function GenerateDevProgressReport(ByVal showCompletionMessage As Boolean) As String
    GenerateDevProgressReport = GenerateScopedDevProgressReport( _
                                    showCompletionMessage, _
                                    REPORT_SCOPE_TEAM, _
                                    "")
End Function

Private Function GenerateScopedDevProgressReport(ByVal showCompletionMessage As Boolean, _
                                                 ByVal reportScope As String, _
                                                 ByVal selectedValue As String) As String
    Dim ws As Worksheet
    Dim historyWs As Worksheet
    Dim previousStatusDict As Object
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim reportItems As Collection
    Dim snapshots As Collection
    Dim reportDate As Date
    Dim previousReportDate As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim statusText As String
    Dim taskText As String
    Dim taskKey As String
    Dim moduleText As String
    Dim ownerText As String
    Dim rawModuleText As String
    Dim rawOwnerText As String
    Dim hierarchyPath As Variant
    Dim includeItem As Boolean
    Dim sortDate As Double
    Dim completedCount As Long
    Dim inProgressCount As Long
    Dim plannedCount As Long
    Dim reportText As String
    Dim reportPath As String
    Dim reportTitle As String
    Dim snapshot As Variant
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = REPORT_HISTORY_SHEET_NAME Then
        Err.Raise vbObjectError + 7401, "GenerateScopedDevProgressReport", "업무 시트에서 실행하세요."
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 7402, "GenerateScopedDevProgressReport", "통합문서를 먼저 저장하세요."
    End If

    EnsureConfigSheet
    Set historyWs = EnsureReportHistorySheet()
    reportDate = GetScheduledReportDate(Date)
    previousReportDate = GetPreviousReportDate(historyWs, ws.Name, reportDate)

    Set previousStatusDict = CreateObject("Scripting.Dictionary")
    previousStatusDict.CompareMode = vbTextCompare
    LoadPreviousStatuses historyWs, ws.Name, previousReportDate, previousStatusDict

    Set reportItems = New Collection
    Set snapshots = New Collection

    lastRow = GetLastDataRow(ws)
    SynchronizeTaskHierarchyModules ws, lastRow, True
    LoadHolidaySettings holidayDict, workdayDict
    UpdateDevelopmentProgressStatuses ws, lastRow, holidayDict, workdayDict

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If Not HasChildTask(ws, r, lastRow) Then
                statusText = Trim$(CStr(ws.Cells(r, COL_DEV_PROGRESS).Value2))

                If StrComp(statusText, REPORT_STATUS_PLANNED, vbTextCompare) = 0 Or _
                   StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Or _
                   StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
                    taskText = CleanReportTaskText(CStr(ws.Cells(r, COL_TASK).Value2))
                    rawModuleText = GetEffectiveReportField(ws, r, COL_MODULE)
                    rawOwnerText = GetEffectiveReportField(ws, r, COL_OWNER)
                    includeItem = ReportItemMatchesScope( _
                                      reportScope, _
                                      selectedValue, _
                                      rawModuleText, _
                                      rawOwnerText)

                    moduleText = rawModuleText
                    ownerText = rawOwnerText
                    If Len(moduleText) = 0 Then moduleText = REPORT_UNASSIGNED_MODULE
                    If Len(ownerText) = 0 Then ownerText = REPORT_UNASSIGNED_OWNER

                    hierarchyPath = BuildReportHierarchyPath(ws, r)
                    taskKey = BuildReportTaskKey(ws, r, taskText)

                    If StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_COMPLETED
                        If includeItem Then
                            If Not previousStatusDict.Exists(taskKey) Then
                                sortDate = GetReportSortDate(ws.Cells(r, COL_ACTUAL_END).Value, ws.Cells(r, COL_PLAN_END).Value)
                                AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, False, r, hierarchyPath)
                                completedCount = completedCount + 1
                            ElseIf StrComp(CStr(previousStatusDict(taskKey)), REPORT_STATUS_COMPLETED, vbTextCompare) <> 0 Then
                                sortDate = GetReportSortDate(ws.Cells(r, COL_ACTUAL_END).Value, ws.Cells(r, COL_PLAN_END).Value)
                                AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, False, r, hierarchyPath)
                                completedCount = completedCount + 1
                            End If
                        End If
                    ElseIf StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_IN_PROGRESS
                        If includeItem Then
                            sortDate = GetReportSortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                            AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, False, r, hierarchyPath)
                            inProgressCount = inProgressCount + 1
                        End If
                    Else
                        statusText = REPORT_STATUS_PLANNED
                        If includeItem Then
                            sortDate = GetReportSortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                            AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, False, r, hierarchyPath)
                            plannedCount = plannedCount + 1
                        End If
                    End If

                    snapshots.Add Array( _
                        reportDate, _
                        Now, _
                        ws.Name, _
                        taskKey, _
                        ws.Cells(r, COL_NO).Value2, _
                        taskText, _
                        statusText, _
                        ws.Cells(r, COL_PLAN_START).Value, _
                        ws.Cells(r, COL_PLAN_END).Value)
                End If
            End If
        End If
    Next r

    reportTitle = BuildReportScopeTitle(reportScope, selectedValue)
    reportText = BuildDevProgressReportText(ws, reportDate, reportTitle, reportItems)
    reportPath = BuildScopedReportPath(reportScope, selectedValue, reportDate)

    WriteUtf8TextFile reportPath, reportText

    DeleteExistingSnapshot historyWs, ws.Name, reportDate
    For Each snapshot In snapshots
        AppendHistorySnapshot historyWs, snapshot
    Next snapshot
    FormatReportHistorySheet historyWs

    GenerateScopedDevProgressReport = reportPath

    If showCompletionMessage Then
        MsgBox reportTitle & " 개발 보고 생성 완료" & vbCrLf & _
               "보고 기준일: " & Format$(reportDate, "yyyy-mm-dd") & vbCrLf & _
               "완료 건: " & completedCount & "개" & vbCrLf & _
               "진행 중: " & inProgressCount & "개" & vbCrLf & _
               "예정: " & plannedCount & "개" & vbCrLf & vbCrLf & _
               reportPath, vbInformation
    End If
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    If showCompletionMessage Then
        MsgBox "개발 진행 보고를 생성할 수 없습니다: " & errDescription, vbExclamation
        GenerateScopedDevProgressReport = ""
    Else
        Err.Raise errNumber, "GenerateScopedDevProgressReport", errDescription
    End If
End Function

Private Function PromptReportFilterValue(ByVal ws As Worksheet, _
                                         ByVal columnAddress As String, _
                                         ByVal valueLabel As String) As String
    Dim values As Collection
    Dim valueSeen As Object
    Dim promptText As String
    Dim answer As Variant
    Dim selectedIndex As Long
    Dim i As Long
    Dim useDirectInput As Boolean

    If ws Is Nothing Then Exit Function
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = REPORT_HISTORY_SHEET_NAME Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Function
    End If

    Set values = CollectReportFilterValues(ws, columnAddress)
    If values.Count = 0 Then
        MsgBox "선택할 " & valueLabel & " 값이 없습니다.", vbExclamation
        Exit Function
    End If

    promptText = valueLabel & "를 선택하세요." & vbCrLf & vbCrLf
    For i = 1 To values.Count
        If Len(promptText) + Len(CStr(values(i))) + 10 > 850 Then
            useDirectInput = True
            Exit For
        End If
        promptText = promptText & CStr(i) & ". " & CStr(values(i)) & vbCrLf
    Next i

    If useDirectInput Then
        Set valueSeen = CreateObject("Scripting.Dictionary")
        valueSeen.CompareMode = vbTextCompare
        For i = 1 To values.Count
            valueSeen(CStr(values(i))) = True
        Next i

        answer = Application.InputBox( _
                    Prompt:=valueLabel & " 목록이 많습니다." & vbCrLf & _
                            "정확한 " & valueLabel & " 이름을 입력하세요.", _
                    Title:=valueLabel & " 선택", _
                    Type:=2)
        If VarType(answer) = vbBoolean Then Exit Function
        If Not valueSeen.Exists(Trim$(CStr(answer))) Then
            MsgBox "목록에 없는 " & valueLabel & "입니다.", vbExclamation
            Exit Function
        End If
        PromptReportFilterValue = Trim$(CStr(answer))
        Exit Function
    End If

    answer = Application.InputBox( _
                Prompt:=promptText, _
                Title:=valueLabel & " 선택", _
                Type:=1)
    If VarType(answer) = vbBoolean Then Exit Function
    If Not IsNumeric(answer) Then Exit Function
    If CDbl(answer) <> Fix(CDbl(answer)) Then
        MsgBox "목록의 번호를 입력하세요.", vbExclamation
        Exit Function
    End If

    selectedIndex = CLng(answer)
    If selectedIndex < 1 Or selectedIndex > values.Count Then
        MsgBox "목록의 번호를 입력하세요.", vbExclamation
        Exit Function
    End If

    PromptReportFilterValue = CStr(values(selectedIndex))
End Function

Private Function CollectReportFilterValues(ByVal ws As Worksheet, _
                                           ByVal columnAddress As String) As Collection
    Dim result As Collection
    Dim valueSeen As Object
    Dim lastRow As Long
    Dim r As Long
    Dim fieldValue As String

    Set result = New Collection
    Set valueSeen = CreateObject("Scripting.Dictionary")
    valueSeen.CompareMode = vbTextCompare
    lastRow = GetLastDataRow(ws)

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) And Not HasChildTask(ws, r, lastRow) Then
            fieldValue = GetEffectiveReportField(ws, r, columnAddress)
            If Len(fieldValue) > 0 And Not valueSeen.Exists(fieldValue) Then
                valueSeen.Add fieldValue, True
                result.Add fieldValue
            End If
        End If
    Next r

    Set CollectReportFilterValues = result
End Function

Private Function GetEffectiveReportField(ByVal ws As Worksheet, _
                                         ByVal rowNum As Long, _
                                         ByVal columnAddress As String) As String
    Dim fieldValue As String
    Dim currentLevel As Long
    Dim candidateLevel As Long
    Dim r As Long

    fieldValue = CleanReportTaskText(CStr(ws.Cells(rowNum, columnAddress).Value2))
    If Len(fieldValue) > 0 Then
        GetEffectiveReportField = fieldValue
        Exit Function
    End If

    currentLevel = GetTaskLevel(ws, rowNum)
    For r = rowNum - 1 To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            candidateLevel = GetTaskLevel(ws, r)
            If candidateLevel < currentLevel Then
                fieldValue = CleanReportTaskText(CStr(ws.Cells(r, columnAddress).Value2))
                If Len(fieldValue) > 0 Then
                    GetEffectiveReportField = fieldValue
                    Exit Function
                End If
                currentLevel = candidateLevel
                If currentLevel = 1 Then Exit For
            End If
        End If
    Next r
End Function

Private Function ReportItemMatchesScope(ByVal reportScope As String, _
                                        ByVal selectedValue As String, _
                                        ByVal moduleText As String, _
                                        ByVal ownerText As String) As Boolean
    Select Case UCase$(reportScope)
        Case REPORT_SCOPE_PERSONAL
            ReportItemMatchesScope = _
                (StrComp(ownerText, selectedValue, vbTextCompare) = 0)
        Case REPORT_SCOPE_MODULE
            ReportItemMatchesScope = _
                (StrComp(moduleText, selectedValue, vbTextCompare) = 0)
        Case Else
            ReportItemMatchesScope = True
    End Select
End Function

Private Function BuildReportScopeTitle(ByVal reportScope As String, _
                                       ByVal selectedValue As String) As String
    Select Case UCase$(reportScope)
        Case REPORT_SCOPE_PERSONAL
            BuildReportScopeTitle = selectedValue & " 개인"
        Case REPORT_SCOPE_MODULE
            BuildReportScopeTitle = selectedValue & " 모듈"
        Case Else
            BuildReportScopeTitle = "팀"
    End Select
End Function

Private Function BuildScopedReportPath(ByVal reportScope As String, _
                                       ByVal selectedValue As String, _
                                       ByVal reportDate As Date) As String
    Dim filePrefix As String

    Select Case UCase$(reportScope)
        Case REPORT_SCOPE_PERSONAL
            filePrefix = "개인개발보고_" & SanitizeReportFilePart(selectedValue)
        Case REPORT_SCOPE_MODULE
            filePrefix = "모듈개발보고_" & SanitizeReportFilePart(selectedValue)
        Case Else
            filePrefix = "팀개발보고"
    End Select

    BuildScopedReportPath = ThisWorkbook.Path & Application.PathSeparator & _
                            filePrefix & "_" & Format$(reportDate, "yyyy-mm-dd") & ".txt"
End Function

Private Function SanitizeReportFilePart(ByVal filePart As String) As String
    Dim invalidCharacter As Variant

    For Each invalidCharacter In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        filePart = Replace$(filePart, CStr(invalidCharacter), "_")
    Next invalidCharacter

    filePart = Trim$(filePart)
    If Len(filePart) > 60 Then filePart = Left$(filePart, 60)
    If Len(filePart) = 0 Then filePart = "미지정"
    SanitizeReportFilePart = filePart
End Function

Private Function EnsureReportHistorySheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REPORT_HISTORY_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = REPORT_HISTORY_SHEET_NAME
    End If

    With ws
        .Cells(HISTORY_HEADER_ROW, 1).Value = "Report Date"
        .Cells(HISTORY_HEADER_ROW, 2).Value = "Generated At"
        .Cells(HISTORY_HEADER_ROW, 3).Value = "Source Sheet"
        .Cells(HISTORY_HEADER_ROW, 4).Value = "Task Key"
        .Cells(HISTORY_HEADER_ROW, 5).Value = "No."
        .Cells(HISTORY_HEADER_ROW, 6).Value = "Task Content"
        .Cells(HISTORY_HEADER_ROW, 7).Value = "Dev Status"
        .Cells(HISTORY_HEADER_ROW, 8).Value = "Plan Start"
        .Cells(HISTORY_HEADER_ROW, 9).Value = "Plan End"
    End With

    Set EnsureReportHistorySheet = ws
End Function

Private Function GetScheduledReportDate(ByVal targetDate As Date) As Date
    Select Case Weekday(targetDate, vbMonday)
        Case 1
            GetScheduledReportDate = targetDate - 4
        Case 2
            GetScheduledReportDate = targetDate
        Case 3
            GetScheduledReportDate = targetDate - 1
        Case 4
            GetScheduledReportDate = targetDate
        Case 5
            GetScheduledReportDate = targetDate - 1
        Case 6
            GetScheduledReportDate = targetDate - 2
        Case 7
            GetScheduledReportDate = targetDate - 3
    End Select
End Function

Private Function GetPreviousReportDate(ByVal historyWs As Worksheet, _
                                       ByVal sourceSheetName As String, _
                                       ByVal reportDate As Date) As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim candidateDate As Variant
    Dim latestDate As Date
    Dim hasDate As Boolean

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row

    For r = HISTORY_DATA_START_ROW To lastRow
        If StrComp(Trim$(CStr(historyWs.Cells(r, 3).Value2)), sourceSheetName, vbTextCompare) = 0 Then
            candidateDate = historyWs.Cells(r, 1).Value
            If IsDate(candidateDate) Then
                If CLng(CDate(candidateDate)) < CLng(reportDate) Then
                    If Not hasDate Or CLng(CDate(candidateDate)) > CLng(latestDate) Then
                        latestDate = CDate(candidateDate)
                        hasDate = True
                    End If
                End If
            End If
        End If
    Next r

    If hasDate Then
        GetPreviousReportDate = latestDate
    Else
        GetPreviousReportDate = Empty
    End If
End Function

Private Sub LoadPreviousStatuses(ByVal historyWs As Worksheet, _
                                 ByVal sourceSheetName As String, _
                                 ByVal previousReportDate As Variant, _
                                 ByVal statusDict As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim taskKey As String

    If Not IsDate(previousReportDate) Then Exit Sub

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row

    For r = HISTORY_DATA_START_ROW To lastRow
        If IsDate(historyWs.Cells(r, 1).Value) Then
            If CLng(CDate(historyWs.Cells(r, 1).Value)) = CLng(CDate(previousReportDate)) And _
               StrComp(Trim$(CStr(historyWs.Cells(r, 3).Value2)), sourceSheetName, vbTextCompare) = 0 Then
                taskKey = Trim$(CStr(historyWs.Cells(r, 4).Value2))
                If Len(taskKey) > 0 Then
                    statusDict(taskKey) = Trim$(CStr(historyWs.Cells(r, 7).Value2))
                End If
            End If
        End If
    Next r
End Sub

Private Function BuildReportTaskKey(ByVal ws As Worksheet, _
                                    ByVal rowNum As Long, _
                                    ByVal taskText As String) As String
    Dim planStartKey As String

    If IsDate(ws.Cells(rowNum, COL_PLAN_START).Value) Then
        planStartKey = Format$(CDate(ws.Cells(rowNum, COL_PLAN_START).Value), "yyyymmdd")
    Else
        planStartKey = ""
    End If

    BuildReportTaskKey = UCase$(Trim$(taskText)) & "|" & planStartKey
End Function

Private Function GetReportSortDate(ByVal primaryDate As Variant, _
                                   ByVal fallbackDate As Variant) As Double
    If IsDate(primaryDate) Then
        GetReportSortDate = CDbl(CDate(primaryDate))
    ElseIf IsDate(fallbackDate) Then
        GetReportSortDate = CDbl(CDate(fallbackDate))
    Else
        GetReportSortDate = CDbl(DateSerial(9999, 12, 31))
    End If
End Function

Private Function GetReportStatusRank(ByVal statusText As String) As Long
    If StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
        GetReportStatusRank = 1
    ElseIf StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Then
        GetReportStatusRank = 2
    Else
        GetReportStatusRank = 3
    End If
End Function

Private Sub AddSortedDevReportItem(ByVal items As Collection, ByVal newItem As Variant)
    Dim i As Long
    Dim existingItem As Variant
    Dim newRank As Long
    Dim existingRank As Long

    newRank = GetReportStatusRank(CStr(newItem(3)))
    For i = 1 To items.Count
        existingItem = items(i)
        existingRank = GetReportStatusRank(CStr(existingItem(3)))

        If newRank < existingRank Or _
           (newRank = existingRank And CDbl(newItem(4)) < CDbl(existingItem(4))) Then
            items.Add newItem, Before:=i
            Exit Sub
        End If
    Next i

    items.Add newItem
End Sub

Private Function CleanReportTaskText(ByVal taskText As String) As String
    taskText = Replace$(taskText, vbCr, " ")
    taskText = Replace$(taskText, vbLf, " ")
    taskText = Replace$(taskText, vbTab, " ")

    Do While InStr(taskText, "  ") > 0
        taskText = Replace$(taskText, "  ", " ")
    Loop

    CleanReportTaskText = Trim$(taskText)
End Function

Private Function BuildReportHierarchyPath(ByVal ws As Worksheet, _
                                          ByVal rowNum As Long) As Variant
    Dim reversePath As Collection
    Dim pathValues() As String
    Dim currentLevel As Long
    Dim candidateLevel As Long
    Dim r As Long
    Dim i As Long

    Set reversePath = New Collection
    reversePath.Add BuildHierarchyPathToken(ws, rowNum)
    currentLevel = GetTaskLevel(ws, rowNum)

    For r = rowNum - 1 To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            candidateLevel = GetTaskLevel(ws, r)
            If candidateLevel < currentLevel Then
                reversePath.Add BuildHierarchyPathToken(ws, r)
                currentLevel = candidateLevel
                If currentLevel = 1 Then Exit For
            End If
        End If
    Next r

    ReDim pathValues(0 To reversePath.Count - 1)
    For i = 1 To reversePath.Count
        pathValues(i - 1) = CStr(reversePath(reversePath.Count - i + 1))
    Next i

    BuildReportHierarchyPath = pathValues
End Function

Private Function BuildHierarchyPathToken(ByVal ws As Worksheet, _
                                         ByVal rowNum As Long) As String
    BuildHierarchyPathToken = CStr(rowNum) & vbTab & _
                              CleanReportTaskText(CStr(ws.Cells(rowNum, COL_TASK).Value2))
End Function

Private Function BuildDevProgressReportText(ByVal ws As Worksheet, _
                                            ByVal reportDate As Date, _
                                            ByVal reportTitle As String, _
                                            ByVal reportItems As Collection) As String
    Dim textValue As String

    textValue = Format$(reportDate, "yyyy-mm-dd") & " " & _
                reportTitle & " 개발 진행 현황입니다." & vbCrLf & vbCrLf

    If reportItems.Count = 0 Then
        textValue = textValue & "보고 내역이 없습니다." & vbCrLf
    ElseIf GetDevReportSeparateStatusFlag() Then
        AppendStatusReportSection textValue, ws, reportItems, REPORT_STATUS_COMPLETED
        AppendStatusReportSection textValue, ws, reportItems, REPORT_STATUS_IN_PROGRESS
        AppendStatusReportSection textValue, ws, reportItems, REPORT_STATUS_PLANNED
    Else
        AppendReportModules textValue, ws, reportItems, "", True
    End If

    BuildDevProgressReportText = textValue
End Function

Private Sub AppendStatusReportSection(ByRef textValue As String, _
                                      ByVal ws As Worksheet, _
                                      ByVal reportItems As Collection, _
                                      ByVal statusFilter As String)
    textValue = textValue & GetReportStatusLabel(statusFilter) & vbCrLf & vbCrLf

    If CountReportItemsByStatus(reportItems, statusFilter) = 0 Then
        textValue = textValue & Space$(5) & "내역 없음" & vbCrLf & vbCrLf
    Else
        AppendReportModules textValue, ws, reportItems, statusFilter, False
    End If
End Sub

Private Sub AppendReportModules(ByRef textValue As String, _
                                ByVal ws As Worksheet, _
                                ByVal reportItems As Collection, _
                                ByVal statusFilter As String, _
                                ByVal includeStatusLabel As Boolean)
    Dim item As Variant
    Dim moduleNames As Collection
    Dim ownerNames As Collection
    Dim moduleSeen As Object
    Dim ownerSeen As Object
    Dim moduleName As Variant
    Dim moduleIndex As Long

    Set moduleNames = New Collection
    Set moduleSeen = CreateObject("Scripting.Dictionary")
    moduleSeen.CompareMode = vbTextCompare

    For Each item In reportItems
        If ReportItemMatchesStatus(item, statusFilter) Then
            If Not moduleSeen.Exists(CStr(item(0))) Then
                moduleSeen.Add CStr(item(0)), True
                moduleNames.Add CStr(item(0))
            End If
        End If
    Next item

    moduleIndex = 0
    For Each moduleName In moduleNames
        moduleIndex = moduleIndex + 1
        Set ownerNames = New Collection
        Set ownerSeen = CreateObject("Scripting.Dictionary")
        ownerSeen.CompareMode = vbTextCompare

        For Each item In reportItems
            If ReportItemMatchesStatus(item, statusFilter) And _
               StrComp(CStr(item(0)), CStr(moduleName), vbTextCompare) = 0 Then
                If Not ownerSeen.Exists(CStr(item(1))) Then
                    ownerSeen.Add CStr(item(1)), True
                    ownerNames.Add CStr(item(1))
                End If
            End If
        Next item

        textValue = textValue & GetCircledReportNumber(moduleIndex) & "     " & _
                    CStr(moduleName) & " (" & JoinReportText(ownerNames, ", ") & ")" & vbCrLf

        AppendGroupedHierarchyItems textValue, ws, reportItems, CStr(moduleName), _
                                    statusFilter, includeStatusLabel
        textValue = textValue & vbCrLf
    Next moduleName
End Sub

Private Function CountReportItemsByStatus(ByVal reportItems As Collection, _
                                          ByVal statusFilter As String) As Long
    Dim item As Variant

    For Each item In reportItems
        If ReportItemMatchesStatus(item, statusFilter) Then
            CountReportItemsByStatus = CountReportItemsByStatus + 1
        End If
    Next item
End Function

Private Function ReportItemMatchesStatus(ByVal item As Variant, _
                                         ByVal statusFilter As String) As Boolean
    ReportItemMatchesStatus = _
        (Len(statusFilter) = 0 Or _
         StrComp(CStr(item(3)), statusFilter, vbTextCompare) = 0)
End Function

Private Sub AppendGroupedHierarchyItems(ByRef textValue As String, _
                                        ByVal ws As Worksheet, _
                                        ByVal reportItems As Collection, _
                                        ByVal moduleName As String, _
                                        ByVal statusFilter As String, _
                                        ByVal includeStatusLabel As Boolean)
    Dim lastRow As Long
    Dim r As Long
    Dim item As Variant
    Dim currentPath As Variant
    Dim previousPath As Variant
    Dim leafSuffix As String

    lastRow = GetLastDataRow(ws)

    For r = DATA_START_ROW To lastRow
        For Each item In reportItems
            If CLng(item(6)) = r And _
               StrComp(CStr(item(0)), moduleName, vbTextCompare) = 0 And _
               ReportItemMatchesStatus(item, statusFilter) Then
                currentPath = item(7)
                leafSuffix = " (" & CStr(item(1)) & ")"
                If includeStatusLabel Then
                    leafSuffix = leafSuffix & ": " & _
                                 GetReportStatusLabel(CStr(item(3)))
                End If
                leafSuffix = leafSuffix & _
                             BuildPlannedDateSuffix(CStr(item(3)), CDbl(item(4)))
                AppendHierarchyPath textValue, currentPath, previousPath, leafSuffix, 5
                previousPath = currentPath
                Exit For
            End If
        Next item
    Next r
End Sub

Private Sub AppendHierarchyPath(ByRef textValue As String, _
                                ByVal currentPath As Variant, _
                                ByVal previousPath As Variant, _
                                ByVal leafSuffix As String, _
                                ByVal baseIndent As Long)
    Dim commonDepth As Long
    Dim depth As Long
    Dim marker As String
    Dim suffix As String

    commonDepth = GetCommonHierarchyDepth(previousPath, currentPath)

    For depth = commonDepth To UBound(currentPath)
        marker = GetDevReportLevelBullet(depth + 1) & " "

        suffix = ""
        If depth = UBound(currentPath) Then suffix = leafSuffix

        textValue = textValue & Space$(baseIndent + (depth * 4)) & _
                    marker & GetHierarchyPathText(CStr(currentPath(depth))) & suffix & vbCrLf
    Next depth
End Sub

Private Function GetHierarchyPathText(ByVal pathToken As String) As String
    Dim separatorPosition As Long

    separatorPosition = InStr(1, pathToken, vbTab, vbBinaryCompare)
    If separatorPosition > 0 Then
        GetHierarchyPathText = Mid$(pathToken, separatorPosition + 1)
    Else
        GetHierarchyPathText = pathToken
    End If
End Function

Private Function GetCommonHierarchyDepth(ByVal previousPath As Variant, _
                                         ByVal currentPath As Variant) As Long
    Dim maxDepth As Long
    Dim depth As Long

    If Not IsArray(previousPath) Then Exit Function
    If Not IsArray(currentPath) Then Exit Function

    maxDepth = UBound(previousPath)
    If UBound(currentPath) < maxDepth Then maxDepth = UBound(currentPath)

    For depth = 0 To maxDepth
        If StrComp(CStr(previousPath(depth)), CStr(currentPath(depth)), vbTextCompare) <> 0 Then Exit For
        GetCommonHierarchyDepth = depth + 1
    Next depth
End Function

Private Function GetCircledReportNumber(ByVal indexValue As Long) As String
    If indexValue >= 1 And indexValue <= 20 Then
        GetCircledReportNumber = ChrW(&H2460 + indexValue - 1)
    Else
        GetCircledReportNumber = CStr(indexValue) & "."
    End If
End Function

Private Function JoinReportText(ByVal items As Collection, _
                                ByVal delimiter As String) As String
    Dim item As Variant
    Dim result As String

    For Each item In items
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next item

    JoinReportText = result
End Function

Private Function GetReportStatusLabel(ByVal statusText As String) As String
    If StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
        GetReportStatusLabel = "완료"
    ElseIf StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Then
        GetReportStatusLabel = "진행중"
    Else
        GetReportStatusLabel = "예정"
    End If
End Function

Private Function BuildPlannedDateSuffix(ByVal statusText As String, _
                                        ByVal sortDate As Double) As String
    If StrComp(statusText, REPORT_STATUS_PLANNED, vbTextCompare) <> 0 Then Exit Function
    If sortDate >= CDbl(DateSerial(9999, 12, 31)) Then Exit Function
    BuildPlannedDateSuffix = " (" & Format$(CDate(sortDate), "mm/dd") & ")"
End Function

Private Sub WriteUtf8TextFile(ByVal filePath As String, ByVal textValue As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textValue
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Sub DeleteExistingSnapshot(ByVal historyWs As Worksheet, _
                                   ByVal sourceSheetName As String, _
                                   ByVal reportDate As Date)
    Dim lastRow As Long
    Dim r As Long

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row

    For r = lastRow To HISTORY_DATA_START_ROW Step -1
        If IsDate(historyWs.Cells(r, 1).Value) Then
            If CLng(CDate(historyWs.Cells(r, 1).Value)) = CLng(reportDate) And _
               StrComp(Trim$(CStr(historyWs.Cells(r, 3).Value2)), sourceSheetName, vbTextCompare) = 0 Then
                historyWs.Rows(r).Delete
            End If
        End If
    Next r
End Sub

Private Sub AppendHistorySnapshot(ByVal historyWs As Worksheet, ByVal snapshot As Variant)
    Dim targetRow As Long
    Dim i As Long

    targetRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < HISTORY_DATA_START_ROW Then targetRow = HISTORY_DATA_START_ROW

    For i = LBound(snapshot) To UBound(snapshot)
        historyWs.Cells(targetRow, i + 1).Value = snapshot(i)
    Next i
End Sub

Private Sub FormatReportHistorySheet(ByVal historyWs As Worksheet)
    Dim lastRow As Long

    lastRow = historyWs.Cells(historyWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < HISTORY_HEADER_ROW Then lastRow = HISTORY_HEADER_ROW

    With historyWs
        .Visible = xlSheetVisible
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(217, 225, 242)
        .Range("A1:I" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A1:I" & lastRow).Borders.Color = RGB(210, 210, 210)
        .Columns("A").NumberFormat = "yyyy-mm-dd"
        .Columns("B").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Columns("H:I").NumberFormat = "yyyy-mm-dd"
        .Columns("A:B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 36
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 70
        .Columns("G").ColumnWidth = 14
        .Columns("H:I").ColumnWidth = 14

        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:I" & lastRow).AutoFilter
    End With
End Sub
