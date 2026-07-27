Attribute VB_Name = "modDevProgressReport"
Option Explicit

Private Const HISTORY_HEADER_ROW As Long = 1
Private Const HISTORY_DATA_START_ROW As Long = 2

Public Sub 개발진행보고_텍스트생성()
    Call GenerateDevProgressReport(True)
End Sub

Public Function GenerateDevProgressReport(ByVal showCompletionMessage As Boolean) As String
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
    Dim useLegacyLayout As Boolean
    Dim sortDate As Double
    Dim completedCount As Long
    Dim inProgressCount As Long
    Dim plannedCount As Long
    Dim reportText As String
    Dim reportPath As String
    Dim snapshot As Variant
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = REPORT_HISTORY_SHEET_NAME Then
        Err.Raise vbObjectError + 7401, "GenerateDevProgressReport", "업무 시트에서 실행하세요."
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 7402, "GenerateDevProgressReport", "통합문서를 먼저 저장하세요."
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
                    moduleText = CleanReportTaskText(CStr(ws.Cells(r, COL_MODULE).Value2))
                    ownerText = CleanReportTaskText(CStr(ws.Cells(r, COL_OWNER).Value2))
                    useLegacyLayout = (Len(moduleText) = 0 Or Len(ownerText) = 0)
                    taskKey = BuildReportTaskKey(ws, r, taskText)

                    If StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_COMPLETED
                        If Not previousStatusDict.Exists(taskKey) Then
                            sortDate = GetReportSortDate(ws.Cells(r, COL_ACTUAL_END).Value, ws.Cells(r, COL_PLAN_END).Value)
                            AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, useLegacyLayout)
                            completedCount = completedCount + 1
                        ElseIf StrComp(CStr(previousStatusDict(taskKey)), REPORT_STATUS_COMPLETED, vbTextCompare) <> 0 Then
                            sortDate = GetReportSortDate(ws.Cells(r, COL_ACTUAL_END).Value, ws.Cells(r, COL_PLAN_END).Value)
                            AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, useLegacyLayout)
                            completedCount = completedCount + 1
                        End If
                    ElseIf StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_IN_PROGRESS
                        sortDate = GetReportSortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                        AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, useLegacyLayout)
                        inProgressCount = inProgressCount + 1
                    Else
                        statusText = REPORT_STATUS_PLANNED
                        sortDate = GetReportSortDate(ws.Cells(r, COL_PLAN_END).Value, Empty)
                        AddSortedDevReportItem reportItems, Array(moduleText, ownerText, taskText, statusText, sortDate, useLegacyLayout)
                        plannedCount = plannedCount + 1
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

    reportText = BuildDevProgressReportText(reportDate, previousReportDate, reportItems)
    reportPath = ThisWorkbook.Path & Application.PathSeparator & _
                 "개발진행보고_" & Format$(reportDate, "yyyy-mm-dd") & ".txt"

    WriteUtf8TextFile reportPath, reportText

    DeleteExistingSnapshot historyWs, ws.Name, reportDate
    For Each snapshot In snapshots
        AppendHistorySnapshot historyWs, snapshot
    Next snapshot
    FormatReportHistorySheet historyWs

    GenerateDevProgressReport = reportPath

    If showCompletionMessage Then
        MsgBox "개발 진행 보고 생성 완료" & vbCrLf & _
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
        GenerateDevProgressReport = ""
    Else
        Err.Raise errNumber, "GenerateDevProgressReport", errDescription
    End If
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

Private Function BuildDevProgressReportText(ByVal reportDate As Date, _
                                            ByVal previousReportDate As Variant, _
                                            ByVal reportItems As Collection) As String
    Dim textValue As String
    Dim item As Variant
    Dim moduleNames As Collection
    Dim ownerNames As Collection
    Dim moduleSeen As Object
    Dim ownerSeen As Object
    Dim moduleName As Variant
    Dim moduleIndex As Long
    Dim hasGroupedItems As Boolean
    Dim hasLegacyItems As Boolean

    For Each item In reportItems
        If CBool(item(5)) Then
            hasLegacyItems = True
        Else
            hasGroupedItems = True
        End If
    Next item

    If hasGroupedItems Then
        textValue = Format$(reportDate, "yyyy-mm-dd") & " 개발 진행 현황입니다." & vbCrLf & vbCrLf
    Else
        textValue = BuildLegacyReportHeader(reportDate, previousReportDate)
    End If

    If reportItems.Count = 0 Then
        textValue = textValue & "보고 내역이 없습니다." & vbCrLf
    Else
        Set moduleNames = New Collection
        Set moduleSeen = CreateObject("Scripting.Dictionary")
        moduleSeen.CompareMode = vbTextCompare

        For Each item In reportItems
            If Not CBool(item(5)) And Not moduleSeen.Exists(CStr(item(0))) Then
                moduleSeen.Add CStr(item(0)), True
                moduleNames.Add CStr(item(0))
            End If
        Next item

        moduleIndex = 0
        For Each moduleName In moduleNames
            moduleIndex = moduleIndex + 1
            Set ownerNames = New Collection
            Set ownerSeen = CreateObject("Scripting.Dictionary")
            ownerSeen.CompareMode = vbTextCompare

            For Each item In reportItems
                If Not CBool(item(5)) And _
                   StrComp(CStr(item(0)), CStr(moduleName), vbTextCompare) = 0 Then
                    If Not ownerSeen.Exists(CStr(item(1))) Then
                        ownerSeen.Add CStr(item(1)), True
                        ownerNames.Add CStr(item(1))
                    End If
                End If
            Next item

            textValue = textValue & GetCircledReportNumber(moduleIndex) & "     " & _
                        CStr(moduleName) & " (" & JoinReportText(ownerNames, ", ") & ")" & vbCrLf

            For Each item In reportItems
                If Not CBool(item(5)) And _
                   StrComp(CStr(item(0)), CStr(moduleName), vbTextCompare) = 0 Then
                    textValue = textValue & "     -     " & CStr(item(2)) & _
                                " (" & CStr(item(1)) & "): " & _
                                GetReportStatusLabel(CStr(item(3))) & _
                                BuildPlannedDateSuffix(CStr(item(3)), CDbl(item(4))) & vbCrLf
                End If
            Next item
            textValue = textValue & vbCrLf
        Next moduleName

        If hasLegacyItems Then AppendLegacyReportItems textValue, reportItems
    End If

    BuildDevProgressReportText = textValue
End Function

Private Function BuildLegacyReportHeader(ByVal reportDate As Date, _
                                         ByVal previousReportDate As Variant) As String
    Dim textValue As String

    textValue = "개발 진행 보고" & vbCrLf
    textValue = textValue & "보고 기준일: " & Format$(reportDate, "yyyy-mm-dd") & _
                " (" & GetKoreanWeekdayName(reportDate) & ")" & vbCrLf

    If IsDate(previousReportDate) Then
        textValue = textValue & "비교 기준: " & _
                    Format$(CDate(previousReportDate), "yyyy-mm-dd") & " 이후" & vbCrLf
    Else
        textValue = textValue & "비교 기준: 최초 보고" & vbCrLf
    End If

    BuildLegacyReportHeader = textValue & vbCrLf
End Function

Private Sub AppendLegacyReportItems(ByRef textValue As String, _
                                    ByVal reportItems As Collection)
    AppendLegacyStatusSection textValue, reportItems, REPORT_STATUS_COMPLETED, "완료 건"
    AppendLegacyStatusSection textValue, reportItems, REPORT_STATUS_IN_PROGRESS, "진행 중"
    AppendLegacyStatusSection textValue, reportItems, REPORT_STATUS_PLANNED, "예정"
End Sub

Private Sub AppendLegacyStatusSection(ByRef textValue As String, _
                                      ByVal reportItems As Collection, _
                                      ByVal targetStatus As String, _
                                      ByVal sectionTitle As String)
    Dim item As Variant
    Dim itemCount As Long

    textValue = textValue & sectionTitle & vbCrLf & vbCrLf
    For Each item In reportItems
        If CBool(item(5)) And _
           StrComp(CStr(item(3)), targetStatus, vbTextCompare) = 0 Then
            textValue = textValue & ChrW(&H2022) & " " & CStr(item(2)) & vbCrLf
            itemCount = itemCount + 1
        End If
    Next item

    If itemCount = 0 Then textValue = textValue & ChrW(&H2022) & " 없음" & vbCrLf
    textValue = textValue & vbCrLf
End Sub

Private Function GetKoreanWeekdayName(ByVal targetDate As Date) As String
    Select Case Weekday(targetDate, vbSunday)
        Case vbSunday: GetKoreanWeekdayName = "일요일"
        Case vbMonday: GetKoreanWeekdayName = "월요일"
        Case vbTuesday: GetKoreanWeekdayName = "화요일"
        Case vbWednesday: GetKoreanWeekdayName = "수요일"
        Case vbThursday: GetKoreanWeekdayName = "목요일"
        Case vbFriday: GetKoreanWeekdayName = "금요일"
        Case vbSaturday: GetKoreanWeekdayName = "토요일"
    End Select
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
