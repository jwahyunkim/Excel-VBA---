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
    Dim completedItems As Collection
    Dim inProgressItems As Collection
    Dim plannedItems As Collection
    Dim snapshots As Collection
    Dim reportDate As Date
    Dim previousReportDate As Variant
    Dim lastRow As Long
    Dim r As Long
    Dim statusText As String
    Dim taskText As String
    Dim taskKey As String
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

    Set completedItems = New Collection
    Set inProgressItems = New Collection
    Set plannedItems = New Collection
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
                    taskKey = BuildReportTaskKey(ws, r, taskText)

                    If StrComp(statusText, REPORT_STATUS_COMPLETED, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_COMPLETED
                        If Not previousStatusDict.Exists(taskKey) Then
                            completedItems.Add taskText
                        ElseIf StrComp(CStr(previousStatusDict(taskKey)), REPORT_STATUS_COMPLETED, vbTextCompare) <> 0 Then
                            completedItems.Add taskText
                        End If
                    ElseIf StrComp(statusText, REPORT_STATUS_IN_PROGRESS, vbTextCompare) = 0 Then
                        statusText = REPORT_STATUS_IN_PROGRESS
                        inProgressItems.Add taskText
                    Else
                        statusText = REPORT_STATUS_PLANNED
                        plannedItems.Add taskText
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

    reportText = BuildDevProgressReportText(reportDate, previousReportDate, completedItems, inProgressItems, plannedItems)
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
               "완료 건: " & completedItems.Count & "개" & vbCrLf & _
               "진행 중: " & inProgressItems.Count & "개" & vbCrLf & _
               "예정: " & plannedItems.Count & "개" & vbCrLf & vbCrLf & _
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
                                            ByVal completedItems As Collection, _
                                            ByVal inProgressItems As Collection, _
                                            ByVal plannedItems As Collection) As String
    Dim textValue As String
    Dim item As Variant

    textValue = "개발 진행 보고" & vbCrLf
    textValue = textValue & "보고 기준일: " & Format$(reportDate, "yyyy-mm-dd") & " (" & GetKoreanWeekdayName(reportDate) & ")" & vbCrLf

    If IsDate(previousReportDate) Then
        textValue = textValue & "비교 기준: " & Format$(CDate(previousReportDate), "yyyy-mm-dd") & " 이후" & vbCrLf
    Else
        textValue = textValue & "비교 기준: 최초 보고" & vbCrLf
    End If

    textValue = textValue & vbCrLf & "완료 건" & vbCrLf & vbCrLf
    AppendReportCollection textValue, completedItems

    textValue = textValue & vbCrLf & "진행 중" & vbCrLf & vbCrLf
    AppendReportCollection textValue, inProgressItems

    textValue = textValue & vbCrLf & "예정" & vbCrLf & vbCrLf
    AppendReportCollection textValue, plannedItems

    BuildDevProgressReportText = textValue
End Function

Private Sub AppendReportCollection(ByRef textValue As String, ByVal items As Collection)
    Dim item As Variant

    If items.Count = 0 Then
        textValue = textValue & ChrW(&H2022) & " 없음" & vbCrLf
    Else
        For Each item In items
            textValue = textValue & ChrW(&H2022) & " " & CStr(item) & vbCrLf
        Next item
    End If
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