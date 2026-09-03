Attribute VB_Name = "modGanttSheet"
Option Explicit

Private Const TASK_VALIDATION_MIN_LAST_ROW As Long = 5000
Private Const TASK_VALIDATION_EXTRA_ROWS As Long = 1000

Public Sub SetupDataHeaders(ws As Worksheet)
    NormalizeSheetStructure ws

    ws.Cells(HEADER_ROW, COL_NO).Value = "No."
    ws.Cells(HEADER_ROW, COL_LEVEL).Value = "Level"
    ws.Cells(HEADER_ROW, COL_TYPE).Value = "타입"
    ws.Cells(HEADER_ROW, COL_MAJOR_CATEGORY).Value = "대분류"
    ws.Cells(HEADER_ROW, COL_MIDDLE_CATEGORY).Value = "중분류"
    ws.Cells(HEADER_ROW, COL_MINOR_CATEGORY).Value = "소분류"
    ws.Cells(HEADER_ROW, COL_TASK).Value = "내용"
    ws.Cells(HEADER_ROW, COL_OWNER).Value = "담당"
    ws.Cells(HEADER_ROW, COL_NOTE).Value = "비고"
    ws.Cells(HEADER_ROW, COL_PLAN_START).Value = "계획 시작일"
    ws.Cells(HEADER_ROW, COL_PLAN_END).Value = "계획 종료일"
    ws.Cells(HEADER_ROW, COL_ACTUAL_START).Value = "실제 시작일"
    ws.Cells(HEADER_ROW, COL_ACTUAL_END).Value = "실제 종료일"
    ws.Cells(HEADER_ROW, COL_PROGRESS).Value = "진행률"
    ws.Cells(HEADER_ROW, COL_NORMAL_PROGRESS).Value = "정상 진행률"
    ws.Cells(HEADER_ROW, COL_MANUAL_PROGRESS).Value = "진행률 수동"
    ws.Cells(HEADER_ROW, COL_MANUAL_STATUS).Value = "상태 수동"
    ws.Cells(HEADER_ROW, COL_WEEKLY_REPORT).Value = "주간보고"
    ws.Cells(HEADER_ROW, COL_PLAN_DAYS).Value = "계획일수"
    ws.Cells(HEADER_ROW, COL_ACTUAL_DAYS).Value = "실소요일수"
    ws.Cells(HEADER_ROW, COL_STATUS).Value = "상태"
End Sub

Private Sub NormalizeReportStatusCell(ByVal targetCell As Range)
    Dim statusText As String

    statusText = UCase$(Trim$(CStr(targetCell.Value2)))

    Select Case statusText
        Case "Y", "TRUE"
            targetCell.Value = REPORT_STATUS_IN_PROGRESS
        Case "N", "FALSE"
            targetCell.ClearContents
        Case UCase$(REPORT_STATUS_PLANNED)
            targetCell.Value = REPORT_STATUS_PLANNED
        Case UCase$(REPORT_STATUS_IN_PROGRESS)
            targetCell.Value = REPORT_STATUS_IN_PROGRESS
        Case UCase$(REPORT_STATUS_COMPLETED)
            targetCell.Value = REPORT_STATUS_COMPLETED
    End Select
End Sub

Public Sub UpdateTaskNumbers(ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim seqNo As Long

    If lastRow < DATA_START_ROW Then Exit Sub

    seqNo = 1

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            ws.Cells(r, COL_NO).Value = seqNo
            ws.Cells(r, COL_NO).NumberFormat = "0"
            seqNo = seqNo + 1
        Else
            ws.Cells(r, COL_NO).ClearContents
        End If
    Next r
End Sub

Private Sub NormalizeSheetStructure(ws As Worksheet)
    Dim i As Long
    Dim r As Long
    Dim lastMigrationRow As Long
    Dim manualStatusText As String
    Dim targetRange As Range
    Dim lo As ListObject

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_NO).Value)) <> "No." Then
        ws.Columns(COL_NO).Insert Shift:=xlToRight
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_TYPE).Value)) <> "타입" Then
        ws.Columns(COL_TYPE).Insert Shift:=xlToRight
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_MAJOR_CATEGORY).Value)) <> "대분류" Then
        If Trim$(CStr(ws.Cells(HEADER_ROW, COL_MAJOR_CATEGORY).Value)) = "모듈" Then
            ws.Cells(HEADER_ROW, COL_MAJOR_CATEGORY).Value = "대분류"
        Else
            ws.Columns(COL_MAJOR_CATEGORY).Insert Shift:=xlToRight
        End If
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_MIDDLE_CATEGORY).Value)) <> "중분류" Then
        ws.Columns(COL_MIDDLE_CATEGORY).Insert Shift:=xlToRight
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_MINOR_CATEGORY).Value)) <> "소분류" Then
        If Trim$(CStr(ws.Cells(HEADER_ROW, COL_MINOR_CATEGORY).Value)) = "프로그램" Then
            ws.Cells(HEADER_ROW, COL_MINOR_CATEGORY).Value = "소분류"
        Else
            ws.Columns(COL_MINOR_CATEGORY).Insert Shift:=xlToRight
        End If
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_OWNER).Value)) <> "담당" Then
        ws.Columns(COL_OWNER).Insert Shift:=xlToRight
    End If

    ' Newly inserted classification columns can inherit validation from adjacent cells.
    On Error Resume Next
    ws.Range(COL_TYPE & DATA_START_ROW & ":" & _
             COL_MINOR_CATEGORY & ws.Rows.Count).Validation.Delete
    On Error GoTo 0

    ' Remove the retired development-report column from existing workbooks.
    If Trim$(CStr(ws.Range(COL_WEEKLY_REPORT & HEADER_ROW).Value)) = "개발진행" Then
        ws.Columns(COL_WEEKLY_REPORT).Delete Shift:=xlToLeft
    End If

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_WEEKLY_REPORT).Value)) <> "주간보고" Then
        ws.Columns(COL_WEEKLY_REPORT).Insert Shift:=xlToRight
    End If

    lastMigrationRow = Application.Max( _
        ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, COL_MANUAL_STATUS).End(xlUp).Row)

    For r = DATA_START_ROW To lastMigrationRow
        manualStatusText = Trim$(CStr(ws.Cells(r, COL_MANUAL_STATUS).Value))

        If manualStatusText = STATUS_WEEKLY_REPORT Then
            ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_IN_PROGRESS
            ws.Cells(r, COL_MANUAL_STATUS).ClearContents
        ElseIf manualStatusText = "개발진행" Or manualStatusText = "개발 진행" Then
            ws.Cells(r, COL_MANUAL_STATUS).ClearContents
        End If
        NormalizeReportStatusCell ws.Cells(r, COL_WEEKLY_REPORT)
    Next r

    Set targetRange = ws.Range(COL_NO & HEADER_ROW & ":" & ws.Cells(HEADER_ROW, ws.Columns.Count).Address(False, False))

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = ws.ListObjects(i)
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            lo.Unlist
        End If
    Next i
End Sub

Public Sub ClearCalculatedArea(ws As Worksheet, ByVal lastRow As Long)
    Dim normalProgressCol As Long
    Dim planDaysCol As Long
    Dim actualDaysCol As Long
    Dim statusCol As Long
    
    If lastRow < DATA_START_ROW Then Exit Sub
    
    normalProgressCol = FindHeaderColumnByName(ws, "정상 진행률")
    planDaysCol = FindHeaderColumnByName(ws, "계획일수")
    actualDaysCol = FindHeaderColumnByName(ws, "실소요일수")
    statusCol = FindHeaderColumnByName(ws, "상태")
    
    If normalProgressCol > 0 And planDaysCol > 0 And actualDaysCol > 0 And statusCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(lastRow, normalProgressCol)).ClearContents
        ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(lastRow, normalProgressCol)).ClearFormats
        
        ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(lastRow, actualDaysCol)).ClearContents
        ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(lastRow, actualDaysCol)).ClearFormats
        
        ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(lastRow, statusCol)).ClearContents
        ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(lastRow, statusCol)).ClearFormats
    Else
        ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).ClearContents
        ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).ClearFormats
        
        ws.Range(COL_PLAN_DAYS & DATA_START_ROW & ":" & COL_ACTUAL_DAYS & lastRow).ClearContents
        ws.Range(COL_PLAN_DAYS & DATA_START_ROW & ":" & COL_ACTUAL_DAYS & lastRow).ClearFormats
        
        ws.Range(COL_STATUS & DATA_START_ROW & ":" & COL_STATUS & lastRow).ClearContents
        ws.Range(COL_STATUS & DATA_START_ROW & ":" & COL_STATUS & lastRow).ClearFormats
    End If
End Sub

Public Sub ClearGanttArea(ws As Worksheet)
    Dim startCol As Long
    Dim lastCol As Long

    startCol = ws.Range(COL_GANTT_START & "1").Column
    lastCol = ws.Columns.Count

    On Error Resume Next
    ws.Range(ws.Cells(1, startCol), ws.Cells(ws.Rows.Count, lastCol)).UnMerge
    On Error GoTo 0

    ws.Range(ws.Cells(1, startCol), ws.Cells(ws.Rows.Count, lastCol)).ClearFormats
    ws.Range(ws.Cells(1, startCol), ws.Cells(ws.Rows.Count, lastCol)).ClearContents
End Sub

Public Sub DrawDateHeader(ws As Worksheet, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim startCol As Long
    Dim endCol As Long
    Dim curCol As Long
    Dim d As Date
    Dim weekStartCol As Long
    Dim currentWeekLabel As String
    Dim nextWeekLabel As String
    Dim monthStartCol As Long
    Dim currentMonthLabel As String
    Dim nextMonthLabel As String

    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = startCol + CLng(chartEndDate - chartStartDate)

    ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol)).ClearContents
    ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol)).ClearFormats

    curCol = startCol
    weekStartCol = startCol
    monthStartCol = startCol
    currentWeekLabel = ""
    currentMonthLabel = ""

    For d = chartStartDate To chartEndDate
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Value = GetWeekdayKorShort(d)
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).HorizontalAlignment = xlCenter
        ws.Cells(GANTT_HEADER_ROW_DAY, curCol).VerticalAlignment = xlCenter

        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Value = d
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).NumberFormat = "mm-dd"
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Orientation = 90
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).HorizontalAlignment = xlCenter
        ws.Cells(GANTT_HEADER_ROW_DATE, curCol).VerticalAlignment = xlCenter

        nextWeekLabel = DatePart("ww", d, vbMonday, vbFirstFourDays) & " 주차"
        nextMonthLabel = Month(d) & "월"

        If currentMonthLabel = "" Then
            currentMonthLabel = nextMonthLabel
            monthStartCol = curCol
        ElseIf currentMonthLabel <> nextMonthLabel Then
            With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, monthStartCol), ws.Cells(GANTT_HEADER_ROW_MONTH, curCol - 1))
                .Merge
                .Value = currentMonthLabel
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
            End With

            currentMonthLabel = nextMonthLabel
            monthStartCol = curCol
        End If

        If currentWeekLabel = "" Then
            currentWeekLabel = nextWeekLabel
            weekStartCol = curCol
        ElseIf currentWeekLabel <> nextWeekLabel Then
            With ws.Range(ws.Cells(GANTT_HEADER_ROW_WEEK, weekStartCol), ws.Cells(GANTT_HEADER_ROW_WEEK, curCol - 1))
                .Merge
                .Value = currentWeekLabel
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
            End With

            currentWeekLabel = nextWeekLabel
            weekStartCol = curCol
        End If

        If Not IsWorkingDay(d, holidayDict, workdayDict) Then
            ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Interior.Color = RGB(230, 230, 230)
            ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Interior.Color = RGB(230, 230, 230)
        End If

        If CLng(d) = CLng(Date) Then
            ws.Cells(GANTT_HEADER_ROW_DAY, curCol).Interior.Color = RGB(255, 120, 120)
            ws.Cells(GANTT_HEADER_ROW_DATE, curCol).Interior.Color = RGB(255, 120, 120)
        End If

        ws.Columns(curCol).ColumnWidth = 3
        curCol = curCol + 1
    Next d

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, monthStartCol), ws.Cells(GANTT_HEADER_ROW_MONTH, curCol - 1))
        .Merge
        .Value = currentMonthLabel
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_WEEK, weekStartCol), ws.Cells(GANTT_HEADER_ROW_WEEK, curCol - 1))
        .Merge
        .Value = currentWeekLabel
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Public Sub UpdateWeeklyReportStatuses(ByVal ws As Worksheet, _
                                      ByVal lastRow As Long)
    Dim r As Long
    Dim actualStart As Variant
    Dim actualEnd As Variant

    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            actualStart = ws.Cells(r, COL_ACTUAL_START).Value
            actualEnd = ws.Cells(r, COL_ACTUAL_END).Value

            If IsDate(actualEnd) And CLng(CDate(actualEnd)) <= CLng(Date) Then
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_COMPLETED
            ElseIf IsDate(actualStart) And CLng(CDate(actualStart)) <= CLng(Date) Then
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_IN_PROGRESS
            Else
                ws.Cells(r, COL_WEEKLY_REPORT).Value = REPORT_STATUS_PLANNED
            End If
        Else
            ws.Cells(r, COL_WEEKLY_REPORT).ClearContents
        End If
    Next r
End Sub
Public Sub DrawTaskBars(ws As Worksheet, ByVal lastRow As Long, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim r As Long
    Dim d As Date
    Dim colIdx As Long
    Dim startCol As Long
    Dim planS As Variant
    Dim planE As Variant
    Dim actS As Variant
    Dim actE As Variant
    Dim statusText As String
    Dim actualColor As Long
    Dim actualDrawEnd As Variant
    Dim markerDate As Variant
    Dim markerText As String

    ApplyHierarchySummaryValues ws, lastRow, holidayDict, workdayDict
    UpdateWeeklyReportStatuses ws, lastRow

    startCol = ws.Range(COL_GANTT_START & "1").Column

    For r = DATA_START_ROW To lastRow
        If Not HasTaskContent(ws, r) Then GoTo ContinueNextRow
        If HasTaskError(ws, r) Then GoTo ContinueNextRow

        planS = ws.Cells(r, COL_PLAN_START).Value
        planE = ws.Cells(r, COL_PLAN_END).Value
        actS = ws.Cells(r, COL_ACTUAL_START).Value
        actE = ws.Cells(r, COL_ACTUAL_END).Value
        statusText = CStr(ws.Cells(r, COL_STATUS).Value)
        markerDate = FindLevelMarkerDate(ws, r, holidayDict, workdayDict)
        markerText = GetTaskLevelMarker(ws, r)

        actualColor = GetTaskStatusFillColor(statusText)


        If IsDate(planS) And IsDate(planE) Then
            For d = CDate(planS) To CDate(planE)
                If d >= chartStartDate And d <= chartEndDate Then
                    If IsWorkingDay(d, holidayDict, workdayDict) Then
                        colIdx = startCol + CLng(d - chartStartDate)
                        With ws.Cells(r, colIdx)
                            .Interior.Color = RGB(180, 180, 180)
                            If Not IsEmpty(markerDate) Then
                                If CLng(d) = CLng(markerDate) Then
                                    .Value = markerText
                                    .Font.Color = RGB(0, 0, 0)
                                    .Font.Bold = True
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlCenter
                                End If
                            End If
                        End With
                    End If
                End If
            Next d
        End If

        actualDrawEnd = Empty
        If IsDate(actS) Then
            If IsDate(actE) Then
                actualDrawEnd = CDate(actE)
            ElseIf CDate(actS) <= Date Then
                actualDrawEnd = Date
            End If
        End If

        If IsDate(actS) And IsDate(actualDrawEnd) Then
            For d = CDate(actS) To CDate(actualDrawEnd)
                If d >= chartStartDate And d <= chartEndDate Then
                    If IsWorkingDay(d, holidayDict, workdayDict) Then
                        colIdx = startCol + CLng(d - chartStartDate)
                        With ws.Cells(r, colIdx)
                            .Interior.Color = actualColor
                            .Value = markerText
                            .Font.Color = RGB(0, 0, 0)
                            .Font.Bold = True
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                    End If
                End If
            Next d
        End If

ContinueNextRow:
    Next r
End Sub

Private Sub ApplyHierarchySummaryValues(ws As Worksheet, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim r As Long
    
    For r = lastRow To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            If HasChildTask(ws, r, lastRow) Then
                SetParentSummaryValues ws, r, lastRow, holidayDict, workdayDict
            Else
                NormalizeLeafProgressByActualEnd ws, r
            End If
            
            SetNormalProgressValue ws, r, holidayDict, workdayDict
            SetDurationValues ws, r, holidayDict, workdayDict
            SetStatusValue ws, r, lastRow, holidayDict, workdayDict
        End If
    Next r
End Sub

Private Sub NormalizeLeafProgressByActualEnd(ws As Worksheet, ByVal rowNum As Long)
    If IsDate(ws.Cells(rowNum, COL_ACTUAL_END).Value) Then
        ws.Cells(rowNum, COL_PROGRESS).Value = 1
        ws.Cells(rowNum, COL_PROGRESS).NumberFormat = "0%"
    End If
End Sub

Private Sub SetParentSummaryValues(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim currentLevel As Long
    Dim childRow As Long
    Dim childEndRow As Long
    
    Dim minPlanStart As Date
    Dim maxPlanEnd As Date
    Dim minActStart As Date
    Dim maxActEnd As Date
    
    Dim hasPlanStart As Boolean
    Dim hasPlanEnd As Boolean
    Dim hasActStart As Boolean
    Dim hasActEnd As Boolean
    Dim allChildActualDone As Boolean
    
    Dim vPlanS As Variant
    Dim vPlanE As Variant
    Dim vActS As Variant
    Dim vActE As Variant
    
    Dim childProgress As Double
    Dim childWeight As Double
    Dim weightedProgressSum As Double
    Dim totalWeight As Double
    
    currentLevel = GetTaskLevel(ws, rowNum)
    allChildActualDone = True
    
    childRow = rowNum + 1
    
    Do While childRow <= lastRow
        If HasTaskContent(ws, childRow) Then
            If GetTaskLevel(ws, childRow) <= currentLevel Then Exit Do
            
            vPlanS = ws.Cells(childRow, COL_PLAN_START).Value
            vPlanE = ws.Cells(childRow, COL_PLAN_END).Value
            vActS = ws.Cells(childRow, COL_ACTUAL_START).Value
            vActE = ws.Cells(childRow, COL_ACTUAL_END).Value
            
            If IsDate(vPlanS) Then
                If Not hasPlanStart Then
                    minPlanStart = CDate(vPlanS)
                    hasPlanStart = True
                ElseIf CDate(vPlanS) < minPlanStart Then
                    minPlanStart = CDate(vPlanS)
                End If
            End If
            
            If IsDate(vPlanE) Then
                If Not hasPlanEnd Then
                    maxPlanEnd = CDate(vPlanE)
                    hasPlanEnd = True
                ElseIf CDate(vPlanE) > maxPlanEnd Then
                    maxPlanEnd = CDate(vPlanE)
                End If
            End If
            
            If IsDate(vActS) Then
                If Not hasActStart Then
                    minActStart = CDate(vActS)
                    hasActStart = True
                ElseIf CDate(vActS) < minActStart Then
                    minActStart = CDate(vActS)
                End If
            End If
            
            If IsDate(vActE) Then
                If Not hasActEnd Then
                    maxActEnd = CDate(vActE)
                    hasActEnd = True
                ElseIf CDate(vActE) > maxActEnd Then
                    maxActEnd = CDate(vActE)
                End If
            Else
                allChildActualDone = False
            End If
            
            childProgress = GetTaskProgressValue(ws, childRow)
            childWeight = GetProgressWeight(ws, childRow, holidayDict, workdayDict)
            
            weightedProgressSum = weightedProgressSum + (childProgress * childWeight)
            totalWeight = totalWeight + childWeight
            
            childEndRow = GetTaskSubtreeEndRow(ws, childRow, lastRow)
            childRow = childEndRow + 1
        Else
            childRow = childRow + 1
        End If
    Loop
    
    If hasPlanStart Then
        ws.Cells(rowNum, COL_PLAN_START).Value = minPlanStart
    Else
        ws.Cells(rowNum, COL_PLAN_START).ClearContents
    End If
    
    If hasPlanEnd Then
        ws.Cells(rowNum, COL_PLAN_END).Value = maxPlanEnd
    Else
        ws.Cells(rowNum, COL_PLAN_END).ClearContents
    End If
    
    If hasActStart Then
        ws.Cells(rowNum, COL_ACTUAL_START).Value = minActStart
    Else
        ws.Cells(rowNum, COL_ACTUAL_START).ClearContents
    End If
    
    If allChildActualDone And hasActEnd Then
        ws.Cells(rowNum, COL_ACTUAL_END).Value = maxActEnd
    Else
        ws.Cells(rowNum, COL_ACTUAL_END).ClearContents
    End If
    
    If Not IsManualProgressEnabled(ws, rowNum) Then
        If totalWeight > 0 Then
            ws.Cells(rowNum, COL_PROGRESS).Value = weightedProgressSum / totalWeight
        Else
            ws.Cells(rowNum, COL_PROGRESS).ClearContents
        End If
    End If
End Sub

Private Sub SetNormalProgressValue(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim planS As Variant
    Dim planE As Variant
    Dim actE As Variant
    Dim requiredProgress As Double
    
    actE = ws.Cells(rowNum, COL_ACTUAL_END).Value
    
    If IsDate(actE) Then
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).Value = 1
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).NumberFormat = "0%"
        Exit Sub
    End If
    
    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value
    
    If IsDate(planS) And IsDate(planE) Then
        requiredProgress = GetRequiredNormalProgress(CDate(planS), CDate(planE), holidayDict, workdayDict)
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).Value = requiredProgress
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).NumberFormat = "0%"
    Else
        ws.Cells(rowNum, COL_NORMAL_PROGRESS).ClearContents
    End If
End Sub

Private Function IsManualProgressEnabled(ws As Worksheet, ByVal rowNum As Long) As Boolean
    IsManualProgressEnabled = (UCase$(Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_PROGRESS).Value))) = "Y")
End Function

Private Function GetTaskNoKey(ws As Worksheet, ByVal rowNum As Long) As String
    If IsNumeric(ws.Cells(rowNum, COL_NO).Value) Then
        GetTaskNoKey = Trim$(CStr(CLng(ws.Cells(rowNum, COL_NO).Value)))
    Else
        GetTaskNoKey = ""
    End If
End Function

Private Function GetProgressWeight(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object) As Double
    Dim v As Variant
    Dim calcDays As Variant
    
    v = ws.Cells(rowNum, COL_PLAN_DAYS).Value
    
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            GetProgressWeight = CDbl(v)
            Exit Function
        End If
    End If
    
    calcDays = CountWorkingDaysInclusive(ws.Cells(rowNum, COL_PLAN_START).Value, ws.Cells(rowNum, COL_PLAN_END).Value, holidayDict, workdayDict)
    
    If IsNumeric(calcDays) Then
        If CDbl(calcDays) > 0 Then
            GetProgressWeight = CDbl(calcDays)
            Exit Function
        End If
    End If
    
    GetProgressWeight = 1
End Function

Private Sub SetDurationValues(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim planDays As Variant
    Dim actualDays As Variant
    Dim actualStart As Variant
    Dim actualEnd As Variant
    Dim compareEnd As Date

    planDays = CountWorkingDaysInclusive(ws.Cells(rowNum, COL_PLAN_START).Value, ws.Cells(rowNum, COL_PLAN_END).Value, holidayDict, workdayDict)

    actualStart = ws.Cells(rowNum, COL_ACTUAL_START).Value
    actualEnd = ws.Cells(rowNum, COL_ACTUAL_END).Value
    
    If IsDate(actualStart) Then
        If IsDate(actualEnd) Then
            actualDays = CountWorkingDaysInclusive(actualStart, actualEnd, holidayDict, workdayDict)
        Else
            If CDate(actualStart) <= Date Then
                compareEnd = Date
                actualDays = CountWorkingDaysInclusive(actualStart, compareEnd, holidayDict, workdayDict)
            Else
                actualDays = ""
            End If
        End If
    Else
        actualDays = ""
    End If

    ws.Cells(rowNum, COL_PLAN_DAYS).Value = planDays
    ws.Cells(rowNum, COL_ACTUAL_DAYS).Value = actualDays

    ws.Cells(rowNum, COL_PLAN_DAYS).NumberFormat = "0"
    ws.Cells(rowNum, COL_ACTUAL_DAYS).NumberFormat = "0"
End Sub

Private Sub SetStatusValue(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long, ByVal holidayDict As Object, ByVal workdayDict As Object)
    If HasChildTask(ws, rowNum, lastRow) Then
        ws.Cells(rowNum, COL_STATUS).Value = BuildParentStatusFromChildren(ws, rowNum, lastRow)
    Else
        ws.Cells(rowNum, COL_STATUS).Value = GetTaskStatus(ws, rowNum, holidayDict, workdayDict)
    End If
End Sub

Private Function BuildParentStatusFromChildren(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long) As String
    Dim currentLevel As Long
    Dim childRow As Long
    Dim childEndRow As Long
    Dim childStatus As String
    
    Dim hasAnyChild As Boolean
    Dim hasError As Boolean
    Dim hasDelay As Boolean
    Dim hasCaution As Boolean
    Dim hasHold As Boolean
    Dim allDone As Boolean
    
    currentLevel = GetTaskLevel(ws, rowNum)
    allDone = True
    childRow = rowNum + 1
    
    Do While childRow <= lastRow
        If HasTaskContent(ws, childRow) Then
            If GetTaskLevel(ws, childRow) <= currentLevel Then Exit Do
            
            childStatus = CStr(ws.Cells(childRow, COL_STATUS).Value)
            hasAnyChild = True
            
            If IsDoneStatusText(childStatus) Then
                If HasHoldIncludedStatusText(childStatus) Then hasHold = True
            Else
                Select Case childStatus
                    Case STATUS_HOLD
                        hasHold = True
                        allDone = False
                    Case STATUS_ERROR
                        hasError = True
                        allDone = False
                    Case STATUS_DELAY
                        hasDelay = True
                        allDone = False
                    Case STATUS_CAUTION
                        hasCaution = True
                        allDone = False
                    Case Else
                        allDone = False
                End Select
            End If
            
            childEndRow = GetTaskSubtreeEndRow(ws, childRow, lastRow)
            childRow = childEndRow + 1
        Else
            childRow = childRow + 1
        End If
    Loop
    
    If Not hasAnyChild Then
        BuildParentStatusFromChildren = ""
    ElseIf hasError Then
        BuildParentStatusFromChildren = STATUS_ERROR
    ElseIf allDone Then
        If hasHold Then
            BuildParentStatusFromChildren = STATUS_DONE & STATUS_DONE_WITH_HOLD_SUFFIX
        Else
            BuildParentStatusFromChildren = STATUS_DONE
        End If
    ElseIf hasDelay Then
        BuildParentStatusFromChildren = STATUS_DELAY
    ElseIf hasCaution Then
        BuildParentStatusFromChildren = STATUS_CAUTION
    ElseIf hasHold Then
        BuildParentStatusFromChildren = STATUS_HOLD
    Else
        BuildParentStatusFromChildren = STATUS_NORMAL
    End If
End Function

Private Function GetTaskStatusFillColor(ByVal statusText As String) As Long
    If IsDoneStatusText(statusText) Then
        GetTaskStatusFillColor = RGB(102, 255, 51)
        Exit Function
    End If

    Select Case statusText
        Case STATUS_NORMAL
            GetTaskStatusFillColor = RGB(91, 155, 213)
        Case STATUS_CAUTION
            GetTaskStatusFillColor = RGB(255, 192, 0)
        Case STATUS_HOLD
            GetTaskStatusFillColor = RGB(153, 102, 204)
        Case STATUS_DELAY
            GetTaskStatusFillColor = RGB(255, 99, 99)
        Case STATUS_ERROR
            GetTaskStatusFillColor = RGB(237, 125, 49)
        Case Else
            GetTaskStatusFillColor = RGB(91, 155, 213)
    End Select
End Function

Private Function IsTaskPlannedOnDate(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim planS As Variant
    Dim planE As Variant

    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value

    If Not IsDate(planS) Or Not IsDate(planE) Then Exit Function
    If Not IsWorkingDay(targetDate, holidayDict, workdayDict) Then Exit Function

    IsTaskPlannedOnDate = (CLng(targetDate) >= CLng(CDate(planS)) And _
                           CLng(targetDate) <= CLng(CDate(planE)))
End Function

Private Function IsTaskActualOnDate(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim actualS As Variant
    Dim actualE As Variant
    Dim actualDrawEnd As Date

    actualS = ws.Cells(rowNum, COL_ACTUAL_START).Value
    actualE = ws.Cells(rowNum, COL_ACTUAL_END).Value

    If Not IsDate(actualS) Then Exit Function
    If Not IsWorkingDay(targetDate, holidayDict, workdayDict) Then Exit Function

    If IsDate(actualE) Then
        actualDrawEnd = CDate(actualE)
    ElseIf CDate(actualS) <= Date Then
        actualDrawEnd = Date
    Else
        Exit Function
    End If

    IsTaskActualOnDate = (CLng(targetDate) >= CLng(CDate(actualS)) And _
                          CLng(targetDate) <= CLng(actualDrawEnd))
End Function

Private Function IsDoneStatusText(ByVal statusText As String) As Boolean
    IsDoneStatusText = (statusText = STATUS_DONE Or statusText = STATUS_DONE & STATUS_DONE_WITH_HOLD_SUFFIX)
End Function

Private Function HasHoldIncludedStatusText(ByVal statusText As String) As Boolean
    HasHoldIncludedStatusText = (InStr(1, statusText, STATUS_DONE_WITH_HOLD_SUFFIX, vbTextCompare) > 0)
End Function

Public Sub FormatBaseArea(ws As Worksheet, ByVal lastRow As Long, ByVal chartStartDate As Date, ByVal chartEndDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object)
    Dim startCol As Long
    Dim endCol As Long
    Dim r As Long
    Dim c As Long
    Dim d As Date
    Dim prevDate As Date
    Dim filterLastRow As Long
    Dim statusText As String
    Dim dataLastCol As String
    Dim levelValue As Long
    Dim baseTaskText As String
    
    dataLastCol = COL_STATUS
    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = startCol + CLng(chartEndDate - chartStartDate)
    filterLastRow = IIf(lastRow < HEADER_ROW + 1, HEADER_ROW + 1, lastRow)

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).Borders.LineStyle = xlNone
    ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).Interior.Pattern = xlNone
    ws.Range(ws.Cells(DATA_START_ROW, startCol), ws.Cells(filterLastRow, endCol)).Borders.LineStyle = xlNone

    With ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & HEADER_ROW)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With

    With ws.Range(COL_PLAN_DAYS & HEADER_ROW & ":" & COL_STATUS & HEADER_ROW)
        .Interior.Color = RGB(217, 217, 217)
    End With

    ws.Range(COL_NO & HEADER_ROW & ":" & dataLastCol & filterLastRow).AutoFilter

    With ws.Range(ws.Cells(GANTT_HEADER_ROW_MONTH, startCol), ws.Cells(GANTT_HEADER_ROW_DATE, endCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 210)
    End With

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            With ws.Range(COL_NO & r & ":" & dataLastCol & r)
                .Borders.LineStyle = xlContinuous
                .VerticalAlignment = xlCenter
            End With
            
            With ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(210, 210, 210)
            End With
            
            levelValue = GetTaskLevel(ws, r)
            baseTaskText = RemoveTaskLevelPrefix(CStr(ws.Cells(r, COL_TASK).Value))
            ws.Cells(r, COL_TASK).Value = BuildTaskDisplayText(baseTaskText, levelValue)
            ws.Cells(r, COL_TASK).HorizontalAlignment = xlLeft
            
            ws.Cells(r, COL_NO).HorizontalAlignment = xlCenter
            ws.Cells(r, COL_NO).VerticalAlignment = xlCenter
            ws.Cells(r, COL_LEVEL).HorizontalAlignment = xlCenter
            ws.Cells(r, COL_LEVEL).VerticalAlignment = xlCenter
            
            statusText = CStr(ws.Cells(r, COL_STATUS).Value)
            
            If IsDoneStatusText(statusText) Then
                ws.Cells(r, COL_STATUS).Interior.Color = RGB(198, 239, 206)
            Else
                Select Case statusText
                    Case STATUS_NORMAL
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(221, 235, 247)
                    Case STATUS_CAUTION
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(255, 235, 156)
                    Case STATUS_HOLD
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(217, 210, 233)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(180, 167, 214)
                    Case STATUS_DELAY
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(255, 235, 235)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(255, 199, 206)
                    Case STATUS_ERROR
                        ws.Range(COL_NO & r & ":" & dataLastCol & r).Interior.Color = RGB(255, 242, 204)
                        ws.Cells(r, COL_STATUS).Interior.Color = RGB(244, 176, 132)
                End Select
            End If
        Else
            ws.Range(COL_NO & r & ":" & dataLastCol & r).Borders.LineStyle = xlNone
            If endCol >= startCol Then
                ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.LineStyle = xlNone
            End If
        End If
    Next r

    For c = startCol To endCol
        d = chartStartDate + (c - startCol)

        If Not IsWorkingDay(d, holidayDict, workdayDict) Then
            For r = DATA_START_ROW To lastRow
                If HasTaskContent(ws, r) Then
                    If ws.Cells(r, c).Value = "" Then
                        ws.Cells(r, c).Interior.Color = RGB(245, 245, 245)
                    End If
                End If
            Next r
        End If

        If CLng(d) = CLng(Date) Then
            For r = DATA_START_ROW To lastRow
                If HasTaskContent(ws, r) Then
                    If ws.Cells(r, c).Value = "" Then
                        If Not IsTaskPlannedOnDate(ws, r, d, holidayDict, workdayDict) And _
                           Not IsTaskActualOnDate(ws, r, d, holidayDict, workdayDict) Then
                            ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)
                        End If
                    End If
                End If
            Next r
        End If
        
        If c > startCol Then
            prevDate = chartStartDate + (c - startCol - 1)
            
            If Month(prevDate) <> Month(d) Or Year(prevDate) <> Year(d) Then
                With ws.Range(ws.Cells(DATA_START_ROW, c), ws.Cells(lastRow, c)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(160, 160, 160)
                End With
            End If
        End If
    Next c

    ws.Columns(COL_NO).ColumnWidth = 6
    ws.Columns(COL_LEVEL).ColumnWidth = 7
    ws.Columns(COL_TASK).WrapText = False
    ws.Columns(COL_TYPE).ColumnWidth = 12
    ws.Columns(COL_MAJOR_CATEGORY).ColumnWidth = 14
    ws.Columns(COL_MIDDLE_CATEGORY).ColumnWidth = 14
    ws.Columns(COL_MINOR_CATEGORY).ColumnWidth = 14
    ws.Columns(COL_TASK).AutoFit
    ws.Columns(COL_OWNER).ColumnWidth = 12
    ws.Columns(COL_NOTE).ColumnWidth = 4.5
    ws.Columns(COL_PLAN_START).ColumnWidth = 11
    ws.Columns(COL_PLAN_END).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_START).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_END).ColumnWidth = 11
    ws.Columns(COL_PROGRESS).ColumnWidth = 9
    ws.Columns(COL_NORMAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_MANUAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_MANUAL_STATUS).ColumnWidth = 11
    ws.Columns(COL_WEEKLY_REPORT).ColumnWidth = 12
    ws.Columns(COL_PLAN_DAYS).ColumnWidth = 9
    ws.Columns(COL_ACTUAL_DAYS).ColumnWidth = 10
    ws.Columns(COL_STATUS).ColumnWidth = 15

    ws.Rows(GANTT_HEADER_ROW_MONTH).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_WEEK).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_DAY).RowHeight = 20
    ws.Rows(GANTT_HEADER_ROW_DATE).RowHeight = 70

    ws.Range(COL_PLAN_START & DATA_START_ROW & ":" & COL_ACTUAL_END & lastRow).NumberFormat = "yyyy-mm-dd"
    ws.Range(COL_PROGRESS & DATA_START_ROW & ":" & COL_PROGRESS & lastRow).NumberFormat = "0%"
    ws.Range(COL_NORMAL_PROGRESS & DATA_START_ROW & ":" & COL_NORMAL_PROGRESS & lastRow).NumberFormat = "0%"
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub HandleTaskHierarchyChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim watchedRange As Range
    Dim changedRange As Range
    Dim lastRow As Long
    Dim r As Long
    Dim changedArea As Range
    Dim issueMessages As String
    Dim changedClassificationOrLevel As Boolean

    If ws Is Nothing Or Target Is Nothing Then Exit Sub
    If ws.Name = CONFIG_SHEET_NAME Then Exit Sub

    Set watchedRange = Union( _
        ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_LEVEL & ws.Rows.Count), _
        ws.Range(COL_TYPE & DATA_START_ROW & ":" & COL_MINOR_CATEGORY & ws.Rows.Count), _
        ws.Range(COL_TASK & DATA_START_ROW & ":" & COL_TASK & ws.Rows.Count), _
        ws.Range(COL_OWNER & DATA_START_ROW & ":" & COL_OWNER & ws.Rows.Count))
    Set changedRange = Intersect(Target, watchedRange)
    If changedRange Is Nothing Then Exit Sub

    lastRow = GetLastDataRow(ws)
    For Each changedArea In changedRange.Areas
        If changedArea.Row + changedArea.Rows.Count - 1 > lastRow Then
            lastRow = changedArea.Row + changedArea.Rows.Count - 1
        End If
    Next changedArea

    For r = DATA_START_ROW To lastRow
        changedClassificationOrLevel = _
            Not Intersect(changedRange, ws.Cells(r, COL_LEVEL)) Is Nothing Or _
            Not Intersect(changedRange, ws.Range(COL_TYPE & r & ":" & _
                                                 COL_MINOR_CATEGORY & r)) Is Nothing

        If changedClassificationOrLevel Then
            EnforceTaskModuleFromParent ws, r, issueMessages
            PropagateTaskModuleToDescendants ws, r, lastRow
        End If
    Next r

    SynchronizeTaskHierarchyOwners ws, lastRow

    If Len(issueMessages) > 0 Then
        MsgBox "분류 계층 정합성을 자동으로 수정했습니다." & vbCrLf & vbCrLf & _
               issueMessages & vbCrLf & _
               "하위 작업의 네 분류 값은 가장 가까운 부모 작업과 같아야 합니다.", _
               vbExclamation, _
               "분류 정합성 오류"
    End If
End Sub

Public Sub SynchronizeTaskHierarchyModules(ByVal ws As Worksheet, _
                                           ByVal lastRow As Long, _
                                           Optional ByVal showIssueMessage As Boolean = True)
    Dim r As Long
    Dim issueMessages As String

    If ws Is Nothing Then Exit Sub
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Or _
           Len(Trim$(CStr(ws.Cells(r, COL_LEVEL).Value2))) > 0 Then
            EnforceTaskModuleFromParent ws, r, issueMessages
        End If
    Next r

    SynchronizeTaskHierarchyOwners ws, lastRow

    If showIssueMessage And Len(issueMessages) > 0 Then
        MsgBox "기존 분류 계층 오류를 자동으로 수정했습니다." & vbCrLf & vbCrLf & _
               issueMessages, _
               vbExclamation, _
               "분류 정합성 오류"
    End If
End Sub

Public Sub SynchronizeTaskHierarchyOwners(ByVal ws As Worksheet, _
                                          ByVal lastRow As Long)
    Dim r As Long
    Dim ownerText As String

    If ws Is Nothing Then Exit Sub
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = lastRow To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) And HasChildTask(ws, r, lastRow) Then
            ownerText = GetDescendantLeafOwnerText(ws, r, lastRow)
            If Len(ownerText) = 0 Then
                ws.Cells(r, COL_OWNER).ClearContents
            Else
                ws.Cells(r, COL_OWNER).Value = ownerText
            End If
        End If
    Next r
End Sub

Private Function GetDescendantLeafOwnerText(ByVal ws As Worksheet, _
                                            ByVal parentRow As Long, _
                                            ByVal lastRow As Long) As String
    Dim ownerSet As Object
    Dim parentLevel As Long
    Dim r As Long

    Set ownerSet = CreateObject("Scripting.Dictionary")
    ownerSet.CompareMode = vbTextCompare
    parentLevel = GetTaskLevel(ws, parentRow)

    For r = parentRow + 1 To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) <= parentLevel Then Exit For
            If Not HasChildTask(ws, r, lastRow) Then
                AddDistinctOwnerNames ownerSet, CStr(ws.Cells(r, COL_OWNER).Value2)
            End If
        End If
    Next r

    GetDescendantLeafOwnerText = JoinOwnerNameSet(ownerSet, ", ")
End Function

Public Sub AddDistinctOwnerNames(ByVal ownerSet As Object, _
                                 ByVal rawOwnerText As String)
    Dim ownerNames As Variant
    Dim ownerName As Variant
    Dim normalizedText As String

    normalizedText = Trim$(rawOwnerText)
    If Len(normalizedText) = 0 Then Exit Sub

    If Left$(normalizedText, 1) = "(" And Right$(normalizedText, 1) = ")" Then
        normalizedText = Trim$(Mid$(normalizedText, 2, Len(normalizedText) - 2))
    End If
    normalizedText = Replace$(normalizedText, "，", ",")
    ownerNames = Split(normalizedText, ",")

    For Each ownerName In ownerNames
        normalizedText = Trim$(CStr(ownerName))
        If Len(normalizedText) > 0 Then
            If Not ownerSet.Exists(normalizedText) Then ownerSet.Add normalizedText, True
        End If
    Next ownerName
End Sub

Public Function JoinOwnerNameSet(ByVal ownerSet As Object, _
                                 ByVal delimiter As String) As String
    Dim ownerName As Variant
    Dim result As String

    For Each ownerName In ownerSet.Keys
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(ownerName)
    Next ownerName

    JoinOwnerNameSet = result
End Function

Public Sub ShowTaskInputErrorReasons(ByVal ws As Worksheet, ByVal lastRow As Long)
    Const MAX_DISPLAY_ERROR_COUNT As Long = 20
    Dim r As Long
    Dim errorCount As Long
    Dim errorText As String
    Dim reasonText As String
    Dim taskText As String
    Dim taskNo As String

    If ws Is Nothing Then Exit Sub

    For r = DATA_START_ROW To lastRow
        reasonText = GetTaskErrorReason(ws, r)
        If Len(reasonText) > 0 Then
            errorCount = errorCount + 1
            If errorCount <= MAX_DISPLAY_ERROR_COUNT Then
                taskText = Trim$(CStr(ws.Cells(r, COL_TASK).Value2))
                taskNo = Trim$(CStr(ws.Cells(r, COL_NO).Value2))
                If Len(taskText) = 0 Then taskText = "내용 미입력"
                If Len(taskNo) = 0 Then taskNo = "-"

                If Len(errorText) > 0 Then errorText = errorText & vbCrLf
                errorText = errorText & "- 행 " & r & " / No. " & taskNo & _
                            " [" & taskText & "]: " & reasonText
            End If
        End If
    Next r

    If errorCount > MAX_DISPLAY_ERROR_COUNT Then
        errorText = errorText & vbCrLf & "- 그 외 " & _
                    (errorCount - MAX_DISPLAY_ERROR_COUNT) & "건"
    End If

    If errorCount > 0 Then
        MsgBox "입력 오류 " & errorCount & "건이 있습니다." & vbCrLf & vbCrLf & _
               errorText, _
               vbExclamation, _
               "간트 입력 오류 상세"
    End If
End Sub

Public Function GetNearestParentTaskRow(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long) As Long
    Dim currentLevel As Long
    Dim candidateLevel As Long
    Dim r As Long

    currentLevel = GetTaskLevel(ws, rowNum)
    If currentLevel <= 1 Then Exit Function

    For r = rowNum - 1 To DATA_START_ROW Step -1
        If HasTaskContent(ws, r) Then
            candidateLevel = GetTaskLevel(ws, r)
            If candidateLevel < currentLevel Then
                GetNearestParentTaskRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub EnforceTaskModuleFromParent(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long, _
                                        ByRef issueMessages As String)
    Dim taskLevel As Long
    Dim parentRow As Long
    Dim parentTask As String
    Dim childTask As String
    Dim parentValues As Variant
    Dim childValues As Variant

    taskLevel = GetTaskLevel(ws, rowNum)
    If taskLevel <= 1 Then Exit Sub

    parentRow = GetNearestParentTaskRow(ws, rowNum)
    childTask = Trim$(CStr(ws.Cells(rowNum, COL_TASK).Value2))

    If parentRow = 0 Then
        ws.Range(COL_TYPE & rowNum & ":" & COL_MINOR_CATEGORY & rowNum).ClearContents
        AppendModuleConsistencyIssue issueMessages, rowNum, childTask, _
            "Level " & taskLevel & "이지만 위쪽에 부모 작업이 없습니다. " & _
            "부모 작업을 먼저 만들거나 Level을 1로 변경하세요."
        Exit Sub
    End If

    parentTask = Trim$(CStr(ws.Cells(parentRow, COL_TASK).Value2))
    parentValues = ws.Range(COL_TYPE & parentRow & ":" & _
                            COL_MINOR_CATEGORY & parentRow).Value2
    childValues = ws.Range(COL_TYPE & rowNum & ":" & _
                           COL_MINOR_CATEGORY & rowNum).Value2

    If Not ClassificationValuesEqual(parentValues, childValues) Then
        ws.Range(COL_TYPE & rowNum & ":" & _
                 COL_MINOR_CATEGORY & rowNum).Value2 = parentValues
        AppendModuleConsistencyIssue issueMessages, rowNum, childTask, _
            "입력한 분류가 부모 작업 '" & parentTask & "'(행 " & parentRow & _
            ")과 달라 부모의 타입/대분류/중분류/소분류로 맞췄습니다."
    End If
End Sub

Private Function ClassificationValuesEqual(ByVal leftValues As Variant, _
                                             ByVal rightValues As Variant) As Boolean
    Dim i As Long

    For i = 1 To 4
        If StrComp(Trim$(CStr(leftValues(1, i))), _
                   Trim$(CStr(rightValues(1, i))), vbTextCompare) <> 0 Then Exit Function
    Next i
    ClassificationValuesEqual = True
End Function

Private Sub PropagateTaskModuleToDescendants(ByVal ws As Worksheet, _
                                             ByVal rowNum As Long, _
                                             ByVal lastRow As Long)
    Dim parentLevel As Long
    Dim subtreeEndRow As Long
    Dim classificationValues As Variant
    Dim r As Long

    parentLevel = GetTaskLevel(ws, rowNum)
    subtreeEndRow = GetTaskSubtreeEndRow(ws, rowNum, lastRow)
    If subtreeEndRow <= rowNum Then Exit Sub

    classificationValues = ws.Range(COL_TYPE & rowNum & ":" & _
                                    COL_MINOR_CATEGORY & rowNum).Value2
    For r = rowNum + 1 To subtreeEndRow
        If GetTaskLevel(ws, r) > parentLevel And _
           (HasTaskContent(ws, r) Or _
            Len(Trim$(CStr(ws.Cells(r, COL_LEVEL).Value2))) > 0) Then
            ws.Range(COL_TYPE & r & ":" & _
                     COL_MINOR_CATEGORY & r).Value2 = classificationValues
        End If
    Next r
End Sub

Private Sub AppendModuleConsistencyIssue(ByRef issueMessages As String, _
                                         ByVal rowNum As Long, _
                                         ByVal taskText As String, _
                                         ByVal reasonText As String)
    If Len(taskText) = 0 Then taskText = "내용 미입력"
    If Len(issueMessages) > 0 Then issueMessages = issueMessages & vbCrLf
    issueMessages = issueMessages & "- 행 " & rowNum & " [" & taskText & "]: " & reasonText
End Sub

Public Sub ApplyTaskInputValidation(ws As Worksheet)
    Dim lastSheetRow As Long
    Dim rngDate As Range
    Dim rngClassification As Range
    Dim rngLevel As Range
    Dim rngProgress As Range
    Dim rngManualProgress As Range
    Dim rngWeeklyReport As Range
    
    lastSheetRow = ws.Rows.Count

    Set rngClassification = ws.Range(COL_TYPE & DATA_START_ROW & ":" & _
                                     COL_MINOR_CATEGORY & lastSheetRow)
    On Error Resume Next
    rngClassification.Validation.Delete
    On Error GoTo 0

    ApplyTaskTextLengthValidation ws
    
    Set rngDate = Union( _
        ws.Range(COL_PLAN_START & DATA_START_ROW & ":" & COL_PLAN_START & lastSheetRow), _
        ws.Range(COL_PLAN_END & DATA_START_ROW & ":" & COL_PLAN_END & lastSheetRow), _
        ws.Range(COL_ACTUAL_START & DATA_START_ROW & ":" & COL_ACTUAL_START & lastSheetRow), _
        ws.Range(COL_ACTUAL_END & DATA_START_ROW & ":" & COL_ACTUAL_END & lastSheetRow))
    
    On Error Resume Next
    rngDate.Validation.Delete
    On Error GoTo 0
    
    rngDate.NumberFormat = "yyyy-mm-dd"
    
    rngDate.Validation.Add Type:=xlValidateDate, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:="2000-01-01", _
                           Formula2:="2100-12-31"
    
    rngDate.Validation.IgnoreBlank = True
    rngDate.Validation.InCellDropdown = True
    rngDate.Validation.InputTitle = "날짜 입력"
    rngDate.Validation.InputMessage = "날짜 형식으로 입력하세요."
    rngDate.Validation.ErrorTitle = "입력 오류"
    rngDate.Validation.ErrorMessage = "올바른 날짜만 입력할 수 있습니다."
    
    Set rngLevel = ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_LEVEL & lastSheetRow)
    
    On Error Resume Next
    rngLevel.Validation.Delete
    On Error GoTo 0
    
    rngLevel.Validation.Add Type:=xlValidateWholeNumber, _
                            AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, _
                            Formula1:="1", _
                            Formula2:="3"
    
    rngLevel.Validation.IgnoreBlank = True
    rngLevel.Validation.InputTitle = "Level 입력"
    rngLevel.Validation.InputMessage = "1~3 사이 정수만 입력하세요."
    rngLevel.Validation.ErrorTitle = "입력 오류"
    rngLevel.Validation.ErrorMessage = "Level은 1~3 사이 정수만 입력할 수 있습니다."
    
    Set rngProgress = ws.Range(COL_PROGRESS & DATA_START_ROW & ":" & COL_PROGRESS & lastSheetRow)
    
    On Error Resume Next
    rngProgress.Validation.Delete
    On Error GoTo 0
    
    rngProgress.NumberFormat = "0%"
    
    rngProgress.Validation.Add Type:=xlValidateDecimal, _
                               AlertStyle:=xlValidAlertStop, _
                               Operator:=xlBetween, _
                               Formula1:="0", _
                               Formula2:="1"
    
    rngProgress.Validation.IgnoreBlank = True
    rngProgress.Validation.InCellDropdown = True
    rngProgress.Validation.InputTitle = "진행률 입력"
    rngProgress.Validation.InputMessage = "0% ~ 100% 사이로 입력하세요. 예: 50%"
    rngProgress.Validation.ErrorTitle = "입력 오류"
    rngProgress.Validation.ErrorMessage = "진행률은 0% ~ 100% 사이만 입력할 수 있습니다."

    Set rngManualProgress = ws.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & COL_MANUAL_PROGRESS & lastSheetRow)

    On Error Resume Next
    rngManualProgress.Validation.Delete
    On Error GoTo 0

    rngManualProgress.Validation.Add Type:=xlValidateList, _
                                     AlertStyle:=xlValidAlertStop, _
                                     Operator:=xlBetween, _
                                     Formula1:="Y,N"

    rngManualProgress.Validation.IgnoreBlank = True
    rngManualProgress.Validation.InCellDropdown = True
    rngManualProgress.Validation.InputTitle = "진행률 수동 입력"
    rngManualProgress.Validation.InputMessage = "수동 입력 허용은 Y, 자동 계산은 N으로 입력하세요."
    rngManualProgress.Validation.ErrorTitle = "입력 오류"
    rngManualProgress.Validation.ErrorMessage = "Y 또는 N만 입력할 수 있습니다."
    
    Set rngWeeklyReport = ws.Range(COL_WEEKLY_REPORT & DATA_START_ROW & ":" & COL_WEEKLY_REPORT & lastSheetRow)

    On Error Resume Next
    rngWeeklyReport.Validation.Delete
    On Error GoTo 0

    rngWeeklyReport.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=REPORT_STATUS_PLANNED & "," & REPORT_STATUS_IN_PROGRESS & "," & REPORT_STATUS_COMPLETED
    rngWeeklyReport.Validation.IgnoreBlank = True
    rngWeeklyReport.Validation.InCellDropdown = True
    rngWeeklyReport.Validation.InputTitle = "주간보고 상태"
    rngWeeklyReport.Validation.InputMessage = "Planned, In Progress, Completed를 선택하세요."
    rngWeeklyReport.Validation.ErrorTitle = "입력 오류"
    rngWeeklyReport.Validation.ErrorMessage = "Planned, In Progress, Completed만 입력할 수 있습니다."

    ApplyManualStatusValidation ws, lastSheetRow
End Sub

Public Sub ApplyTaskTextLengthValidation(ByVal ws As Worksheet)
    Dim lastValidationRow As Long
    Dim rngTask As Range
    Dim validationFormula As String

    If ws Is Nothing Then Exit Sub

    lastValidationRow = GetLastDataRow(ws) + TASK_VALIDATION_EXTRA_ROWS
    If lastValidationRow < TASK_VALIDATION_MIN_LAST_ROW Then
        lastValidationRow = TASK_VALIDATION_MIN_LAST_ROW
    End If
    If lastValidationRow > ws.Rows.Count Then lastValidationRow = ws.Rows.Count

    Set rngTask = ws.Range(COL_TASK & DATA_START_ROW & ":" & _
                           COL_TASK & lastValidationRow)

    On Error Resume Next
    rngTask.Validation.Delete
    On Error GoTo 0

    validationFormula = "=LEN(" & COL_TASK & DATA_START_ROW & ")<=" & _
        "IF(" & COL_LEVEL & DATA_START_ROW & "=2," & _
        "INDIRECT(""'" & CONFIG_SHEET_NAME & "'!$" & _
        TASK_MAX_LENGTH_LEVEL2_VALUE_CELL & """)," & _
        "IF(" & COL_LEVEL & DATA_START_ROW & "=3," & _
        "INDIRECT(""'" & CONFIG_SHEET_NAME & "'!$" & _
        TASK_MAX_LENGTH_LEVEL3_VALUE_CELL & """)," & _
        "INDIRECT(""'" & CONFIG_SHEET_NAME & "'!$" & _
        TASK_MAX_LENGTH_LEVEL1_VALUE_CELL & """)))"

    rngTask.Validation.Add Type:=xlValidateCustom, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=validationFormula
    rngTask.Validation.IgnoreBlank = True
    rngTask.Validation.InputTitle = "내용 입력"
    rngTask.Validation.InputMessage = _
        "내용은 config 시트의 Level별 최대 글자 수 이내로 입력하세요."
    rngTask.Validation.ErrorTitle = "내용 글자 수 초과"
    rngTask.Validation.ErrorMessage = _
        "입력한 내용이 현재 Level에 설정된 최대 글자 수를 초과했습니다. " & _
        "config 시트의 입력 제한 설정을 확인하세요."
End Sub

Private Sub ApplyManualStatusValidation(ws As Worksheet, ByVal lastSheetRow As Long)
    Dim rngManualStatus As Range

    Set rngManualStatus = ws.Range(COL_MANUAL_STATUS & DATA_START_ROW & ":" & COL_MANUAL_STATUS & lastSheetRow)

    On Error Resume Next
    rngManualStatus.Validation.Delete
    On Error GoTo 0

    rngManualStatus.Validation.Add Type:=xlValidateList, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:=STATUS_HOLD

    rngManualStatus.Validation.IgnoreBlank = True
    rngManualStatus.Validation.InCellDropdown = True
End Sub

Private Function GetWeekdayKorShort(ByVal targetDate As Date) As String
    Select Case Weekday(targetDate, vbSunday)
        Case 1: GetWeekdayKorShort = "일"
        Case 2: GetWeekdayKorShort = "월"
        Case 3: GetWeekdayKorShort = "화"
        Case 4: GetWeekdayKorShort = "수"
        Case 5: GetWeekdayKorShort = "목"
        Case 6: GetWeekdayKorShort = "금"
        Case 7: GetWeekdayKorShort = "토"
    End Select
End Function

Private Function BuildTaskDisplayText(ByVal baseTaskText As String, ByVal levelValue As Long) As String
    Select Case levelValue
        Case 2
            BuildTaskDisplayText = Space$(4) & baseTaskText
        Case 3
            BuildTaskDisplayText = Space$(8) & baseTaskText
        Case Else
            BuildTaskDisplayText = baseTaskText
    End Select
End Function

Private Function RemoveTaskLevelPrefix(ByVal taskText As String) As String
    If Left$(taskText, 8) = Space$(8) Then
        RemoveTaskLevelPrefix = Mid$(taskText, 9)
    ElseIf Left$(taskText, 4) = Space$(4) Then
        RemoveTaskLevelPrefix = Mid$(taskText, 5)
    Else
        RemoveTaskLevelPrefix = taskText
    End If
End Function

Private Function FindHeaderColumnByName(ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(HEADER_ROW, c).Value)) = headerText Then
            FindHeaderColumnByName = c
            Exit Function
        End If
    Next c
End Function

Public Sub ShowAllTaskRows(ws As Worksheet, ByVal lastRow As Long)
    If lastRow < DATA_START_ROW Then Exit Sub
    ws.Rows(DATA_START_ROW & ":" & lastRow).Hidden = False
End Sub

Public Sub ApplyDisplayTaskRowFilter(ws As Worksheet, ByVal lastRow As Long, ByVal chartStartDate As Date, ByVal chartEndDate As Date)
    Dim ganttOnlyFlag As Boolean
    Dim reportOnlyFlag As String
    Dim r As Long
    Dim showRow As Boolean

    If lastRow < DATA_START_ROW Then Exit Sub

    ganttOnlyFlag = GetDisplayGanttOnlyFlag()
    reportOnlyFlag = GetDisplayReportOnlyFlag()

    If Not ganttOnlyFlag And (reportOnlyFlag = "" Or reportOnlyFlag = REPORT_FILTER_ALL) Then Exit Sub

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            showRow = True

            If ganttOnlyFlag Then
                showRow = HasTaskDateInDisplayRange(ws, r, chartStartDate, chartEndDate)
            End If

            If showRow Then
                Select Case reportOnlyFlag
                    Case STATUS_WEEKLY_REPORT
                        showRow = IsTaskManualStatus(ws, r, reportOnlyFlag)
                    Case REPORT_FILTER_EMPTY
                        showRow = IsTaskManualStatusEmpty(ws, r)
                End Select
            End If

            ws.Rows(r).Hidden = Not showRow
        End If
    Next r
End Sub

Private Function HasTaskDateInDisplayRange(ws As Worksheet, ByVal rowNum As Long, ByVal displayStartDate As Date, ByVal displayEndDate As Date) As Boolean
    HasTaskDateInDisplayRange = _
        IsTaskDateRangeOverlap(ws.Cells(rowNum, COL_PLAN_START).Value, ws.Cells(rowNum, COL_PLAN_END).Value, displayStartDate, displayEndDate) Or _
        IsTaskDateRangeOverlap(ws.Cells(rowNum, COL_ACTUAL_START).Value, ws.Cells(rowNum, COL_ACTUAL_END).Value, displayStartDate, displayEndDate)
End Function

Private Function IsTaskDateRangeOverlap(ByVal startValue As Variant, ByVal endValue As Variant, ByVal displayStartDate As Date, ByVal displayEndDate As Date) As Boolean
    If Not IsDate(startValue) Or Not IsDate(endValue) Then Exit Function

    IsTaskDateRangeOverlap = (CLng(CDate(startValue)) <= CLng(displayEndDate) And CLng(CDate(endValue)) >= CLng(displayStartDate))
End Function

Private Function IsTaskManualStatus(ws As Worksheet, ByVal rowNum As Long, ByVal statusText As String) As Boolean
    Select Case statusText
        Case STATUS_WEEKLY_REPORT
            IsTaskManualStatus = HasReportStatusValue(ws.Cells(rowNum, COL_WEEKLY_REPORT).Value)
        Case Else
            IsTaskManualStatus = (Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_STATUS).Value)) = statusText)
    End Select
End Function

Private Function IsTaskManualStatusEmpty(ws As Worksheet, ByVal rowNum As Long) As Boolean
    IsTaskManualStatusEmpty = _
        Not HasReportStatusValue(ws.Cells(rowNum, COL_WEEKLY_REPORT).Value)
End Function

Private Function HasReportStatusValue(ByVal statusValue As Variant) As Boolean
    Dim statusText As String

    statusText = UCase$(Trim$(CStr(statusValue)))
    HasReportStatusValue = _
        (statusText = UCase$(REPORT_STATUS_PLANNED) Or _
         statusText = UCase$(REPORT_STATUS_IN_PROGRESS) Or _
         statusText = UCase$(REPORT_STATUS_COMPLETED))
End Function
Public Sub ShowAllDateColumns(ws As Worksheet)
    Dim startCol As Long
    Dim endCol As Long

    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = ws.Cells(GANTT_HEADER_ROW_DATE, ws.Columns.Count).End(xlToLeft).Column

    If endCol < startCol Then Exit Sub
    ws.Range(ws.Cells(1, startCol), ws.Cells(1, endCol)).EntireColumn.Hidden = False
End Sub

Public Sub HideCompletedTaskRows(ws As Worksheet, ByVal lastRow As Long)
    Dim excludeNoDict As Object
    Dim maxHideLevel As Long
    Dim r As Long
    Dim noKey As String

    ShowAllTaskRows ws, lastRow
    LoadExcludedRowNos excludeNoDict
    maxHideLevel = GetHideCompletedMaxLevel()

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            noKey = GetTaskNoKey(ws, r)

            If IsDoneStatusText(CStr(ws.Cells(r, COL_STATUS).Value)) Then
                If GetTaskLevel(ws, r) <= maxHideLevel Then
                    If Len(noKey) = 0 Or Not excludeNoDict.Exists(noKey) Then
                        ws.Rows(r).Hidden = True
                    End If
                End If
            End If
        End If
    Next r
End Sub

Public Sub ApplyCalculatedColumnsProtection(ws As Worksheet, ByVal lastRow As Long)
    Dim normalProgressCol As Long
    Dim planDaysCol As Long
    Dim actualDaysCol As Long
    Dim statusCol As Long
    Dim r As Long

    ws.Cells.Locked = False

    ws.Range(ws.Cells(DATA_START_ROW, ws.Range(COL_NO & "1").Column), ws.Cells(ws.Rows.Count, ws.Range(COL_NO & "1").Column)).Locked = True

    normalProgressCol = FindHeaderColumnByName(ws, "정상 진행률")
    planDaysCol = FindHeaderColumnByName(ws, "계획일수")
    actualDaysCol = FindHeaderColumnByName(ws, "실소요일수")
    statusCol = FindHeaderColumnByName(ws, "상태")

    If normalProgressCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, normalProgressCol), ws.Cells(ws.Rows.Count, normalProgressCol)).Locked = True
    End If

    If planDaysCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, planDaysCol), ws.Cells(ws.Rows.Count, planDaysCol)).Locked = True
    End If

    If actualDaysCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, actualDaysCol), ws.Cells(ws.Rows.Count, actualDaysCol)).Locked = True
    End If

    If statusCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, statusCol), ws.Cells(ws.Rows.Count, statusCol)).Locked = True
    End If

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If HasChildTask(ws, r, lastRow) Then
                ws.Cells(r, COL_PROGRESS).Locked = Not IsManualProgressEnabled(ws, r)
            End If
        End If
    Next r

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ws.Rows("1:" & HEADER_ROW).Locked = True
    ws.Rows(DATA_START_ROW & ":" & ws.Rows.Count).Locked = False

    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFiltering:=True, _
               AllowInsertingRows:=True, AllowDeletingRows:=True
End Sub

Public Sub UnprotectTaskSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
End Sub



