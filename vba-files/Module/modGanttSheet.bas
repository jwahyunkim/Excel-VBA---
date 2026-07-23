Attribute VB_Name = "modGanttSheet"
Option Explicit

Public Sub SetupDataHeaders(ws As Worksheet)
    NormalizeSheetStructure ws

    ws.Cells(HEADER_ROW, COL_NO).Value = "No."
    ws.Cells(HEADER_ROW, COL_LEVEL).Value = "Level"
    ws.Cells(HEADER_ROW, COL_TASK).Value = "내용"
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
    ws.Cells(HEADER_ROW, COL_DEV_PROGRESS).Value = "개발진행"
    ws.Cells(HEADER_ROW, COL_PLAN_DAYS).Value = "계획일수"
    ws.Cells(HEADER_ROW, COL_ACTUAL_DAYS).Value = "실소요일수"
    ws.Cells(HEADER_ROW, COL_STATUS).Value = "상태"
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

    If Trim$(CStr(ws.Cells(HEADER_ROW, COL_WEEKLY_REPORT).Value)) <> "주간보고" And _
       Trim$(CStr(ws.Cells(HEADER_ROW, COL_DEV_PROGRESS).Value)) <> "개발진행" Then
        ws.Range(COL_WEEKLY_REPORT & ":" & COL_DEV_PROGRESS).EntireColumn.Insert Shift:=xlToRight
    Else
        If Trim$(CStr(ws.Cells(HEADER_ROW, COL_WEEKLY_REPORT).Value)) <> "주간보고" Then
            ws.Columns(COL_WEEKLY_REPORT).Insert Shift:=xlToRight
        End If
        If Trim$(CStr(ws.Cells(HEADER_ROW, COL_DEV_PROGRESS).Value)) <> "개발진행" Then
            ws.Columns(COL_DEV_PROGRESS).Insert Shift:=xlToRight
        End If
    End If

    lastMigrationRow = Application.Max( _
        ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, COL_MANUAL_STATUS).End(xlUp).Row)

    For r = DATA_START_ROW To lastMigrationRow
        manualStatusText = Trim$(CStr(ws.Cells(r, COL_MANUAL_STATUS).Value))

        If manualStatusText = STATUS_WEEKLY_REPORT Then
            ws.Cells(r, COL_WEEKLY_REPORT).Value = "Y"
            ws.Cells(r, COL_MANUAL_STATUS).ClearContents
        ElseIf manualStatusText = STATUS_DEV_PROGRESS Or manualStatusText = "개발 진행" Then
            ws.Cells(r, COL_DEV_PROGRESS).Value = "Y"
            ws.Cells(r, COL_MANUAL_STATUS).ClearContents
        End If
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

        If statusText = STATUS_HOLD Then
            actualColor = RGB(191, 191, 191)
        ElseIf statusText = STATUS_DELAY Then
            actualColor = RGB(255, 102, 102)
        ElseIf statusText = STATUS_CAUTION Then
            actualColor = RGB(255, 192, 0)
        Else
            actualColor = RGB(102, 255, 51)
        End If

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
                        ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)
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
    ws.Columns(COL_TASK).AutoFit
    ws.Columns(COL_NOTE).ColumnWidth = 14
    ws.Columns(COL_PLAN_START).ColumnWidth = 11
    ws.Columns(COL_PLAN_END).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_START).ColumnWidth = 11
    ws.Columns(COL_ACTUAL_END).ColumnWidth = 11
    ws.Columns(COL_PROGRESS).ColumnWidth = 9
    ws.Columns(COL_NORMAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_MANUAL_PROGRESS).ColumnWidth = 11
    ws.Columns(COL_MANUAL_STATUS).ColumnWidth = 11
    ws.Columns(COL_WEEKLY_REPORT).ColumnWidth = 9
    ws.Columns(COL_DEV_PROGRESS).ColumnWidth = 9
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

Public Sub ApplyTaskInputValidation(ws As Worksheet)
    Dim lastSheetRow As Long
    Dim rngDate As Range
    Dim rngLevel As Range
    Dim rngProgress As Range
    Dim rngManualProgress As Range
    Dim rngReportFlags As Range
    
    lastSheetRow = ws.Rows.Count
    
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
    
    Set rngReportFlags = Union( _
        ws.Range(COL_WEEKLY_REPORT & DATA_START_ROW & ":" & COL_WEEKLY_REPORT & lastSheetRow), _
        ws.Range(COL_DEV_PROGRESS & DATA_START_ROW & ":" & COL_DEV_PROGRESS & lastSheetRow))

    On Error Resume Next
    rngReportFlags.Validation.Delete
    On Error GoTo 0

    rngReportFlags.Validation.Add Type:=xlValidateList, _
                                  AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlBetween, _
                                  Formula1:="Y,N"

    rngReportFlags.Validation.IgnoreBlank = True
    rngReportFlags.Validation.InCellDropdown = True
    rngReportFlags.Validation.InputTitle = "체크"
    rngReportFlags.Validation.InputMessage = "해당하면 Y, 아니면 N을 선택하세요."
    rngReportFlags.Validation.ErrorTitle = "입력 오류"
    rngReportFlags.Validation.ErrorMessage = "Y 또는 N만 입력할 수 있습니다."

    ApplyManualStatusValidation ws, lastSheetRow
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
                    Case STATUS_WEEKLY_REPORT, STATUS_DEV_PROGRESS
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
            IsTaskManualStatus = IsCheckedFlag(ws.Cells(rowNum, COL_WEEKLY_REPORT).Value)
        Case STATUS_DEV_PROGRESS
            IsTaskManualStatus = IsCheckedFlag(ws.Cells(rowNum, COL_DEV_PROGRESS).Value)
        Case Else
            IsTaskManualStatus = (Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_STATUS).Value)) = statusText)
    End Select
End Function

Private Function IsTaskManualStatusEmpty(ws As Worksheet, ByVal rowNum As Long) As Boolean
    IsTaskManualStatusEmpty = _
        Not IsCheckedFlag(ws.Cells(rowNum, COL_WEEKLY_REPORT).Value) And _
        Not IsCheckedFlag(ws.Cells(rowNum, COL_DEV_PROGRESS).Value)
End Function

Private Function IsCheckedFlag(ByVal flagValue As Variant) As Boolean
    Dim flagText As String

    flagText = UCase$(Trim$(CStr(flagValue)))
    IsCheckedFlag = (flagText = "Y" Or flagText = "TRUE")
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

Public Sub HideIdleDateColumns(ws As Worksheet, ByVal lastRow As Long)
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim excludeDateDict As Object
    Dim startCol As Long
    Dim endCol As Long
    Dim c As Long
    Dim seqStartCol As Long
    Dim seqCount As Long
    Dim thresholdDays As Long
    Dim targetDate As Date
    Dim canHideThisCol As Boolean

    ShowAllDateColumns ws
    EnsureHolidaySheet
    LoadHolidaySettings holidayDict, workdayDict
    LoadExcludedDates excludeDateDict
    thresholdDays = GetHideIdlePeriodDays()

    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = ws.Cells(GANTT_HEADER_ROW_DATE, ws.Columns.Count).End(xlToLeft).Column

    If endCol < startCol Then Exit Sub

    seqStartCol = 0
    seqCount = 0

    For c = startCol To endCol
        If IsDate(ws.Cells(GANTT_HEADER_ROW_DATE, c).Value) Then
            targetDate = CDate(ws.Cells(GANTT_HEADER_ROW_DATE, c).Value)
            canHideThisCol = CanHideDateColumn(ws, lastRow, targetDate, holidayDict, workdayDict, excludeDateDict)
        Else
            canHideThisCol = False
        End If

        If canHideThisCol Then
            If seqStartCol = 0 Then seqStartCol = c
            seqCount = seqCount + 1
        Else
            If seqCount >= thresholdDays Then
                ws.Range(ws.Cells(1, seqStartCol), ws.Cells(1, c - 1)).EntireColumn.Hidden = True
            End If
            seqStartCol = 0
            seqCount = 0
        End If
    Next c

    If seqCount >= thresholdDays Then
        ws.Range(ws.Cells(1, seqStartCol), ws.Cells(1, endCol)).EntireColumn.Hidden = True
    End If
End Sub

Private Function CanHideDateColumn(ws As Worksheet, ByVal lastRow As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object, ByVal excludeDateDict As Object) As Boolean
    Dim dateKey As String

    dateKey = NormalizeDateKey(targetDate)

    If excludeDateDict.Exists(dateKey) Then
        CanHideDateColumn = False
        Exit Function
    End If

    CanHideDateColumn = Not HasAnyIncompleteActualTaskOnDate(ws, lastRow, targetDate, holidayDict, workdayDict)
End Function

Private Function HasAnyIncompleteActualTaskOnDate(ws As Worksheet, ByVal lastRow As Long, ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim r As Long
    Dim actS As Variant
    Dim actE As Variant
    Dim compareEnd As Date
    Dim statusText As String

    If Not IsWorkingDay(targetDate, holidayDict, workdayDict) Then
        HasAnyIncompleteActualTaskOnDate = False
        Exit Function
    End If

    For r = DATA_START_ROW To lastRow
        If HasTaskContent(ws, r) Then
            If Not HasChildTask(ws, r, lastRow) Then
                statusText = CStr(ws.Cells(r, COL_STATUS).Value)

                If Not IsDoneStatusText(statusText) And statusText <> STATUS_ERROR And statusText <> STATUS_HOLD Then
                    actS = ws.Cells(r, COL_ACTUAL_START).Value
                    actE = ws.Cells(r, COL_ACTUAL_END).Value

                    If IsDate(actS) Then
                        If IsDate(actE) Then
                            compareEnd = CDate(actE)
                        Else
                            compareEnd = Date
                        End If

                        If CDate(actS) <= targetDate And compareEnd >= targetDate Then
                            HasAnyIncompleteActualTaskOnDate = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r

    HasAnyIncompleteActualTaskOnDate = False
End Function

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



