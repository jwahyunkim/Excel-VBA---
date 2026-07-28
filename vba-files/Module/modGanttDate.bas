Attribute VB_Name = "modGanttDate"

Option Explicit

Public Function GetLastDataRow(ws As Worksheet) As Long
    Dim lastRow As Long
    Dim colAddr As Variant
    Dim colLastRow As Long

    lastRow = HEADER_ROW

    For Each colAddr In Array(COL_LEVEL, COL_TASK, COL_NOTE, COL_PLAN_START, COL_PLAN_END, COL_ACTUAL_START, COL_ACTUAL_END, COL_MANUAL_STATUS, COL_WEEKLY_REPORT, COL_DEV_PROGRESS)
        colLastRow = ws.Cells(ws.Rows.Count, CStr(colAddr)).End(xlUp).Row
        If colLastRow > lastRow Then lastRow = colLastRow
    Next colAddr

    GetLastDataRow = lastRow
End Function

Public Function GetMinMaxDate(ws As Worksheet, lastRow As Long, ByRef minDate As Date, ByRef maxDate As Date) As Boolean
    Dim r As Long
    Dim hasDate As Boolean
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    Dim v4 As Variant
    
    hasDate = False
    
    For r = DATA_START_ROW To lastRow
        v1 = ws.Cells(r, COL_PLAN_START).Value
        v2 = ws.Cells(r, COL_PLAN_END).Value
        v3 = ws.Cells(r, COL_ACTUAL_START).Value
        v4 = ws.Cells(r, COL_ACTUAL_END).Value
        
        UpdateMinMaxDate v1, hasDate, minDate, maxDate
        UpdateMinMaxDate v2, hasDate, minDate, maxDate
        UpdateMinMaxDate v3, hasDate, minDate, maxDate
        UpdateMinMaxDate v4, hasDate, minDate, maxDate
        If IsDate(v3) And Not IsDate(v4) Then
            If CDate(v3) <= Date Then UpdateMinMaxDate Date, hasDate, minDate, maxDate
        End If
    Next r
    
    If hasDate Then
        minDate = minDate - 2
        maxDate = maxDate + 2
    End If
    
    GetMinMaxDate = hasDate
End Function

Private Sub UpdateMinMaxDate(ByVal v As Variant, ByRef hasDate As Boolean, ByRef minDate As Date, ByRef maxDate As Date)
    Dim d As Date
    
    If Not IsDate(v) Then Exit Sub
    
    d = CDate(v)
    
    If Not hasDate Then
        minDate = d
        maxDate = d
        hasDate = True
    Else
        If d < minDate Then minDate = d
        If d > maxDate Then maxDate = d
    End If
End Sub

Public Function IsWorkingDay(ByVal targetDate As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Boolean
    Dim key As String
    
    key = NormalizeDateKey(targetDate)
    
    If workdayDict.Exists(key) Then
        IsWorkingDay = True
        Exit Function
    End If
    
    If holidayDict.Exists(key) Then
        IsWorkingDay = False
        Exit Function
    End If
    
    If Weekday(targetDate, vbMonday) >= 6 Then
        IsWorkingDay = False
    Else
        IsWorkingDay = True
    End If
End Function

Public Function CountWorkingDaysInclusive(ByVal startDate As Variant, ByVal endDate As Variant, ByVal holidayDict As Object, ByVal workdayDict As Object) As Variant
    Dim d As Date
    Dim cnt As Long
    Dim s As Date
    Dim e As Date
    
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        CountWorkingDaysInclusive = ""
        Exit Function
    End If
    
    s = CDate(startDate)
    e = CDate(endDate)
    
    If s > e Then
        CountWorkingDaysInclusive = ""
        Exit Function
    End If
    
    cnt = 0
    
    For d = s To e
        If IsWorkingDay(d, holidayDict, workdayDict) Then
            cnt = cnt + 1
        End If
    Next d
    
    CountWorkingDaysInclusive = cnt
End Function

Public Function HasTaskContent(ws As Worksheet, ByVal rowNum As Long) As Boolean
    HasTaskContent = Len(Trim$(CStr(ws.Cells(rowNum, COL_TASK).Value))) > 0
End Function

Public Function HasAnyTaskDate(ws As Worksheet, ByVal rowNum As Long) As Boolean
    HasAnyTaskDate = _
        IsDate(ws.Cells(rowNum, COL_PLAN_START).Value) Or _
        IsDate(ws.Cells(rowNum, COL_PLAN_END).Value) Or _
        IsDate(ws.Cells(rowNum, COL_ACTUAL_START).Value) Or _
        IsDate(ws.Cells(rowNum, COL_ACTUAL_END).Value)
End Function

Public Function HasAnyTaskInput(ws As Worksheet, ByVal rowNum As Long) As Boolean
    HasAnyTaskInput = _
        Len(Trim$(CStr(ws.Cells(rowNum, COL_LEVEL).Value))) > 0 Or _
        HasTaskContent(ws, rowNum) Or _
        Len(Trim$(CStr(ws.Cells(rowNum, COL_NOTE).Value))) > 0 Or _
        Len(Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_STATUS).Value))) > 0 Or _
        Len(Trim$(CStr(ws.Cells(rowNum, COL_WEEKLY_REPORT).Value))) > 0 Or _
        Len(Trim$(CStr(ws.Cells(rowNum, COL_DEV_PROGRESS).Value))) > 0 Or _
        HasAnyTaskDate(ws, rowNum)
End Function

Public Function GetTaskLevel(ws As Worksheet, ByVal rowNum As Long) As Long
    Dim v As Variant
    
    v = ws.Cells(rowNum, COL_LEVEL).Value
    
    If IsNumeric(v) Then
        GetTaskLevel = CLng(v)
    Else
        GetTaskLevel = 1
    End If
    
    If GetTaskLevel < 1 Then GetTaskLevel = 1
    If GetTaskLevel > 3 Then GetTaskLevel = 3
End Function

Public Function GetTaskLevelMarker(ws As Worksheet, ByVal rowNum As Long) As String
    Select Case GetTaskLevel(ws, rowNum)
        Case 1
            GetTaskLevelMarker = "■"
        Case 2
            GetTaskLevelMarker = "▣"
        Case 3
            GetTaskLevelMarker = "◆"
        Case Else
            GetTaskLevelMarker = "■"
    End Select
End Function

Public Function HasChildTask(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long) As Boolean
    Dim currentLevel As Long
    Dim r As Long
    
    currentLevel = GetTaskLevel(ws, rowNum)
    
    For r = rowNum + 1 To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) <= currentLevel Then Exit For
            HasChildTask = True
            Exit Function
        End If
    Next r
    
    HasChildTask = False
End Function

Public Function GetTaskSubtreeEndRow(ws As Worksheet, ByVal rowNum As Long, ByVal lastRow As Long) As Long
    Dim currentLevel As Long
    Dim r As Long
    
    currentLevel = GetTaskLevel(ws, rowNum)
    GetTaskSubtreeEndRow = lastRow
    
    For r = rowNum + 1 To lastRow
        If HasTaskContent(ws, r) Then
            If GetTaskLevel(ws, r) <= currentLevel Then
                GetTaskSubtreeEndRow = r - 1
                Exit Function
            End If
        End If
    Next r
End Function

Public Function HasTaskError(ws As Worksheet, ByVal rowNum As Long) As Boolean
    HasTaskError = (Len(GetTaskErrorReason(ws, rowNum)) > 0)
End Function

Public Function GetTaskErrorReason(ws As Worksheet, ByVal rowNum As Long) As String
    Dim planS As Variant
    Dim planE As Variant
    Dim actS As Variant
    Dim actE As Variant
    Dim hasPlanS As Boolean
    Dim hasPlanE As Boolean
    Dim hasActS As Boolean
    Dim hasActE As Boolean
    Dim taskText As String
    Dim rawLevel As Variant
    Dim taskLevel As Long
    Dim parentRow As Long
    Dim parentModule As String
    Dim childModule As String
    Dim maxTaskLength As Long
    
    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value
    actS = ws.Cells(rowNum, COL_ACTUAL_START).Value
    actE = ws.Cells(rowNum, COL_ACTUAL_END).Value
    taskText = CStr(ws.Cells(rowNum, COL_TASK).Value2)
    rawLevel = ws.Cells(rowNum, COL_LEVEL).Value2
    
    If Not HasAnyTaskInput(ws, rowNum) Then
        Exit Function
    End If

    If Len(Trim$(CStr(rawLevel))) > 0 Then
        If Not IsNumeric(rawLevel) Then
            GetTaskErrorReason = "Level 값이 숫자가 아닙니다. 1~3 사이 정수를 입력하세요."
            Exit Function
        End If
        If CDbl(rawLevel) <> Fix(CDbl(rawLevel)) Or _
           CLng(rawLevel) < 1 Or CLng(rawLevel) > 3 Then
            GetTaskErrorReason = "Level은 1~3 사이 정수여야 합니다."
            Exit Function
        End If
    End If

    taskLevel = GetTaskLevel(ws, rowNum)
    maxTaskLength = GetTaskMaxLength(taskLevel)
    If Len(taskText) > maxTaskLength Then
        GetTaskErrorReason = "Level " & taskLevel & _
                             " 내용이 config 시트의 최대 글자 수(" & _
                             maxTaskLength & "자)를 초과했습니다. 현재 " & _
                             Len(taskText) & "자입니다."
        Exit Function
    End If
    
    If Not HasTaskContent(ws, rowNum) And HasAnyTaskDate(ws, rowNum) Then
        GetTaskErrorReason = "날짜가 입력되어 있지만 내용이 비어 있습니다."
        Exit Function
    End If
    
    hasPlanS = Len(Trim$(CStr(planS))) > 0
    hasPlanE = Len(Trim$(CStr(planE))) > 0
    hasActS = Len(Trim$(CStr(actS))) > 0
    hasActE = Len(Trim$(CStr(actE))) > 0
    
    If (hasPlanS And Not IsDate(planS)) Or (hasPlanE And Not IsDate(planE)) Or _
       (hasActS And Not IsDate(actS)) Or (hasActE And Not IsDate(actE)) Then
        GetTaskErrorReason = "날짜 열에 날짜 형식이 아닌 값이 있습니다."
        Exit Function
    End If
    
    If hasPlanS <> hasPlanE Then
        GetTaskErrorReason = "계획 시작일과 계획 종료일은 둘 다 입력하거나 둘 다 비워야 합니다."
        Exit Function
    End If
    
    If hasActE And Not hasActS Then
        GetTaskErrorReason = "실제 종료일이 있지만 실제 시작일이 없습니다."
        Exit Function
    End If
    
    If IsDate(planS) And IsDate(planE) Then
        If CDate(planS) > CDate(planE) Then
            GetTaskErrorReason = "계획 시작일이 계획 종료일보다 늦습니다."
            Exit Function
        End If
    End If
    
    If IsDate(actS) And IsDate(actE) Then
        If CDate(actS) > CDate(actE) Then
            GetTaskErrorReason = "실제 시작일이 실제 종료일보다 늦습니다."
            Exit Function
        End If
    End If

    If taskLevel > 1 Then
        parentRow = GetNearestParentTaskRow(ws, rowNum)
        If parentRow = 0 Then
            GetTaskErrorReason = "Level " & taskLevel & _
                                 "이지만 위쪽에 부모 작업이 없습니다."
            Exit Function
        End If

        parentModule = Trim$(CStr(ws.Cells(parentRow, COL_MODULE).Value2))
        childModule = Trim$(CStr(ws.Cells(rowNum, COL_MODULE).Value2))
        If Len(parentModule) = 0 Then
            GetTaskErrorReason = "부모 작업(행 " & parentRow & ")의 모듈이 비어 있습니다."
            Exit Function
        End If
        If StrComp(parentModule, childModule, vbTextCompare) <> 0 Then
            GetTaskErrorReason = "하위 모듈 '" & childModule & _
                                 "'이 부모 모듈 '" & parentModule & "'과 다릅니다."
            Exit Function
        End If
    End If
End Function

Public Function GetTaskProgressValue(ws As Worksheet, ByVal rowNum As Long) As Double
    Dim v As Variant
    
    v = ws.Cells(rowNum, COL_PROGRESS).Value
    
    If IsNumeric(v) Then
        GetTaskProgressValue = CDbl(v)
    Else
        GetTaskProgressValue = 0
    End If
    
    If GetTaskProgressValue < 0 Then GetTaskProgressValue = 0
    If GetTaskProgressValue > 1 Then GetTaskProgressValue = 1
End Function

Public Function GetExpectedProgress(ByVal planStart As Date, ByVal planEnd As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Double
    Dim totalDays As Variant
    Dim elapsedDays As Variant
    Dim compareEndDate As Date
    
    totalDays = CountWorkingDaysInclusive(planStart, planEnd, holidayDict, workdayDict)
    
    If totalDays = "" Then
        GetExpectedProgress = 0
        Exit Function
    End If
    
    If CLng(totalDays) <= 0 Then
        GetExpectedProgress = 0
        Exit Function
    End If
    
    If Date < planStart Then
        GetExpectedProgress = 0
        Exit Function
    End If
    
    If Date >= planEnd Then
        GetExpectedProgress = 1
        Exit Function
    End If
    
    compareEndDate = Date
    elapsedDays = CountWorkingDaysInclusive(planStart, compareEndDate, holidayDict, workdayDict)
    
    If elapsedDays = "" Then
        GetExpectedProgress = 0
    Else
        GetExpectedProgress = CDbl(elapsedDays) / CDbl(totalDays)
        If GetExpectedProgress < 0 Then GetExpectedProgress = 0
        If GetExpectedProgress > 1 Then GetExpectedProgress = 1
    End If
End Function

Public Function GetRequiredNormalProgress(ByVal planStart As Date, ByVal planEnd As Date, ByVal holidayDict As Object, ByVal workdayDict As Object) As Double
    Dim expectedProgress As Double
    Dim requiredProgress As Double
    
    expectedProgress = GetExpectedProgress(planStart, planEnd, holidayDict, workdayDict)
    requiredProgress = expectedProgress - 0.1
    
    If requiredProgress < 0 Then requiredProgress = 0
    If requiredProgress > 1 Then requiredProgress = 1
    
    GetRequiredNormalProgress = requiredProgress
End Function

Public Function IsTaskManualHold(ws As Worksheet, ByVal rowNum As Long) As Boolean
    IsTaskManualHold = (Trim$(CStr(ws.Cells(rowNum, COL_MANUAL_STATUS).Value)) = STATUS_HOLD)
End Function

Public Function GetTaskStatus(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object) As String
    Dim planS As Variant
    Dim planE As Variant
    Dim actE As Variant
    Dim actualProgress As Double
    Dim expectedProgress As Double
    Dim progressGap As Double
    
    If Not HasAnyTaskInput(ws, rowNum) Then
        GetTaskStatus = ""
        Exit Function
    End If
    
    If HasTaskError(ws, rowNum) Then
        GetTaskStatus = STATUS_ERROR
        Exit Function
    End If
    
    If IsTaskManualHold(ws, rowNum) Then
        GetTaskStatus = STATUS_HOLD
        Exit Function
    End If
    
    actE = ws.Cells(rowNum, COL_ACTUAL_END).Value
    actualProgress = GetTaskProgressValue(ws, rowNum)
    
    If actualProgress >= 1 Or IsDate(actE) Then
        GetTaskStatus = STATUS_DONE
        Exit Function
    End If
    
    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    planE = ws.Cells(rowNum, COL_PLAN_END).Value
    
    If Not IsDate(planS) Or Not IsDate(planE) Then
        GetTaskStatus = ""
        Exit Function
    End If
    
    expectedProgress = GetExpectedProgress(CDate(planS), CDate(planE), holidayDict, workdayDict)
    progressGap = expectedProgress - actualProgress
    
    If progressGap <= 0.1 Then
        GetTaskStatus = STATUS_NORMAL
    ElseIf progressGap <= 0.3 Then
        GetTaskStatus = STATUS_CAUTION
    Else
        GetTaskStatus = STATUS_DELAY
    End If
End Function

Public Function FindLevelMarkerDate(ws As Worksheet, ByVal rowNum As Long, ByVal holidayDict As Object, ByVal workdayDict As Object) As Variant
    Dim actS As Variant
    Dim planS As Variant
    Dim d As Date
    
    actS = ws.Cells(rowNum, COL_ACTUAL_START).Value
    If IsDate(actS) Then
        d = CDate(actS)
        If IsWorkingDay(d, holidayDict, workdayDict) Then
            FindLevelMarkerDate = d
            Exit Function
        End If
    End If
    
    planS = ws.Cells(rowNum, COL_PLAN_START).Value
    If IsDate(planS) Then
        d = CDate(planS)
        If IsWorkingDay(d, holidayDict, workdayDict) Then
            FindLevelMarkerDate = d
            Exit Function
        End If
    End If
    
    FindLevelMarkerDate = Empty
End Function

