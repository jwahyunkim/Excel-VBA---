Attribute VB_Name = "modHoliday"
Option Explicit

Public Sub EnsureConfigSheet()
    Dim ws As Worksheet
    Dim lastSheetRow As Long
    Dim rngType As Range
    Dim rngHideLevel As Range
    Dim rngExcludeNo As Range
    Dim rngExcludeDate As Range
    Dim rngDisplayStart As Range
    Dim rngDisplayEnd As Range
    Dim rngDisplayGanttOnly As Range
    Dim rngDisplayReportOnly As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(LEGACY_CONFIG_SHEET_NAME)
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Name = CONFIG_SHEET_NAME
        End If
    End If

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
    End If

    ws.Range(HOLIDAY_COL_DATE & HOLIDAY_HEADER_ROW).Value = "날짜"
    ws.Range(HOLIDAY_COL_TYPE & HOLIDAY_HEADER_ROW).Value = "구분"
    ws.Range(HOLIDAY_COL_DESC & HOLIDAY_HEADER_ROW).Value = "설명"

    ws.Range("E1").Value = "입력 예시"
    ws.Range("E2").Value = "A열: 날짜"
    ws.Range("E3").Value = "B열: 휴일 또는 근무일"
    ws.Range("E4").Value = "C열: 설명(선택)"

    ws.Range(HIDE_SETTING_TITLE_CELL).Value = "숨김 설정"
    ws.Range(HIDE_SETTING_LEVEL_LABEL_CELL).Value = "완료 숨김 레벨"
    ws.Range(DISPLAY_SETTING_TITLE_CELL).Value = "표시 기간 설정"
    ws.Range(DISPLAY_SETTING_START_LABEL_CELL).Value = "표시 시작일"
    ws.Range(DISPLAY_SETTING_END_LABEL_CELL).Value = "표시 종료일"
    ws.Range(DISPLAY_SETTING_GANTT_ONLY_LABEL_CELL).Value = "간트 only"
    ws.Range(DISPLAY_SETTING_REPORT_ONLY_LABEL_CELL).Value = "보고 only"
    ws.Range(HIDE_EXCLUDE_NO_HEADER_CELL).Value = "숨김 제외 No."
    ws.Range(HIDE_EXCLUDE_DATE_HEADER_CELL).Value = "숨김 제외 날짜"

    If Trim$(CStr(ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value)) = "" Then
        ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value = 3
    End If


    ws.Columns(HOLIDAY_COL_DATE).ColumnWidth = 14
    ws.Columns(HOLIDAY_COL_TYPE).ColumnWidth = 12
    ws.Columns(HOLIDAY_COL_DESC).ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 26
    ws.Columns("J").ColumnWidth = 14

    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(242, 242, 242)
    ws.Range("A1:C1").Borders.LineStyle = xlContinuous

    ws.Range("F1:G1").Font.Bold = True
    ws.Range("F1:G1").Interior.Color = RGB(242, 242, 242)
    ws.Range("F1:G1").Borders.LineStyle = xlContinuous

    ws.Range("I1:J1").Font.Bold = True
    ws.Range("I1:J1").Interior.Color = RGB(242, 242, 242)
    ws.Range("I1:J1").Borders.LineStyle = xlContinuous

    ws.Range("F2:G2").Borders.LineStyle = xlContinuous
    ws.Range("I2:J5").Borders.LineStyle = xlContinuous
    ws.Range("F5:G5").Font.Bold = True
    ws.Range("F5:G5").Interior.Color = RGB(242, 242, 242)
    ws.Range("F5:G5").Borders.LineStyle = xlContinuous
    ws.Range("F6:G200").Borders.LineStyle = xlContinuous

    ws.Columns(HOLIDAY_COL_DATE).NumberFormat = "yyyy-mm-dd"
    ws.Columns("G").NumberFormat = "General"
    ws.Range(DISPLAY_SETTING_START_VALUE_CELL).NumberFormat = "yyyy-mm-dd"
    ws.Range(DISPLAY_SETTING_END_VALUE_CELL).NumberFormat = "yyyy-mm-dd"
    ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL).NumberFormat = "General"
    ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).NumberFormat = "General"

    lastSheetRow = ws.Rows.Count
    Set rngType = ws.Range(HOLIDAY_COL_TYPE & HOLIDAY_DATA_START_ROW & ":" & HOLIDAY_COL_TYPE & lastSheetRow)
    Set rngHideLevel = ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL)
    Set rngExcludeNo = ws.Range(HIDE_EXCLUDE_NO_START_CELL & ":F" & lastSheetRow)
    Set rngExcludeDate = ws.Range(HIDE_EXCLUDE_DATE_START_CELL & ":G" & lastSheetRow)
    Set rngDisplayStart = ws.Range(DISPLAY_SETTING_START_VALUE_CELL)
    Set rngDisplayEnd = ws.Range(DISPLAY_SETTING_END_VALUE_CELL)
    Set rngDisplayGanttOnly = ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL)
    Set rngDisplayReportOnly = ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL)

    On Error Resume Next
    rngType.Validation.Delete
    rngHideLevel.Validation.Delete
    rngExcludeNo.Validation.Delete
    rngExcludeDate.Validation.Delete
    rngDisplayStart.Validation.Delete
    rngDisplayEnd.Validation.Delete
    rngDisplayGanttOnly.Validation.Delete
    rngDisplayReportOnly.Validation.Delete
    On Error GoTo 0

    rngType.Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:=HOLIDAY_TYPE_HOLIDAY & "," & HOLIDAY_TYPE_WORKDAY

    rngType.Validation.IgnoreBlank = True
    rngType.Validation.InCellDropdown = True
    rngType.Validation.InputTitle = "구분 선택"
    rngType.Validation.InputMessage = "휴일 또는 근무일만 선택할 수 있습니다."
    rngType.Validation.ErrorTitle = "입력 오류"
    rngType.Validation.ErrorMessage = "휴일 또는 근무일만 입력할 수 있습니다."

    rngHideLevel.Validation.Add Type:=xlValidateWholeNumber, _
                                AlertStyle:=xlValidAlertStop, _
                                Operator:=xlBetween, _
                                Formula1:="1", _
                                Formula2:="3"

    rngExcludeNo.Validation.Add Type:=xlValidateWholeNumber, _
                                AlertStyle:=xlValidAlertStop, _
                                Operator:=xlBetween, _
                                Formula1:="1", _
                                Formula2:="100000"
    rngExcludeNo.Validation.IgnoreBlank = True

    rngExcludeDate.NumberFormat = "yyyy-mm-dd"
    rngExcludeDate.Validation.Add Type:=xlValidateDate, _
                                  AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlBetween, _
                                  Formula1:="2000-01-01", _
                                  Formula2:="2100-12-31"
    rngExcludeDate.Validation.IgnoreBlank = True

    rngDisplayStart.Validation.Add Type:=xlValidateDate, _
                                   AlertStyle:=xlValidAlertStop, _
                                   Operator:=xlBetween, _
                                   Formula1:="2000-01-01", _
                                   Formula2:="2100-12-31"
    rngDisplayStart.Validation.IgnoreBlank = True
    rngDisplayStart.Validation.InputTitle = "표시 시작일"
    rngDisplayStart.Validation.InputMessage = "간트에서 처음 보여줄 날짜를 입력하세요."
    rngDisplayStart.Validation.ErrorTitle = "입력 오류"
    rngDisplayStart.Validation.ErrorMessage = "올바른 날짜를 입력하세요."

    rngDisplayEnd.Validation.Add Type:=xlValidateDate, _
                                 AlertStyle:=xlValidAlertStop, _
                                 Operator:=xlBetween, _
                                 Formula1:="2000-01-01", _
                                 Formula2:="2100-12-31"
    rngDisplayEnd.Validation.IgnoreBlank = True
    rngDisplayEnd.Validation.InputTitle = "표시 종료일"
    rngDisplayEnd.Validation.InputMessage = "간트에서 마지막으로 보여줄 날짜를 입력하세요."
    rngDisplayEnd.Validation.ErrorTitle = "입력 오류"
    rngDisplayEnd.Validation.ErrorMessage = "올바른 날짜를 입력하세요."

    rngDisplayGanttOnly.Validation.Add Type:=xlValidateList, _
                                        AlertStyle:=xlValidAlertStop, _
                                        Operator:=xlBetween, _
                                        Formula1:="Y,N"
    rngDisplayGanttOnly.Validation.IgnoreBlank = True
    rngDisplayGanttOnly.Validation.InCellDropdown = True

    rngDisplayReportOnly.Validation.Add Type:=xlValidateList, _
                                         AlertStyle:=xlValidAlertStop, _
                                         Operator:=xlBetween, _
                                         Formula1:=STATUS_WEEKLY_REPORT & "," & STATUS_DEV_PROGRESS & "," & REPORT_FILTER_ALL & "," & REPORT_FILTER_EMPTY
    rngDisplayReportOnly.Validation.IgnoreBlank = True
    rngDisplayReportOnly.Validation.InCellDropdown = True
End Sub

Public Sub LoadHolidaySettings(ByRef holidayDict As Object, ByRef workdayDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim vDate As Variant
    Dim parsedDate As Date
    Dim vType As String
    Dim key As String

    Set holidayDict = CreateObject("Scripting.Dictionary")
    Set workdayDict = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, HOLIDAY_COL_DATE).End(xlUp).Row

    If lastRow < HOLIDAY_DATA_START_ROW Then Exit Sub

    For r = HOLIDAY_DATA_START_ROW To lastRow
        vDate = ws.Cells(r, HOLIDAY_COL_DATE).Value
        vType = Trim$(CStr(ws.Cells(r, HOLIDAY_COL_TYPE).Value))

        If TryParseHolidayDate(vDate, parsedDate) Then
            key = NormalizeDateKey(parsedDate)

            If vType = HOLIDAY_TYPE_HOLIDAY Then
                If Not holidayDict.Exists(key) Then holidayDict.Add key, True
            ElseIf vType = HOLIDAY_TYPE_WORKDAY Then
                If Not workdayDict.Exists(key) Then workdayDict.Add key, True
            End If
        End If
    Next r
End Sub


Public Function TryGetDisplayDateRange(ByRef displayStartDate As Date, ByRef displayEndDate As Date) As Boolean
    Dim ws As Worksheet
    Dim startValue As Variant
    Dim endValue As Variant

    TryGetDisplayDateRange = False

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)

    startValue = ws.Range(DISPLAY_SETTING_START_VALUE_CELL).Value
    endValue = ws.Range(DISPLAY_SETTING_END_VALUE_CELL).Value

    If Trim$(CStr(startValue)) = "" And Trim$(CStr(endValue)) = "" Then Exit Function

    If Trim$(CStr(startValue)) = "" Or Trim$(CStr(endValue)) = "" Then
        Err.Raise vbObjectError + 7101, "TryGetDisplayDateRange", "config 시트의 표시 시작일과 표시 종료일을 모두 입력해야 합니다."
    End If

    If Not IsDate(startValue) Or Not IsDate(endValue) Then
        Err.Raise vbObjectError + 7102, "TryGetDisplayDateRange", "config 시트의 표시 기간은 날짜만 입력할 수 있습니다."
    End If

    displayStartDate = CDate(startValue)
    displayEndDate = CDate(endValue)

    If CLng(displayStartDate) > CLng(displayEndDate) Then
        Err.Raise vbObjectError + 7103, "TryGetDisplayDateRange", "config 시트의 표시 시작일은 표시 종료일보다 늦을 수 없습니다."
    End If

    TryGetDisplayDateRange = True
End Function

Public Function GetDisplayGanttOnlyFlag() As Boolean
    Dim ws As Worksheet
    Dim v As String

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = UCase$(Trim$(CStr(ws.Range(DISPLAY_SETTING_GANTT_ONLY_VALUE_CELL).Value)))

    GetDisplayGanttOnlyFlag = (v = "Y")
End Function

Public Function GetDisplayReportOnlyFlag() As String
    Dim ws As Worksheet
    Dim v As String

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = Trim$(CStr(ws.Range(DISPLAY_SETTING_REPORT_ONLY_VALUE_CELL).Value))

    If v = STATUS_WEEKLY_REPORT Or v = STATUS_DEV_PROGRESS Or v = REPORT_FILTER_ALL Or v = REPORT_FILTER_EMPTY Then
        GetDisplayReportOnlyFlag = v
    Else
        GetDisplayReportOnlyFlag = ""
    End If
End Function

Public Function GetHideCompletedMaxLevel() As Long
    Dim ws As Worksheet
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    v = ws.Range(HIDE_SETTING_LEVEL_VALUE_CELL).Value

    If IsNumeric(v) Then
        GetHideCompletedMaxLevel = CLng(v)
    Else
        GetHideCompletedMaxLevel = 3
    End If

    If GetHideCompletedMaxLevel < 1 Then GetHideCompletedMaxLevel = 1
    If GetHideCompletedMaxLevel > 3 Then GetHideCompletedMaxLevel = 3
End Function

Public Sub LoadExcludedRowNos(ByRef excludeNoDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    Dim key As String

    Set excludeNoDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    If lastRow < 6 Then Exit Sub

    For r = 6 To lastRow
        v = ws.Cells(r, "F").Value
        If IsNumeric(v) Then
            key = Trim$(CStr(CLng(v)))
            If Len(key) > 0 Then
                If Not excludeNoDict.Exists(key) Then excludeNoDict.Add key, True
            End If
        End If
    Next r
End Sub

Public Sub LoadExcludedDates(ByRef excludeDateDict As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    Dim parsedDate As Date
    Dim key As String

    Set excludeDateDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    If lastRow < 6 Then Exit Sub

    For r = 6 To lastRow
        v = ws.Cells(r, "G").Value
        If TryParseHolidayDate(v, parsedDate) Then
            key = NormalizeDateKey(parsedDate)
            If Not excludeDateDict.Exists(key) Then excludeDateDict.Add key, True
        End If
    Next r
End Sub

Private Function TryParseHolidayDate(ByVal vDate As Variant, ByRef parsedDate As Date) As Boolean
    Dim s As String
    Dim yyyyPart As Long
    Dim mmPart As Long
    Dim ddPart As Long
    Dim tmpDate As Date

    TryParseHolidayDate = False

    If IsDate(vDate) Then
        parsedDate = CDate(vDate)
        TryParseHolidayDate = True
        Exit Function
    End If

    s = Trim$(CStr(vDate))
    s = Replace$(s, "-", "")
    s = Replace$(s, ".", "")
    s = Replace$(s, "/", "")

    If Len(s) = 8 And IsNumeric(s) Then
        yyyyPart = CLng(Left$(s, 4))
        mmPart = CLng(Mid$(s, 5, 2))
        ddPart = CLng(Right$(s, 2))

        On Error GoTo InvalidDate
        tmpDate = DateSerial(yyyyPart, mmPart, ddPart)
        On Error GoTo 0

        If Year(tmpDate) <> yyyyPart Then Exit Function
        If Month(tmpDate) <> mmPart Then Exit Function
        If Day(tmpDate) <> ddPart Then Exit Function

        parsedDate = tmpDate
        TryParseHolidayDate = True
        Exit Function
    End If

    If IsDate(s) Then
        parsedDate = CDate(s)
        TryParseHolidayDate = True
    End If

    Exit Function

InvalidDate:
    On Error GoTo 0
End Function

Public Function NormalizeDateKey(ByVal targetDate As Date) As String
    NormalizeDateKey = Format$(CLng(targetDate), "0")
End Function
