' Append to modTextReport only in a disposable workbook copy.
Public Function RunTextReportRegression(ByVal outputDirectory As String) As String
    Dim configWs As Worksheet, taskWs As Worksheet
    Dim files As Collection, rows As Collection, items As Collection, dates As Collection
    Dim levels As Collection, plans As Collection, groups As Collection, assignment As Object
    Dim reportStart As Date, reportEnd As Date, text As String, secondText As String
    Dim checked As Long, stage As String, snapshot As String, path As String, i As Long
    Dim expected As Variant, numbers As Variant, firstNumbers As Variant, secondNumbers As Variant
    Dim modulePath As String, errorsExpected As Boolean
    On Error GoTo Failed

    stage = "configure output fixture"
    EnsureConfigSheet
    Set configWs = ThisWorkbook.Worksheets(WEEKLY_REPORT_CONFIG_SHEET_NAME)
    reportStart = Date - 20
    reportEnd = Date - 10
    configWs.Range("B17").Value = reportStart
    configWs.Range("B18").Value = reportEnd
    configWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).ClearContents
    TextRegressionAssert GetOutputShowModificationFlag(), "blank modification flag defaults to shown", checked
    configWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "Y"
    configWs.Range("B2:G5").Value2 = "N"
    configWs.Range("C2:F2").Value2 = "Y"
    configWs.Range("F3").Value2 = "Y"
    configWs.Range("F5").Value2 = "Y"
    configWs.Range("G5").Value2 = "전체"
    configWs.Range("B7").Value2 = WEEKLY_REPORT_PAGE_MODE_ALL
    configWs.Range("B8").Value2 = WEEKLY_REPORT_CATEGORY_DEPTH_MAJOR
    configWs.Range("G8:G14").ClearContents
    configWs.Range("E8:E14").Value = Application.Transpose(Array("ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ"))
    configWs.Range("F8:F14").Value2 = "기호"
    configWs.Range("F9").Value2 = "글머리 없음"
    configWs.Range("I3:J197").ClearContents
    Set taskWs = ThisWorkbook.Worksheets.Add
    taskWs.Name = "TextReportFixture"
    SetupDataHeaders taskWs
    BuildTextRegressionFixture taskWs, reportStart, reportEnd

    stage = "inclusive period and request modification export"
    path = outputDirectory & Application.PathSeparator & "whole"
    EnsureReportOutputDirectory path
    Application.EnableEvents = True
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionAssert Application.EnableEvents, "success must restore enabled events", checked
    Application.EnableEvents = False
    TextRegressionEqual CStr(files.Count), "1", "whole grouping file count", checked
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionContains text, "업무 현황" & vbCrLf, "request header", checked
    TextRegressionAbsent text, "Job Plan", "obsolete report title removed", checked
    TextRegressionAbsent text, "Digital MFG", "obsolete team title removed", checked
    TextRegressionAbsent text, "( 주요 계획 )", "planned heading removed", checked
    TextRegressionAbsent text, "(개발 항목 - 요청 내용)", "planned request heading removed", checked
    TextRegressionContains text, "기간 : " & Format$(reportStart, "yyyy.mm.dd") & _
        "~" & Format$(reportEnd, "yyyy.mm.dd"), "configured period title", checked
    TextRegressionContains text, "기간시작완료", "completed inclusive start", checked
    TextRegressionContains text, "기간종료완료", "completed inclusive end", checked
    TextRegressionContains text, "진행중기한초과", "overdue progress remains", checked
    TextRegressionContains text, "계획시작경계", "planned end meets start", checked
    TextRegressionContains text, "계획종료경계", "planned start meets end", checked
    TextRegressionContains text, "계획날짜미정", "undated plan", checked
    TextRegressionAbsent text, "완료기간이전", "completed before period", checked
    TextRegressionAbsent text, "완료기간이후", "completed after period", checked
    TextRegressionAbsent text, "진행시작이후", "progress after period", checked
    TextRegressionAbsent text, "계획기간이전", "plan before period", checked
    TextRegressionAbsent text, "계획기간이후", "plan after period", checked
    TextRegressionModification text, "상위요청", "상위 한글 수정", "current ancestor", checked
    TextRegressionModification text, "하위요청", "하위 한글 수정", "current middle ancestor", checked
    TextRegressionModification text, "기간시작완료", "경계 검증 수정", "current leaf", checked
    TextRegressionModification text, "계획상위요청", "계획 상위 수정", "planned ancestor", checked
    TextRegressionModification text, "계획하위요청", "계획 하위 수정", "planned middle ancestor", checked
    TextRegressionModification text, "계획세부요청", "계획 세부 수정", "planned leaf", checked
    TextRegressionContains text, "└ 중복 업무 첫 수정", "duplicate name first modification", checked
    TextRegressionContains text, "└ 중복 업무 둘째 수정", "duplicate name second modification", checked
    TextRegressionContains text, "(요청 내용 없음)", "modification only request anchor", checked
    TextRegressionContains text, "└ 수정만 있는 기간 내 업무", "modification only completed inside period", checked
    TextRegressionAbsent text, "수정만 있는 기간 외 업무", "modification only completed outside period", checked
    TextRegressionContains text, "~" & CStr(Month(reportStart - 1)) & "/" & CStr(Day(reportStart - 1)), "in progress date text", checked
    TextRegressionContains text, CStr(Month(reportStart)) & "/" & CStr(Day(reportStart)), "completed date text", checked
    TextRegressionAbsent text, ChrW(&H200B), "no invisible date characters", checked
    TextRegressionAbsent text, ChrW(11), "plain text line endings", checked
    TextRegressionAbsent text, "(유지보수 항목)", "empty maintenance section removed", checked
    TextRegressionAbsent text, "(기타 항목)", "empty other section removed", checked
    TextRegressionAbsent text, "(이슈 사항)", "empty issue section removed", checked
    TextRegressionAbsent text, "수정 내용:", "obsolete modification label removed", checked
    TextRegressionContains text, vbCrLf & "분류A" & vbCrLf & Space$(4) & "ㄷ 중분류A", "major and middle category slots", checked
    TextRegressionContains text, Space$(12) & "ㅁ 상위요청", "level 1 hierarchy indent", checked
    TextRegressionContains text, Space$(16) & "ㅂ 하위요청", "level 2 hierarchy indent", checked
    TextRegressionContains text, Space$(20) & "ㅅ 기간시작완료 (한글 담당자)", "level 3 owner", checked
    TextRegressionContains text, vbCrLf & Space$(22) & "└ 경계 검증 수정" & vbCrLf & _
        Space$(22) & "└ 두 번째 수정" & vbCrLf & vbCrLf & Space$(22) & "└ 세 번째 수정", _
        "current multiline modifications preserve aligned lines", checked
    TextRegressionContains text, vbCrLf & Space$(22) & "└ 계획 세부 수정" & vbCrLf & _
        Space$(22) & "└ 계획 두 번째 수정", "planned multiline modifications preserve aligned lines", checked

    stage = "modification visibility toggle"
    configWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "N"
    Set files = ExportPeriodTextReport(taskWs, path)
    secondText = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionAbsent secondText, "└", "disabled modification flag hides every child marker", checked
    TextRegressionAbsent secondText, "상위 한글 수정", "disabled flag hides current ancestor modification", checked
    TextRegressionAbsent secondText, "경계 검증 수정", "disabled flag hides current leaf modification", checked
    TextRegressionAbsent secondText, "계획 상위 수정", "disabled flag hides planned ancestor modification", checked
    TextRegressionAbsent secondText, "계획 세부 수정", "disabled flag hides planned leaf modification", checked
    TextRegressionContains secondText, "상위요청", "disabled flag retains current requests", checked
    TextRegressionContains secondText, "계획세부요청", "disabled flag retains planned requests", checked
    configWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = "Y"

    stage = "custom leading spaces and hidden categories"
    configWs.Range("D2").Value2 = "N"
    configWs.Range("G12:G14").Value = Application.Transpose(Array(2, 7, 10))
    configWs.Range("G10").Value2 = 0
    snapshot = TextRegressionSnapshot(configWs)
    EnsureWeeklyReportConfigSheet
    EnsureWeeklyReportConfigSheet
    TextRegressionEqual TextRegressionSnapshot(configWs), snapshot, "settings preserved by initialization", checked
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionAssert Not Application.EnableEvents, "success must preserve disabled events", checked
    secondText = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionContains secondText, vbCrLf & Space$(2) & "ㅁ 상위요청", "custom first indent", checked
    TextRegressionContains secondText, vbCrLf & Space$(7) & "ㅂ 하위요청", "custom second indent", checked
    TextRegressionContains secondText, vbCrLf & Space$(10) & "ㅅ 기간시작완료", "custom third indent", checked
    TextRegressionContains secondText, vbCrLf & Space$(12) & "└ 경계 검증 수정", "custom modification indent", checked
    TextRegressionModification secondText, "상위요청", "상위 한글 수정", "custom current indent", checked
    TextRegressionModification secondText, "계획상위요청", "계획 상위 수정", "custom planned indent", checked
    TextRegressionAbsent secondText, "중분류A", "hidden category", checked
    TextRegressionAssert CStr(files(1)) <> path & Application.PathSeparator & _
        "기간보고_" & Format$(reportStart, "yyyymmdd") & "_" & Format$(reportEnd, "yyyymmdd") & "_001_전체.txt", _
        "second export must avoid overwrite", checked
    configWs.Range("D2").Value2 = "Y"
    configWs.Range("G8:G14").ClearContents

    stage = "all automatic bullet number formats"
    numbers = Array("번호 매기기", "원 숫자 ①", "괄호 숫자 (1)", "반괄호 숫자 1)", "원 알파벳 ⓐ", "원 한글 자음 ㉠")
    firstNumbers = Array("1.", "①", "(1)", "1)", "ⓐ", "㉠")
    secondNumbers = Array("2.", "②", "(2)", "2)", "ⓑ", "㉡")
    For i = 0 To UBound(numbers)
        configWs.Range("F12").Value2 = CStr(numbers(i))
        ResetWeeklyReportNumbering
        TextRegressionEqual GetWeeklyReportLevelBullet(1), CStr(firstNumbers(i)), "number first " & CStr(i), checked
        TextRegressionEqual GetWeeklyReportLevelBullet(1), CStr(secondNumbers(i)), "number second " & CStr(i), checked
    Next i
    configWs.Range("F12").Value2 = "번호 매기기"
    Set files = ExportPeriodTextReport(taskWs, path)
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionContains text, Space$(12) & "1. 상위요청", "export first numbering", checked
    TextRegressionContains text, Space$(12) & "2. 기간종료완료", "modification does not consume numbering", checked
    TextRegressionContains text, Space$(12) & "1. 계획시작경계", "plan numbering restarts", checked
    TextRegressionModification text, "상위요청", "상위 한글 수정", "numbered current request", checked
    TextRegressionModification text, "계획상위요청", "계획 상위 수정", "numbered planned request", checked
    configWs.Range("F12").Value2 = "기호"

    stage = "unbulleted requests and optional level labels"
    configWs.Range("F12:F14").Value2 = "글머리 없음"
    Set files = ExportPeriodTextReport(taskWs, path)
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionModification text, "상위요청", "상위 한글 수정", "unbulleted current request", checked
    TextRegressionModification text, "계획세부요청", "계획 세부 수정", "unbulleted planned request", checked
    configWs.Range("G2").Value2 = "Y"
    Set files = ExportPeriodTextReport(taskWs, path)
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionContains text, "Level 1 - 상위요청", "current level label shown", checked
    TextRegressionContains text, "Level 3 - 계획세부요청", "planned level label shown", checked
    TextRegressionModification text, "상위요청", "상위 한글 수정", "current marker follows level label", checked
    TextRegressionModification text, "계획세부요청", "계획 세부 수정", "planned marker follows level label", checked
    configWs.Range("G2").Value2 = "N"
    configWs.Range("F12:F14").Value2 = "기호"

    stage = "classification files and custom grouping"
    configWs.Range("B7").Value2 = WEEKLY_REPORT_PAGE_MODE_MODULE
    path = outputDirectory & Application.PathSeparator & "classification"
    EnsureReportOutputDirectory path
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionEqual CStr(files.Count), "2", "major classification file count", checked
    text = ReadTextRegressionUtf8(CStr(files(1))) & ReadTextRegressionUtf8(CStr(files(2)))
    TextRegressionContains text, "분류B업무", "second classification retained", checked
    configWs.Range("B7").Value2 = WEEKLY_REPORT_PAGE_MODE_CUSTOM
    configWs.Range("I3").Value2 = 7
    configWs.Range("J3").Value2 = " > 분류A > 중분류A"
    configWs.Range("I4").Value2 = 7
    configWs.Range("J4").Value2 = " > 분류B > 중분류B"
    path = outputDirectory & Application.PathSeparator & "custom"
    EnsureReportOutputDirectory path
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionEqual CStr(files.Count), "1", "custom shared file grouping", checked
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionContains text, "기간시작완료", "first custom category", checked
    TextRegressionContains text, "분류B업무", "second custom category", checked
    configWs.Range("I4:J4").ClearContents
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionEqual CStr(files.Count), "2", "unassigned categories get own file", checked

    stage = "long file remains a single complete output"
    configWs.Range("B7").Value2 = WEEKLY_REPORT_PAGE_MODE_ALL
    For i = 1 To 45
        TextRegressionAddLongTask taskWs, 30 + i, i, reportStart
    Next i
    For i = 1 To 12
        TextRegressionAddLongPlan taskWs, 100 + i, i, reportStart
    Next i
    path = outputDirectory & Application.PathSeparator & "long"
    EnsureReportOutputDirectory path
    Set files = ExportPeriodTextReport(taskWs, path)
    TextRegressionEqual CStr(files.Count), "1", "long text must not split", checked
    text = ReadTextRegressionUtf8(CStr(files(1)))
    For i = 1 To 45
        TextRegressionContains text, "장문업무" & Format$(i, "000"), "long file task " & CStr(i), checked
    Next i

    stage = "multi digit automatic numbering alignment"
    configWs.Range("F12").Value2 = "번호 매기기"
    Set files = ExportPeriodTextReport(taskWs, path)
    text = ReadTextRegressionUtf8(CStr(files(1)))
    TextRegressionAssert InStr(text, "10. 장문업무") > 0, "current numbers reach two digits", checked
    TextRegressionAssert InStr(text, "10. 계획번호요청") > 0, "planned numbers reach two digits", checked
    For i = 1 To 45
        TextRegressionModification text, "장문업무" & Format$(i, "000"), _
            "장문 수정 " & Format$(i, "000"), "long numbered current " & CStr(i), checked
    Next i
    For i = 1 To 12
        TextRegressionModification text, "계획번호요청" & Format$(i, "000"), _
            "계획 번호 수정 " & Format$(i, "000"), "long numbered plan " & CStr(i), checked
    Next i

    stage = "invalid configured period fails without writing"
    configWs.Range("B17").Value = reportEnd
    configWs.Range("B18").Value = reportStart
    On Error Resume Next
    Err.Clear
    Application.EnableEvents = True
    Set files = ExportPeriodTextReport(taskWs, path)
    errorsExpected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    TextRegressionAssert Application.EnableEvents, "reversed period must restore enabled events", checked
    Application.EnableEvents = False
    TextRegressionAssert errorsExpected, "reversed period accepted", checked
    configWs.Range("B17").Value2 = "invalid date"
    configWs.Range("B18").Value = reportEnd
    On Error Resume Next
    Err.Clear
    Application.EnableEvents = True
    Set files = ExportPeriodTextReport(taskWs, path)
    errorsExpected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    TextRegressionAssert Application.EnableEvents, "invalid date must restore enabled events", checked
    Application.EnableEvents = False
    TextRegressionAssert errorsExpected, "invalid date accepted", checked
    On Error Resume Next
    Err.Clear
    Set files = ExportPeriodTextReport(taskWs, path)
    errorsExpected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    TextRegressionAssert errorsExpected And Not Application.EnableEvents, "failure must preserve disabled events", checked

    RunTextReportRegression = "PASS: " & CStr(checked) & " assertions; " & outputDirectory
    Exit Function
Failed:
    RunTextReportRegression = "FAIL: " & stage & " (" & CStr(Err.Number) & "): " & Err.Description
    Application.EnableEvents = False
End Function

Private Sub BuildTextRegressionFixture(ByVal ws As Worksheet, ByVal firstDate As Date, ByVal lastDate As Date)
    Dim names As Variant, i As Long, r As Long
    names = Array("상위요청", "하위요청", "기간시작완료", "기간종료완료", _
        "완료기간이전", "완료기간이후", "진행중기한초과", "진행시작이후", _
        "계획시작경계", "계획종료경계", "계획기간이전", "계획기간이후", "계획날짜미정", _
        "분류B업무", "중복요청", "중복요청")
    For i = 0 To UBound(names)
        r = DATA_START_ROW + i
        ws.Cells(r, COL_LEVEL).Value2 = 1
        ws.Cells(r, COL_TYPE).Value2 = "개발"
        ws.Cells(r, COL_MAJOR_CATEGORY).Value2 = "분류A"
        ws.Cells(r, COL_MIDDLE_CATEGORY).Value2 = "중분류A"
        ws.Cells(r, COL_MINOR_CATEGORY).Value2 = "소분류A"
        ws.Cells(r, COL_TASK).Value2 = CStr(names(i))
        ws.Cells(r, COL_OWNER).Value2 = "한글 담당자"
    Next i
    ws.Cells(DATA_START_ROW + 1, COL_LEVEL).Value2 = 2
    ws.Cells(DATA_START_ROW + 2, COL_LEVEL).Value2 = 3
    ws.Cells(DATA_START_ROW, COL_MODIFICATION).Value2 = "상위 한글 수정"
    ws.Cells(DATA_START_ROW + 1, COL_MODIFICATION).Value2 = "하위 한글 수정"
    ws.Cells(DATA_START_ROW + 2, COL_MODIFICATION).Value2 = "경계 검증 수정" & vbCrLf & _
        "두 번째 수정" & vbLf & vbLf & "세 번째 수정"
    For i = 2 To 7
        ws.Cells(DATA_START_ROW + i, COL_ACTUAL_START).Value = firstDate - 5
    Next i
    ws.Cells(DATA_START_ROW + 2, COL_ACTUAL_END).Value = firstDate + TimeSerial(23, 59, 0)
    ws.Cells(DATA_START_ROW + 3, COL_ACTUAL_END).Value = lastDate + TimeSerial(23, 59, 0)
    ws.Cells(DATA_START_ROW + 4, COL_ACTUAL_END).Value = firstDate - 1
    ws.Cells(DATA_START_ROW + 5, COL_ACTUAL_END).Value = lastDate + 1
    ws.Cells(DATA_START_ROW + 6, COL_PLAN_END).Value = firstDate - 1
    ws.Cells(DATA_START_ROW + 7, COL_ACTUAL_START).Value = lastDate + 1
    ws.Cells(DATA_START_ROW + 8, COL_PLAN_START).Value = firstDate - 5
    ws.Cells(DATA_START_ROW + 8, COL_PLAN_END).Value = firstDate
    ws.Cells(DATA_START_ROW + 9, COL_PLAN_START).Value = lastDate
    ws.Cells(DATA_START_ROW + 9, COL_PLAN_END).Value = lastDate + 5
    ws.Cells(DATA_START_ROW + 10, COL_PLAN_START).Value = firstDate - 5
    ws.Cells(DATA_START_ROW + 10, COL_PLAN_END).Value = firstDate - 1
    ws.Cells(DATA_START_ROW + 11, COL_PLAN_START).Value = lastDate + 1
    ws.Cells(DATA_START_ROW + 11, COL_PLAN_END).Value = lastDate + 5
    ws.Cells(DATA_START_ROW + 13, COL_MAJOR_CATEGORY).Value2 = "분류B"
    ws.Cells(DATA_START_ROW + 13, COL_MIDDLE_CATEGORY).Value2 = "중분류B"
    ws.Cells(DATA_START_ROW + 13, COL_MINOR_CATEGORY).Value2 = "소분류B"
    For i = 14 To 15
        ws.Cells(DATA_START_ROW + i, COL_ACTUAL_START).Value = firstDate
        ws.Cells(DATA_START_ROW + i, COL_MODIFICATION).Value2 = _
            IIf(i = 14, "중복 업무 첫 수정", "중복 업무 둘째 수정")
    Next i
    For i = 16 To 17
        r = DATA_START_ROW + i
        ws.Cells(r, COL_LEVEL).Value2 = 1
        ws.Cells(r, COL_TYPE).Value2 = "개발"
        ws.Cells(r, COL_MAJOR_CATEGORY).Value2 = "분류A"
        ws.Cells(r, COL_MIDDLE_CATEGORY).Value2 = "중분류A"
        ws.Cells(r, COL_MINOR_CATEGORY).Value2 = "소분류A"
        ws.Cells(r, COL_MODIFICATION).Value2 = _
            IIf(i = 16, "수정만 있는 기간 내 업무", "수정만 있는 기간 외 업무")
        ws.Cells(r, COL_ACTUAL_START).Value = firstDate - 5
        ws.Cells(r, COL_ACTUAL_END).Value = IIf(i = 16, firstDate, firstDate - 1)
    Next i
    names = Array("계획상위요청", "계획하위요청", "계획세부요청")
    For i = 0 To 2
        r = DATA_START_ROW + 18 + i
        ws.Cells(r, COL_LEVEL).Value2 = i + 1
        ws.Cells(r, COL_TYPE).Value2 = "개발"
        ws.Cells(r, COL_MAJOR_CATEGORY).Value2 = "분류A"
        ws.Cells(r, COL_MIDDLE_CATEGORY).Value2 = "중분류A"
        ws.Cells(r, COL_MINOR_CATEGORY).Value2 = "소분류A"
        ws.Cells(r, COL_TASK).Value2 = CStr(names(i))
        ws.Cells(r, COL_PLAN_START).Value = firstDate
        ws.Cells(r, COL_PLAN_END).Value = lastDate
    Next i
    ws.Cells(DATA_START_ROW + 18, COL_MODIFICATION).Value2 = "계획 상위 수정"
    ws.Cells(DATA_START_ROW + 19, COL_MODIFICATION).Value2 = "계획 하위 수정"
    ws.Cells(DATA_START_ROW + 20, COL_MODIFICATION).Value2 = "계획 세부 수정" & vbCrLf & "계획 두 번째 수정"
End Sub

Private Sub TextRegressionAddLongTask(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                                     ByVal taskNumber As Long, ByVal firstDate As Date)
    ws.Cells(rowNumber, COL_LEVEL).Value2 = 1
    ws.Cells(rowNumber, COL_TYPE).Value2 = "개발"
    ws.Cells(rowNumber, COL_MAJOR_CATEGORY).Value2 = "분류A"
    ws.Cells(rowNumber, COL_MIDDLE_CATEGORY).Value2 = "중분류A"
    ws.Cells(rowNumber, COL_MINOR_CATEGORY).Value2 = "소분류A"
    ws.Cells(rowNumber, COL_TASK).Value2 = "장문업무" & Format$(taskNumber, "000") & _
        " 긴 요청 내용을 파일 개수 제한 없이 보존하는 검증"
    ws.Cells(rowNumber, COL_ACTUAL_START).Value = firstDate
    ws.Cells(rowNumber, COL_MODIFICATION).Value2 = "장문 수정 " & Format$(taskNumber, "000")
End Sub

Private Sub TextRegressionAddLongPlan(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                                     ByVal taskNumber As Long, ByVal firstDate As Date)
    ws.Cells(rowNumber, COL_LEVEL).Value2 = 1
    ws.Cells(rowNumber, COL_TYPE).Value2 = "개발"
    ws.Cells(rowNumber, COL_MAJOR_CATEGORY).Value2 = "분류A"
    ws.Cells(rowNumber, COL_MIDDLE_CATEGORY).Value2 = "중분류A"
    ws.Cells(rowNumber, COL_MINOR_CATEGORY).Value2 = "소분류A"
    ws.Cells(rowNumber, COL_TASK).Value2 = "계획번호요청" & Format$(taskNumber, "000")
    ws.Cells(rowNumber, COL_MODIFICATION).Value2 = "계획 번호 수정 " & Format$(taskNumber, "000")
    ws.Cells(rowNumber, COL_PLAN_START).Value = firstDate
    ws.Cells(rowNumber, COL_PLAN_END).Value = firstDate + 1
End Sub

' Verify the public TXT contract: directly following the request, marker at the request text column.
Private Sub TextRegressionModification(ByVal text As String, ByVal requestText As String, _
                                       ByVal modificationText As String, ByVal label As String, _
                                       ByRef checked As Long)
    Dim lines As Variant, i As Long, requestPosition As Long
    lines = Split(text, vbCrLf)
    For i = LBound(lines) To UBound(lines) - 1
        requestPosition = InStr(1, CStr(lines(i)), requestText, vbBinaryCompare)
        If requestPosition > 0 Then
            TextRegressionEqual CStr(lines(i + 1)), Space$(requestPosition - 1) & "└ " & modificationText, _
                label & " marker alignment and immediate placement", checked
            Exit Sub
        End If
    Next i
    Err.Raise vbObjectError + 7593, "TextReportRegression", label & ": request not found [" & requestText & "]"
End Sub

Private Function ReadTextRegressionUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextRegressionUtf8 = stream.ReadText
    stream.Close
End Function

Private Function TextRegressionSnapshot(ByVal ws As Worksheet) As String
    Dim cell As Range, result As String, value As String
    For Each cell In ws.Range("B2:G5,B7:B8,E8:G14,I3:J4,B17:B19")
        value = CStr(cell.Value2)
        result = result & CStr(Len(value)) & ":" & value & ";"
    Next cell
    TextRegressionSnapshot = result
End Function

Private Sub TextRegressionEqual(ByVal actual As String, ByVal expected As String, _
                                ByVal label As String, ByRef checked As Long)
    If actual <> expected Then Err.Raise vbObjectError + 7591, "TextReportRegression", _
        label & ": expected [" & expected & "], got [" & actual & "]"
    checked = checked + 1
End Sub

Private Sub TextRegressionAssert(ByVal condition As Boolean, ByVal label As String, ByRef checked As Long)
    If Not condition Then Err.Raise vbObjectError + 7592, "TextReportRegression", label
    checked = checked + 1
End Sub

Private Sub TextRegressionContains(ByVal text As String, ByVal needle As String, _
                                   ByVal label As String, ByRef checked As Long)
    TextRegressionAssert InStr(1, text, needle, vbBinaryCompare) > 0, label & ": missing [" & needle & "]", checked
End Sub

Private Sub TextRegressionAbsent(ByVal text As String, ByVal needle As String, _
                                 ByVal label As String, ByRef checked As Long)
    TextRegressionAssert InStr(1, text, needle, vbBinaryCompare) = 0, label & ": unexpectedly found [" & needle & "]", checked
End Sub
