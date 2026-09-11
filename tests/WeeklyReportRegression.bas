' Append to modWeeklyPptReport only in a disposable workbook copy.
' No module declarations: this harness calls the production private builders.
Public Function RunWeeklyReportRegression(ByVal outputPath As String) As String
    Dim ws As Worksheet
    Dim rows As Collection, items As Collection, dates As Collection
    Dim levels As Collection, plans As Collection
    Dim pptApp As Object, presentation As Object, slide As Object, tableShape As Object
    Dim checked As Long, i As Long, snapshot As String, stage As String
    Dim errNumber As Long, errDescription As String
    Dim planLines As Variant, expected As Variant
    Dim itemPages As New Collection, datePages As New Collection
    Dim levelPages As New Collection, planPages As New Collection
    Dim duplicatedSlides As Object

    On Error GoTo Failed
    stage = "configure fixture"
    EnsureWeeklyReportConfigSheet
    Set ws = ThisWorkbook.Worksheets(WEEKLY_REPORT_CONFIG_SHEET_NAME)
    ws.Range("B2:G5").Value2 = "N"
    ws.Range("C2").Value2 = "Y"
    ws.Range("F2").Value2 = "Y"
    ws.Range("F5").Value2 = "Y"
    ws.Range("G5").Value2 = "전체"
    ws.Range("B7").Value2 = WEEKLY_REPORT_PAGE_MODE_ALL
    ws.Range("B8").Value2 = WEEKLY_REPORT_CATEGORY_DEPTH_TYPE
    ws.Range("B9").Value2 = WEEKLY_REPORT_OVERFLOW_MODE_NEW_SLIDE
    ws.Range("E8:E14").Value = Application.Transpose(Array("ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ"))
    ws.Range("F8:F14").Value = Application.Transpose( _
        Array("표시 안 함", "표시 안 함", "기호", "기호", "기호", "표시 안 함", "표시 안 함"))
    Set rows = WeeklyRegressionRows()

    stage = "hidden categories and legacy aliases"
    ResetWeeklyReportNumbering
    WeeklyRegressionEqual GetWeeklyReportCategoryBullet(2), "", "legacy no bullet", checked
    WeeklyRegressionBuild rows, items, dates, levels, plans
    WeeklyRegressionEqual CStr(items.Count), "10", "hidden current count", checked
    WeeklyRegressionEqual CStr(plans.Count), "2", "hidden plan blocks", checked
    WeeklyRegressionEqual CStr(levels(1)), "-2", "major category retained", checked
    For i = 2 To levels.Count
        WeeklyRegressionAssert CLng(levels(i)) > 0, "N category appeared in current", checked
    Next i
    WeeklyRegressionAssert InStr(CStr(plans(1)) & CStr(plans(2)), "중분류") = 0, "N middle appeared in plan", checked
    WeeklyRegressionAssert InStr(CStr(plans(1)) & CStr(plans(2)), "소분류") = 0, "N minor appeared in plan", checked

    stage = "preview formulas and idempotent initialization"
    EnsureWeeklyReportConfigSheet
    WeeklyRegressionEqual CStr(ws.Range("F9").Value2), "글머리 없음", "legacy mode normalization", checked
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G9").Value2), "대분류", "no bullet preview", checked
    WeeklyRegressionEqual CStr(ws.Range("G10").Value2), "항목 숨김", "hidden middle preview", checked
    WeeklyRegressionEqual CStr(ws.Range("G12").Value2), "ㅁ 업무명", "task preview", checked
    For i = 8 To 14
        WeeklyRegressionAssert ws.Cells(i, 7).HasFormula, "missing preview formula " & CStr(i), checked
    Next i
    ws.Range("D2:E2").Value2 = "Y"
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G10").Value2), "ㄷ 중분류", "visible middle preview", checked
    WeeklyRegressionEqual CStr(ws.Range("G11").Value2), "ㄹ 소분류", "visible minor preview", checked
    ws.Range("F2").Value2 = "N"
    ws.Range("G2").Value2 = "Y"
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G12").Value2), "ㅁ Level 1", "level only preview", checked
    ws.Range("F2").Value2 = "Y"
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G12").Value2), "ㅁ Level 1 - 업무명", "name and level preview", checked
    ws.Range("F2:G2").Value2 = "N"
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G12").Value2), "항목 숨김", "hidden task preview", checked
    ws.Range("F2").Value2 = "Y"
    snapshot = WeeklyRegressionConfigSnapshot(ws)
    EnsureWeeklyReportConfigSheet
    EnsureWeeklyReportConfigSheet
    WeeklyRegressionEqual WeeklyRegressionConfigSnapshot(ws), snapshot, "initialization preserves settings", checked

    stage = "visible current and plan builders"
    WeeklyRegressionBuild rows, items, dates, levels, plans
    WeeklyRegressionEqual CStr(items.Count), "14", "visible current count", checked
    WeeklyRegressionEqual CStr(dates.Count), CStr(items.Count), "dates aligned with items", checked
    WeeklyRegressionEqual CStr(levels.Count), CStr(items.Count), "levels aligned with items", checked
    expected = Array("REO", "    ㄷ 중분류A", "        ㄹ 소분류A", _
        "            ㅁ 업무A", "                하위A", "                    상세A", _
        "            ㅁ 업무B", "                하위B", "                    상세B")
    planLines = Split(CStr(plans(1)), ChrW(11))
    WeeklyRegressionEqual CStr(UBound(planLines) + 1), CStr(UBound(expected) + 1), "first plan lines", checked
    For i = 0 To UBound(expected)
        WeeklyRegressionEqual CStr(planLines(i)), CStr(expected(i)), "first plan line " & CStr(i), checked
    Next i
    planLines = Split(CStr(plans(2)), ChrW(11))
    WeeklyRegressionEqual CStr(planLines(0)), "    ㄷ 중분류B", "new plan block preserves middle indent", checked
    WeeklyRegressionEqual CStr(planLines(1)), "        ㄹ 소분류B", "second minor indent", checked
    WeeklyRegressionEqual CStr(planLines(4)), "                    상세C", "level 3 no extra space", checked

    stage = "render current and plan into PowerPoint"
    Set pptApp = CreateObject("PowerPoint.Application")
    Set presentation = CreateCodeBasedWeeklyPptPresentation(pptApp)
    Set slide = presentation.Slides(1)
    FillWeeklyReportCurrentTable slide, items, dates, levels, GetWeeklyReportVisibleCategoryCount()
    FillWeeklyReportPlanArea slide, plans, False
    Set tableShape = FindFirstTableShape(slide)
    For i = 0 To UBound(expected)
        WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, i + 1), _
            CStr(expected(i)), "current rendered line " & CStr(i), checked
    Next i
    WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, 10), "    ㄷ 중분류B", "current second middle indent", checked
    WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, 14), "                    상세C", "current last leaf", checked
    WeeklyRegressionAssert InStr(CStr(FindTextShape(slide, "(개발 항목)").TextFrame.TextRange.Text), _
        "    ㄷ 중분류B") > 0, "rendered plan loses second block indent", checked

    stage = "numbering and aliases"
    ws.Range("F12").Value2 = "레벨 번호"
    ResetWeeklyReportNumbering
    WeeklyRegressionEqual GetWeeklyReportLevelBullet(1), "1.", "legacy numbering first", checked
    WeeklyRegressionEqual GetWeeklyReportLevelBullet(1), "2.", "legacy numbering second", checked
    ws.Calculate
    WeeklyRegressionEqual CStr(ws.Range("G12").Value2), "1. 업무명", "legacy numbering preview", checked
    ws.Range("F12").Value2 = "번호 매기기"
    ResetWeeklyReportNumbering
    ws.Calculate
    WeeklyRegressionEqual GetWeeklyReportLevelBullet(1), "1.", "preview does not consume numbering", checked
    WeeklyRegressionBuild rows, items, dates, levels, plans
    planLines = Split(CStr(plans(1)), ChrW(11))
    WeeklyRegressionEqual CStr(planLines(3)), "            1. 업무A", "plan first number", checked
    WeeklyRegressionEqual CStr(planLines(6)), "            2. 업무B", "plan second number", checked
    planLines = Split(CStr(plans(2)), ChrW(11))
    WeeklyRegressionEqual CStr(planLines(2)), "            1. 업무C", "plan new category numbering reset", checked
    FillWeeklyReportCurrentTable slide, items, dates, levels, GetWeeklyReportVisibleCategoryCount()
    WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, 4), "            1. 업무A", "current first number", checked
    WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, 7), "            2. 업무B", "current second number", checked
    WeeklyRegressionEqual WeeklyRegressionParagraph(tableShape, 12), "            1. 업무C", "current new category numbering reset", checked

    stage = "save symbol sample"
    ws.Range("F12").Value2 = "기호"
    WeeklyRegressionBuild rows, items, dates, levels, plans
    Set tableShape = Nothing
    Set slide = Nothing
    presentation.Close
    Set presentation = CreateCodeBasedWeeklyPptPresentation(pptApp)
    Set slide = presentation.Slides(1)
    AppendWeeklyOutputPages items, dates, levels, plans, WEEKLY_REPORT_OVERFLOW_MODE_NEW_SLIDE, _
        GetWeeklyReportCurrentPageCapacity(slide), GetWeeklyReportPlanPageCapacity(slide), _
        itemPages, datePages, levelPages, planPages
    For i = 2 To itemPages.Count
        Set duplicatedSlides = presentation.Slides(1).Duplicate
        Set duplicatedSlides = Nothing
    Next i
    For i = 1 To itemPages.Count
        Set slide = presentation.Slides(i)
        Set items = itemPages(i)
        Set dates = datePages(i)
        Set levels = levelPages(i)
        Set plans = planPages(i)
        FillWeeklyReportCurrentTable slide, items, dates, levels, GetWeeklyReportVisibleCategoryCount()
        FillWeeklyReportPlanArea slide, plans, False
        FillWeeklyReportPeriodText slide, DateSerial(2026, 9, 7), DateSerial(2026, 9, 11), _
            DateSerial(2026, 9, 14), DateSerial(2026, 9, 18)
    Next i
    presentation.SaveAs outputPath, PPT_SAVE_AS_OPEN_XML_PRESENTATION
    presentation.Slides(1).Export outputPath & ".png", "PNG", 1300, 900
    Set tableShape = Nothing
    Set slide = Nothing
    presentation.Close
    Set presentation = Nothing
    pptApp.Quit
    Set pptApp = Nothing
    RunWeeklyReportRegression = "PASS: " & CStr(checked) & " assertions; " & outputPath
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    Set duplicatedSlides = Nothing
    Set tableShape = Nothing
    Set slide = Nothing
    If Not presentation Is Nothing Then presentation.Close
    Set presentation = Nothing
    If Not pptApp Is Nothing Then pptApp.Quit
    Set pptApp = Nothing
    On Error GoTo 0
    Err.Raise errNumber, "RunWeeklyReportRegression: " & stage, errDescription
End Function

Private Function WeeklyRegressionRows() As Collection
    Dim result As New Collection
    Dim names As Variant, middles As Variant, minors As Variant
    Dim i As Long, sourceRow As Long, hierarchy As Variant, moduleName As String
    names = Array("A", "B", "C")
    middles = Array("중분류A", "중분류A", "중분류B")
    minors = Array("소분류A", "소분류A", "소분류B")
    For i = 0 To 2
        sourceRow = DATA_START_ROW + i * 3
        hierarchy = Array(CStr(sourceRow) & vbTab & "업무" & CStr(names(i)), _
            CStr(sourceRow + 1) & vbTab & "하위" & CStr(names(i)), _
            CStr(sourceRow + 2) & vbTab & "상세" & CStr(names(i)))
        moduleName = "개발 > REO > " & CStr(middles(i))
        result.Add Array(moduleName, "상세" & CStr(names(i)), 2, _
            CDbl(DateSerial(2026, 9, 17)), "~9/17", False, 3, "담당자", _
            sourceRow + 2, hierarchy, CStr(minors(i)), "개발")
    Next i
    Set WeeklyRegressionRows = result
End Function

Private Sub WeeklyRegressionBuild(ByVal rows As Collection, ByRef items As Collection, _
                                  ByRef dates As Collection, ByRef levels As Collection, _
                                  ByRef plans As Collection)
    Set items = New Collection
    Set dates = New Collection
    Set levels = New Collection
    Set plans = New Collection
    BuildWeeklyGroupedCurrentItems rows, items, dates, levels, False, False, False, _
        0, GetWeeklyReportVisibleCategoryCount()
    BuildWeeklyGroupedPlanItems rows, plans, False, False, False, _
        0, GetWeeklyReportVisibleCategoryCount()
End Sub

Private Function WeeklyRegressionParagraph(ByVal tableShape As Object, ByVal index As Long) As String
    Dim result As String
    result = CStr(tableShape.Table.Cell(2, 2).Shape.TextFrame.TextRange.Paragraphs(index).Text)
    If Len(result) > 0 Then
        If Right$(result, 1) = vbCr Then result = Left$(result, Len(result) - 1)
    End If
    WeeklyRegressionParagraph = result
End Function

Private Function WeeklyRegressionConfigSnapshot(ByVal ws As Worksheet) As String
    Dim cell As Range, result As String, value As String
    For Each cell In ws.Range("B2:G5,B7:B9,E8:F14,H3:I3")
        value = CStr(cell.Value2)
        result = result & CStr(Len(value)) & ":" & value & ";"
    Next cell
    WeeklyRegressionConfigSnapshot = result
End Function

Private Sub WeeklyRegressionEqual(ByVal actual As String, ByVal expected As String, _
                                  ByVal label As String, ByRef checked As Long)
    If actual <> expected Then
        Err.Raise vbObjectError + 7591, "WeeklyReportRegression", _
            label & ": expected [" & expected & "], got [" & actual & "]"
    End If
    checked = checked + 1
End Sub

Private Sub WeeklyRegressionAssert(ByVal condition As Boolean, ByVal label As String, _
                                   ByRef checked As Long)
    If Not condition Then Err.Raise vbObjectError + 7592, "WeeklyReportRegression", label
    checked = checked + 1
End Sub
