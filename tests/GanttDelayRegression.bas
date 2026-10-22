Attribute VB_Name = "GanttDelayRegression"
Option Explicit
' 2026-10-22 김좌현 수정 : 확대된 기본 노란색 모서리 배경 검사
' 2026-10-22 김좌현 수정 : 셀 배경 모서리 표시, 기존 완료색 및 도형/메모/값 미생성과 숨김/펼침 검증
' 2026-10-22 김좌현 수정 : 짧은 상태/완료 초록/빈 막대와 지연 완료 각 셀의 단일 사선 검증
' 2026-10-22 김좌현 수정 : 오늘 민트 헤더/배경과 지연 완료 빗금 및 상태 전환 회귀 검사
' 2026-10-21 김좌현 수정 : 지연 완료 이력과 상위 전파 및 분홍 구간의 실제 Excel 회귀 검사

Private Const PINK As Long = 15461375 ' RGB(255, 235, 235)

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 953, "GanttDelayRegression", message
End Sub

Private Sub CheckCompletedDelay(ByVal cell As Range, ByVal message As String)
    Dim gradient As LinearGradient, baseColor As Long
    Select Case GetTaskLevel(cell.Parent, cell.Row)
        Case 1: baseColor = RGB(36, 59, 83)
        Case 2: baseColor = RGB(52, 120, 212)
        Case Else: baseColor = RGB(0, 140, 134)
    End Select
    Check cell.Interior.Pattern = xlPatternLinearGradient, message & " native corner fill missing"
    Set gradient = cell.Interior.Gradient
    Check gradient.Degree = 45, message & " corner is not top-left"
    Check gradient.ColorStops.Count = 4, message & " corner color stops"
    Check gradient.ColorStops(1).Color = RGB(255, 255, 0) And _
          gradient.ColorStops(2).Color = RGB(255, 255, 0), message & " standard yellow corner color"
    Check gradient.ColorStops(3).Color = baseColor And _
          gradient.ColorStops(4).Color = baseColor, message & " original completion color"
    Check gradient.ColorStops(1).Position = 0 And gradient.ColorStops(4).Position = 1, message & " incomplete background fill"
    Check gradient.ColorStops(2).Position > 0.11 And gradient.ColorStops(2).Position <= 0.15, message & " enlarged corner size"
    Check gradient.ColorStops(3).Position > gradient.ColorStops(2).Position And _
          gradient.ColorStops(3).Position - gradient.ColorStops(2).Position < 0.001, message & " blurred corner edge"
    Check cell.Borders(xlDiagonalUp).LineStyle = xlNone, message & " stale diagonal up"
    Check cell.Borders(xlDiagonalDown).LineStyle = xlNone, message & " stale diagonal down"
    Check Len(CStr(cell.Value2)) = 0 And Not cell.HasFormula, message & " unexpected cell content"
End Sub

Private Function CornerFillCount(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    For c = 24 To lastCol
        If ws.Cells(rowNum, c).Interior.Pattern = xlPatternLinearGradient Then CornerFillCount = CornerFillCount + 1
    Next c
End Function

Private Sub Seed(ByVal ws As Worksheet, ByVal r As Long, ByVal level As Long, _
                 ByVal caption As String, ByVal planStart As Date, ByVal planEnd As Date, _
                 ByVal actualStart As Variant, ByVal actualEnd As Variant, ByVal progress As Double)
    ws.Cells(r, "C").Value2 = level
    ws.Cells(r, "H").Value2 = caption
    ws.Cells(r, "L").Value = planStart
    ws.Cells(r, "M").Value = planEnd
    If IsDate(actualStart) Then ws.Cells(r, "N").Value = CDate(actualStart)
    If IsDate(actualEnd) Then ws.Cells(r, "O").Value = CDate(actualEnd)
    ws.Cells(r, "P").Value2 = progress
End Sub

Private Function DayCell(ByVal ws As Worksheet, ByVal r As Long, ByVal d As Date) As Range
    Set DayCell = ws.Cells(r, 24 + CLng(d - DateSerial(2026, 10, 12)))
End Function

Public Function RunGanttDelayRegression() As String
    Dim ws As Worksheet, dateProbe As Worksheet, holidays As Object, workdays As Object
    Dim r As Long, c As Long, lastCol As Long, result As String
    Dim startDate As Date, endDate As Date, minimumDate As Date, maximumDate As Date
    Dim inputs As Variant, afterInputs As Variant, rowIndex As Long, colIndex As Long
    Dim rowHeight As Double
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "GanttDelayPreview"
    Set holidays = CreateObject("Scripting.Dictionary")
    Set workdays = CreateObject("Scripting.Dictionary")
    holidays.Add NormalizeDateKey(DateSerial(2026, 10, 14)), True
    workdays.Add NormalizeDateKey(DateSerial(2026, 10, 17)), True
    startDate = DateSerial(2026, 10, 12)
    endDate = DateSerial(2026, 10, 30)
    If Date + 12 > endDate Then endDate = Date + 12
    lastCol = 24 + CLng(endDate - startDate)
    DrawDateHeader ws, startDate, endDate, holidays, workdays
    Check DayCell(ws, 3, Date).Value2 = "오늘", "today header caption"
    Check DayCell(ws, 3, Date).Interior.Color = RGB(45, 212, 191), "today header mint"
    Check DayCell(ws, 4, Date).Interior.Color = RGB(45, 212, 191), "today date mint"
    Check CLng(DayCell(ws, 4, Date).Value2) = CLng(Date), "today date must remain numeric"
    ws.Rows(4).RowHeight = 58
    ws.Range("H4").Value = "요청 내용"
    ws.Range("W4").Value = "상태 / 지연 일수"
    ws.Columns("H").ColumnWidth = 28
    ws.Columns("W").ColumnWidth = 13
    ws.Rows("5:21").RowHeight = 25
    Seed ws, 5, 1, "기한 내 완료", startDate, DateSerial(2026, 10, 16), startDate, DateSerial(2026, 10, 16), 0
    Seed ws, 6, 1, "지연 완료 (늦게 시작, 짧게 수행)", startDate, DateSerial(2026, 10, 13), DateSerial(2026, 10, 19), DateSerial(2026, 10, 20), 0
    Seed ws, 7, 1, "진행 중 90%, 마감 초과", Date - 14, Date - 7, Date - 14, Empty, 0.9
    Seed ws, 8, 1, "시작하지 못한 마감 초과", Date - 14, Date - 7, Empty, Empty, 0
    Seed ws, 9, 1, "100%, 실제 완료일 미입력", startDate, DateSerial(2026, 10, 13), startDate, Empty, 1
    Seed ws, 10, 1, "앞으로 진행할 정상 작업", Date + 1, Date + 10, Empty, Empty, 0
    Seed ws, 11, 1, "상위 완료, 자체 기한 준수", startDate, DateSerial(2026, 10, 30), Empty, Empty, 0
    Seed ws, 12, 2, "중간 완료, 하위 지연 이력", startDate, DateSerial(2026, 10, 30), Empty, Empty, 0
    Seed ws, 13, 3, "하위 지연 완료 7일", startDate, DateSerial(2026, 10, 13), DateSerial(2026, 10, 19), DateSerial(2026, 10, 20), 0
    Seed ws, 14, 3, "하위 기한 내 완료", startDate, DateSerial(2026, 10, 30), startDate, DateSerial(2026, 10, 21), 0
    Seed ws, 15, 1, "진행 중 상위, 하위 지연 이력", startDate, Date + 10, Empty, Empty, 0
    Seed ws, 16, 2, "진행 중 중간 작업", startDate, Date + 10, Empty, Empty, 0
    Seed ws, 17, 3, "완료된 하위 지연 7일", startDate, DateSerial(2026, 10, 13), DateSerial(2026, 10, 19), DateSerial(2026, 10, 20), 0
    Seed ws, 18, 3, "아직 기한이 남은 하위", Date + 1, Date + 10, Empty, Empty, 0
    Seed ws, 19, 1, "보류 상태 유지", Date - 14, Date - 7, Empty, Empty, 0
    ws.Range("S19").Value = STATUS_HOLD
    Seed ws, 20, 1, "잘못된 날짜는 오류 유지", startDate + 1, startDate, Empty, Empty, 0
    Seed ws, 21, 1, "계획일 없이 완료", startDate, startDate, startDate, startDate, 1
    ws.Range("L21:M21").ClearContents

    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    Check CStr(ws.Range("W5").Value2) = STATUS_DONE, "on-time completion status"
    Check ws.Range("W5").Text = STATUS_DONE, "on-time completion gained delay detail"
    Check GetTaskDelayDays(ws, 5, holidays, workdays) = 0, "on-time completion delay"
    Check CStr(ws.Range("W6").Value2) = STATUS_DONE, "late completion lost base done status"
    Check GetTaskDelayDays(ws, 6, holidays, workdays) = 7, "late completion must use finish dates, not duration difference"
    Check ws.Range("W6").Text = STATUS_DONE & "(7)", "compact late completion days"
    Check ws.Range("V6").Value2 = 2 And ws.Range("U6").Value2 = 2, "late-start fixture must have equal durations"
    Check CStr(ws.Range("W7").Value2) = STATUS_DELAY, "90% overdue task remained normal"
    Check ws.Range("W7").Text = STATUS_DELAY & "(7)", "compact ongoing delay days"
    Check CStr(ws.Range("W8").Value2) = STATUS_DELAY, "unstarted overdue task remained normal"
    Check GetTaskDelayDays(ws, 7, holidays, workdays) = 7, "running delay must include all seven calendar days"
    Check CStr(ws.Range("W9").Value2) = STATUS_DONE, "100% completion status"
    Check ws.Range("W9").Text = STATUS_DONE, "unknown completion must have no invented days"
    Check GetTaskDelayDays(ws, 9, holidays, workdays) = "", "missing completion date used today"
    Check CStr(ws.Range("W10").Value2) = STATUS_NORMAL, "future task status"
    Check CStr(ws.Range("W11").Value2) = STATUS_DONE And CStr(ws.Range("W12").Value2) = STATUS_DONE, "completed parents lost completion"
    Check GetTaskDelayDays(ws, 11, holidays, workdays) = 0, "child delays were summed into parent"
    Check ws.Range("W11").Text = STATUS_DONE, "root gained verbose child detail"
    Check ws.Range("W12").Text = STATUS_DONE, "middle gained verbose child detail"
    Check CStr(ws.Range("W15").Value2) = STATUS_DELAY And CStr(ws.Range("W16").Value2) = STATUS_DELAY, "completed child delay did not reach unfinished ancestors"
    Check CStr(ws.Range("W19").Value2) = STATUS_HOLD, "manual hold changed"
    Check CStr(ws.Range("W20").Value2) = STATUS_ERROR, "invalid dates changed"
    Check GetTaskDelayDays(ws, 21, holidays, workdays) = "", "missing plan dates fabricated delay"
    ws.Range("Z6").Borders(xlDiagonalUp).LineStyle = xlContinuous
    ws.Range("Z6").Borders(xlDiagonalDown).LineStyle = xlContinuous
    ws.Range("Z6").Interior.Pattern = xlLightUp
    inputs = ws.Range("B5:W21").Value2
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    afterInputs = ws.Range("B5:W21").Value2
    For rowIndex = 1 To UBound(inputs, 1)
        For colIndex = 1 To UBound(inputs, 2)
            Check CStr(inputs(rowIndex, colIndex)) = CStr(afterInputs(rowIndex, colIndex)), "renderer changed task inputs"
        Next colIndex
    Next rowIndex
    CheckCompletedDelay DayCell(ws, 6, DateSerial(2026, 10, 15)), "gap before late start"
    CheckCompletedDelay DayCell(ws, 6, DateSerial(2026, 10, 17)), "configured Saturday workday"
    CheckCompletedDelay DayCell(ws, 6, DateSerial(2026, 10, 14)), "holiday calendar delay"
    CheckCompletedDelay DayCell(ws, 6, DateSerial(2026, 10, 18)), "Sunday calendar delay"
    Check DayCell(ws, 6, DateSerial(2026, 10, 20)).Interior.Gradient.ColorStops(4).Color = RGB(36, 59, 83), "background retains existing level color"
    For c = 26 To 32
        CheckCompletedDelay ws.Cells(6, c), "each delayed calendar cell"
    Next c
    Check CornerFillCount(ws, 6, lastCol) = 7, "one background corner per delayed calendar date"
    Check DayCell(ws, 6, DateSerial(2026, 10, 13)).Interior.Pattern = xlSolid, "planned finish gained corner"
    Check DayCell(ws, 6, DateSerial(2026, 10, 21)).Interior.Pattern = xlSolid, "corner passed actual finish"
    Check DayCell(ws, 6, DateSerial(2026, 10, 21)).Interior.Color <> PINK, "completed delay kept growing"
    CheckCompletedDelay DayCell(ws, 11, DateSerial(2026, 10, 15)), "root inherited completion"
    CheckCompletedDelay DayCell(ws, 16, DateSerial(2026, 10, 15)), "unfinished parent with completed delayed child"
    Check Len(DayCell(ws, 12, DateSerial(2026, 10, 20)).Text) = 0, "parent inherited bar text"
    Check DayCell(ws, 8, Date).Interior.Color = PINK, "unstarted task has no delay bar"
    Check DayCell(ws, 8, Date).Interior.Pattern = xlSolid, "ongoing delay should be solid pink"
    Check DayCell(ws, 10, Date).Interior.Color = RGB(224, 252, 244), "today empty background is not pale mint"
    Check DayCell(ws, 7, Date).Interior.Color = PINK, "today mint overwrote running delay"
    Check DayCell(ws, 5, DateSerial(2026, 10, 12)).Interior.Color = RGB(36, 59, 83), "original actual bar color changed"
    For c = 24 To lastCol
        Check ws.Cells(5, c).Interior.Color <> PINK, "on-time completion painted pink"
        Check ws.Cells(5, c).Borders(xlDiagonalUp).LineStyle = xlNone, "on-time completion gained slash"
        Check ws.Cells(9, c).Interior.Color <> PINK, "unknown completion painted pink"
        Check ws.Cells(9, c).Borders(xlDiagonalUp).LineStyle = xlNone, "unknown completion gained slash"
        Check ws.Cells(20, c).Interior.Color <> PINK, "invalid task painted pink"
        Check ws.Cells(5, c).Interior.Pattern = xlSolid, "on-time completion gained corner"
        Check ws.Cells(7, c).Interior.Pattern = xlSolid, "ongoing delay gained corner"
        Check ws.Cells(9, c).Interior.Pattern = xlSolid, "unknown completion gained corner"
        Check ws.Cells(20, c).Interior.Pattern = xlSolid, "invalid task gained corner"
    Next c
    Check ws.Shapes.Count = 0, "delay renderer created shapes"
    Check ws.Comments.Count = 0, "delay renderer created notes"
    ws.Rows(6).RowHeight = 31
    rowHeight = ws.Rows(6).RowHeight
    CheckCompletedDelay ws.Range("AF6"), "resized native background"
    ws.Rows(6).Hidden = True
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check ws.Rows(6).Hidden, "native background render expanded hidden row"
    ws.Rows(6).Hidden = False
    Check ws.Rows(6).RowHeight = rowHeight, "native background changed row height"
    CheckCompletedDelay ws.Range("AF6"), "restored hidden row"
    ws.Outline.SummaryRow = xlAbove
    ws.Rows("12:14").Group
    ws.Rows(11).ShowDetail = False
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check ws.Rows(12).Hidden And ws.Rows(13).Hidden, "native background render expanded outline"
    ws.Rows(11).ShowDetail = True
    CheckCompletedDelay ws.Range("AF13"), "restored outline child"
    ' A live overrun changes from pink to a native corner fill and back on reopening.
    ws.Range("O7").Value = Date
    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    CheckCompletedDelay DayCell(ws, 7, Date), "newly completed delay"
    Check GetTaskDelayDays(ws, 7, holidays, workdays) = 7, "completion changed calendar delay"
    ws.Range("O7").ClearContents
    ws.Range("P7").Value2 = 0.9
    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check DayCell(ws, 7, Date).Interior.Color = PINK, "reopened delay not pink"
    Check DayCell(ws, 7, Date).Interior.Pattern = xlSolid, "reopened delay left stale hatch"
    Check DayCell(ws, 7, Date).Borders(xlDiagonalUp).LineStyle = xlNone, "reopened delay left stale slash"
    Check CornerFillCount(ws, 7, lastCol) = 0, "reopened delay left stale background corners"
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check Application.CountA(ws.Range(ws.Cells(5, 24), ws.Cells(21, lastCol))) = 0, "repeated render added bar text"
    ws.Range("O14").Value = DateSerial(2026, 10, 20)
    ws.Range("M14").Value = DateSerial(2026, 10, 13)
    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    Check GetTaskDelayDays(ws, 11, holidays, workdays) = 7, "two overlapping child delays must not add to fourteen"
    ws.Range("O14").Value = DateSerial(2026, 10, 21)
    ws.Range("M14").Value = DateSerial(2026, 10, 30)
    ws.Range("N6").Value = startDate
    ws.Range("O6").Value = DateSerial(2026, 10, 13)
    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check ws.Range("W6").Text = STATUS_DONE, "corrected dates left stale status suffix"
    Check CornerFillCount(ws, 6, lastCol) = 0, "corrected dates left stale background corners"
    For c = 24 To lastCol
        Check ws.Cells(6, c).Interior.Color <> PINK, "corrected dates left stale pink"
        Check ws.Cells(6, c).Borders(xlDiagonalUp).LineStyle = xlNone, "corrected dates left stale slash"
        Check Len(CStr(ws.Cells(6, c).Value2)) = 0, "corrected dates left stale delay label"
    Next c
    ' Leave a delayed completion in the preview and verify clipped display totals.
    ws.Range("N6").Value = DateSerial(2026, 10, 19)
    ws.Range("O6").Value = DateSerial(2026, 10, 20)
    ApplyHierarchySummaryValues ws, 21, holidays, workdays
    RenderHierarchyTaskBars ws, 21, 27, 28, holidays, workdays
    CheckCompletedDelay ws.Cells(6, 27), "clipped first delayed cell"
    CheckCompletedDelay ws.Cells(6, 28), "clipped last delayed cell"
    Check ws.Range("W6").Text = STATUS_DONE & "(7)", "clipping changed status delay days"
    FormatBaseArea ws, 21, startDate, endDate, holidays, workdays
    Check ws.Range("W5").Interior.Color = RGB(198, 239, 206), "on-time completion green changed"
    Check ws.Range("W6").Interior.Color = RGB(198, 239, 206), "late completion must stay green"
    Check ws.Range("W11").Interior.Color = RGB(198, 239, 206), "completed parent must stay green"
    RenderHierarchyTaskBars ws, 21, 24, lastCol, holidays, workdays
    Check Application.CountA(ws.Range(ws.Cells(5, 24), ws.Cells(21, lastCol))) = 0, "timeline must contain no labels"
    Check DayCell(ws, 8, Date).Interior.Color = ws.Range("H8").Interior.Color, "ongoing delay must match project description pink"
    Check ws.Shapes.Count = 0 And ws.Comments.Count = 0, "final background created objects or notes"
    ws.Range("L5:O21").NumberFormat = "yyyy-mm-dd"
    ws.Range("P5:P21").NumberFormat = "0%"
    ws.Columns("L:O").ColumnWidth = 11
    ws.Columns("I:K").Hidden = True
    ws.Columns("Q:V").Hidden = True
    ws.Columns("A:G").Hidden = True
    ws.Rows("1:2").Hidden = True
    ws.Range("H3").Value = "민트 오늘 / 분홍 지연 / 배경 모서리 지연 완료"
    Set dateProbe = ThisWorkbook.Worksheets.Add
    Seed dateProbe, 5, 1, "Unstarted overdue", Date - 14, Date - 7, Empty, Empty, 0
    Check GetMinMaxDate(dateProbe, 5, minimumDate, maximumDate), "unstarted date range missing"
    Check maximumDate = Date + 2, "unstarted overdue timeline does not reach today"
    dateProbe.Range("P5").Value2 = 1
    Check GetMinMaxDate(dateProbe, 5, minimumDate, maximumDate), "unknown completion range missing"
    Check maximumDate = Date - 5, "unknown completion extended timeline to today"
    dateProbe.Delete
    ws.Activate
    ws.Range("H3").Select
    RunGanttDelayRegression = "PASS: mint today header/background, original bar colors, compact status, green completion, text-free bars, pale pink ongoing delay, enlarged native yellow corners, no shapes/notes/values, hide/outline preservation, calendar days including weekends/holidays, hierarchy history, completion/reopen transitions, clipping, stale cleanup and input preservation"
    Exit Function
Failed:
    RunGanttDelayRegression = "FAIL: " & Err.Description
End Function
