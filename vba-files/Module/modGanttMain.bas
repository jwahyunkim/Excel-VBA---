Attribute VB_Name = "modGanttMain"
Option Explicit
' 2026-10-27 김좌현 수정 : 모두 새로고침 후 등록 기준 문제 행을 완료 메시지에 안내
' 2026-10-23 김좌현 수정 : 간트 전용 새로고침도 이전 웹등록 헤더를 안전하게 WEB UPDATE로 갱신
' 2026-10-22 김좌현 수정 : 새로고침 서명 계산의 열 번호를 한 번만 조회해 반복 Range 접근 제거
' 2026-10-22 김좌현 수정 : 중복 새로고침 매크로를 제거하고 모두 새로고침으로 버튼 연결 통일
' 2026-10-22 김좌현 수정 : 모두/간트 새로고침 분리, 중복 렌더 제거와 변경된 레이아웃만 재구성
' 2026-10-22 김좌현 수정 : 웹 유틸리티 시트 제외 및 웹등록 설정/버튼 초기화
' 2026-10-20 김좌현 수정 : 새로고침 실패 시 완료 숨김과 수동/계층 보기 및 보호 복원
' 2026-10-20 김좌현 수정 : 완료 숨김 버튼과 새로고침 시 행 보기 상태 유지
' 2026-10-20 김좌현 수정 : 미사용 간트 only/보고 only 행 필터 호출 제거
' 2026-10-20 김좌현 수정 : 미사용 항목 숨기기 매크로와 전용 버튼 생성 제거

' 2026-10-20 김좌현 수정 : 생성/새로고침을 새로고침 진입점으로 통합
Private mTaskLayoutSignatures As Object
Private mTaskStructureSignatures As Object
Private mTaskVisualSignatures As Object

Public Sub InvalidateGanttRenderCache(ByVal ws As Worksheet)
    If ws Is Nothing Or mTaskVisualSignatures Is Nothing Then Exit Sub
    If mTaskVisualSignatures.Exists(ws.CodeName) Then mTaskVisualSignatures.Remove ws.CodeName
End Sub

Private Function GanttDateHeaderMatches(ByVal ws As Worksheet, ByVal firstDate As Date, ByVal lastDate As Date) As Boolean
    Dim startCol As Long, endCol As Long
    startCol = ws.Range(COL_GANTT_START & "1").Column
    endCol = startCol + CLng(lastDate - firstDate)
    If Not IsDate(ws.Cells(GANTT_HEADER_ROW_DATE, startCol).Value) Then Exit Function
    If Not IsDate(ws.Cells(GANTT_HEADER_ROW_DATE, endCol).Value) Then Exit Function
    GanttDateHeaderMatches = (DateValue(ws.Cells(GANTT_HEADER_ROW_DATE, startCol).Value) = DateValue(firstDate) And _
                              DateValue(ws.Cells(GANTT_HEADER_ROW_DATE, endCol).Value) = DateValue(lastDate))
End Function

Public Sub 칸트차트_모두새로고침()
    RefreshGanttSheet
End Sub

Public Sub 칸트차트_간트새로고침()
    RefreshGanttOnly
End Sub

Public Sub RefreshGanttSheet(Optional ByVal showCompletionMessage As Boolean = True)
    RefreshGanttCore True, showCompletionMessage
End Sub

Public Sub RefreshGanttOnly(Optional ByVal showCompletionMessage As Boolean = True)
    RefreshGanttCore False, showCompletionMessage
End Sub

Private Function TaskRefreshSignature(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                      ByVal structureOnly As Boolean) As String
    Dim values As Variant, r As Long, c As Long, result As String, value As Variant
    Dim requestOffset As Long
    requestOffset = ws.Range(COL_REQUEST & "1").Column - ws.Range(COL_LEVEL & "1").Column + 1
    values = ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_OWNER & lastRow).Value2
    result = CStr(lastRow) & "|"
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If c = 1 Or (Not structureOnly And _
                (c <= 5 Or c >= requestOffset)) Then
                value = values(r, c)
                If IsError(value) Then
                    result = result & "#ERR|"
                Else
                    result = result & Len(CStr(value)) & ":" & CStr(value) & "|"
                End If
            End If
        Next c
        If structureOnly Then result = result & IIf(HasTaskContent(ws, r + DATA_START_ROW - 1), "T", "-")
    Next r
    TaskRefreshSignature = result
End Function

Private Function TaskVisualSignature(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                     ByVal firstDate As Date, ByVal lastDate As Date, _
                                     ByVal holidays As Object, ByVal workdays As Object) As String
    Dim values As Variant, r As Long, c As Long, columnNumber As Long, result As String, value As Variant
    Dim levelColumn As Long, minorCategoryColumn As Long, requestColumn As Long, workMinutesColumn As Long
    levelColumn = ws.Range(COL_LEVEL & "1").Column
    minorCategoryColumn = ws.Range(COL_MINOR_CATEGORY & "1").Column
    requestColumn = ws.Range(COL_REQUEST & "1").Column
    workMinutesColumn = ws.Range(COL_WORK_MINUTES & "1").Column
    values = ws.Range(COL_LEVEL & DATA_START_ROW & ":" & COL_STATUS & lastRow).Value2
    result = CStr(lastRow) & "|" & CLng(firstDate) & "|" & CLng(lastDate) & "|" & CLng(Date) & "|"
    result = result & Join(holidays.Keys, ",") & "|" & Join(workdays.Keys, ",") & "|"
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            columnNumber = c + levelColumn - 1
            If columnNumber <= minorCategoryColumn Or _
               (columnNumber >= requestColumn And columnNumber <> workMinutesColumn) Then
                value = values(r, c)
                If IsError(value) Then
                    result = result & "#ERR|"
                Else
                    result = result & Len(CStr(value)) & ":" & CStr(value) & "|"
                End If
            End If
        Next c
    Next r
    TaskVisualSignature = result
End Function

Private Sub RefreshGanttCore(ByVal includeWebSetup As Boolean, ByVal showCompletionMessage As Boolean)
    Dim ws As Worksheet, config As Worksheet
    Dim lastRow As Long, minDate As Date, maxDate As Date
    Dim chartStartDate As Date, chartEndDate As Date
    Dim holidayDict As Object, workdayDict As Object
    Dim taskSheetWasProtected As Boolean, previousScreenUpdating As Boolean, previousEnableEvents As Boolean
    Dim errorText As String, detailColumnHidden As Variant, webBasisIssues As String
    Dim detailScrollColumn As Long, detailScrollRow As Long
    Dim rowView As Object, originalRowView As Object
    Dim originalLastRow As Long, taskViewCaptured As Boolean
    Dim layoutSignature As String, structureSignature As String, fitLayout As Boolean, rebuildHierarchy As Boolean
    Dim visualSignature As String, redrawChart As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    If Not TypeOf ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 2120, , "업무 시트에서 실행해 주세요."
    Set ws = ActiveSheet
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(ws) Then _
        Err.Raise vbObjectError + 2120, , "config 시트에서는 실행할 수 없습니다."
    If includeWebSetup Then
        EnsureTaskSheetLayout ws
    Else
        On Error Resume Next
        Set config = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
        On Error GoTo EH
        If config Is Nothing Or Not HasTaskWebHeader(ws) Or _
           CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2) <> "요청 내용" Then _
            Err.Raise vbObjectError + 2120, , "모두 새로고침을 먼저 실행해 주세요."
    End If
    detailColumnHidden = CaptureTaskDetailColumnHidden(ws)
    If Not ActiveWindow Is Nothing Then
        detailScrollColumn = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn
        detailScrollRow = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow
    End If
    taskSheetWasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    originalLastRow = GetLastDataRow(ws)
    Set originalRowView = CaptureTaskRowView(ws, originalLastRow)
    taskViewCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    UnprotectTaskSheet ws
    If Not includeWebSetup And CStr(ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2) <> "WEB UPDATE" Then _
        ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2 = "WEB UPDATE"
    ' 2026-10-22 김좌현 수정 : 삭제한 웹등록 이력의 연결과 잠금을 두 새로고침에 반영
    RefreshWebHistoryState ws
    If includeWebSetup Then
        EnsureConfigSheet
        SetupDataHeaders ws
        UpgradeGanttRefreshButtons ws
        webBasisIssues = WebTaskBasisIssues(ws)
        PrepareWebTaskSheet ws
        NormalizeGanttPptArtifactNames ws
        ApplyTaskInputValidation ws
    End If
    LoadHolidaySettings holidayDict, workdayDict
    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then Err.Raise vbObjectError + 2120, , "데이터가 없습니다."
    Set rowView = CaptureTaskRowView(ws, lastRow)
    SynchronizeTaskHierarchyModules ws, lastRow, showCompletionMessage
    If showCompletionMessage Then ShowTaskInputErrorReasons ws, lastRow
    UpdateTaskNumbers ws, lastRow

    ' Calculate once. The old DrawTaskBars painted every cell before the final renderer replaced it.
    ClearCalculatedArea ws, lastRow, includeWebSetup
    ApplyHierarchySummaryValues ws, lastRow, holidayDict, workdayDict
    UpdateWeeklyReportStatuses ws, lastRow
    If Not GetMinMaxDate(ws, lastRow, minDate, maxDate) Then _
        Err.Raise vbObjectError + 2120, , "시작일/종료일 데이터가 없습니다."
    chartStartDate = minDate: chartEndDate = maxDate
    TryGetDisplayDateRange chartStartDate, chartEndDate
    If mTaskLayoutSignatures Is Nothing Then Set mTaskLayoutSignatures = CreateObject("Scripting.Dictionary")
    If mTaskStructureSignatures Is Nothing Then Set mTaskStructureSignatures = CreateObject("Scripting.Dictionary")
    layoutSignature = TaskRefreshSignature(ws, lastRow, False)
    structureSignature = TaskRefreshSignature(ws, lastRow, True)
    fitLayout = True: rebuildHierarchy = True
    If Not includeWebSetup Then
        If mTaskLayoutSignatures.Exists(ws.CodeName) Then _
            fitLayout = (mTaskLayoutSignatures(ws.CodeName) <> layoutSignature)
        If mTaskStructureSignatures.Exists(ws.CodeName) Then _
            rebuildHierarchy = (mTaskStructureSignatures(ws.CodeName) <> structureSignature)
    End If
    If mTaskVisualSignatures Is Nothing Then Set mTaskVisualSignatures = CreateObject("Scripting.Dictionary")
    visualSignature = TaskVisualSignature(ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict)
    redrawChart = True
    If Not includeWebSetup And mTaskVisualSignatures.Exists(ws.CodeName) Then _
        redrawChart = (mTaskVisualSignatures(ws.CodeName) <> visualSignature)
    If Not GanttDateHeaderMatches(ws, chartStartDate, chartEndDate) Then redrawChart = True
    ShowAllTaskRows ws, lastRow
    If redrawChart Then
        ClearGanttArea ws
        DrawDateHeader ws, chartStartDate, chartEndDate, holidayDict, workdayDict
        FormatBaseArea ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict, fitLayout, rebuildHierarchy
    End If
    ShowAllDateColumns ws
    RestoreTaskRowView ws, lastRow, rowView
    ApplyCompletedTaskVisibility ws, lastRow
    RestoreTaskDetailColumnHidden ws, detailColumnHidden
    If detailScrollColumn > 0 And Not ActiveWindow Is Nothing Then
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = detailScrollColumn
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow = detailScrollRow
    End If
    ApplyCalculatedColumnsProtection ws, lastRow
    RefreshWebModeView ws, True
    mTaskLayoutSignatures(ws.CodeName) = layoutSignature
    mTaskStructureSignatures(ws.CodeName) = structureSignature
    mTaskVisualSignatures(ws.CodeName) = visualSignature
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    AlignAllGanttPptArtifactsToNoteCells ws
    If showCompletionMessage Then
        If includeWebSetup Then
            If Len(webBasisIssues) > 0 Then
                MsgBox "모두 새로고침 완료" & vbCrLf & vbCrLf & _
                       "웹 등록 기준에 문제가 있습니다. 해당 행을 수정한 뒤 웹등록 미리보기를 실행하세요." & _
                       vbCrLf & webBasisIssues, vbExclamation, "모두 새로고침"
            Else
                MsgBox "모두 새로고침 완료", vbInformation
            End If
        Else
            MsgBox "간트 새로고침 완료", vbInformation
        End If
    End If
    Exit Sub
EH:
    errorText = Err.Description
    On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If taskViewCaptured Then
        UnprotectTaskSheet ws
        RestoreTaskRowViewSnapshot ws, originalLastRow, originalRowView
        RestoreTaskDetailColumnHidden ws, detailColumnHidden
        If detailScrollColumn > 0 And Not ActiveWindow Is Nothing Then
            ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = detailScrollColumn
            ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow = detailScrollRow
        End If
        If taskSheetWasProtected Then
            ApplyCalculatedColumnsProtection ws, Application.Max(DATA_START_ROW, originalLastRow)
        Else
            UnprotectTaskSheet ws
        End If
        RefreshWebModeView ws, True
    End If
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    If showCompletionMessage Then
        MsgBox "오류가 발생했습니다: " & errorText, vbExclamation
    Else
        Err.Raise vbObjectError + 2121, "RefreshGanttCore", errorText
    End If
End Sub

Public Sub 칸트차트_완료숨김()
    On Error GoTo Failed
    ToggleCompletedTaskFilter ActiveSheet
    Exit Sub
Failed:
    MsgBox "완료 항목 표시를 변경할 수 없습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_초기화()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(ws) Then
        MsgBox "config 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws

    If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW

    ClearCompletedTaskFilter ws, lastRow
    ShowAllTaskRows ws, lastRow
    ClearCalculatedArea ws, lastRow
    ClearGanttArea ws
    ShowAllDateColumns ws
    ApplyCalculatedColumnsProtection ws, lastRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "차트 초기화 완료", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_새로고침버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttRefresh", "모두 새로고침", "칸트차트_모두새로고침", 1
    EnsureMainActionButtons ActiveSheet
End Sub

Public Sub 칸트차트_초기화버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttReset", "초기화", "칸트차트_초기화", 2
End Sub

Public Sub 칸트차트_개체삽입버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttObjectInsert", "개체삽입", "칸트차트_개체삽입", 3
End Sub

Public Sub 칸트차트_개체삽입()
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim oleObj As OLEObject
    Dim oleName As String
    Dim lastRow As Long
    Dim wasProtected As Boolean

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(ws) Then
        MsgBox "config 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "데이터 행의 셀을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If

    If ActiveCell.Row < DATA_START_ROW Then
        MsgBox "데이터 행의 셀을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If

    targetRow = ActiveCell.Row
    oleName = GetGanttPptOleName(targetRow)

    If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
    If HasEmbeddedPptObject(ws, targetRow) Then
        MsgBox "해당 행에는 이미 PPT 개체가 삽입되어 있습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If wasProtected Then
        UnprotectTaskSheet ws
    End If

    RemoveEmbeddedPptArtifacts ws, targetRow

    Set oleObj = ws.OLEObjects.Add( _
        ClassType:="PowerPoint.Show.12", _
        Link:=False, _
        DisplayAsIcon:=True, _
        IconLabel:="", _
        Left:=0, _
        Top:=0, _
        Width:=1, _
        Height:=1)

    With oleObj
        .Name = oleName
        .Placement = xlMoveAndSize
        .PrintObject = False
        .Width = 1
        .Height = 1
        .Visible = False
    End With

    PositionHiddenGanttPptOle ws, oleObj, targetRow
    SetGanttPptCellIcon ws, targetRow

    If wasProtected Then
        If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "PPT 개체 삽입 완료 (비고 셀의 빨간 사각형을 클릭하세요.)", vbInformation
    Exit Sub

EH:
    On Error Resume Next
    If Not ws Is Nothing And targetRow >= DATA_START_ROW Then
        RemoveEmbeddedPptArtifacts ws, targetRow
    End If
    If wasProtected Then
        If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If
    On Error GoTo 0

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_PPT셀아이콘열기(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim oleObj As OLEObject
    Dim oleName As String

    On Error GoTo EH

    If ws Is Nothing Then Exit Sub
    If targetRow < DATA_START_ROW Then Exit Sub

    oleName = GetGanttPptOleName(targetRow)

    If Not OLEObjectExistsOnSheet(ws, oleName) Then
        ClearGanttPptCellIcon ws, targetRow
        MsgBox "Excel에 포함된 PPT 개체를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Set oleObj = ws.OLEObjects(oleName)
    oleObj.Visible = False
    oleObj.Verb Verb:=xlOpen
    Exit Sub

EH:
    MsgBox "PPT 편집 화면을 여는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_PPT개체열기()
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim callerName As String

    On Error GoTo EH

    Set ws = ActiveSheet
    callerName = CStr(Application.Caller)

    If ShapeExistsOnSheet(ws, callerName) Then
        targetRow = ws.Shapes(callerName).TopLeftCell.Row
    Else
        targetRow = ActiveCell.Row
    End If

    칸트차트_PPT셀아이콘열기 ws, targetRow
    Exit Sub

EH:
    MsgBox "PPT 편집 화면을 여는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub
Private Sub CreateGanttActionButton(ByVal ws As Worksheet, _
                                    ByVal buttonName As String, _
                                    ByVal buttonCaption As String, _
                                    ByVal macroName As String, _
                                    ByVal buttonOrder As Long, _
                                    Optional ByVal showCompletionMessage As Boolean = True)
    Dim btn As Shape
    Dim anchorCell As Range
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim btnGap As Double
    Dim wasProtected As Boolean
    Dim lastRow As Long

    On Error GoTo EH

    If ws Is Nothing Then Exit Sub

    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(ws) Then
        MsgBox "config 시트에는 버튼을 생성할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    wasProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    If wasProtected Then
        UnprotectTaskSheet ws
    End If

    Set anchorCell = ws.Range("B2")

    btnWidth = 72
    btnHeight = 22
    btnGap = 8

    btnLeft = anchorCell.Left + ((buttonOrder - 1) * (btnWidth + btnGap))
    btnTop = anchorCell.Top + anchorCell.Height - btnHeight

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)

    With btn
        .Name = buttonName
        .OnAction = macroName
        .Placement = xlFreeFloating

        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(212, 208, 200)
        .Fill.Transparency = 0

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(128, 128, 128)
        .Line.Weight = 1

        .Shadow.Visible = msoFalse
        .Adjustments.Item(1) = 0.05

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 1
            .MarginBottom = 1

            With .TextRange
                .Characters.Text = buttonCaption
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "굴림"
                .Font.Size = 9
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With

        .Glow.Radius = 0
        .SoftEdge.Radius = 0
        .Rotation = 0
    End With

    If wasProtected Then
        If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If

    If showCompletionMessage Then
        MsgBox "버튼 생성 완료: " & buttonCaption, vbInformation
    End If
    Exit Sub

EH:
    If wasProtected Then
        On Error Resume Next
        If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
        On Error GoTo 0
    End If
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function GetGanttPptCellIconText() As String
    GetGanttPptCellIconText = ChrW(&H25A0)
End Function

Private Function GetGanttPptShapeName(ByVal targetRow As Long) As String
    GetGanttPptShapeName = "shpGanttPpt_" & CStr(targetRow)
End Function

Private Function GetGanttPptOleName(ByVal targetRow As Long) As String
    GetGanttPptOleName = "oleGanttPpt_" & CStr(targetRow)
End Function

Private Function ShapeExistsOnSheet(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    ShapeExistsOnSheet = Not ws.Shapes(shapeName) Is Nothing
    On Error GoTo 0
End Function

Private Function OLEObjectExistsOnSheet(ByVal ws As Worksheet, ByVal oleName As String) As Boolean
    On Error Resume Next
    OLEObjectExistsOnSheet = Not ws.OLEObjects(oleName) Is Nothing
    On Error GoTo 0
End Function

Private Function GetGanttPptNoteCell(ByVal ws As Worksheet, ByVal targetRow As Long) As Range
    Set GetGanttPptNoteCell = ws.Cells(targetRow, ws.Range(COL_NOTE & "1").Column)
End Function

Private Sub ApplyGanttPptNoteCellBorder(ByVal noteCell As Range)
    With noteCell.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
End Sub

Private Sub SetGanttPptCellIcon(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim noteCell As Range
    Dim iconFontSize As Double
    Dim subAddress As String

    Set noteCell = GetGanttPptNoteCell(ws, targetRow)
    subAddress = "'" & Replace(ws.Name, "'", "''") & "'!" & noteCell.Address

    On Error Resume Next
    noteCell.Hyperlinks.Delete
    On Error GoTo 0

    noteCell.Value = GetGanttPptCellIconText()
    ws.Hyperlinks.Add Anchor:=noteCell, _
                      Address:="", _
                      subAddress:=subAddress, _
                      ScreenTip:="PPT 편집 화면 열기", _
                      TextToDisplay:=GetGanttPptCellIconText()

    iconFontSize = noteCell.Height * 0.72
    If iconFontSize < 9 Then iconFontSize = 9
    If iconFontSize > 16 Then iconFontSize = 16

    With noteCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .ShrinkToFit = False
        .Font.Name = "Arial"
        .Font.Size = iconFontSize
        .Font.Bold = False
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 0, 0)
        .Font.Underline = xlUnderlineStyleNone
    End With
    ApplyGanttPptNoteCellBorder noteCell
End Sub

Private Sub ClearGanttPptCellIcon(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim noteCell As Range

    Set noteCell = GetGanttPptNoteCell(ws, targetRow)

    If CStr(noteCell.Value2) = GetGanttPptCellIconText() Then
        On Error Resume Next
        noteCell.Hyperlinks.Delete
        On Error GoTo 0

        noteCell.ClearContents
        noteCell.HorizontalAlignment = xlGeneral
        noteCell.VerticalAlignment = xlCenter
        noteCell.Interior.Pattern = xlNone
        noteCell.Font.ColorIndex = xlAutomatic
        noteCell.Font.Underline = xlUnderlineStyleNone
    End If
    ApplyGanttPptNoteCellBorder noteCell
End Sub

Private Sub PositionHiddenGanttPptOle(ByVal ws As Worksheet, ByVal oleObj As OLEObject, ByVal targetRow As Long)
    Dim noteCell As Range

    Set noteCell = GetGanttPptNoteCell(ws, targetRow)

    With oleObj
        .Placement = xlMoveAndSize
        .Left = noteCell.Left + 1
        .Top = noteCell.Top + 1
        .Width = 1
        .Height = 1
        .Visible = False
    End With
End Sub

Private Function HasEmbeddedPptObject(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    Dim oleName As String

    oleName = GetGanttPptOleName(targetRow)
    HasEmbeddedPptObject = OLEObjectExistsOnSheet(ws, oleName)

    If HasEmbeddedPptObject Then
        SetGanttPptCellIcon ws, targetRow
    Else
        ClearGanttPptCellIcon ws, targetRow
    End If
End Function

Private Sub RemoveEmbeddedPptArtifacts(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim shapeName As String
    Dim oleName As String

    shapeName = GetGanttPptShapeName(targetRow)
    oleName = GetGanttPptOleName(targetRow)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    ws.OLEObjects(oleName).Delete
    On Error GoTo 0

    ClearGanttPptCellIcon ws, targetRow
End Sub

Private Sub DeleteLegacyGanttPptShapes(ByVal ws As Worksheet)
    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1
        If IsGanttPptShape(ws.Shapes(i)) Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Private Sub ClearOrphanGanttPptCellIcons(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim noteCell As Range

    If CStr(ws.Range(COL_MODIFICATION & HEADER_ROW).Value2) <> "수정 내용" Then EnsureTaskSheetLayout ws
    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then Exit Sub

    For r = DATA_START_ROW To lastRow
        Set noteCell = GetGanttPptNoteCell(ws, r)
        If CStr(noteCell.Value2) = GetGanttPptCellIconText() Then
            If Not OLEObjectExistsOnSheet(ws, GetGanttPptOleName(r)) Then
                ClearGanttPptCellIcon ws, r
            End If
        End If
    Next r
End Sub

Private Sub AlignAllGanttPptArtifactsToNoteCells(ByVal ws As Worksheet)
    Dim oleObj As OLEObject
    Dim targetRow As Long

    DeleteLegacyGanttPptShapes ws

    For Each oleObj In ws.OLEObjects
        If Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Then
            targetRow = CLng(Val(Mid$(oleObj.Name, Len("oleGanttPpt_") + 1)))
            If targetRow >= DATA_START_ROW Then
                PositionHiddenGanttPptOle ws, oleObj, targetRow
                SetGanttPptCellIcon ws, targetRow
            End If
        End If
    Next oleObj

    ClearOrphanGanttPptCellIcons ws
End Sub

Private Sub NormalizeGanttPptArtifactNames(ByVal ws As Worksheet)
    Dim oleObj As OLEObject
    Dim targetRow As Long
    Dim finalName As String
    Dim tempIndex As Long

    tempIndex = 1

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            oleObj.Name = "tmpGanttPptOle_" & Format$(Now, "hhnnss") & "_" & CStr(tempIndex)
            tempIndex = tempIndex + 1
        End If
    Next oleObj

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            targetRow = oleObj.TopLeftCell.Row
            If targetRow >= DATA_START_ROW Then
                finalName = GetGanttPptOleName(targetRow)
                If Not OLEObjectExistsOnSheet(ws, finalName) Then
                    oleObj.Name = finalName
                    PositionHiddenGanttPptOle ws, oleObj, targetRow
                    SetGanttPptCellIcon ws, targetRow
                End If
            End If
        End If
    Next oleObj

    DeleteLegacyGanttPptShapes ws
    ClearOrphanGanttPptCellIcons ws
End Sub

Private Function IsGanttPptShape(ByVal shp As Shape) As Boolean
    If Left$(shp.Name, Len("shpGanttPpt_")) = "shpGanttPpt_" Or _
       Left$(shp.Name, Len("tmpGanttPptShape_")) = "tmpGanttPptShape_" Then
        IsGanttPptShape = True
        Exit Function
    End If

    ' Some embedded/OLE shapes raise error 1004 when OnAction is queried.
    On Error Resume Next
    IsGanttPptShape = (CStr(shp.OnAction) = "칸트차트_PPT개체열기")
    On Error GoTo 0
End Function

Private Function IsGanttPptOleObject(ByVal oleObj As OLEObject) As Boolean
    IsGanttPptOleObject = (Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Or _
                           Left$(oleObj.Name, Len("tmpGanttPptOle_")) = "tmpGanttPptOle_")
End Function
