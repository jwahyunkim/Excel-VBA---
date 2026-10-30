Attribute VB_Name = "modGanttMain"
Option Explicit
' 2026-10-29 김좌현 수정 : 새로고침 단일화 및 미사용 캐시·버튼 생성·구형 PPT 진입점 제거
' 2026-10-29 김좌현 수정 : 비고 개체의 행 연결을 이름 범위로 보존하고 숨기기 전에 위치 정렬
' 2026-10-27 김좌현 수정 : 새로고침 후 등록 기준 문제 행을 완료 메시지에 안내
' 2026-10-23 김좌현 수정 : 간트 전용 새로고침도 이전 웹등록 헤더를 안전하게 WEB UPDATE로 갱신
' 2026-10-22 김좌현 수정 : 새로고침 서명 계산의 열 번호를 한 번만 조회해 반복 Range 접근 제거
' 2026-10-22 김좌현 수정 : 중복 새로고침 매크로를 제거하고 새로고침으로 버튼 연결 통일
' 2026-10-22 김좌현 수정 : 모두/간트 새로고침 분리, 중복 렌더 제거와 변경된 레이아웃만 재구성
' 2026-10-22 김좌현 수정 : 웹 유틸리티 시트 제외 및 웹등록 설정/버튼 초기화
' 2026-10-20 김좌현 수정 : 새로고침 실패 시 완료 숨김과 수동/계층 보기 및 보호 복원
' 2026-10-20 김좌현 수정 : 완료 숨김 버튼과 새로고침 시 행 보기 상태 유지
' 2026-10-20 김좌현 수정 : 미사용 간트 only/보고 only 행 필터 호출 제거
' 2026-10-20 김좌현 수정 : 미사용 항목 숨기기 매크로와 전용 버튼 생성 제거

' 2026-10-20 김좌현 수정 : 생성/새로고침을 새로고침 진입점으로 통합

Public Sub 칸트차트_새로고침()
    RefreshGanttSheet
End Sub

Public Sub RefreshGanttSheet(Optional ByVal showCompletionMessage As Boolean = True)
    Dim ws As Worksheet
    Dim lastRow As Long, minDate As Date, maxDate As Date
    Dim chartStartDate As Date, chartEndDate As Date
    Dim holidayDict As Object, workdayDict As Object
    Dim taskSheetWasProtected As Boolean, previousScreenUpdating As Boolean, previousEnableEvents As Boolean
    Dim errorText As String, detailColumnHidden As Variant, webBasisIssues As String
    Dim typeIssues As String, structureIssue As String, validationIssues As String
    Dim detailScrollColumn As Long, detailScrollRow As Long
    Dim rowView As Object, originalRowView As Object
    Dim originalLastRow As Long, taskViewCaptured As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    If Not TypeOf ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 2120, , "업무 시트에서 실행해 주세요."
    Set ws = ActiveSheet
    If ws.Name = CONFIG_SHEET_NAME Or ws.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(ws) Then _
        Err.Raise vbObjectError + 2120, , "config 시트에서는 실행할 수 없습니다."
    EnsureTaskSheetLayout ws
    EnsureConfigSheet
    If HasTaskWebHeader(ws) Then RefreshWebHistoryState ws
    ' 2026-10-27 김좌현 수정 : 새로고침이 기존 입력값을 바꾸기 전에 오류를 확인
    If HasTaskWebHeader(ws) Then
        typeIssues = TaskTypeIssues(ws)
        webBasisIssues = WebTaskBasisIssues(ws)
        structureIssue = WebTaskStructureIssue(ws)
        If Len(typeIssues) > 0 Then validationIssues = "타입 오류:" & vbCrLf & typeIssues
        If Len(webBasisIssues) > 0 Then validationIssues = validationIssues & vbCrLf & "등록 기준 오류:" & vbCrLf & webBasisIssues
        If Len(structureIssue) > 0 Then validationIssues = validationIssues & vbCrLf & "웹등록 구조 오류: " & structureIssue
        If Len(validationIssues) > 0 Then Err.Raise vbObjectError + 2121, "RefreshGanttSheet", validationIssues
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
    SetupDataHeaders ws
    UpgradeGanttRefreshButtons ws
    webBasisIssues = WebTaskBasisIssues(ws)
    PrepareWebTaskSheet ws
    ApplyTaskInputValidation ws
    NormalizeGanttPptArtifactNames ws
    LoadHolidaySettings holidayDict, workdayDict
    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then Err.Raise vbObjectError + 2120, , "데이터가 없습니다."
    Set rowView = CaptureTaskRowView(ws, lastRow)
    SynchronizeTaskHierarchyModules ws, lastRow, showCompletionMessage
    If showCompletionMessage Then ShowTaskInputErrorReasons ws, lastRow
    UpdateTaskNumbers ws, lastRow

    ' Calculate once. The old DrawTaskBars painted every cell before the final renderer replaced it.
    ClearCalculatedArea ws, lastRow
    ApplyHierarchySummaryValues ws, lastRow, holidayDict, workdayDict
    UpdateWeeklyReportStatuses ws, lastRow
    If Not GetMinMaxDate(ws, lastRow, minDate, maxDate) Then _
        Err.Raise vbObjectError + 2120, , "시작일/종료일 데이터가 없습니다."
    chartStartDate = minDate: chartEndDate = maxDate
    TryGetDisplayDateRange chartStartDate, chartEndDate
    ShowAllTaskRows ws, lastRow
    ClearGanttArea ws
    DrawDateHeader ws, chartStartDate, chartEndDate, holidayDict, workdayDict
    FormatBaseArea ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict
    ShowAllDateColumns ws
    AlignAllGanttPptArtifactsToNoteCells ws
    RestoreTaskRowView ws, lastRow, rowView
    ApplyCompletedTaskVisibility ws, lastRow
    RestoreTaskDetailColumnHidden ws, detailColumnHidden
    If detailScrollColumn > 0 And Not ActiveWindow Is Nothing Then
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = detailScrollColumn
        ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow = detailScrollRow
    End If
    ApplyCalculatedColumnsProtection ws, lastRow
    RefreshWebModeView ws, True
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    If showCompletionMessage Then
        If Len(webBasisIssues) > 0 Then
            MsgBox "새로고침 완료" & vbCrLf & vbCrLf & _
                   "웹 등록 기준에 문제가 있습니다. 해당 행을 수정한 뒤 웹등록 미리보기를 실행하세요." & _
                   vbCrLf & webBasisIssues, vbExclamation, "새로고침"
        Else
            MsgBox "새로고침 완료", vbInformation
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
        Err.Raise vbObjectError + 2121, "RefreshGanttSheet", errorText
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

Private Function GetGanttPptCellIconText() As String
    GetGanttPptCellIconText = ChrW(&H25A0)
End Function

Private Function GetGanttPptShapeName(ByVal targetRow As Long) As String
    GetGanttPptShapeName = "shpGanttPpt_" & CStr(targetRow)
End Function

Private Function GetGanttPptOleName(ByVal targetRow As Long) As String
    GetGanttPptOleName = "oleGanttPpt_" & CStr(targetRow)
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

' A local range follows inserted/deleted rows even when Excel reports the next visible cell.
Private Function GanttPptRowAnchorName(ByVal ws As Worksheet, ByVal oleObj As OLEObject) As String
    GanttPptRowAnchorName = "'" & Replace(ws.Name, "'", "''") & "'!_GanttPptRow_" & CStr(oleObj.ShapeRange.Item(1).ID)
End Function

Public Function GetGanttPptObjectRow(ByVal ws As Worksheet, ByVal oleObj As OLEObject) As Long
    Dim anchor As Name, cell As Range, namedRow As Long, suffix As String
    On Error Resume Next
    Set anchor = ws.Names(GanttPptRowAnchorName(ws, oleObj))
    If Not anchor Is Nothing Then Set cell = anchor.RefersToRange
    On Error GoTo 0
    If Not cell Is Nothing Then
        If cell.Parent Is ws Then
            GetGanttPptObjectRow = cell.Row
            Exit Function
        End If
    End If
    If Not anchor Is Nothing Then
        Err.Raise vbObjectError + 2130, "GetGanttPptObjectRow", "첨부 개체의 연결 행이 삭제되었습니다: " & oleObj.Name
    End If
    ' Upgrade legacy attachments by their managed row name, including already displaced objects.
    If Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Then
        suffix = Mid$(oleObj.Name, Len("oleGanttPpt_") + 1)
        If Len(suffix) > 0 And Not suffix Like "*[!0-9]*" Then
            namedRow = CLng(Val(suffix))
            If namedRow >= DATA_START_ROW And namedRow <= ws.Rows.Count Then
                GetGanttPptObjectRow = namedRow
                Exit Function
            End If
        End If
    End If
    GetGanttPptObjectRow = oleObj.TopLeftCell.Row
End Function

Public Sub BindGanttPptObjectRow(ByVal ws As Worksheet, ByVal oleObj As OLEObject, ByVal targetRow As Long)
    Dim anchorName As String, anchor As Name, cell As Range
    If targetRow < DATA_START_ROW Or targetRow > ws.Rows.Count Then _
        Err.Raise vbObjectError + 2131, "BindGanttPptObjectRow", "첨부 개체의 업무 행이 올바르지 않습니다."
    anchorName = GanttPptRowAnchorName(ws, oleObj)
    On Error Resume Next
    Set anchor = ws.Names(anchorName)
    If Not anchor Is Nothing Then Set cell = anchor.RefersToRange
    On Error GoTo 0
    If Not cell Is Nothing Then
        If cell.Parent Is ws Then
            If cell.Row = targetRow Then Exit Sub
        End If
    End If
    ws.Names.Add Name:=anchorName, _
        RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(targetRow, COL_NOTE).Address, Visible:=False
End Sub


Private Sub PositionHiddenGanttPptOle(ByVal ws As Worksheet, ByVal oleObj As OLEObject, ByVal targetRow As Long)
    Dim noteCell As Range

    BindGanttPptObjectRow ws, oleObj, targetRow
    If ws.Rows(targetRow).Hidden Then Exit Sub
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

Public Sub AlignAllGanttPptArtifactsToNoteCells(ByVal ws As Worksheet)
    Dim oleObj As OLEObject
    Dim targetRow As Long

    DeleteLegacyGanttPptShapes ws

    For Each oleObj In ws.OLEObjects
        If Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Then
            targetRow = GetGanttPptObjectRow(ws, oleObj)
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
            BindGanttPptObjectRow ws, oleObj, GetGanttPptObjectRow(ws, oleObj)
            oleObj.Name = "tmpGanttPptOle_" & Format$(Now, "hhnnss") & "_" & CStr(tempIndex)
            tempIndex = tempIndex + 1
        End If
    Next oleObj

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            targetRow = GetGanttPptObjectRow(ws, oleObj)
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
