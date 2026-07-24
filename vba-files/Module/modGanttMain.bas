Attribute VB_Name = "modGanttMain"
Option Explicit

Public Sub 칸트차트_생성()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim minDate As Date
    Dim maxDate As Date
    Dim chartStartDate As Date
    Dim chartEndDate As Date
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim hasDisplayRange As Boolean

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Then
        MsgBox "config 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    EnsureConfigSheet
    LoadHolidaySettings holidayDict, workdayDict

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then
        MsgBox "데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    If Not GetMinMaxDate(ws, lastRow, minDate, maxDate) Then
        MsgBox "시작일/종료일 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    chartStartDate = minDate
    chartEndDate = maxDate

    hasDisplayRange = TryGetDisplayDateRange(chartStartDate, chartEndDate)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws
    SetupDataHeaders ws
    UpdateTaskNumbers ws, lastRow
    NormalizeGanttPptArtifactNames ws
    ApplyTaskInputValidation ws
    ShowAllTaskRows ws, lastRow
    ClearCalculatedArea ws, lastRow
    ClearGanttArea ws
    DrawDateHeader ws, chartStartDate, chartEndDate, holidayDict, workdayDict
    DrawTaskBars ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict
    FormatBaseArea ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict
    ShowAllDateColumns ws
    ApplyDisplayTaskRowFilter ws, lastRow, chartStartDate, chartEndDate
    ApplyCalculatedColumnsProtection ws, lastRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    DoEvents
    AlignAllGanttPptArtifactsToNoteCells ws

    MsgBox "칸트차트 생성 완료", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_새로고침()
    칸트차트_생성
End Sub

Public Sub 칸트차트_초기화()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Then
        MsgBox "config 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW

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

Public Sub 칸트차트_항목숨기기()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = CONFIG_SHEET_NAME Then
        MsgBox "config 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    EnsureConfigSheet

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then
        MsgBox "데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws
    UpdateTaskNumbers ws, lastRow
    HideCompletedTaskRows ws, lastRow
    ApplyCalculatedColumnsProtection ws, lastRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "항목 숨김 완료", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_생성버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttCreate", "생성", "칸트차트_생성", 1
End Sub

Public Sub 칸트차트_새로고침버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttRefresh", "새로고침", "칸트차트_새로고침", 2
End Sub

Public Sub 칸트차트_초기화버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttReset", "초기화", "칸트차트_초기화", 3
End Sub

Public Sub 칸트차트_항목숨기기버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttHideTask", "항목 숨김", "칸트차트_항목숨기기", 4
End Sub

Public Sub 개발진행보고_버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnDevProgressReport", "개발 보고", "개발진행보고_텍스트생성", 6
End Sub

Public Sub 칸트차트_개체삽입버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttObjectInsert", "개체삽입", "칸트차트_개체삽입", 5
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

    If ws.Name = CONFIG_SHEET_NAME Then
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
Private Sub CreateGanttActionButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal buttonCaption As String, ByVal macroName As String, ByVal buttonOrder As Long)
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

    If ws.Name = CONFIG_SHEET_NAME Then
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
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If

    MsgBox "버튼 생성 완료: " & buttonCaption, vbInformation
    Exit Sub

EH:
    If wasProtected Then
        On Error Resume Next
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
    IsGanttPptShape = (Left$(shp.Name, Len("shpGanttPpt_")) = "shpGanttPpt_" Or _
                       Left$(shp.Name, Len("tmpGanttPptShape_")) = "tmpGanttPptShape_" Or _
                       shp.OnAction = "칸트차트_PPT개체열기")
End Function

Private Function IsGanttPptOleObject(ByVal oleObj As OLEObject) As Boolean
    IsGanttPptOleObject = (Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Or _
                           Left$(oleObj.Name, Len("tmpGanttPptOle_")) = "tmpGanttPptOle_")
End Function
