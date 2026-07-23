Attribute VB_Name = "modGanttMain"
Option Explicit

Private mRefreshPptFileNamesOnCreate As Boolean

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

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    EnsureHolidaySheet
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
    If mRefreshPptFileNamesOnCreate Then
        RefreshGanttPptFileNames ws, lastRow
    End If
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

    MsgBox "칸트차트 생성 완료", vbInformation
    Exit Sub

EH:
    mRefreshPptFileNamesOnCreate = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_새로고침()
    mRefreshPptFileNamesOnCreate = True
    칸트차트_생성
    mRefreshPptFileNamesOnCreate = False
End Sub

Public Sub 칸트차트_초기화()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
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

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    EnsureHolidaySheet

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

Public Sub 칸트차트_기간숨기기()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    EnsureHolidaySheet

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then
        MsgBox "데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws
    HideIdleDateColumns ws, lastRow
    ApplyCalculatedColumnsProtection ws, lastRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "기간 숨김 완료", vbInformation
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

Public Sub 칸트차트_기간숨기기버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttHidePeriod", "기간 숨김", "칸트차트_기간숨기기", 5
End Sub

Public Sub 칸트차트_행추가()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim insertedRow As Long

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    If ActiveCell.Row < DATA_START_ROW Then
        MsgBox "데이터 행을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws
    insertedRow = InsertBlankTaskRowAbove(ws, ActiveCell.Row)
    lastRow = Application.Max(GetLastDataRow(ws), insertedRow)
    UpdateTaskNumbers ws, lastRow
    NormalizeGanttPptArtifactNames ws
    RefreshGanttPptFileNames ws, lastRow
    ApplyTaskInputValidation ws
    ApplyCalculatedColumnsProtection ws, lastRow

    ws.Cells(insertedRow, ws.Range(COL_LEVEL & "1").Column).Select

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "행 추가 완료", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 칸트차트_행삭제()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNo As Long
    Dim deleteRows As Object
    Dim deleteCount As Long
    Dim firstDeleteRow As Long
    Dim nextSelectRow As Long
    Dim answer As VbMsgBoxResult

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then
        MsgBox "데이터 행이 없습니다.", vbExclamation
        Exit Sub
    End If

    Set deleteRows = GetSelectedTaskRowMap(ws, lastRow)
    If deleteRows.Count = 0 Then
        MsgBox "삭제할 데이터 행을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox("선택한 " & deleteRows.Count & "개 행을 삭제하시겠습니까?", vbQuestion + vbYesNo, "행삭제")
    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    UnprotectTaskSheet ws

    For rowNo = lastRow To DATA_START_ROW Step -1
        If deleteRows.Exists(CStr(rowNo)) Then
            RemoveEmbeddedPptArtifacts ws, rowNo
            DeleteTaskRow ws, rowNo
            deleteCount = deleteCount + 1
            If firstDeleteRow = 0 Or rowNo < firstDeleteRow Then firstDeleteRow = rowNo
        End If
    Next rowNo

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW

    UpdateTaskNumbers ws, lastRow
    NormalizeGanttPptArtifactNames ws
    RefreshGanttPptFileNames ws, lastRow
    ApplyTaskInputValidation ws
    ApplyCalculatedColumnsProtection ws, lastRow

    nextSelectRow = firstDeleteRow
    If nextSelectRow > lastRow Then nextSelectRow = lastRow
    If nextSelectRow < DATA_START_ROW Then nextSelectRow = DATA_START_ROW
    ws.Cells(nextSelectRow, ws.Range(COL_LEVEL & "1").Column).Select

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox deleteCount & "개 행 삭제 완료", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function GetSelectedTaskRowMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim rowMap As Object
    Dim sourceRange As Range
    Dim selectedRange As Range
    Dim area As Range
    Dim rowNo As Long

    Set rowMap = CreateObject("Scripting.Dictionary")
    Set GetSelectedTaskRowMap = rowMap

    Set sourceRange = GetTaskDeleteSourceRange(ws)
    If sourceRange Is Nothing Then Exit Function

    Set selectedRange = Intersect(sourceRange.EntireRow, ws.Rows(CStr(DATA_START_ROW) & ":" & CStr(lastRow)))
    If selectedRange Is Nothing Then Exit Function

    For Each area In selectedRange.Areas
        For rowNo = area.Row To area.Row + area.Rows.Count - 1
            If Not rowMap.Exists(CStr(rowNo)) Then
                rowMap.Add CStr(rowNo), rowNo
            End If
        Next rowNo
    Next area
End Function

Private Function GetTaskDeleteSourceRange(ByVal ws As Worksheet) As Range
    Dim candidateRange As Range

    On Error Resume Next

    If TypeName(Selection) = "Range" Then
        Set candidateRange = Selection
        If Not candidateRange Is Nothing Then
            If candidateRange.Worksheet.Name = ws.Name Then
                Set GetTaskDeleteSourceRange = candidateRange
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If

    Set candidateRange = Nothing
    Set candidateRange = ActiveWindow.RangeSelection
    If Not candidateRange Is Nothing Then
        If candidateRange.Worksheet.Name = ws.Name Then
            Set GetTaskDeleteSourceRange = candidateRange
        End If
    End If

    On Error GoTo 0
End Function

Public Sub 칸트차트_행추가버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttRowAdd", "행추가", "칸트차트_행추가", 6
End Sub

Public Sub 칸트차트_행삭제버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttRowDelete", "행삭제", "칸트차트_행삭제", 7
End Sub

Public Sub 칸트차트_개체삽입버튼_생성()
    CreateGanttActionButton ActiveSheet, "btnGanttObjectInsert", "개체삽입", "칸트차트_개체삽입", 8
End Sub

Public Sub 칸트차트_개체삽입()
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim noteCell As Range
    Dim squareSize As Double
    Dim shapeLeft As Double
    Dim shapeTop As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim innerPad As Double
    Dim shp As Shape
    Dim shapeName As String
    Dim pptPath As String
    Dim pptFileName As String
    Dim lastRow As Long
    Dim wasProtected As Boolean

    On Error GoTo EH

    Set ws = ActiveSheet

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에서는 실행할 수 없습니다.", vbExclamation
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
    shapeName = GetGanttPptShapeName(targetRow)

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

    pptFileName = GetGanttPptFileName(ws, targetRow)
    If Len(pptFileName) = 0 Then Err.Raise vbObjectError + 513, , "PPT 파일명 생성에 실패했습니다."

    pptPath = GetGanttPptFilePathByFileName(pptFileName)
    If Len(pptPath) = 0 Then Err.Raise vbObjectError + 514, , "엑셀 파일을 먼저 저장한 뒤 실행하세요."

    If Not CreateBlankPptFile(pptPath) Then Err.Raise vbObjectError + 515, , "빈 PPT 파일 생성에 실패했습니다."

    Set noteCell = ws.Cells(targetRow, ws.Range(COL_NOTE & "1").Column)

    squareSize = noteCell.Height
    If squareSize > noteCell.Width Then squareSize = noteCell.Width
    If squareSize < 6 Then squareSize = 6

    innerPad = 1.2

    shapeLeft = noteCell.Left + innerPad
    shapeTop = noteCell.Top + innerPad
    shapeWidth = squareSize - (innerPad * 2)
    shapeHeight = squareSize - (innerPad * 2)

    If shapeWidth < 1 Then shapeWidth = 1
    If shapeHeight < 1 Then shapeHeight = 1

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)

    With shp
        .Name = shapeName
        .OnAction = "칸트차트_PPT개체열기"
        .Placement = xlMoveAndSize
        .AlternativeText = pptFileName

        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Fill.Transparency = 0

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(192, 0, 0)
        .Line.Weight = 0.75

        .TextFrame2.TextRange.Characters.Text = ""
        .LockAspectRatio = msoFalse
    End With

    If wasProtected Then
        lastRow = GetLastDataRow(ws)
        If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
        ApplyCalculatedColumnsProtection ws, lastRow
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "PPT 개체 삽입 완료", vbInformation
    Exit Sub

EH:
    On Error Resume Next
    If Len(pptPath) > 0 Then DeleteFileIfExists pptPath
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

Public Sub 칸트차트_PPT개체열기()
    Dim ws As Worksheet
    Dim shpName As String
    Dim pptFileName As String
    Dim pptPath As String

    On Error GoTo EH

    shpName = CStr(Application.Caller)
    Set ws = ActiveSheet

    If Not ShapeExistsOnSheet(ws, shpName) Then
        MsgBox "도형을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    pptFileName = Trim$(ws.Shapes(shpName).AlternativeText)
    If Len(pptFileName) = 0 Then
        MsgBox "연결된 PPT 파일 정보가 없습니다.", vbExclamation
        Exit Sub
    End If

    pptPath = GetGanttPptFilePathByFileName(pptFileName)
    If Len(pptPath) = 0 Then
        MsgBox "엑셀 파일을 먼저 저장한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If

    If FileExists(pptPath) Then
        OpenPowerPointFile pptPath
        Exit Sub
    End If

    MsgBox "연결된 PPT 파일을 찾을 수 없습니다." & vbCrLf & pptPath, vbExclamation
    Exit Sub

EH:
    MsgBox "오류가 발생했습니다: " & Err.Description, vbExclamation
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

    If ws.Name = HOLIDAY_SHEET_NAME Then
        MsgBox "휴일설정 시트에는 버튼을 생성할 수 없습니다.", vbExclamation
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

Private Function HasEmbeddedPptObject(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    Dim shapeName As String
    Dim oleName As String
    Dim pptFileName As String
    Dim pptPath As String
    Dim hasShape As Boolean
    Dim hasOle As Boolean

    shapeName = GetGanttPptShapeName(targetRow)
    oleName = GetGanttPptOleName(targetRow)

    hasShape = ShapeExistsOnSheet(ws, shapeName)
    hasOle = OLEObjectExistsOnSheet(ws, oleName)

    If hasShape Then
        pptFileName = Trim$(ws.Shapes(shapeName).AlternativeText)
        pptPath = GetGanttPptFilePathByFileName(pptFileName)

        If FileExists(pptPath) Then
            HasEmbeddedPptObject = True
            Exit Function
        End If
    End If

    If hasShape And hasOle Then
        HasEmbeddedPptObject = True
    Else
        If hasShape Or hasOle Then
            RemoveEmbeddedPptArtifacts ws, targetRow
        End If
        HasEmbeddedPptObject = False
    End If
End Function

Private Sub RemoveEmbeddedPptArtifacts(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim shapeName As String
    Dim oleName As String
    Dim pptFileName As String
    Dim pptPath As String

    shapeName = GetGanttPptShapeName(targetRow)
    oleName = GetGanttPptOleName(targetRow)

    On Error Resume Next
    If ShapeExistsOnSheet(ws, shapeName) Then
        pptFileName = Trim$(ws.Shapes(shapeName).AlternativeText)
        pptPath = GetGanttPptFilePathByFileName(pptFileName)
    End If

    ws.Shapes(shapeName).Delete
    ws.OLEObjects(oleName).Delete

    If FileExists(pptPath) Then
        DeleteFileIfExists pptPath
    End If
    On Error GoTo 0
End Sub

Private Function CreateBlankPptFile(ByVal pptPath As String) As Boolean
    Dim ppApp As Object
    Dim ppPres As Object

    On Error GoTo EH

    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = True

    Set ppPres = ppApp.Presentations.Add

    Do While ppPres.Slides.Count > 0
        ppPres.Slides(1).Delete
    Loop

    ppPres.Slides.Add 1, 12
    ppPres.SaveAs pptPath, 24
    ppPres.Close
    ppApp.Quit

    Set ppPres = Nothing
    Set ppApp = Nothing

    CreateBlankPptFile = True
    Exit Function

EH:
    On Error Resume Next
    If Not ppPres Is Nothing Then ppPres.Close
    If Not ppApp Is Nothing Then ppApp.Quit
    Set ppPres = Nothing
    Set ppApp = Nothing
    CreateBlankPptFile = False
End Function

Private Sub NormalizeGanttPptArtifactNames(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim oleObj As OLEObject
    Dim targetRow As Long
    Dim finalName As String
    Dim tempIndex As Long

    tempIndex = 1

    For Each shp In ws.Shapes
        If IsGanttPptShape(shp) Then
            shp.Name = "tmpGanttPptShape_" & Format$(Now, "hhnnss") & "_" & CStr(tempIndex)
            tempIndex = tempIndex + 1
        End If
    Next shp

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            oleObj.Name = "tmpGanttPptOle_" & Format$(Now, "hhnnss") & "_" & CStr(tempIndex)
            tempIndex = tempIndex + 1
        End If
    Next oleObj

    For Each shp In ws.Shapes
        If IsGanttPptShape(shp) Then
            targetRow = shp.TopLeftCell.Row
            If targetRow >= DATA_START_ROW Then
                finalName = GetGanttPptShapeName(targetRow)
                If Not ShapeExistsOnSheet(ws, finalName) Then
                    shp.Name = finalName
                End If
            End If
        End If
    Next shp

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            targetRow = oleObj.TopLeftCell.Row
            If targetRow >= DATA_START_ROW Then
                finalName = GetGanttPptOleName(targetRow)
                If Not OLEObjectExistsOnSheet(ws, finalName) Then
                    oleObj.Name = finalName
                End If
            End If
        End If
    Next oleObj
End Sub

Private Function HasGanttPptArtifacts(ByVal ws As Worksheet) As Boolean
    Dim shp As Shape
    Dim oleObj As OLEObject

    For Each shp In ws.Shapes
        If IsGanttPptShape(shp) Then
            HasGanttPptArtifacts = True
            Exit Function
        End If
    Next shp

    For Each oleObj In ws.OLEObjects
        If IsGanttPptOleObject(oleObj) Then
            HasGanttPptArtifacts = True
            Exit Function
        End If
    Next oleObj
End Function

Private Function IsGanttPptShape(ByVal shp As Shape) As Boolean
    IsGanttPptShape = (Left$(shp.Name, Len("shpGanttPpt_")) = "shpGanttPpt_" Or _
                       Left$(shp.Name, Len("tmpGanttPptShape_")) = "tmpGanttPptShape_" Or _
                       shp.OnAction = "칸트차트_PPT개체열기")
End Function

Private Function IsGanttPptOleObject(ByVal oleObj As OLEObject) As Boolean
    IsGanttPptOleObject = (Left$(oleObj.Name, Len("oleGanttPpt_")) = "oleGanttPpt_" Or _
                           Left$(oleObj.Name, Len("tmpGanttPptOle_")) = "tmpGanttPptOle_")
End Function

Private Sub RefreshGanttPptFileNames(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim rowNo As Long
    Dim shapeName As String
    Dim oldFileName As String
    Dim oldPath As String
    Dim dateTimeText As String
    Dim newFileName As String
    Dim newPath As String
    Dim renamedCount As Long
    Dim skippedCount As Long
    Dim missingCount As Long
    Dim failedCount As Long

    On Error GoTo EH

    NormalizeGanttPptArtifactNames ws

    If Not HasGanttPptArtifacts(ws) Then Exit Sub

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "엑셀 파일을 먼저 저장한 뒤 새로고침하세요.", vbExclamation
        Exit Sub
    End If

    For rowNo = DATA_START_ROW To lastRow
        shapeName = GetGanttPptShapeName(rowNo)

        If ShapeExistsOnSheet(ws, shapeName) Then
            On Error GoTo RowEH

            oldFileName = Trim$(ws.Shapes(shapeName).AlternativeText)
            If Len(oldFileName) = 0 Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            oldPath = GetGanttPptFilePathByFileName(oldFileName)
            If Not FileExists(oldPath) Then
                missingCount = missingCount + 1
                GoTo NextRow
            End If

            dateTimeText = ExtractGanttPptDateTimeText(oldFileName)
            If Len(dateTimeText) = 0 Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            newFileName = GetGanttPptFileName(ws, rowNo, dateTimeText)
            If Len(newFileName) = 0 Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            If StrComp(oldFileName, newFileName, vbTextCompare) <> 0 Then
                newFileName = GetUniqueGanttPptFileName(newFileName, oldFileName)
                newPath = GetGanttPptFilePathByFileName(newFileName)

                Name oldPath As newPath
                ws.Shapes(shapeName).AlternativeText = newFileName
                renamedCount = renamedCount + 1
            End If
        End If

NextRow:
        On Error GoTo EH
    Next rowNo

    If skippedCount > 0 Or missingCount > 0 Or failedCount > 0 Then
        MsgBox "PPT 파일명 갱신 중 일부 항목은 처리하지 못했습니다." & vbCrLf & _
               "변경: " & renamedCount & "건" & vbCrLf & _
               "건너뜀: " & skippedCount & "건" & vbCrLf & _
               "파일 없음: " & missingCount & "건" & vbCrLf & _
               "실패: " & failedCount & "건", vbExclamation
    End If

    Exit Sub

RowEH:
    failedCount = failedCount + 1
    Resume NextRow

EH:
    MsgBox "PPT 파일명 갱신 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function ExtractGanttPptDateTimeText(ByVal pptFileName As String) As String
    Dim baseName As String
    Dim dotPos As Long
    Dim dateTimeText As String

    pptFileName = Trim$(pptFileName)
    dotPos = InStrRev(pptFileName, ".")

    If dotPos > 0 Then
        baseName = Left$(pptFileName, dotPos - 1)
    Else
        baseName = pptFileName
    End If

    If Len(baseName) < 15 Then Exit Function

    dateTimeText = Right$(baseName, 15)

    If Mid$(dateTimeText, 9, 1) <> "_" Then Exit Function
    If Not IsNumeric(Left$(dateTimeText, 8)) Then Exit Function
    If Not IsNumeric(Right$(dateTimeText, 6)) Then Exit Function

    ExtractGanttPptDateTimeText = dateTimeText
End Function

Private Function GetUniqueGanttPptFileName(ByVal desiredFileName As String, ByVal oldFileName As String) As String
    Dim folderPath As String
    Dim dotPos As Long
    Dim baseName As String
    Dim extName As String
    Dim candidateFileName As String
    Dim candidatePath As String
    Dim indexNo As Long

    folderPath = EnsureGanttPptFolder()
    If Len(folderPath) = 0 Then Exit Function

    dotPos = InStrRev(desiredFileName, ".")
    If dotPos > 0 Then
        baseName = Left$(desiredFileName, dotPos - 1)
        extName = Mid$(desiredFileName, dotPos)
    Else
        baseName = desiredFileName
        extName = ""
    End If

    candidateFileName = desiredFileName
    indexNo = 1

    Do
        candidatePath = folderPath & "\" & candidateFileName

        If StrComp(candidateFileName, oldFileName, vbTextCompare) = 0 Then
            GetUniqueGanttPptFileName = candidateFileName
            Exit Function
        End If

        If Not FileExists(candidatePath) Then
            GetUniqueGanttPptFileName = candidateFileName
            Exit Function
        End If

        candidateFileName = baseName & "_" & CStr(indexNo) & extName
        indexNo = indexNo + 1
    Loop
End Function

Private Function GetGanttPptFolderPath() As String
    Dim folderName As String
    Dim dotPos As Long

    If Len(ThisWorkbook.Path) = 0 Then Exit Function

    folderName = ThisWorkbook.Name
    dotPos = InStrRev(folderName, ".")
    If dotPos > 0 Then
        folderName = Left$(folderName, dotPos - 1)
    End If

    GetGanttPptFolderPath = ThisWorkbook.Path & "\" & folderName
End Function

Private Function EnsureGanttPptFolder() As String
    Dim folderPath As String

    folderPath = GetGanttPptFolderPath()
    If Len(folderPath) = 0 Then Exit Function

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If

    EnsureGanttPptFolder = folderPath
End Function

Private Function GetGanttPptFileName(ByVal ws As Worksheet, ByVal targetRow As Long, Optional ByVal dateTimeText As String = vbNullString) As String
    Dim taskNo As String
    Dim taskNoValue As String
    Dim taskLevel As String
    Dim taskTitle As String

    taskNoValue = Trim$(CStr(ws.Cells(targetRow, ws.Range(COL_NO & "1").Column).Value))
    If Len(taskNoValue) = 0 Then Exit Function
    If Not IsNumeric(taskNoValue) Then Exit Function

    taskNo = Format$(CLng(taskNoValue), "000")
    taskLevel = SanitizeFileName(CStr(ws.Cells(targetRow, ws.Range(COL_LEVEL & "1").Column).Value))
    taskTitle = SanitizeFileName(CStr(ws.Cells(targetRow, ws.Range(COL_TASK & "1").Column).Value))

    If Len(taskLevel) = 0 Then taskLevel = "Level"
    If Len(taskTitle) = 0 Then taskTitle = "NoTitle"
    If Len(dateTimeText) = 0 Then dateTimeText = Format$(Now, "yyyymmdd_hhnnss")

    GetGanttPptFileName = taskNo & "_" & taskLevel & "_" & taskTitle & "_" & dateTimeText & ".pptx"
End Function

Private Function GetGanttPptFilePathByFileName(ByVal pptFileName As String) As String
    Dim folderPath As String

    pptFileName = Trim$(pptFileName)
    If Len(pptFileName) = 0 Then Exit Function

    folderPath = EnsureGanttPptFolder()
    If Len(folderPath) = 0 Then Exit Function

    GetGanttPptFilePathByFileName = folderPath & "\" & pptFileName
End Function

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long

    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")

    fileName = Trim$(fileName)

    For i = LBound(invalidChars) To UBound(invalidChars)
        fileName = Replace(fileName, invalidChars(i), "_")
    Next i

    fileName = Replace(fileName, vbCr, "_")
    fileName = Replace(fileName, vbLf, "_")
    fileName = Replace(fileName, vbTab, "_")

    Do While InStr(fileName, "__") > 0
        fileName = Replace(fileName, "__", "_")
    Loop

    Do While Len(fileName) > 0 And (Right$(fileName, 1) = "." Or Right$(fileName, 1) = " ")
        fileName = Left$(fileName, Len(fileName) - 1)
    Loop

    SanitizeFileName = fileName
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    If Len(filePath) > 0 Then
        FileExists = (Len(Dir$(filePath)) > 0)
    End If
    On Error GoTo 0
End Function

Private Sub OpenPowerPointFile(ByVal pptPath As String)
    Dim ppApp As Object

    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    If ppApp Is Nothing Then
        Set ppApp = CreateObject("PowerPoint.Application")
    End If

    ppApp.Visible = True
    ppApp.Presentations.Open pptPath, False, False, True
End Sub

Private Sub DeleteFileIfExists(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub



