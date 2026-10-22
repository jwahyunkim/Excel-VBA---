Attribute VB_Name = "modWebView"
Option Explicit
' 2026-10-22 김좌현 수정 : 모두 새로고침 버튼의 전용 매크로 연결 통일
' 2026-10-22 김좌현 수정 : 3단계 웹등록 모드의 버튼/헤더와 공통 작업 버튼 배치

Private Const WEB_MODE_NAME As String = "_GanttWebMode"
Private mWebModeCache As Object
Private mWebViewUpdating As Boolean

Private Function HasMainTaskHeaders(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If IsWebUtilitySheet(ws) Then Exit Function
    HasMainTaskHeaders = (CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "No." And _
                          CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Level")
End Function

Private Function FindActionShape(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set FindActionShape = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Sub ApplyActionStyle(ByVal button As Shape, ByVal leftValue As Double, _
                             ByVal topValue As Double, ByVal widthValue As Double, _
                             ByVal caption As String, ByVal action As String, ByVal isWeb As Boolean)
    Dim phase As String
    On Error GoTo Failed
    With button
        phase = "위치"
        If Abs(.Left - leftValue) > 0.1 Then .Left = leftValue
        If Abs(.Top - topValue) > 0.1 Then .Top = topValue
        If Abs(.Width - widthValue) > 0.1 Then .Width = widthValue
        If Abs(.Height - 22) > 0.1 Then .Height = 22
        .Placement = xlFreeFloating
        phase = "동작"
        If Len(action) > 0 Then
            If .OnAction <> action Then .OnAction = action
        End If
        If .TextFrame2.TextRange.Text <> caption Then .TextFrame2.TextRange.Text = caption
        phase = "색상"
        .Fill.Visible = msoTrue
        If isWeb Then
            .Fill.ForeColor.RGB = RGB(229, 239, 249)
            .Line.ForeColor.RGB = RGB(139, 164, 185)
        Else
            .Fill.ForeColor.RGB = RGB(212, 208, 200)
            .Line.ForeColor.RGB = RGB(128, 128, 128)
        End If
        .Fill.Transparency = 0
        .Line.Visible = msoTrue
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
        .Adjustments.Item(1) = 0.05
        .Glow.Radius = 0
        .SoftEdge.Radius = 0
        .Rotation = 0
        phase = "문자"
        With .TextFrame2
            .AutoSize = msoAutoSizeNone
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2
            .MarginTop = 1: .MarginBottom = 1
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "굴림"
                .Font.Size = 9
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With
        phase = "크기 보정"
        ' Excel snaps drawing coordinates to screen pixels at the current display scale.
        .Top = topValue
        .Height = 22
    End With
    Exit Sub
Failed:
    Err.Raise Err.Number, "ApplyActionStyle", button.Name & " / " & phase & ": " & Err.Description
End Sub

Public Sub EnsureMainActionButtons(ByVal ws As Worksheet)
    Dim button As Shape, refreshButton As Shape, fastButton As Shape, temp As Shape
    Dim managed As Variant, ordered() As Shape, name As Variant
    Dim count As Long, i As Long, j As Long
    Dim leftValue As Double, topValue As Double, caption As String, action As String
    If Not HasMainTaskHeaders(ws) Then Exit Sub
    Set refreshButton = FindActionShape(ws, "btnGanttRefresh")
    If refreshButton Is Nothing Then
        Set refreshButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B2").Left, ws.Range("B2").Top, 72, 22)
        refreshButton.Name = "btnGanttRefresh"
    End If
    Set fastButton = FindActionShape(ws, "btnGanttRefreshOnly")
    If fastButton Is Nothing Then
        Set fastButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, refreshButton.Left + 0.5, refreshButton.Top, 72, 22)
        fastButton.Name = "btnGanttRefreshOnly"
    End If
    managed = Array("btnDataImport", "btnGanttRefresh", "btnGanttRefreshOnly", "btnGanttReset", _
                    "btnGanttObjectInsert", "btnPeriodTextReport", COMPLETED_TASK_BUTTON_NAME)
    ReDim ordered(1 To UBound(managed) + 1)
    For Each name In managed
        Set button = FindActionShape(ws, CStr(name))
        If Not button Is Nothing Then
            count = count + 1
            Set ordered(count) = button
        End If
    Next name
    ' Preserve the user-selected left-to-right order, including the new fast button.
    For i = 2 To count
        Set temp = ordered(i)
        j = i - 1
        Do While j >= 1
            If ordered(j).Left <= temp.Left Then Exit Do
            Set ordered(j + 1) = ordered(j)
            j = j - 1
        Loop
        Set ordered(j + 1) = temp
    Next i
    leftValue = ws.Range("B2").Left
    topValue = ws.Range("B2").Top + ws.Range("B2").Height - 22
    For i = 1 To count
        Set button = ordered(i)
        caption = button.TextFrame2.TextRange.Text
        action = vbNullString
        If button.Name = "btnGanttRefresh" Then
            caption = "모두 새로고침": action = "칸트차트_모두새로고침"
        ElseIf button.Name = "btnGanttRefreshOnly" Then
            caption = "간트 새로고침": action = "칸트차트_간트새로고침"
        End If
        ApplyActionStyle button, leftValue, topValue, 72, caption, action, False
        leftValue = leftValue + 80
    Next i
    managed = Array("btnWebPreview", "btnWebUpload", "btnWebConfig")
    For Each name In managed
        Set button = FindActionShape(ws, CStr(name))
        If Not button Is Nothing Then
            ApplyActionStyle button, leftValue, topValue, button.Width, button.TextFrame2.TextRange.Text, vbNullString, True
            leftValue = leftValue + button.Width + 8
        End If
    Next name
End Sub

Public Function IsWebRegistrationMode(ByVal ws As Worksheet) As Boolean
    Dim c As Long, visibleCount As Long, fieldCount As Long, previousMode As Boolean
    Dim modeName As Name
    If Not HasMainTaskHeaders(ws) Then Exit Function
    If mWebModeCache Is Nothing Then Set mWebModeCache = CreateObject("Scripting.Dictionary")
    If mWebModeCache.Exists(ws.CodeName) Then
        previousMode = CBool(mWebModeCache(ws.CodeName))
    Else
        On Error Resume Next
        Set modeName = ws.Names(WEB_MODE_NAME)
        If Not modeName Is Nothing Then previousMode = (modeName.RefersTo = "=TRUE")
        On Error GoTo 0
    End If
    For c = ws.Range(COL_WEB_ENABLED & "1").Column To ws.Range(COL_WEB_RESULT & "1").Column
        fieldCount = fieldCount + 1
        If Not ws.Columns(c).Hidden Then visibleCount = visibleCount + 1
    Next c
    fieldCount = fieldCount + 1
    If Not ws.Columns(COL_WORK_MINUTES).Hidden Then visibleCount = visibleCount + 1
    ' A partial manual hide/unhide keeps the last explicit outline mode.
    If visibleCount = fieldCount Then
        IsWebRegistrationMode = True
    ElseIf visibleCount > 0 Then
        IsWebRegistrationMode = previousMode
    End If
End Function

Public Sub RefreshWebModeView(ByVal ws As Worksheet, Optional ByVal force As Boolean = False)
    Dim enabled As Boolean, previousMode As Boolean, knownMode As Boolean, wasSaved As Boolean
    Dim button As Shape, name As Variant, headers As Range, area As Range, colorValue As Long
    Dim modeName As Name, savedError As Long, savedDescription As String, phase As String
    If mWebViewUpdating Then Exit Sub
    If Not HasMainTaskHeaders(ws) Then Exit Sub
    mWebViewUpdating = True
    On Error GoTo Failed
    wasSaved = ws.Parent.Saved
    phase = "모드 판단"
    enabled = IsWebRegistrationMode(ws)
    knownMode = mWebModeCache.Exists(ws.CodeName)
    If knownMode Then previousMode = CBool(mWebModeCache(ws.CodeName))
    If Not ws.Columns(COL_WEB_TASK_ID).Hidden Then ws.Columns(COL_WEB_TASK_ID).Hidden = True
    If Not force And knownMode And previousMode = enabled Then GoTo Finished
    phase = "버튼 배치"
    If force Then EnsureMainActionButtons ws
    phase = "버튼 표시"
    For Each name In Array("btnWebPreview", "btnWebUpload", "btnWebConfig")
        Set button = FindActionShape(ws, CStr(name))
        If Not button Is Nothing Then
            If CBool(button.Visible) <> enabled Then button.Visible = IIf(enabled, msoTrue, msoFalse)
        End If
    Next name
    phase = "헤더 색상"
    Set headers = Union(ws.Range(COL_MINOR_CATEGORY & HEADER_ROW & ":" & COL_WEB_RESULT & HEADER_ROW), _
                        ws.Range(COL_REQUEST & HEADER_ROW & ":" & COL_MODIFICATION & HEADER_ROW), _
                        ws.Range(COL_ACTUAL_START & HEADER_ROW & ":" & COL_WORK_MINUTES & HEADER_ROW))
    colorValue = RGB(242, 242, 242)
    If enabled Then colorValue = RGB(255, 235, 156)
    For Each area In headers.Areas
        area.Interior.Pattern = xlSolid
        area.Interior.Color = colorValue
    Next area
    phase = "모드 저장"
    mWebModeCache(ws.CodeName) = enabled
    On Error Resume Next
    Set modeName = ws.Names(WEB_MODE_NAME)
    On Error GoTo Failed
    If modeName Is Nothing Then
        ws.Names.Add Name:="'" & Replace(ws.Name, "'", "''") & "'!" & WEB_MODE_NAME, _
                     RefersTo:=IIf(enabled, "=TRUE", "=FALSE"), Visible:=False
    ElseIf modeName.RefersTo <> IIf(enabled, "=TRUE", "=FALSE") Then
        modeName.RefersTo = IIf(enabled, "=TRUE", "=FALSE")
    End If
Finished:
    If wasSaved Then ws.Parent.Saved = True
    mWebViewUpdating = False
    Exit Sub
Failed:
    savedError = Err.Number: savedDescription = Err.Description
    mWebViewUpdating = False
    Err.Raise savedError, "RefreshWebModeView", phase & ": " & savedDescription
End Sub
