Attribute VB_Name = "WebViewRegression"
Option Explicit
' 2026-10-23 김좌현 수정 : WEB UPDATE 이전·재열기 보존과 중복 웹 버튼 제거 검증
' 2026-10-22 김좌현 수정 : 웹 모드 표시와 웹 초기화 없는 간트 새로고침 회귀 검증
' 2026-10-22 김좌현 수정 : 암호로 보호된 비업무 시트의 버튼 이관 제외와 재열기 보존 검증

Private Const UTILITY_FIXTURE_MARKER As String = "WebViewRegression protected utility fixture"
Private Const UTILITY_FIXTURE_PASSWORD As String = "WebViewRegressionFixture"
Private Const UTILITY_FIXTURE_SHEET As String = "ProtectedUtilityFixture"

Public ConfigurationCalls As Long
Public WebSetupCalls As Long
Public RefreshPhase As String
Public RenderCalls As Long

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 9280, "WebViewRegression", label
End Sub

Private Sub AddProtectedUtilityFixture(ByVal sheetName As String)
    Dim ws As Worksheet, button As Shape
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = sheetName
    ws.Range("A1").Value2 = UTILITY_FIXTURE_MARKER
    ws.Range("A1").Interior.Color = RGB(231, 219, 247)
    ws.Range("B2").Value2 = "editable utility value"
    ws.Range("B2").Locked = False
    ws.Cells(HEADER_ROW, COL_NO).Value2 = "Release information"
    ws.Cells(HEADER_ROW, COL_LEVEL).Value2 = "Protected utility"
    Set button = ws.Shapes.AddShape(msoShapeRectangle, 12, 18, 120, 24)
    button.Name = "btnGanttRefresh"
    button.TextFrame2.TextRange.Text = "Utility fixture button"
    button.OnAction = "WebViewFixtureNoAction"
    button.Fill.ForeColor.RGB = RGB(208, 229, 241)
    ws.Protect Password:=UTILITY_FIXTURE_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Sub CreateProtectedUtilityFixtures()
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets("배포정보")
    On Error GoTo 0
    If existing Is Nothing Then AddProtectedUtilityFixture "배포정보"
    AddProtectedUtilityFixture UTILITY_FIXTURE_SHEET
End Sub

Private Sub CheckProtectedUtilityFixtures()
    Dim candidate As Variant, ws As Worksheet, button As Shape, fixtureCount As Long, wrongPasswordError As Long
    For Each candidate In Array("배포정보", UTILITY_FIXTURE_SHEET)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(candidate))
        On Error GoTo 0
        If CStr(candidate) = UTILITY_FIXTURE_SHEET Then Check Not ws Is Nothing, "generic protected utility fixture missing"
        If Not ws Is Nothing Then
            If CStr(candidate) = UTILITY_FIXTURE_SHEET Then Check Not IsError(ws.Range("A1").Value2), "generic utility marker has a cell error"
            If Not IsError(ws.Range("A1").Value2) Then
                If CStr(candidate) = UTILITY_FIXTURE_SHEET Or CStr(ws.Range("A1").Value2) = UTILITY_FIXTURE_MARKER Then
                    fixtureCount = fixtureCount + 1
                    UpgradeGanttRefreshButtons ws
                    Check ws.ProtectContents And ws.ProtectDrawingObjects And ws.ProtectScenarios, "utility protection changed: " & ws.Name
                    Check CStr(ws.Range("A1").Value2) = UTILITY_FIXTURE_MARKER, "utility content changed: " & ws.Name
                    Check ws.Range("A1").Interior.Color = RGB(231, 219, 247), "utility formatting changed: " & ws.Name
                    Check CStr(ws.Range("B2").Value2) = "editable utility value" And Not ws.Range("B2").Locked, "utility input changed: " & ws.Name
                    Check ws.Range("A1").Locked, "utility locked cell changed: " & ws.Name
                    Check CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) = "Release information", "utility header changed: " & ws.Name
                    Check CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) = "Protected utility", "utility level header changed: " & ws.Name
                    Check ws.Shapes.Count = 1, "utility shapes changed: " & ws.Name
                    Set button = ws.Shapes("btnGanttRefresh")
                    Check button.TextFrame2.TextRange.Text = "Utility fixture button", "utility button caption changed: " & ws.Name
                    Check button.OnAction = "WebViewFixtureNoAction", "utility button action changed: " & ws.Name
                    Check button.Fill.ForeColor.RGB = RGB(208, 229, 241), "utility button fill changed: " & ws.Name
                    Check Abs(button.Left - 12) < 0.75 And Abs(button.Top - 18) < 0.75, "utility button position changed: " & ws.Name
                    Check Abs(button.Width - 120) < 0.75 And Abs(button.Height - 24) < 0.75, "utility button size changed: " & ws.Name
                    On Error Resume Next
                    ws.Unprotect Password:="wrong-fixture-password"
                    wrongPasswordError = Err.Number
                    Err.Clear
                    On Error GoTo 0
                    Check wrongPasswordError <> 0 And ws.ProtectContents, "utility password protection lost: " & ws.Name
                End If
            End If
        End If
    Next candidate
    Check fixtureCount > 0, "protected utility fixture missing"
End Sub

Private Sub CheckMode(ByVal ws As Worksheet, ByVal enabled As Boolean)
    Dim name As Variant, c As Long, expected As Long, button As Shape
    RefreshWebModeView ws, True
    Check IsWebRegistrationMode(ws) = enabled, "mode mismatch: expected=" & enabled & ", actual=" & IsWebRegistrationMode(ws) & ", Hhidden=" & ws.Columns(COL_WEB_ENABLED).Hidden & ", Zhidden=" & ws.Columns(COL_WORK_MINUTES).Hidden
    For Each name In Array("btnWebPreview", "btnWebConfig")
        Check CBool(ws.Shapes(CStr(name)).Visible) = enabled, "button visibility " & CStr(name)
        ' Excel quantizes shape geometry at the current screen scale; allow less than one 96-DPI pixel.
        Check Abs(ws.Shapes(CStr(name)).Height - ws.Shapes("btnGanttRefresh").Height) < 0.75, "button height " & CStr(name)
        Check Abs(ws.Shapes(CStr(name)).Top - ws.Shapes("btnGanttRefresh").Top) < 0.75, "button top " & CStr(name)
        Check ws.Shapes(CStr(name)).TextFrame2.TextRange.Font.Name = "굴림", "button font " & CStr(name)
        Check ws.Shapes(CStr(name)).Fill.ForeColor.RGB = RGB(229, 239, 249), "button blue " & CStr(name)
    Next name
    For Each button In ws.Shapes
        Check button.Name <> "btnWebUpload", "duplicate main web upload button survived"
    Next button
    Check CStr(ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2) = "WEB UPDATE", "WEB UPDATE header missing"
    expected = RGB(242, 242, 242)
    If enabled Then expected = RGB(255, 235, 156)
    For Each name In Array("G4:P4", "R4:S4", "X4:Z4")
        Check ws.Range(CStr(name)).Interior.Color = expected, "header " & CStr(name)
    Next name
    Check ws.Columns(COL_WEB_TASK_ID).Hidden, "private ID visible"
End Sub

Public Function RunWebViewRegression() As String
    Dim ws As Worksheet, config As Worksheet, c As Long, r As Long
    Dim scrollColumn As Long, scrollRow As Long, started As Double, fullSeconds As Double, fastSeconds As Double
    Dim stage As String, legacyWebButton As Shape
    On Error GoTo Failed
    Application.EnableEvents = False
    stage = "protected utility button upgrade"
    CreateProtectedUtilityFixtures
    CheckProtectedUtilityFixtures
    EnsureConfigSheet
    Set config = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    config.Range(DISPLAY_SETTING_START_VALUE_CELL).Value2 = DateSerial(2026, 10, 1)
    config.Range(DISPLAY_SETTING_END_VALUE_CELL).Value2 = DateSerial(2026, 10, 31)
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "WebViewRegression"
    SetupDataHeaders ws
    For r = 5 To 7
        ws.Cells(r, COL_LEVEL).Value2 = IIf(r = 5, 1, 2)
        ws.Cells(r, COL_REQUEST).Value2 = "Regression task " & r
        ws.Cells(r, COL_MINOR_CATEGORY).Value2 = "Test module"
        ws.Cells(r, COL_PLAN_START).Value2 = DateSerial(2026, 10, 19)
        ws.Cells(r, COL_PLAN_END).Value2 = DateSerial(2026, 10, 23)
        ws.Cells(r, COL_ACTUAL_START).Value2 = DateSerial(2026, 10, 21)
        ws.Cells(r, COL_PROGRESS).Value2 = 0.25
    Next r
    stage = "initial full refresh"
    started = Timer
    RefreshGanttSheet False
    fullSeconds = Timer - started
    Check ConfigurationCalls > 0 And WebSetupCalls > 0, "full setup was omitted"
    Check RenderCalls > 0, "full renderer was omitted"
    For c = 1 To 3
        stage = "outline " & c
        ws.Outline.ShowLevels ColumnLevels:=c
        EnsureTaskRequestFreeze ws
        CheckMode ws, (c = 3)
    Next c
    stage = "partial mode"
    ws.Columns(COL_WEB_MODULE).Hidden = True
    CheckMode ws, True
    ws.Outline.ShowLevels ColumnLevels:=2
    CheckMode ws, False
    ws.Outline.ShowLevels ColumnLevels:=3
    CheckMode ws, True
    ws.Columns(COL_WEB_MODULE).Hidden = True
    stage = "hidden row and scroll"
    ws.Rows(7).Hidden = True
    ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = ws.Range(COL_GANTT_START & "1").Column
    scrollColumn = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn
    scrollRow = ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow
    stage = "edit child dates"
    UnprotectTaskSheet ws
    ws.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2 = "웹등록"
    Set legacyWebButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, 1, 1, 96, 22)
    legacyWebButton.Name = "btnWebUpload"
    ws.Cells(6, COL_ACTUAL_END).Value2 = DateSerial(2026, 10, 23)
    ws.Cells(7, COL_ACTUAL_END).Value2 = DateSerial(2026, 10, 24)
    ApplyCalculatedColumnsProtection ws, 7
    ConfigurationCalls = 0: WebSetupCalls = 0
    RenderCalls = 0
    stage = "first fast refresh"
    started = Timer
    RefreshGanttOnly False
    fastSeconds = Timer - started
    Check ConfigurationCalls = 0 And WebSetupCalls = 0, "fast refresh invoked configuration/web setup"
    Check RenderCalls > 0, "changed dates did not redraw chart"
    stage = "verify first fast"
    CheckMode ws, True
    Check ws.Columns(COL_WEB_MODULE).Hidden, "manual column hide lost"
    Check ws.Rows(7).Hidden, "manual row hide lost"
    Check ws.ProtectContents, "protection lost"
    Check CDate(ws.Cells(5, COL_ACTUAL_END).Value2) = DateSerial(2026, 10, 24), "parent actual end stale"
    Check CDbl(ws.Cells(5, COL_PROGRESS).Value2) = 1, "parent progress stale"
    Check ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollColumn = scrollColumn, "horizontal scroll lost"
    Check ActiveWindow.Panes(ActiveWindow.Panes.Count).ScrollRow = scrollRow, "vertical scroll lost"
    c = ws.Range(COL_GANTT_START & "1").Column + 24
    Check ws.Cells(6, c).Interior.Color = RGB(245, 245, 245), "weekend background lost"
    stage = "repeat fast"
    RenderCalls = 0
    RefreshGanttOnly False
    Check ConfigurationCalls = 0 And WebSetupCalls = 0, "repeat fast setup invoked"
    Check RenderCalls = 0, "unchanged fast refresh redrew chart"
    stage = "clear chart then fast"
    UnprotectTaskSheet ws
    ClearGanttArea ws
    ApplyCalculatedColumnsProtection ws, 7
    RefreshGanttOnly False
    Check RenderCalls > 0, "cleared chart was not restored"
    stage = "manual chart cell change"
    RenderCalls = 0
    Application.EnableEvents = True
    ws.Cells(6, c).Value2 = "manual chart change"
    Application.EnableEvents = False
    RefreshGanttOnly False
    Check RenderCalls > 0 And Len(CStr(ws.Cells(6, c).Value2)) = 0, "manual chart edit was not restored"
    stage = "error then normal styling"
    ws.Cells(6, COL_PLAN_END).Value2 = DateSerial(2026, 10, 17)
    RefreshGanttOnly False
    Check HasTaskError(ws, 6), "invalid dates were not recalculated"
    ws.Cells(6, COL_PLAN_END).Value2 = DateSerial(2026, 10, 23)
    RefreshGanttOnly False
    Check Not HasTaskError(ws, 6), "corrected dates retained error"
    Check ws.Cells(6, COL_REQUEST).Interior.Color <> RGB(255, 242, 204), "corrected dates retained error styling"
    stage = "second full"
    RefreshGanttSheet False
    CheckMode ws, True
    ws.Outline.ShowLevels ColumnLevels:=2
    CheckMode ws, False
    RefreshGanttOnly False
    CheckMode ws, False
    RunWebViewRegression = "PASS: web mode, uniform buttons, fast/full separation, hierarchy dates/progress, hidden view, weekend and protection; full=" & _
                           Format$(fullSeconds, "0.00") & "s fast=" & Format$(fastSeconds, "0.00") & "s"
    Exit Function
Failed:
    RunWebViewRegression = "FAIL: " & stage & " / " & RefreshPhase & ": " & Err.Description
End Function

Public Function RunPartialSavedWebViewRegression() As String
    Dim ws As Worksheet
    On Error GoTo Failed
    CheckProtectedUtilityFixtures
    Set ws = ThisWorkbook.Worksheets("WebViewRegression")
    ws.Activate
    CheckMode ws, True
    Check ws.Columns(COL_WEB_MODULE).Hidden, "saved manual web column hide lost"
    RunPartialSavedWebViewRegression = "PASS: partial manual hide preserves web mode after reopen"
    Exit Function
Failed:
    RunPartialSavedWebViewRegression = "FAIL: " & Err.Description
End Function

Public Function RunSavedWebViewRegression() As String
    Dim ws As Worksheet, stage As String
    On Error GoTo Failed
    stage = "saved protected utility button upgrade"
    CheckProtectedUtilityFixtures
    Set ws = ThisWorkbook.Worksheets("WebViewRegression")
    ws.Activate
    StartTaskOutlineViewMonitor
    stage = "saved mode"
    CheckMode ws, False
    ws.Outline.ShowLevels ColumnLevels:=3
    RefreshWebModeView ws
    stage = "native 3"
    CheckMode ws, True
    ws.Columns(COL_WEB_MODULE).Hidden = True
    RefreshWebModeView ws
    stage = "manual hide"
    CheckMode ws, True
    RunSavedWebViewRegression = "PASS: reopen and native outline mode restoration"
    Exit Function
Failed:
    RunSavedWebViewRegression = "FAIL: " & stage & ": " & Err.Description
End Function
