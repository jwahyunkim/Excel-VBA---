Attribute VB_Name = "modVersionMigration"
Option Explicit
' 2026-10-29 김좌현 수정 : 수동 버튼 생성 화면·폐기 보고서 정비 제거 및 새로고침 연결 단일화
' 2026-10-29 김좌현 수정 : 비고 개체의 실제 행 연결을 보존하고 셀을 지정해 OLE 붙여넣기
' 2026-10-23 김좌현 수정 : 현재 웹 업무 열 배치와 업무 ID·등록 이력·요청번호 연결을 함께 이관
' 2026-10-22 김좌현 수정 : 중복 새로고침 매크로 제거 후 기존 버튼을 모두 새로고침으로 이관
' 2026-10-22 김좌현 수정 : 모두/간트 새로고침 버튼 분리와 웹 등록 모드 버튼 배치 연결
' 2026-10-22 김좌현 수정 : 웹 설정/미리보기/이력 시트를 업무 이전 및 버튼 처리에서 제외
' 2026-10-22 김좌현 수정 : 배포정보 등 업무 헤더가 없는 시트의 버튼 정비와 보호 해제 제외
' 2026-10-20 김좌현 수정 : 기존 디자인의 완료 숨김 버튼 선택/생성과 이관 후 상태 반영
' 2026-10-20 김좌현 수정 : 미사용 간트 only/보고 only 행 필터 호출 제거
' 2026-10-20 김좌현 수정 : 항목 숨김 버튼과 선택 항목 제거 및 순서 정리
' 2026-10-20 김좌현 수정 : 새로고침 버튼 통합 및 기존 생성 버튼 연결 전환

Private Const AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3
Private Const FILE_DIALOG_PICKER As Long = 3
Private Const PPT_OLE_PREFIX As String = "oleGanttPpt_"
Private Const TEMP_PPT_OLE_PREFIX As String = "tmpGanttPptOle_"
Private Const LEGACY_BUTTON_SETUP_SHEET_NAME As String = "_버튼생성"

Public Sub 데이터_가져오기()
    Dim targetBook As Workbook
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim originalSheet As Object
    Dim sourcePath As String
    Dim matchingSheetCount As Long
    Dim importedSheetCount As Long
    Dim refreshedSheetCount As Long
    Dim copiedObjectCount As Long
    Dim failedObjectCount As Long
    Dim importedLogCount As Long
    Dim importedStateCount As Long
    Dim webMigrationPlan As Object
    Dim openedByMigration As Boolean
    Dim stateCaptured As Boolean
    Dim previousAutomationSecurity As Long
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim answer As VbMsgBoxResult
    Dim errorText As String

    On Error GoTo EH

    Set targetBook = ThisWorkbook
    If Not ActiveWorkbook Is targetBook Then
        MsgBox "신버전 파일을 활성화한 뒤 다시 실행하세요.", vbExclamation
        Exit Sub
    End If

    sourcePath = PickLegacyWorkbookPath()
    If Len(sourcePath) = 0 Then Exit Sub

    If StrComp(sourcePath, targetBook.FullName, vbTextCompare) = 0 Then
        MsgBox "현재 신버전 파일과 다른 구버전 파일을 선택하세요.", vbExclamation
        Exit Sub
    End If

    previousAutomationSecurity = Application.AutomationSecurity
    Application.AutomationSecurity = AUTOMATION_SECURITY_FORCE_DISABLE

    Set sourceBook = FindOpenWorkbook(sourcePath)
    If sourceBook Is Nothing Then
        Set sourceBook = Workbooks.Open( _
            Filename:=sourcePath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            AddToMru:=False)
        openedByMigration = True
    End If
    Application.AutomationSecurity = previousAutomationSecurity

    matchingSheetCount = CountMatchingTaskSheets(sourceBook, targetBook)
    If matchingSheetCount = 0 Then
        MsgBox "구버전과 신버전에서 이름이 같은 업무 시트를 찾지 못했습니다.", vbExclamation
        GoTo SafeExit
    End If

    answer = MsgBox( _
        "같은 이름의 업무 시트 " & matchingSheetCount & "개를 가져옵니다." & vbCrLf & vbCrLf & _
        "신버전 시트의 기존 입력 내용과 PPT 개체는 구버전 내용으로 교체됩니다." & vbCrLf & _
        "가져오는 업무의 웹 등록 이력과 요청번호 연결도 함께 교체됩니다." & vbCrLf & _
        "신버전의 계산식, VBA 및 기본 서식은 유지됩니다." & vbCrLf & vbCrLf & _
        "계속하시겠습니까?", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, _
        "구버전 데이터 가져오기")
    If answer <> vbYes Then GoTo SafeExit

    Set originalSheet = ActiveSheet
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousCalculation = Application.Calculation
    stateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set webMigrationPlan = BuildWebMigrationPlan(sourceBook, targetBook)
    ImportOutputReportConfig sourceBook, targetBook

    For Each sourceSheet In sourceBook.Worksheets
        If IsMigrationTaskPair(sourceSheet, targetBook) Then
            Set targetSheet = GetWorksheet(targetBook, sourceSheet.Name)
            If Not targetSheet Is Nothing Then
                ImportTaskSheet sourceSheet, targetSheet, copiedObjectCount, failedObjectCount
                importedSheetCount = importedSheetCount + 1
            End If
        End If
    Next sourceSheet

    EnsureWebHistorySheets
    ApplyWebMigrationPlan webMigrationPlan, targetBook, importedLogCount, importedStateCount
    ResetWebPreviewAfterMigration

    For Each sourceSheet In sourceBook.Worksheets
        If IsMigrationTaskPair(sourceSheet, targetBook) Then
            Set targetSheet = GetWorksheet(targetBook, sourceSheet.Name)
            If Not targetSheet Is Nothing Then
                PrepareWebTaskSheet targetSheet
                If RefreshMigratedSheet(targetSheet) Then
                    refreshedSheetCount = refreshedSheetCount + 1
                End If
            End If
        End If
    Next sourceSheet

    originalSheet.Activate
    RestoreApplicationState previousScreenUpdating, previousEnableEvents, previousCalculation
    stateCaptured = False

    MsgBox "구버전 데이터 가져오기가 완료되었습니다." & vbCrLf & vbCrLf & _
           "가져온 시트: " & importedSheetCount & "개" & vbCrLf & _
           "차트 재생성 시트: " & refreshedSheetCount & "개" & vbCrLf & _
           "가져온 웹 이력: " & importedLogCount & "건 / 연결: " & importedStateCount & "건" & vbCrLf & _
           "복사한 PPT 개체: " & copiedObjectCount & "개" & vbCrLf & _
           "복사하지 못한 PPT 개체: " & failedObjectCount & "개", _
           IIf(failedObjectCount = 0, vbInformation, vbExclamation), _
           "구버전 데이터 가져오기"

SafeExit:
    On Error Resume Next
    Application.AutomationSecurity = previousAutomationSecurity
    If openedByMigration And Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Exit Sub

EH:
    errorText = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    Application.AutomationSecurity = previousAutomationSecurity
    If stateCaptured Then
        RestoreApplicationState previousScreenUpdating, previousEnableEvents, previousCalculation
    End If
    If Not originalSheet Is Nothing Then originalSheet.Activate
    If openedByMigration And Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    On Error GoTo 0

    MsgBox "구버전 데이터를 가져오는 중 오류가 발생했습니다." & vbCrLf & _
           "원인: " & errorText, _
           vbExclamation, _
           "구버전 데이터 가져오기 오류"
End Sub

' Converts managed legacy buttons and workbook-qualified custom assignments.
' Called on open and refresh; keeps the existing refresh button when both exist.
Public Sub UpgradeGanttRefreshButtons(ByVal ws As Worksheet)
    Dim createButton As Shape, refreshButton As Shape, fastButton As Shape, button As Shape
    Dim macroName As String, wasProtected As Boolean, needsUpdate As Boolean
    Dim errNumber As Long, errDescription As String

    If Not IsTaskSheet(ws) Then Exit Sub
    If CStr(ws.Cells(HEADER_ROW, COL_NO).Value2) <> "No." Or _
       CStr(ws.Cells(HEADER_ROW, COL_LEVEL).Value2) <> "Level" Then Exit Sub
    On Error GoTo Failed
    On Error Resume Next
    Set createButton = ws.Shapes("btnGanttCreate")
    Set refreshButton = ws.Shapes("btnGanttRefresh")
    Set fastButton = ws.Shapes("btnGanttRefreshOnly")
    On Error GoTo Failed
    ' 2026-10-27 김좌현 수정 : 두 새로고침 버튼을 한 버튼으로 이관
    needsUpdate = refreshButton Is Nothing Or Not createButton Is Nothing Or Not fastButton Is Nothing
    For Each button In ws.Shapes
        macroName = GetGanttButtonMacroName(button)
        If IsLegacyGanttRefreshMacro(macroName) Then needsUpdate = True
        If macroName = "칸트차트_새로고침" Then
            If button.OnAction <> macroName Then needsUpdate = True
        End If
        If button.Name = "btnGanttRefresh" Then
            If macroName <> "칸트차트_새로고침" Or _
               button.TextFrame2.TextRange.Text <> "새로고침" Then needsUpdate = True
        End If
    Next button
    ' Leave already migrated workbooks unchanged when opened.
    If Not needsUpdate Then Exit Sub
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
    If wasProtected Then UnprotectTaskSheet ws
    If Not createButton Is Nothing Then
        If refreshButton Is Nothing Then
            createButton.Name = "btnGanttRefresh"
        Else
            createButton.Delete
        End If
    End If
    If Not fastButton Is Nothing Then fastButton.Delete
    For Each button In ws.Shapes
        macroName = GetGanttButtonMacroName(button)
        If button.Name = "btnGanttRefresh" Or IsLegacyGanttRefreshMacro(macroName) Or macroName = "칸트차트_새로고침" Then
            button.OnAction = "칸트차트_새로고침"
            button.TextFrame2.TextRange.Text = "새로고침"
        End If
    Next button
    EnsureMainActionButtons ws
    If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    Exit Sub
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    If wasProtected Then
        On Error Resume Next
        ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
        On Error GoTo 0
    End If
    Err.Raise errNumber, "UpgradeGanttRefreshButtons", errDescription
End Sub

Private Function IsLegacyGanttRefreshMacro(ByVal macroName As String) As Boolean
    Select Case macroName
        Case "칸트차트_생성", "칸트차트_모두새로고침", "칸트차트_간트새로고침", "RefreshGanttOnly"
            IsLegacyGanttRefreshMacro = True
    End Select
End Function

Private Function GetGanttButtonMacroName(ByVal button As Shape) As String
    Dim macroName As String
    ' Some embedded objects do not expose OnAction.
    On Error Resume Next
    macroName = CStr(button.OnAction)
    On Error GoTo 0
    If InStrRev(macroName, "!") > 0 Then macroName = Mid$(macroName, InStrRev(macroName, "!") + 1)
    If InStrRev(macroName, ".") > 0 Then macroName = Mid$(macroName, InStrRev(macroName, ".") + 1)
    GetGanttButtonMacroName = macroName
End Function

Private Sub ImportTaskSheet(ByVal sourceSheet As Worksheet, _
                            ByVal targetSheet As Worksheet, _
                            ByRef copiedObjectCount As Long, _
                            ByRef failedObjectCount As Long)
    Dim sourceLastRow As Long
    Dim targetLastRow As Long
    Dim clearLastRow As Long
    Dim rowNum As Long
    Dim sourceOle As OLEObject
    Dim sourceHasProgramColumn As Boolean
    Dim sourceHasFourLevelClassification As Boolean
    Dim sourceHasSplitContent As Boolean
    Dim sourceHasWebLayout As Boolean
    Dim sourceHasWorkMinutes As Boolean
    Dim sourceRequestColumn As String
    Dim sourceDetailsStartColumn As String
    Dim sourceDetailsEndColumn As String
    Dim sourceManualStartColumn As String
    Dim sourceManualEndColumn As String

    sourceLastRow = GetMigrationLastRow(sourceSheet)
    targetLastRow = GetMigrationLastRow(targetSheet)
    clearLastRow = sourceLastRow
    If targetLastRow > clearLastRow Then clearLastRow = targetLastRow
    If clearLastRow < DATA_START_ROW Then clearLastRow = DATA_START_ROW

    UnprotectTaskSheet targetSheet
    SetupDataHeaders targetSheet
    DeleteTargetPptObjects targetSheet
    ClearTargetPptIcons targetSheet, clearLastRow
    sourceHasFourLevelClassification = _
        (Trim$(CStr(sourceSheet.Range("D" & HEADER_ROW).Value2)) = "타입" And _
         Trim$(CStr(sourceSheet.Range("E" & HEADER_ROW).Value2)) = "대분류" And _
         Trim$(CStr(sourceSheet.Range("F" & HEADER_ROW).Value2)) = "중분류" And _
         Trim$(CStr(sourceSheet.Range("G" & HEADER_ROW).Value2)) = "소분류")
    sourceHasProgramColumn = _
        (Trim$(CStr(sourceSheet.Range("E" & HEADER_ROW).Value2)) = "프로그램")
    sourceHasSplitContent = _
        (Trim$(CStr(sourceSheet.Range("H" & HEADER_ROW).Value2)) = "요청 내용" And _
         Trim$(CStr(sourceSheet.Range("I" & HEADER_ROW).Value2)) = "수정 내용")
    sourceHasWebLayout = HasMigrationWebLayout(sourceSheet)
    sourceHasWorkMinutes = (Trim$(CStr(sourceSheet.Range(COL_WORK_MINUTES & HEADER_ROW).Value2)) = "작업시간(분)")

    targetSheet.Range(COL_LEVEL & DATA_START_ROW & ":" & _
                      COL_PROGRESS & clearLastRow).ClearContents
    targetSheet.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & _
                      COL_WEEKLY_REPORT & clearLastRow).ClearContents

    If sourceLastRow >= DATA_START_ROW Then
        If sourceHasWebLayout Then
            targetSheet.Range(COL_WEB_ENABLED & DATA_START_ROW & ":" & COL_WEB_TASK_ID & sourceLastRow).NumberFormat = "@"
            If sourceHasWorkMinutes Then
                CopyMigrationColumnBlock sourceSheet, targetSheet, "C", "AA", COL_LEVEL, COL_PROGRESS, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, "AC", "AE", COL_MANUAL_PROGRESS, COL_WEEKLY_REPORT, sourceLastRow
            Else
                CopyMigrationColumnBlock sourceSheet, targetSheet, "C", "Y", COL_LEVEL, COL_ACTUAL_END, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, "Z", "Z", COL_PROGRESS, COL_PROGRESS, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, "AB", "AD", COL_MANUAL_PROGRESS, COL_WEEKLY_REPORT, sourceLastRow
            End If
        Else
            If sourceHasFourLevelClassification Then
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "C", "G", COL_LEVEL, COL_MINOR_CATEGORY, sourceLastRow
                sourceRequestColumn = "H"
                If sourceHasSplitContent Then
                    CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                             "I", "I", COL_MODIFICATION, COL_MODIFICATION, sourceLastRow
                    sourceDetailsStartColumn = "J"
                    sourceDetailsEndColumn = "P"
                    sourceManualStartColumn = "R"
                    sourceManualEndColumn = "T"
                Else
                    sourceDetailsStartColumn = "I"
                    sourceDetailsEndColumn = "O"
                    sourceManualStartColumn = "Q"
                    sourceManualEndColumn = "S"
                End If
            ElseIf sourceHasProgramColumn Then
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "C", "C", COL_LEVEL, COL_LEVEL, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "D", "D", COL_MAJOR_CATEGORY, COL_MAJOR_CATEGORY, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "E", "E", COL_MINOR_CATEGORY, COL_MINOR_CATEGORY, sourceLastRow
                sourceRequestColumn = "F"
                sourceDetailsStartColumn = "G"
                sourceDetailsEndColumn = "M"
                sourceManualStartColumn = "O"
                sourceManualEndColumn = "Q"
            Else
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "C", "C", COL_LEVEL, COL_LEVEL, sourceLastRow
                CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                         "D", "D", COL_MAJOR_CATEGORY, COL_MAJOR_CATEGORY, sourceLastRow
                sourceRequestColumn = "E"
                sourceDetailsStartColumn = "F"
                sourceDetailsEndColumn = "L"
                sourceManualStartColumn = "N"
                sourceManualEndColumn = "P"
            End If

            CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                     sourceRequestColumn, sourceRequestColumn, COL_REQUEST, COL_REQUEST, sourceLastRow
            CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                     sourceDetailsStartColumn, MigrationPreviousColumn(sourceSheet, sourceDetailsEndColumn), _
                                     COL_OWNER, COL_ACTUAL_END, sourceLastRow
            CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                     sourceDetailsEndColumn, sourceDetailsEndColumn, COL_PROGRESS, COL_PROGRESS, sourceLastRow
            CopyMigrationColumnBlock sourceSheet, targetSheet, _
                                     sourceManualStartColumn, sourceManualEndColumn, _
                                     COL_MANUAL_PROGRESS, COL_WEEKLY_REPORT, sourceLastRow
        End If

        For rowNum = DATA_START_ROW To sourceLastRow
            If CStr(targetSheet.Cells(rowNum, COL_NOTE).Value2) = ChrW(&H25A0) Then
                targetSheet.Cells(rowNum, COL_NOTE).ClearContents
            End If
        Next rowNum
    End If

    For Each sourceOle In sourceSheet.OLEObjects
        If IsMigratablePptObject(sourceOle) Then
            rowNum = GetGanttPptObjectRow(sourceSheet, sourceOle)
            If rowNum >= DATA_START_ROW Then
                If CopyPptObject(sourceOle, targetSheet, rowNum) Then
                    copiedObjectCount = copiedObjectCount + 1
                Else
                    failedObjectCount = failedObjectCount + 1
                End If
            End If
        End If
    Next sourceOle

    ApplyTaskInputValidation targetSheet
End Sub

Private Sub CopyMigrationColumnBlock(ByVal sourceSheet As Worksheet, _
                                     ByVal targetSheet As Worksheet, _
                                     ByVal sourceFirstColumn As String, _
                                     ByVal sourceLastColumn As String, _
                                     ByVal targetFirstColumn As String, _
                                     ByVal targetLastColumn As String, _
                                     ByVal lastRow As Long)
    targetSheet.Range(targetFirstColumn & DATA_START_ROW & ":" & _
                      targetLastColumn & lastRow).Value2 = _
        sourceSheet.Range(sourceFirstColumn & DATA_START_ROW & ":" & _
                          sourceLastColumn & lastRow).Value2
End Sub

Private Function CopyPptObject(ByVal sourceOle As OLEObject, _
                               ByVal targetSheet As Worksheet, _
                               ByVal targetRow As Long) As Boolean
    Dim pastedOle As OLEObject
    Dim beforeCount As Long

    On Error GoTo CopyFailed

    sourceOle.Parent.Activate
    sourceOle.Copy
    targetSheet.Activate
    targetSheet.Cells(targetRow, COL_NOTE).Select

    beforeCount = targetSheet.OLEObjects.Count
    targetSheet.Paste
    If targetSheet.OLEObjects.Count <= beforeCount Then GoTo CopyFailed

    Set pastedOle = targetSheet.OLEObjects(targetSheet.OLEObjects.Count)
    pastedOle.Name = PPT_OLE_PREFIX & CStr(targetRow)
    PositionMigratedPptObject targetSheet, pastedOle, targetRow
    SetMigratedPptCellIcon targetSheet, targetRow

    Application.CutCopyMode = False
    CopyPptObject = True
    Exit Function

CopyFailed:
    ' Remove a partial paste so a retry cannot leave an unbound or duplicate object.
    On Error Resume Next
    If Not pastedOle Is Nothing Then pastedOle.Delete
    Application.CutCopyMode = False
    CopyPptObject = False
End Function

Private Function RefreshMigratedSheet(ByVal ws As Worksheet) As Boolean
    Dim lastRow As Long
    Dim minDate As Date
    Dim maxDate As Date
    Dim chartStartDate As Date
    Dim chartEndDate As Date
    Dim holidayDict As Object
    Dim workdayDict As Object
    Dim rowView As Object

    On Error GoTo RefreshFailed

    lastRow = GetLastDataRow(ws)
    If lastRow < DATA_START_ROW Then Exit Function

    EnsureConfigSheet
    LoadHolidaySettings holidayDict, workdayDict
    If Not GetMinMaxDate(ws, lastRow, minDate, maxDate) Then
        ApplyCalculatedColumnsProtection ws, lastRow
        Exit Function
    End If

    chartStartDate = minDate
    chartEndDate = maxDate
    Call TryGetDisplayDateRange(chartStartDate, chartEndDate)

    Set rowView = CaptureTaskRowView(ws, lastRow)
    UnprotectTaskSheet ws
    SetupDataHeaders ws
    SynchronizeTaskHierarchyModules ws, lastRow, False
    UpdateTaskNumbers ws, lastRow
    ApplyTaskInputValidation ws
    ShowAllTaskRows ws, lastRow
    ClearCalculatedArea ws, lastRow
    ClearGanttArea ws
    DrawDateHeader ws, chartStartDate, chartEndDate, holidayDict, workdayDict
    DrawTaskBars ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict
    FormatBaseArea ws, lastRow, chartStartDate, chartEndDate, holidayDict, workdayDict
    ShowAllDateColumns ws
    AlignAllGanttPptArtifactsToNoteCells ws
    RestoreTaskRowView ws, lastRow, rowView
    ApplyCompletedTaskVisibility ws, lastRow
    ApplyCalculatedColumnsProtection ws, lastRow

    RefreshMigratedSheet = True
    Exit Function

RefreshFailed:
    RefreshMigratedSheet = False
End Function

Private Sub PositionMigratedPptObject(ByVal ws As Worksheet, _
                                      ByVal oleItem As OLEObject, _
                                      ByVal targetRow As Long)
    Dim noteCell As Range

    BindGanttPptObjectRow ws, oleItem, targetRow
    If ws.Rows(targetRow).Hidden Then Exit Sub
    Set noteCell = ws.Cells(targetRow, ws.Range(COL_NOTE & "1").Column)
    With oleItem
        .Placement = xlMoveAndSize
        .Left = noteCell.Left + 1
        .Top = noteCell.Top + 1
        .Width = 1
        .Height = 1
        .Visible = False
        .PrintObject = False
    End With
End Sub

Private Sub SetMigratedPptCellIcon(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim noteCell As Range
    Dim iconFontSize As Double
    Dim subAddress As String

    Set noteCell = ws.Cells(targetRow, ws.Range(COL_NOTE & "1").Column)
    subAddress = "'" & Replace(ws.Name, "'", "''") & "'!" & noteCell.Address

    On Error Resume Next
    noteCell.Hyperlinks.Delete
    On Error GoTo 0

    noteCell.Value = ChrW(&H25A0)
    ws.Hyperlinks.Add Anchor:=noteCell, _
                      Address:="", _
                      subAddress:=subAddress, _
                      ScreenTip:="PPT 편집 화면 열기", _
                      TextToDisplay:=ChrW(&H25A0)

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
        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    End With
End Sub

Private Sub DeleteTargetPptObjects(ByVal targetSheet As Worksheet)
    Dim objectIndex As Long
    Dim objectName As String

    For objectIndex = targetSheet.OLEObjects.Count To 1 Step -1
        objectName = targetSheet.OLEObjects(objectIndex).Name
        If Left$(objectName, Len(PPT_OLE_PREFIX)) = PPT_OLE_PREFIX Or _
           Left$(objectName, Len(TEMP_PPT_OLE_PREFIX)) = TEMP_PPT_OLE_PREFIX Then
            targetSheet.OLEObjects(objectIndex).Delete
        End If
    Next objectIndex
End Sub

Private Sub ClearTargetPptIcons(ByVal targetSheet As Worksheet, ByVal lastRow As Long)
    Dim rowNum As Long
    Dim noteCell As Range

    For rowNum = DATA_START_ROW To lastRow
        Set noteCell = targetSheet.Cells(rowNum, targetSheet.Range(COL_NOTE & "1").Column)
        If CStr(noteCell.Value2) = ChrW(&H25A0) Then
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
    Next rowNum
End Sub

Private Function IsMigratablePptObject(ByVal sourceOle As OLEObject) As Boolean
    Dim objectName As String

    objectName = sourceOle.Name
    IsMigratablePptObject = _
        (Left$(objectName, Len(PPT_OLE_PREFIX)) = PPT_OLE_PREFIX Or _
         Left$(objectName, Len(TEMP_PPT_OLE_PREFIX)) = TEMP_PPT_OLE_PREFIX)
End Function

Private Function GetMigrationLastRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Dim oleItem As OLEObject
    Dim lastRow As Long
    Dim inputLastColumn As String

    inputLastColumn = GetMigrationInputLastColumn(ws)

    On Error Resume Next
    Set lastCell = ws.Range(COL_LEVEL & ":" & inputLastColumn).Find( _
        What:="*", _
        After:=ws.Cells(1, ws.Range(COL_LEVEL & "1").Column), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If Not lastCell Is Nothing Then lastRow = lastCell.Row

    For Each oleItem In ws.OLEObjects
        If IsMigratablePptObject(oleItem) Then
            lastRow = Application.Max(lastRow, GetGanttPptObjectRow(ws, oleItem))
        End If
    Next oleItem

    If lastRow < DATA_START_ROW Then
        GetMigrationLastRow = DATA_START_ROW - 1
    Else
        GetMigrationLastRow = lastRow
    End If
End Function

Private Function GetMigrationInputLastColumn(ByVal ws As Worksheet) As String
    If HasMigrationWebLayout(ws) Then
        If Trim$(CStr(ws.Range(COL_WORK_MINUTES & HEADER_ROW).Value2)) = "작업시간(분)" Then
            GetMigrationInputLastColumn = "AE"
        Else
            GetMigrationInputLastColumn = "AD"
        End If
    ElseIf Trim$(CStr(ws.Range("I" & HEADER_ROW).Value2)) = "수정 내용" Then
        GetMigrationInputLastColumn = "T"
    ElseIf Trim$(CStr(ws.Range("D" & HEADER_ROW).Value2)) = "타입" Then
        GetMigrationInputLastColumn = "S"
    ElseIf Trim$(CStr(ws.Range("E" & HEADER_ROW).Value2)) = "프로그램" Then
        GetMigrationInputLastColumn = "Q"
    Else
        GetMigrationInputLastColumn = "P"
    End If
End Function

Private Function PickLegacyWorkbookPath() As String
    Dim picker As FileDialog

    Set picker = Application.FileDialog(FILE_DIALOG_PICKER)
    With picker
        .Title = "가져올 구버전 Excel 파일 선택"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel 매크로 통합 문서", "*.xlsm"
        .Filters.Add "Excel 통합 문서", "*.xlsx;*.xlsb;*.xls"
        If .Show <> -1 Then Exit Function
        PickLegacyWorkbookPath = .SelectedItems(1)
    End With
End Function

Private Function CountMatchingTaskSheets(ByVal sourceBook As Workbook, _
                                         ByVal targetBook As Workbook) As Long
    Dim sourceSheet As Worksheet

    For Each sourceSheet In sourceBook.Worksheets
        If IsMigrationTaskPair(sourceSheet, targetBook) Then
            CountMatchingTaskSheets = CountMatchingTaskSheets + 1
        End If
    Next sourceSheet
End Function

Private Function HasMigrationWebLayout(ByVal ws As Worksheet) As Boolean
    HasMigrationWebLayout = _
        (Trim$(CStr(ws.Cells(HEADER_ROW, COL_WEB_TASK_ID).Value2)) = "웹업무ID" And _
         Trim$(CStr(ws.Cells(HEADER_ROW, COL_REQUEST).Value2)) = "요청 내용" And _
         Trim$(CStr(ws.Cells(HEADER_ROW, COL_MODIFICATION).Value2)) = "수정 내용")
End Function

Private Function MigrationPreviousColumn(ByVal ws As Worksheet, ByVal columnName As String) As String
    MigrationPreviousColumn = Split(ws.Cells(1, ws.Range(columnName & "1").Column - 1).Address(False, False), "1")(0)
End Function

Private Function IsMigrationTaskPair(ByVal sourceSheet As Worksheet, ByVal targetBook As Workbook) As Boolean
    Dim targetSheet As Worksheet, webHeader As String, timeHeader As String
    If Not IsTaskSheet(sourceSheet) Then Exit Function
    If Trim$(CStr(sourceSheet.Cells(HEADER_ROW, COL_NO).Value2)) <> "No." Then Exit Function
    Set targetSheet = GetWorksheet(targetBook, sourceSheet.Name)
    If targetSheet Is Nothing Then Exit Function
    If Not IsTaskSheet(targetSheet) Then Exit Function
    If Trim$(CStr(targetSheet.Cells(HEADER_ROW, COL_NO).Value2)) <> "No." Then Exit Function
    webHeader = Trim$(CStr(sourceSheet.Cells(HEADER_ROW, COL_WEB_ENABLED).Value2))
    If (webHeader = "WEB UPDATE" Or webHeader = "웹등록") And Not HasMigrationWebLayout(sourceSheet) Then _
        Err.Raise vbObjectError + 2640, "IsMigrationTaskPair", sourceSheet.Name & " 시트의 웹 업무 열 배치를 확인하세요."
    If HasMigrationWebLayout(sourceSheet) Then
        timeHeader = Trim$(CStr(sourceSheet.Cells(HEADER_ROW, COL_WORK_MINUTES).Value2))
        If timeHeader <> "작업시간(분)" And timeHeader <> "진행률" And timeHeader <> "진척률" Then _
            Err.Raise vbObjectError + 2640, "IsMigrationTaskPair", sourceSheet.Name & " 시트의 작업시간·진행률 열 배치를 확인하세요."
    End If
    IsMigrationTaskPair = True
End Function

Private Sub ImportOutputReportConfig(ByVal sourceBook As Workbook, _
                                     ByVal targetBook As Workbook)
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet

    Set sourceWs = GetWorksheet(sourceBook, WEEKLY_REPORT_CONFIG_SHEET_NAME)
    If sourceWs Is Nothing Then Set sourceWs = GetWorksheet(sourceBook, "config_주간보고")
    If sourceWs Is Nothing Then Set sourceWs = GetWorksheet(sourceBook, "config_주간 보고")
    If sourceWs Is Nothing Then Exit Sub
    Set targetWs = GetWorksheet(targetBook, WEEKLY_REPORT_CONFIG_SHEET_NAME)
    If targetWs Is Nothing Then Exit Sub

    targetWs.Range("B2:G5").Value2 = sourceWs.Range("B2:G5").Value2
    targetWs.Range("B7:B8").Value2 = sourceWs.Range("B7:B8").Value2
    targetWs.Range("E8:G14").Value2 = sourceWs.Range("E8:G14").Value2
    targetWs.Range(WR_CUSTOM_PAGE_COLUMN & WR_CUSTOM_START_ROW & ":" & _
                   WR_CUSTOM_PATH_COLUMN & WR_CUSTOM_END_ROW).Value2 = _
        sourceWs.Range(WR_CUSTOM_PAGE_COLUMN & WR_CUSTOM_START_ROW & ":" & _
                       WR_CUSTOM_PATH_COLUMN & WR_CUSTOM_END_ROW).Value2

    If IsDate(sourceWs.Range(OUTPUT_REPORT_START_CELL).Value) Then
        targetWs.Range(OUTPUT_REPORT_START_CELL).Value2 = sourceWs.Range(OUTPUT_REPORT_START_CELL).Value2
    End If
    If IsDate(sourceWs.Range(OUTPUT_REPORT_END_CELL).Value) Then
        targetWs.Range(OUTPUT_REPORT_END_CELL).Value2 = sourceWs.Range(OUTPUT_REPORT_END_CELL).Value2
    End If
    Select Case UCase$(Trim$(CStr(sourceWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2)))
        Case "Y", "N"
            targetWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2 = _
                UCase$(Trim$(CStr(sourceWs.Range(OUTPUT_SHOW_MODIFICATION_CELL).Value2)))
    End Select
End Sub

Private Function IsTaskSheet(ByVal ws As Worksheet) As Boolean
    IsTaskSheet = (Not IsWebUtilitySheet(ws) And StrComp(ws.Name, CONFIG_SHEET_NAME, vbTextCompare) <> 0 And _
                   StrComp(ws.Name, WEEKLY_REPORT_CONFIG_SHEET_NAME, vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "config_주간보고", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "config_주간 보고", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "config_일일현황", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "WeeklyPptTemplate", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "_일별진행현황템플릿", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "_일별진척률이력", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, LEGACY_BUTTON_SETUP_SHEET_NAME, vbTextCompare) <> 0)
End Function

Private Function GetWorksheet(ByVal targetBook As Workbook, _
                              ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = targetBook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim item As Workbook

    For Each item In Application.Workbooks
        If StrComp(item.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = item
            Exit Function
        End If
    Next item
End Function

Private Sub RestoreApplicationState(ByVal screenUpdating As Boolean, _
                                    ByVal enableEvents As Boolean, _
                                    ByVal calculation As XlCalculation)
    Application.CutCopyMode = False
    Application.Calculation = calculation
    Application.EnableEvents = enableEvents
    Application.ScreenUpdating = screenUpdating
End Sub

' 2026-10-23 김좌현 수정 : 업무 ID와 이관 시트의 웹 이력·연결을 사전 검증 후 값으로 함께 이전
Private Function BuildWebMigrationPlan(ByVal sourceBook As Workbook, ByVal targetBook As Workbook) As Object
    Dim plan As Object, pairs As Object, sourceIds As Object, targetIds As Object, retainedIds As Object
    Dim sourceOwners As Object, sourceDuplicates As Object, stateKeys As Object, requestOwners As Object
    Dim logRows As New Collection, stateRows As New Collection, sourceLogRows As New Collection
    Dim sourceStateRows As New Collection, ws As Worksheet, log As Worksheet, state As Worksheet
    Dim entry As Object, rowValues As Variant, key As Variant, sourceId As String, sheetName As String
    Dim r As Long, c As Long, lastRow As Long, importedLogs As Long, importedStates As Long
    Set plan = CreateObject("Scripting.Dictionary")
    Set pairs = WebMigrationDictionary()
    Set sourceIds = WebMigrationDictionary()
    Set targetIds = WebMigrationDictionary()
    Set retainedIds = WebMigrationDictionary()
    Set sourceOwners = WebMigrationDictionary()
    Set sourceDuplicates = WebMigrationDictionary()
    Set stateKeys = WebMigrationDictionary()
    Set requestOwners = WebMigrationDictionary()

    For Each ws In sourceBook.Worksheets
        If IsMigrationTaskPair(ws, targetBook) Then
            pairs(ws.Name) = True
            WebMigrationCollectIds ws, sourceIds
            WebMigrationCollectIds GetWorksheet(targetBook, ws.Name), targetIds
        End If
        If WebMigrationHasTaskIds(ws) Then
            For r = DATA_START_ROW To GetMigrationLastRow(ws)
                sourceId = WebMigrationCellText(ws, r, COL_WEB_TASK_ID)
                If Len(sourceId) > 0 And HasTaskContent(ws, r) Then
                    If sourceOwners.Exists(sourceId) Then sourceDuplicates(sourceId) = True
                    sourceOwners(sourceId) = ws.Name & " " & CStr(r) & "행"
                End If
            Next r
        End If
    Next ws
    For Each ws In targetBook.Worksheets
        If Not pairs.Exists(ws.Name) Then WebMigrationCollectIds ws, retainedIds
    Next ws

    Set log = GetWorksheet(sourceBook, WEB_LOG_SHEET_NAME)
    WebMigrationCollectLogIds log, pairs, sourceIds
    Set log = GetWorksheet(targetBook, WEB_LOG_SHEET_NAME)
    WebMigrationCollectLogIds log, pairs, targetIds
    For Each key In sourceIds.Keys
        If sourceDuplicates.Exists(CStr(key)) Then _
            Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "원본 웹업무ID가 중복되어 이력을 가져올 수 없습니다: " & CStr(key)
        If retainedIds.Exists(CStr(key)) Then _
            Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "가져올 웹업무ID가 가져오지 않는 대상 업무 시트와 중복됩니다: " & CStr(key)
    Next key
    For Each key In targetIds.Keys
        If retainedIds.Exists(CStr(key)) Then _
            Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "교체할 웹 이력이 가져오지 않는 대상 업무에서도 사용됩니다: " & CStr(key)
    Next key

    Set log = GetWorksheet(sourceBook, WEB_LOG_SHEET_NAME)
    If Not log Is Nothing Then
        lastRow = WebMigrationLastRow(log, "M")
        For r = 2 To lastRow
            sourceId = WebMigrationCellText(log, r, "C")
            sheetName = WebMigrationCellText(log, r, "B")
            If sourceIds.Exists(sourceId) Or pairs.Exists(sheetName) Then
                sourceLogRows.Add WebMigrationLogEntry(log, r, True)
                importedLogs = importedLogs + 1
            End If
        Next r
    End If
    Set log = GetWorksheet(targetBook, WEB_LOG_SHEET_NAME)
    If Not log Is Nothing Then
        lastRow = WebMigrationLastRow(log, "M")
        For r = 2 To lastRow
            sourceId = WebMigrationCellText(log, r, "C")
            sheetName = WebMigrationCellText(log, r, "B")
            If Not sourceIds.Exists(sourceId) And Not targetIds.Exists(sourceId) And Not pairs.Exists(sheetName) Then _
                logRows.Add WebMigrationLogEntry(log, r)
        Next r
    End If
    For Each entry In sourceLogRows
        logRows.Add entry
    Next entry

    Set state = GetWorksheet(targetBook, WEB_STATE_SHEET_NAME)
    If Not state Is Nothing Then
        lastRow = WebMigrationLastRow(state, "O")
        For r = 2 To lastRow
            sourceId = WebMigrationCellText(state, r, "A")
            If Len(sourceId) > 0 And Not sourceIds.Exists(sourceId) And Not targetIds.Exists(sourceId) Then
                rowValues = state.Range("A" & r & ":O" & r).Value2
                WebMigrationValidateState rowValues, stateKeys, requestOwners
                stateRows.Add rowValues
            End If
        Next r
    End If
    Set state = GetWorksheet(sourceBook, WEB_STATE_SHEET_NAME)
    If Not state Is Nothing Then
        lastRow = WebMigrationLastRow(state, "O")
        For r = 2 To lastRow
            sourceId = WebMigrationCellText(state, r, "A")
            If sourceIds.Exists(sourceId) Then
                rowValues = state.Range("A" & r & ":O" & r).Value2
                For c = 1 To 15
                    sheetName = WebMigrationRowText(rowValues, c)
                Next c
                WebMigrationValidateState rowValues, stateKeys, requestOwners
                stateRows.Add rowValues
                importedStates = importedStates + 1
            End If
        Next r
    End If
    plan.Add "Logs", logRows
    plan.Add "States", stateRows
    plan.Add "ImportedLogCount", importedLogs
    plan.Add "ImportedStateCount", importedStates
    Set BuildWebMigrationPlan = plan
End Function

Private Sub ApplyWebMigrationPlan(ByVal plan As Object, ByVal targetBook As Workbook, _
                                  ByRef importedLogCount As Long, ByRef importedStateCount As Long)
    Dim log As Worksheet, state As Worksheet, entry As Object, rowValues As Variant
    Dim logRows As Collection, stateRows As Collection, r As Long, lastRow As Long
    Dim logProtected As Boolean, stateProtected As Boolean, stateVisibility As XlSheetVisibility
    Dim errorNumber As Long, errorText As String
    Set log = GetWorksheet(targetBook, WEB_LOG_SHEET_NAME)
    Set state = GetWorksheet(targetBook, WEB_STATE_SHEET_NAME)
    If log Is Nothing Or state Is Nothing Then _
        Err.Raise vbObjectError + 2641, "ApplyWebMigrationPlan", "웹 이력과 연결 시트를 먼저 준비하세요."
    Set logRows = plan("Logs")
    Set stateRows = plan("States")
    logProtected = log.ProtectContents
    stateProtected = state.ProtectContents
    stateVisibility = state.Visible
    On Error GoTo Failed
    log.Unprotect
    state.Unprotect
    lastRow = WebMigrationLastRow(log, "M")
    If lastRow >= 2 Then
        log.Range("A2:M" & lastRow).Hyperlinks.Delete
        log.Range("A2:M" & lastRow).ClearContents
    End If
    lastRow = WebMigrationLastRow(state, "O")
    If lastRow >= 2 Then state.Range("A2:O" & lastRow).ClearContents
    If logRows.Count > 0 Then
        log.Range("B2:G" & CStr(logRows.Count + 1)).NumberFormat = "@"
        log.Range("K2:M" & CStr(logRows.Count + 1)).NumberFormat = "@"
    End If
    r = 2
    For Each entry In logRows
        rowValues = entry("Values")
        log.Range("A" & r & ":M" & r).Value2 = rowValues
        If Len(CStr(entry("Link"))) > 0 Then
            log.Hyperlinks.Add Anchor:=log.Cells(r, "D"), Address:=CStr(entry("Link")), _
                TextToDisplay:=WebMigrationRowText(rowValues, 4)
        End If
        r = r + 1
    Next entry
    If stateRows.Count > 0 Then state.Range("A2:O" & CStr(stateRows.Count + 1)).NumberFormat = "@"
    r = 2
    For Each rowValues In stateRows
        state.Range("A" & r & ":O" & r).Value2 = rowValues
        r = r + 1
    Next rowValues
    state.Visible = stateVisibility
    If logProtected Then log.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    If stateProtected Then state.Protect UserInterfaceOnly:=True
    importedLogCount = CLng(plan("ImportedLogCount"))
    importedStateCount = CLng(plan("ImportedStateCount"))
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    state.Visible = stateVisibility
    If logProtected Then log.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    If stateProtected Then state.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise errorNumber, "ApplyWebMigrationPlan", errorText
End Sub

Private Function WebMigrationDictionary() As Object
    Set WebMigrationDictionary = CreateObject("Scripting.Dictionary")
    WebMigrationDictionary.CompareMode = vbTextCompare
End Function

Private Function WebMigrationHasTaskIds(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Not IsTaskSheet(ws) Then Exit Function
    WebMigrationHasTaskIds = (WebMigrationCellText(ws, HEADER_ROW, COL_WEB_TASK_ID) = "웹업무ID" And _
                              WebMigrationCellText(ws, HEADER_ROW, COL_REQUEST) = "요청 내용")
End Function

Private Sub WebMigrationCollectIds(ByVal ws As Worksheet, ByVal ids As Object)
    Dim r As Long, sourceId As String
    If Not WebMigrationHasTaskIds(ws) Then Exit Sub
    For r = DATA_START_ROW To GetMigrationLastRow(ws)
        If HasTaskContent(ws, r) Then
            sourceId = WebMigrationCellText(ws, r, COL_WEB_TASK_ID)
            If Len(sourceId) > 0 Then ids(sourceId) = True
        End If
    Next r
End Sub

Private Sub WebMigrationCollectLogIds(ByVal log As Worksheet, ByVal pairs As Object, ByVal ids As Object)
    Dim r As Long, sourceId As String
    If log Is Nothing Then Exit Sub
    For r = 2 To WebMigrationLastRow(log, "M")
        If pairs.Exists(WebMigrationCellText(log, r, "B")) Then
            sourceId = WebMigrationCellText(log, r, "C")
            If Len(sourceId) > 0 Then ids(sourceId) = True
        End If
    Next r
End Sub

Private Function WebMigrationLastRow(ByVal ws As Worksheet, ByVal lastColumn As String) As Long
    Dim found As Range
    Set found = ws.Range("A:" & lastColumn).Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                  LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    WebMigrationLastRow = 1
    If Not found Is Nothing Then WebMigrationLastRow = found.Row
End Function

Private Function WebMigrationCellText(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal columnName As String) As String
    Dim value As Variant
    value = ws.Cells(rowNum, columnName).Value2
    If IsError(value) Then Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", _
        "웹 이관에 필요한 셀에 오류가 있습니다: " & ws.Name & "!" & columnName & CStr(rowNum)
    WebMigrationCellText = Trim$(CStr(value))
End Function

Private Function WebMigrationRowText(ByRef values As Variant, ByVal columnNum As Long) As String
    If IsError(values(1, columnNum)) Then _
        Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "웹 이력 또는 연결 값에 셀 오류가 있습니다."
    WebMigrationRowText = CStr(values(1, columnNum))
End Function

Private Function WebMigrationLogEntry(ByVal ws As Worksheet, ByVal rowNum As Long, Optional ByVal importing As Boolean = False) As Object
    Dim entry As Object, values As Variant, link As String, site As String, requestNo As String
    Dim sourceConfig As Worksheet, formPosition As Long
    Set entry = CreateObject("Scripting.Dictionary")
    values = ws.Range("A" & rowNum & ":M" & rowNum).Value2
    requestNo = WebMigrationRowText(values, 4)
    If ws.Cells(rowNum, "D").Hyperlinks.Count > 0 Then link = ws.Cells(rowNum, "D").Hyperlinks(1).Address
    If importing And Len(WebMigrationRowText(values, 12)) = 0 And Len(WebMigrationRowText(values, 3)) > 0 Then
        formPosition = InStr(1, link, "/C0/", vbTextCompare)
        If formPosition > 0 Then site = WebMigrationCanonicalSite(Left$(link, formPosition))
        If Len(site) = 0 Then site = WebMigrationLegacyStateSite(ws.Parent, WebMigrationRowText(values, 3), requestNo)
        Set sourceConfig = GetWorksheet(ws.Parent, WEB_CONFIG_SHEET_NAME)
        If Len(site) = 0 And Not sourceConfig Is Nothing Then site = WebMigrationCanonicalSite(WebMigrationCellText(sourceConfig, 2, "B"))
        If Len(site) = 0 Then Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", _
            "이전 웹 이력의 사이트를 확인할 수 없습니다. 원본 파일의 config_web 사이트 주소를 확인하세요."
        values(1, 12) = site
        If Len(WebMigrationRowText(values, 13)) = 0 And Not sourceConfig Is Nothing Then _
            values(1, 13) = WebMigrationCellText(sourceConfig, 3, "B")
    End If
    If Len(link) = 0 And Len(requestNo) > 0 Then
        site = WebMigrationCanonicalSite(WebMigrationRowText(values, 12))
        If Len(site) > 0 Then link = site & "C0/c0200_02.aspx?REQ_NO=" & requestNo
    End If
    entry.Add "Values", values
    entry.Add "Link", link
    Set WebMigrationLogEntry = entry
End Function

Private Function WebMigrationLegacyStateSite(ByVal sourceBook As Workbook, ByVal sourceId As String, ByVal requestNo As String) As String
    Dim state As Worksheet, r As Long, site As String
    If Len(requestNo) = 0 Then Exit Function
    Set state = GetWorksheet(sourceBook, WEB_STATE_SHEET_NAME)
    If state Is Nothing Then Exit Function
    For r = 2 To WebMigrationLastRow(state, "O")
        If WebMigrationCellText(state, r, "A") = sourceId Then
            If WebMigrationCellText(state, r, "D") = requestNo Or WebMigrationCellText(state, r, "E") = requestNo Then
                site = WebMigrationCanonicalSite(WebMigrationCellText(state, r, "B"))
                If Len(site) > 0 Then
                    If Len(WebMigrationLegacyStateSite) > 0 And site <> WebMigrationLegacyStateSite Then _
                        Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "이전 웹 이력의 요청번호가 여러 사이트에 있어 원본 연결을 확인해야 합니다."
                    WebMigrationLegacyStateSite = site
                End If
            End If
        End If
    Next r
End Function

Private Function WebMigrationCanonicalSite(ByVal site As String) As String
    site = LCase$(Trim$(site))
    If Len(site) > 0 Then
        If Right$(site, 1) <> "/" Then site = site & "/"
    End If
    WebMigrationCanonicalSite = site
End Function

Private Sub WebMigrationValidateState(ByRef values As Variant, ByVal stateKeys As Object, ByVal requestOwners As Object)
    Dim sourceId As String, site As String, key As String, requestNo As String, columnNum As Variant
    sourceId = WebMigrationRowText(values, 1)
    site = WebMigrationCanonicalSite(WebMigrationRowText(values, 2))
    key = site & vbNullChar & sourceId
    If stateKeys.Exists(key) Then Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", _
        "같은 웹업무ID와 사이트의 연결 이력이 중복됩니다: " & sourceId
    stateKeys(key) = True
    For Each columnNum In Array(4, 5)
        requestNo = WebMigrationRowText(values, CLng(columnNum))
        If Len(requestNo) > 0 Then
            key = site & vbNullChar & requestNo
            If requestOwners.Exists(key) Then
                If StrComp(CStr(requestOwners(key)), sourceId, vbTextCompare) <> 0 Then _
                    Err.Raise vbObjectError + 2640, "BuildWebMigrationPlan", "같은 사이트의 웹 요청번호가 서로 다른 업무에 연결됩니다: " & requestNo
            Else
                requestOwners(key) = sourceId
            End If
        End If
    Next columnNum
End Sub
