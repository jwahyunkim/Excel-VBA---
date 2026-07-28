Attribute VB_Name = "modVersionMigration"
Option Explicit

Private Const AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3
Private Const FILE_DIALOG_PICKER As Long = 3
Private Const PPT_OLE_PREFIX As String = "oleGanttPpt_"
Private Const TEMP_PPT_OLE_PREFIX As String = "tmpGanttPptOle_"
Private Const BUTTON_SETUP_SHEET_NAME As String = "_버튼생성"

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

    For Each sourceSheet In sourceBook.Worksheets
        If IsTaskSheet(sourceSheet) Then
            Set targetSheet = GetWorksheet(targetBook, sourceSheet.Name)
            If Not targetSheet Is Nothing Then
                ImportTaskSheet sourceSheet, targetSheet, copiedObjectCount, failedObjectCount
                importedSheetCount = importedSheetCount + 1
            End If
        End If
    Next sourceSheet

    For Each sourceSheet In sourceBook.Worksheets
        If IsTaskSheet(sourceSheet) Then
            Set targetSheet = GetWorksheet(targetBook, sourceSheet.Name)
            If Not targetSheet Is Nothing Then
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

Public Sub 데이터_가져오기버튼_생성()
    Dim ws As Worksheet

    On Error GoTo EH

    Set ws = ActiveSheet
    If Not IsTaskSheet(ws) Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    UnprotectTaskSheet ws

    On Error Resume Next
    ws.Shapes("btnLegacyImport").Delete
    ws.Shapes("btnButtonImport").Delete
    On Error GoTo EH

    CreateVersionButton ws, "btnDataImport", "데이터 가져오기", "데이터_가져오기", 1, 72

    ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
    MsgBox "데이터 가져오기 버튼 생성 완료", vbInformation
    Exit Sub

EH:
    MsgBox "버튼을 생성하는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 버튼_생성_선택()
    Dim targetSheet As Worksheet
    Dim setupSheet As Worksheet
    Dim optionCaptions As Variant
    Dim optionIndex As Long
    Dim optionRow As Long
    Dim checkBox As CheckBox
    Dim orderDropDown As DropDown
    Dim orderValue As Long
    Dim actionButton As Shape
    Dim previousEnableEvents As Boolean
    Dim eventsStateCaptured As Boolean

    On Error GoTo EH

    Set targetSheet = ActiveSheet
    If Not IsTaskSheet(targetSheet) Then
        MsgBox "업무 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    previousEnableEvents = Application.EnableEvents
    eventsStateCaptured = True
    Application.EnableEvents = False

    Set setupSheet = GetOrCreateButtonSetupSheet()
    setupSheet.Visible = xlSheetVisible
    setupSheet.Cells.Clear

    On Error Resume Next
    setupSheet.CheckBoxes.Delete
    setupSheet.DropDowns.Delete
    setupSheet.Shapes("btnCreateSelectedButtons").Delete
    setupSheet.Shapes("btnCancelButtonSetup").Delete
    On Error GoTo EH

    setupSheet.Range("A1:B1").Merge
    setupSheet.Range("A1").Value = "생성할 버튼 선택"
    setupSheet.Range("A1").Font.Bold = True
    setupSheet.Range("A1").Font.Size = 14
    setupSheet.Range("A1").HorizontalAlignment = xlCenter

    optionCaptions = Array( _
        "데이터 가져오기", _
        "간트차트 생성", _
        "새로고침", _
        "초기화", _
        "항목 숨김", _
        "개체삽입", _
        "개인 보고", _
        "팀 보고", _
        "모듈 보고", _
        "주간 PPT")

    setupSheet.Range("A2").Value = "버튼"
    setupSheet.Range("B2").Value = "순서"
    setupSheet.Range("A2:B2").Font.Bold = True
    setupSheet.Range("A2:B2").HorizontalAlignment = xlCenter

    For optionIndex = LBound(optionCaptions) To UBound(optionCaptions)
        optionRow = optionIndex + 3

        Set checkBox = setupSheet.CheckBoxes.Add( _
            setupSheet.Cells(optionRow, "A").Left + 6, _
            setupSheet.Cells(optionRow, "A").Top + 2, _
            170, _
            setupSheet.Cells(optionRow, "A").Height - 2)
        With checkBox
            .Name = "chkButtonOption" & CStr(optionIndex + 1)
            .Caption = CStr(optionCaptions(optionIndex))
            .Value = xlOn
        End With

        Set orderDropDown = setupSheet.DropDowns.Add( _
            setupSheet.Cells(optionRow, "B").Left + 5, _
            setupSheet.Cells(optionRow, "B").Top + 2, _
            55, _
            setupSheet.Cells(optionRow, "B").Height - 2)
        With orderDropDown
            .Name = "ddlButtonOrder" & CStr(optionIndex + 1)
            For orderValue = 1 To UBound(optionCaptions) + 1
                .AddItem CStr(orderValue)
            Next orderValue
            .ListIndex = optionIndex + 1
        End With
    Next optionIndex

    setupSheet.Range("Z1").Value = targetSheet.Name
    setupSheet.Columns("Z").Hidden = True
    setupSheet.Columns("A").ColumnWidth = 25
    setupSheet.Columns("B").ColumnWidth = 10
    setupSheet.Rows("1:15").RowHeight = 24

    Set actionButton = setupSheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        setupSheet.Range("A14").Left, _
        setupSheet.Range("A14").Top, _
        90, _
        28)
    ConfigureSetupActionButton actionButton, _
        "btnCreateSelectedButtons", "생성", "선택한_버튼_생성", RGB(91, 155, 213)

    Set actionButton = setupSheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        setupSheet.Range("B14").Left, _
        setupSheet.Range("B14").Top, _
        90, _
        28)
    ConfigureSetupActionButton actionButton, _
        "btnCancelButtonSetup", "취소", "버튼_생성_취소", RGB(166, 166, 166)

    setupSheet.Activate
    setupSheet.Range("A1").Select
    Application.EnableEvents = previousEnableEvents
    Exit Sub

EH:
    If eventsStateCaptured Then Application.EnableEvents = previousEnableEvents
    MsgBox "버튼 선택 화면을 만드는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub 선택한_버튼_생성(Optional ByVal showCompletionMessage As Boolean = True)
    Dim setupSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim targetSheetName As String
    Dim selectedIndexes(1 To 10) As Long
    Dim selectedOrders(1 To 10) As Long
    Dim selectedCount As Long
    Dim optionIndex As Long
    Dim optionOrder As Long
    Dim firstIndex As Long
    Dim secondIndex As Long
    Dim swapValue As Long

    On Error GoTo EH

    Set setupSheet = ThisWorkbook.Worksheets(BUTTON_SETUP_SHEET_NAME)
    targetSheetName = CStr(setupSheet.Range("Z1").Value2)
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)

    For optionIndex = 1 To 10
        If IsSetupOptionChecked(setupSheet, optionIndex) Then
            optionOrder = GetSetupOptionOrder(setupSheet, optionIndex)
            If optionOrder < 1 Then
                MsgBox "선택한 모든 버튼의 순서를 지정하세요.", vbExclamation
                Exit Sub
            End If

            For firstIndex = 1 To selectedCount
                If selectedOrders(firstIndex) = optionOrder Then
                    MsgBox "버튼 순서 " & optionOrder & "번이 중복되었습니다.", vbExclamation
                    Exit Sub
                End If
            Next firstIndex

            selectedCount = selectedCount + 1
            selectedIndexes(selectedCount) = optionIndex
            selectedOrders(selectedCount) = optionOrder
        End If
    Next optionIndex

    If selectedCount = 0 Then
        MsgBox "생성할 버튼을 하나 이상 선택하세요.", vbExclamation
        Exit Sub
    End If

    For firstIndex = 1 To selectedCount - 1
        For secondIndex = firstIndex + 1 To selectedCount
            If selectedOrders(firstIndex) > selectedOrders(secondIndex) Then
                swapValue = selectedOrders(firstIndex)
                selectedOrders(firstIndex) = selectedOrders(secondIndex)
                selectedOrders(secondIndex) = swapValue

                swapValue = selectedIndexes(firstIndex)
                selectedIndexes(firstIndex) = selectedIndexes(secondIndex)
                selectedIndexes(secondIndex) = swapValue
            End If
        Next secondIndex
    Next firstIndex

    UnprotectTaskSheet targetSheet
    DeleteManagedButtons targetSheet

    For firstIndex = 1 To selectedCount
        CreateSelectedButtonByIndex targetSheet, selectedIndexes(firstIndex), firstIndex
    Next firstIndex

    ApplyCalculatedColumnsProtection targetSheet, GetLastDataRow(targetSheet)
    targetSheet.Activate
    setupSheet.Visible = xlSheetVeryHidden

    If showCompletionMessage Then
        MsgBox CStr(selectedCount) & "개의 버튼을 지정한 순서대로 생성했습니다.", vbInformation
    End If
    Exit Sub

EH:
    MsgBox "선택한 버튼을 생성하는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function IsSetupOptionChecked(ByVal setupSheet As Worksheet, _
                                      ByVal optionIndex As Long) As Boolean
    IsSetupOptionChecked = _
        (setupSheet.CheckBoxes("chkButtonOption" & CStr(optionIndex)).Value = xlOn)
End Function

Private Function GetSetupOptionOrder(ByVal setupSheet As Worksheet, _
                                     ByVal optionIndex As Long) As Long
    GetSetupOptionOrder = _
        setupSheet.DropDowns("ddlButtonOrder" & CStr(optionIndex)).ListIndex
End Function

Private Sub CreateSelectedButtonByIndex(ByVal ws As Worksheet, _
                                        ByVal optionIndex As Long, _
                                        ByVal buttonOrder As Long)
    Select Case optionIndex
        Case 1
            CreateVersionButton ws, "btnDataImport", "데이터 가져오기", "데이터_가져오기", buttonOrder, 72
        Case 2
            CreateVersionButton ws, "btnGanttCreate", "생성", "칸트차트_생성", buttonOrder, 72
        Case 3
            CreateVersionButton ws, "btnGanttRefresh", "새로고침", "칸트차트_새로고침", buttonOrder, 72
        Case 4
            CreateVersionButton ws, "btnGanttReset", "초기화", "칸트차트_초기화", buttonOrder, 72
        Case 5
            CreateVersionButton ws, "btnGanttHideTask", "항목 숨김", "칸트차트_항목숨기기", buttonOrder, 72
        Case 6
            CreateVersionButton ws, "btnGanttObjectInsert", "개체삽입", "칸트차트_개체삽입", buttonOrder, 72
        Case 7
            CreateVersionButton ws, "btnPersonalDevReport", "개인 보고", "개인개발보고_텍스트생성", buttonOrder, 72
        Case 8
            CreateVersionButton ws, "btnTeamDevReport", "팀 보고", "팀개발보고_텍스트생성", buttonOrder, 72
        Case 9
            CreateVersionButton ws, "btnModuleDevReport", "모듈 보고", "모듈개발보고_텍스트생성", buttonOrder, 72
        Case 10
            CreateVersionButton ws, "btnWeeklyPptReport", "주간 PPT", "주간보고PPT_생성", buttonOrder, 72
    End Select
End Sub

Public Sub 버튼_생성_취소(Optional ByVal internalCall As Boolean = True)
    Dim setupSheet As Worksheet
    Dim targetSheetName As String

    On Error Resume Next
    Set setupSheet = ThisWorkbook.Worksheets(BUTTON_SETUP_SHEET_NAME)
    targetSheetName = CStr(setupSheet.Range("Z1").Value2)
    ThisWorkbook.Worksheets(targetSheetName).Activate
    setupSheet.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Function GetOrCreateButtonSetupSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateButtonSetupSheet = ThisWorkbook.Worksheets(BUTTON_SETUP_SHEET_NAME)
    On Error GoTo 0

    If GetOrCreateButtonSetupSheet Is Nothing Then
        Set GetOrCreateButtonSetupSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateButtonSetupSheet.Name = BUTTON_SETUP_SHEET_NAME
    End If
End Function

Private Sub DeleteManagedButtons(ByVal ws As Worksheet)
    Dim buttonNames As Variant
    Dim buttonName As Variant

    buttonNames = Array( _
        "btnDataImport", "btnGanttCreate", "btnGanttRefresh", _
        "btnGanttReset", "btnGanttHideTask", "btnGanttObjectInsert", _
        "btnPersonalDevReport", "btnTeamDevReport", "btnModuleDevReport", _
        "btnWeeklyPptReport", _
        "btnLegacyImport", "btnButtonImport", "btnDevProgressReport")

    On Error Resume Next
    For Each buttonName In buttonNames
        ws.Shapes(CStr(buttonName)).Delete
    Next buttonName
    On Error GoTo 0
End Sub

Private Sub ConfigureSetupActionButton(ByVal button As Shape, _
                                       ByVal buttonName As String, _
                                       ByVal buttonCaption As String, _
                                       ByVal macroName As String, _
                                       ByVal fillColor As Long)
    With button
        .Name = buttonName
        .OnAction = macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = RGB(100, 100, 100)
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Characters.Text = buttonCaption
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "맑은 고딕"
                .Font.Size = 10
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
        End With
    End With
End Sub

Private Sub CreateVersionButton(ByVal ws As Worksheet, _
                                ByVal buttonName As String, _
                                ByVal buttonCaption As String, _
                                ByVal macroName As String, _
                                ByVal buttonOrder As Long, _
                                ByVal buttonWidth As Double)
    Dim button As Shape
    Dim anchorCell As Range
    Dim buttonLeft As Double
    Dim buttonTop As Double

    Set anchorCell = ws.Range("B2")

    buttonLeft = anchorCell.Left + ((buttonOrder - 1) * 80)
    buttonTop = anchorCell.Top + anchorCell.Height - 22

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0

    Set button = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, 22)

    With button
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
                .Font.Name = "맑은 고딕"
                .Font.Size = IIf(buttonName = "btnDataImport", 8, 9)
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
        End With
    End With
End Sub

Private Sub ImportTaskSheet(ByVal sourceSheet As Worksheet, _
                            ByVal targetSheet As Worksheet, _
                            ByRef copiedObjectCount As Long, _
                            ByRef failedObjectCount As Long)
    Dim sourceLastRow As Long
    Dim targetLastRow As Long
    Dim clearLastRow As Long
    Dim rowNum As Long
    Dim sourceOle As OLEObject

    sourceLastRow = GetMigrationLastRow(sourceSheet)
    targetLastRow = GetMigrationLastRow(targetSheet)
    clearLastRow = sourceLastRow
    If targetLastRow > clearLastRow Then clearLastRow = targetLastRow
    If clearLastRow < DATA_START_ROW Then clearLastRow = DATA_START_ROW

    UnprotectTaskSheet targetSheet
    DeleteTargetPptObjects targetSheet
    ClearTargetPptIcons targetSheet, clearLastRow

    targetSheet.Range(COL_LEVEL & DATA_START_ROW & ":" & _
                      COL_PROGRESS & clearLastRow).ClearContents
    targetSheet.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & _
                      COL_DEV_PROGRESS & clearLastRow).ClearContents

    If sourceLastRow >= DATA_START_ROW Then
        targetSheet.Range(COL_LEVEL & DATA_START_ROW & ":" & _
                          COL_PROGRESS & sourceLastRow).Value2 = _
            sourceSheet.Range(COL_LEVEL & DATA_START_ROW & ":" & _
                              COL_PROGRESS & sourceLastRow).Value2

        targetSheet.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & _
                          COL_DEV_PROGRESS & sourceLastRow).Value2 = _
            sourceSheet.Range(COL_MANUAL_PROGRESS & DATA_START_ROW & ":" & _
                              COL_DEV_PROGRESS & sourceLastRow).Value2

        For rowNum = DATA_START_ROW To sourceLastRow
            If CStr(targetSheet.Cells(rowNum, COL_NOTE).Value2) = ChrW(&H25A0) Then
                targetSheet.Cells(rowNum, COL_NOTE).ClearContents
            End If
        Next rowNum
    End If

    For Each sourceOle In sourceSheet.OLEObjects
        If IsMigratablePptObject(sourceOle) Then
            rowNum = sourceOle.TopLeftCell.Row
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

Private Function CopyPptObject(ByVal sourceOle As OLEObject, _
                               ByVal targetSheet As Worksheet, _
                               ByVal targetRow As Long) As Boolean
    Dim pastedOle As OLEObject
    Dim beforeCount As Long

    On Error GoTo CopyFailed

    sourceOle.Parent.Activate
    sourceOle.Copy
    targetSheet.Activate

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
    ApplyDisplayTaskRowFilter ws, lastRow, chartStartDate, chartEndDate
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

    On Error Resume Next
    Set lastCell = ws.Range(COL_LEVEL & ":" & COL_DEV_PROGRESS).Find( _
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
            If oleItem.TopLeftCell.Row > lastRow Then lastRow = oleItem.TopLeftCell.Row
        End If
    Next oleItem

    If lastRow < DATA_START_ROW Then
        GetMigrationLastRow = DATA_START_ROW - 1
    Else
        GetMigrationLastRow = lastRow
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
        If IsTaskSheet(sourceSheet) Then
            If Not GetWorksheet(targetBook, sourceSheet.Name) Is Nothing Then
                CountMatchingTaskSheets = CountMatchingTaskSheets + 1
            End If
        End If
    Next sourceSheet
End Function

Private Function IsTaskSheet(ByVal ws As Worksheet) As Boolean
    IsTaskSheet = (StrComp(ws.Name, CONFIG_SHEET_NAME, vbTextCompare) <> 0 And _
                   StrComp(ws.Name, REPORT_HISTORY_SHEET_NAME, vbTextCompare) <> 0 And _
                   StrComp(ws.Name, "WeeklyPptTemplate", vbTextCompare) <> 0 And _
                   StrComp(ws.Name, BUTTON_SETUP_SHEET_NAME, vbTextCompare) <> 0)
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
