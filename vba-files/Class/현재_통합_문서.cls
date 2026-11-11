VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' 2026-11-11 김좌현 수정 : 비고 PPT 셀 우클릭 시 첨부 삭제 메뉴 연결
' 2026-11-04 김좌현 수정 : CONFIG 내용 표시 변경 즉시 반영 및 입력한 내용 행의 줄바꿈 유지
' 2026-11-04 김좌현 수정 : 이벤트를 끈 가져오기·보안 저장 이후 새로고침 단축키 명시적 복원
' 2026-11-04 김좌현 수정 : 닫기 취소 시 단축키 유지, 활성 통합문서만 등록하고 창/시트 전환에서 복원
' 2026-10-29 김좌현 수정 : 새로고침 단축키 연결 통일 및 미사용 캐시 이벤트 제거
' 2026-10-27 김좌현 수정 : 입력한 셀의 업무 구조 변경 여부와 타입 선택목록만 검사해 전체 재계산 축소
' 2026-10-26 김좌현 수정 : 타입 오류로 실행 취소한 행 변경의 웹 설정 비교 상태 복원
' 2026-10-22 김좌현 수정 : 간트 셀 직접 수정 시 간트 전용 새로고침 캐시 무효화
' 2026-10-22 김좌현 수정 : 웹등록 초기화, 보호 비밀번호와 그룹별 등록 기준 이벤트 연결

Private mGanttShortcutsRegistered As Boolean

Private Sub Workbook_Open()
    ' 2026-10-20 김좌현 수정 : 기존 생성 버튼을 새로고침으로 전환
    Dim ws As Worksheet
    ' 2026-10-20 김좌현 수정 : 저장된 완료 숨김 상태에 맞춰 버튼 문구 복원
    InitializeReleaseSecurity
    RegisterGanttShortcuts
    InitializeWebUpload
    For Each ws In ThisWorkbook.Worksheets
        UpgradeGanttRefreshButtons ws
        UpdateCompletedTaskButton ws
    Next ws
    InitializeTaskHierarchyOutlines
    StartTaskOutlineViewMonitor
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim webConfig As Worksheet
    On Error GoTo WebPasswordError
    On Error Resume Next
    Set webConfig = ThisWorkbook.Worksheets(WEB_CONFIG_SHEET_NAME)
    On Error GoTo WebPasswordError
    If Not webConfig Is Nothing Then
        If Len(CStr(webConfig.Range("B4").Value2)) > 0 And CStr(webConfig.Range("B4").Value2) <> "********" Then
            SetWebPasswordValue CStr(webConfig.Range("B4").Value2)
        End If
    End If
    StopTaskOutlineViewMonitor
    PrepareReleaseForSave SaveAsUI, Cancel
    If Cancel And mGanttShortcutsRegistered Then RegisterGanttShortcuts
    If Cancel Then StartTaskOutlineViewMonitor
    Exit Sub
WebPasswordError:
    Cancel = True
    MsgBox "비밀번호를 보호 저장하지 못해 저장을 중단했습니다. config_web에서 다시 입력하세요.", vbExclamation
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next
    If TypeOf Sh Is Worksheet Then
        If Sh.Name = "배포정보" Then HideReleaseInfoSheet
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim duplicateModule As String
    Dim textChanges As Range, textArea As Range

    On Error GoTo EH

    If Not TypeOf Sh Is Worksheet Then Exit Sub

    If IsWebUtilitySheet(Sh) Then
        If Sh.Name = WEB_CONFIG_SHEET_NAME Then
            If Not Intersect(Target, Sh.Range("B4")) Is Nothing Then
                Application.EnableEvents = False
                If CStr(Sh.Range("B4").Value2) <> "********" Then SetWebPasswordValue CStr(Sh.Range("B4").Value2)
            End If
        End If
        GoTo SafeExit
    End If

    If Sh.Name = CONFIG_SHEET_NAME Or _
       Sh.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Then
        If Sh.Name = CONFIG_SHEET_NAME And _
           Not Intersect(Target, Sh.Range(TASK_FOLD_SETTING_VALUE_RANGE)) Is Nothing Then
            Application.enableEvents = False
            ApplyTaskDetailColumnSettings
        End If
        If Sh.Name = CONFIG_SHEET_NAME And _
           Not Intersect(Target, Sh.Range(TASK_TEXT_DISPLAY_VALUE_CELL)) Is Nothing Then
            Application.EnableEvents = False
            ApplyTaskTextDisplaySettings
        End If
        If Sh.Name = CONFIG_SHEET_NAME And _
           Not Intersect(Target, Sh.Range(TASK_TYPE_CONFIG_RANGE)) Is Nothing Then
            Application.enableEvents = False
            RefreshTaskTypeValidation
        End If
        If Sh.Name = CONFIG_SHEET_NAME And _
           Not Intersect(Target, Sh.Range(TASK_MAX_LENGTH_VALUE_RANGE)) Is Nothing Then
            Application.enableEvents = False
            RefreshTaskTextLengthValidation
        ElseIf Sh.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME And _
               Not Intersect( _
                       Target, _
                       Sh.Range(WR_CUSTOM_PATH_COLUMN & _
                                CStr(WR_CUSTOM_START_ROW) & ":" & _
                                WR_CUSTOM_PATH_COLUMN & _
                                CStr(WR_CUSTOM_END_ROW))) Is Nothing Then
            Application.enableEvents = False
            duplicateModule = GetDuplicateWeeklyReportCustomModule(Sh)
            If Len(duplicateModule) > 0 Then
                Application.Undo
                MsgBox "이미 선택한 분류 항목은 다시 선택할 수 없습니다: " & duplicateModule, _
                       vbExclamation, _
                       "커스텀 분류 항목 중복"
            End If
            RefreshWeeklyReportModuleDropdown
        End If
        GoTo SafeExit
    End If

    Application.enableEvents = False
    HandleWebTaskChange Sh, Target
    Set textChanges = Intersect(Target, Sh.Range(COL_REQUEST & DATA_START_ROW & ":" & COL_MODIFICATION & Sh.Rows.Count))
    If Not textChanges Is Nothing Then
        For Each textArea In textChanges.Areas
            ApplyTaskTextDisplay Sh, Application.Min(GetLastDataRow(Sh), textArea.Row + textArea.Rows.Count - 1), textArea.Row
        Next textArea
    End If

SafeExit:
    Application.enableEvents = True
    Exit Sub

EH:
    If Not Sh Is Nothing And Sh.Name = CONFIG_SHEET_NAME Then
        MsgBox "config 설정을 처리하는 중 오류가 발생했습니다." & vbCrLf & _
               "원인: " & Err.Description, _
               vbExclamation, _
               "config 설정 오류"
    Else
        MsgBox "분류 계층을 자동 처리하는 중 오류가 발생했습니다." & vbCrLf & _
               "원인: " & Err.Description, _
               vbExclamation, _
               "분류 자동 상속 오류"
    End If
    Resume SafeExit
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo SafeExit

    If Not TypeOf Sh Is Worksheet Then Exit Sub
    RegisterGanttShortcuts
    StartTaskOutlineViewMonitor
    If Sh.Name <> WEEKLY_REPORT_CONFIG_SHEET_NAME Then Exit Sub

    Application.enableEvents = False
    RefreshWeeklyReportModuleDropdown

SafeExit:
    Application.enableEvents = True
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Not IsGanttPptDeleteCell(Sh, Target) Then Exit Sub
    Cancel = True
    ShowGanttPptDeleteMenu Sh, Target
End Sub

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim anchorCell As Range

    On Error GoTo SafeExit

    If Sh.Name = CONFIG_SHEET_NAME Or _
       Sh.Name = WEEKLY_REPORT_CONFIG_SHEET_NAME Or IsWebUtilitySheet(Sh) Then Exit Sub

    Set anchorCell = Target.Range
    If anchorCell Is Nothing Then Exit Sub
    If anchorCell.Row < DATA_START_ROW Then Exit Sub
    If anchorCell.Column <> Sh.Range(COL_NOTE & "1").Column Then Exit Sub
    If CStr(anchorCell.Value2) <> ChrW(&H25A0) Then Exit Sub

    칸트차트_PPT셀아이콘열기 Sh, anchorCell.Row

SafeExit:
End Sub

Private Sub Workbook_Activate()
    RegisterGanttShortcuts
    StartTaskOutlineViewMonitor
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    RegisterGanttShortcuts
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    UnregisterGanttShortcuts
End Sub

Private Sub Workbook_Deactivate()
    StopTaskOutlineViewMonitor
    UnregisterGanttShortcuts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    StopTaskOutlineViewMonitor
    ' The save prompt may still cancel closing. Release keys on actual deactivation.
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    RegisterGanttShortcuts
    StartTaskOutlineViewMonitor
End Sub

' 단축키는 Excel 전체에 적용되므로 이 통합문서가 활성화된 동안만 연결한다.
Public Sub RestoreGanttShortcutBindings()
    RegisterGanttShortcuts
End Sub

Private Sub RegisterGanttShortcuts()
    Dim workbookPrefix As String
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    workbookPrefix = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    Application.OnKey "^+r", workbookPrefix & "칸트차트_새로고침"
    Application.OnKey "^+h", workbookPrefix & "칸트차트_완료숨김"
    mGanttShortcutsRegistered = True
End Sub

Private Sub UnregisterGanttShortcuts()
    ' A late deactivate event from an inactive copy must not clear the current owner's keys.
    If Not mGanttShortcutsRegistered Then
        If ActiveWorkbook Is Nothing Then Exit Sub
        If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    End If
    Application.OnKey "^+r"
    Application.OnKey "^+h"
    mGanttShortcutsRegistered = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If TypeOf Sh Is Worksheet Then
        RegisterGanttShortcuts
        MaintainTaskOutlineOnSelection Sh
    End If
End Sub
