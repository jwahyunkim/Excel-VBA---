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

Private Sub Workbook_Open()
    InitializeReleaseSecurity
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    PrepareReleaseForSave SaveAsUI, Cancel
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next
    If TypeOf Sh Is Worksheet Then
        If Sh.Name = "배포정보" Then HideReleaseInfoSheet
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim duplicateModule As String

    On Error GoTo EH

    If Not TypeOf Sh Is Worksheet Then Exit Sub

    If Sh.Name = CONFIG_SHEET_NAME Then
        If Not Intersect(Target, Sh.Range(TASK_MAX_LENGTH_VALUE_RANGE)) Is Nothing Then
            Application.enableEvents = False
            RefreshTaskTextLengthValidation
        ElseIf Not Intersect( _
                       Target, _
                       Sh.Range(WEEKLY_REPORT_CUSTOM_MODULE_COLUMN & _
                                CStr(WEEKLY_REPORT_CUSTOM_START_ROW) & ":" & _
                                WEEKLY_REPORT_CUSTOM_MODULE_COLUMN & _
                                CStr(WEEKLY_REPORT_CUSTOM_END_ROW))) Is Nothing Then
            Application.enableEvents = False
            duplicateModule = GetDuplicateWeeklyReportCustomModule(Sh)
            If Len(duplicateModule) > 0 Then
                Application.Undo
                MsgBox "이미 선택한 모듈은 다시 선택할 수 없습니다: " & duplicateModule, _
                       vbExclamation, _
                       "커스텀 모듈 중복"
            End If
            RefreshWeeklyReportModuleDropdown
        End If
        GoTo SafeExit
    End If

    If Sh.Name = REPORT_HISTORY_SHEET_NAME Then Exit Sub

    Application.enableEvents = False
    HandleTaskHierarchyChange Sh, Target

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
        MsgBox "모듈 계층을 자동 처리하는 중 오류가 발생했습니다." & vbCrLf & _
               "원인: " & Err.Description, _
               vbExclamation, _
               "모듈 자동 상속 오류"
    End If
    Resume SafeExit
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo SafeExit

    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Sh.Name <> CONFIG_SHEET_NAME Then Exit Sub

    Application.enableEvents = False
    RefreshWeeklyReportModuleDropdown

SafeExit:
    Application.enableEvents = True
End Sub

Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim anchorCell As Range

    On Error GoTo SafeExit

    If Sh.Name = CONFIG_SHEET_NAME Then Exit Sub

    Set anchorCell = Target.Range
    If anchorCell Is Nothing Then Exit Sub
    If anchorCell.Row < DATA_START_ROW Then Exit Sub
    If anchorCell.Column <> Sh.Range(COL_NOTE & "1").Column Then Exit Sub
    If CStr(anchorCell.Value2) <> ChrW(&H25A0) Then Exit Sub

    칸트차트_PPT셀아이콘열기 Sh, anchorCell.Row

SafeExit:
End Sub
