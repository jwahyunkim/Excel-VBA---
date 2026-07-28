VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo EH

    If Not TypeOf Sh Is Worksheet Then Exit Sub

    If Sh.Name = CONFIG_SHEET_NAME Then
        If Intersect(Target, Sh.Range(TASK_MAX_LENGTH_VALUE_RANGE)) Is Nothing Then Exit Sub
        Application.EnableEvents = False
        RefreshTaskTextLengthValidation
        GoTo SafeExit
    End If

    If Sh.Name = REPORT_HISTORY_SHEET_NAME Then Exit Sub

    Application.EnableEvents = False
    HandleTaskHierarchyChange Sh, Target

SafeExit:
    Application.EnableEvents = True
    Exit Sub

EH:
    MsgBox "모듈 계층을 자동 처리하는 중 오류가 발생했습니다." & vbCrLf & _
           "원인: " & Err.Description, _
           vbExclamation, _
           "모듈 자동 상속 오류"
    Resume SafeExit
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
