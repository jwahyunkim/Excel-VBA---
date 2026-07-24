VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
