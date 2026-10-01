Attribute VB_Name = "modTaskType"
Option Explicit

Public Const TASK_TYPE_TITLE_CELL As String = "L9"
Public Const TASK_TYPE_CONFIG_RANGE As String = "L10:L29"
Public Const TASK_TYPE_LIST_NAME As String = "GanttTaskTypes"
Private Const TASK_TYPE_HELPER_RANGE As String = "AB1:AB20"

Public Sub EnsureTaskTypeConfig()
    Dim ws As Worksheet, item As Range, defaults As Variant
    Dim initialized As Boolean, count As Long, text As String
    Dim seen As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        EnsureConfigSheet
        Exit Sub
    End If

    initialized = (CStr(ws.Range(TASK_TYPE_TITLE_CELL).Value2) = "허용 타입 목록")
    If Not initialized And Application.CountA(ws.Range(TASK_TYPE_CONFIG_RANGE)) = 0 Then
        defaults = Array("개발", "ETC", "유지보수", "프로젝트")
        For count = 0 To UBound(defaults)
            ws.Range(TASK_TYPE_CONFIG_RANGE).Cells(count + 1, 1).Value2 = defaults(count)
        Next count
    End If
    ws.Range(TASK_TYPE_TITLE_CELL).Value2 = "허용 타입 목록"
    ws.Columns("L").ColumnWidth = 22
    With ws.Range(TASK_TYPE_TITLE_CELL)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range(TASK_TYPE_CONFIG_RANGE)
        .NumberFormat = "@"
        .Borders.LineStyle = xlContinuous
        .Locked = False
    End With
    ws.Range("M10").Value2 = "L10:L29에 타입 추가 (최대 20개)"
    ws.Columns("M").ColumnWidth = 38

    ' Keep empty reserve cells and duplicates out of the dropdown.
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    ws.Range(TASK_TYPE_HELPER_RANGE).ClearContents
    count = 0
    For Each item In ws.Range(TASK_TYPE_CONFIG_RANGE)
        If Not IsError(item.Value2) Then
            text = CStr(item.Value2)
            If Len(text) > 0 And Not seen.Exists(text) Then
                seen.Add text, True
                count = count + 1
                ws.Cells(count, "AB").Value2 = text
            End If
        End If
    Next item
    If count = 0 Then count = 1
    ThisWorkbook.Names.Add Name:=TASK_TYPE_LIST_NAME, _
        RefersTo:="='" & CONFIG_SHEET_NAME & "'!$AB$1:$AB$" & CStr(count)
    ws.Columns("AB").Hidden = True
End Sub

Public Sub ApplyTaskTypeValidation(ByVal ws As Worksheet)
    With ws.Range(COL_TYPE & DATA_START_ROW & ":" & COL_TYPE & ws.Rows.Count).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & TASK_TYPE_LIST_NAME
        .IgnoreBlank = False
        .InCellDropdown = True
        .ShowError = True
        .InputTitle = "타입 선택"
        .InputMessage = "config 시트의 허용 타입 목록에서 선택하세요."
        .ErrorTitle = "타입 입력 오류"
        .ErrorMessage = "config 시트에 등록한 타입만 입력할 수 있습니다."
    End With
End Sub

Public Sub RefreshTaskTypeValidation()
    Dim ws As Worksheet, wasProtected As Boolean
    Dim errNumber As Long, errDescription As String
    On Error GoTo EH
    EnsureTaskTypeConfig
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Range(COL_TYPE & HEADER_ROW).Value2) = "타입" And _
           CStr(ws.Range(COL_NO & HEADER_ROW).Value2) = "No." Then
            wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
            If wasProtected Then UnprotectTaskSheet ws
            ApplyTaskTypeValidation ws
            If wasProtected Then ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
            wasProtected = False
        End If
    Next ws
    Exit Sub
EH:
    errNumber = Err.Number
    errDescription = Err.Description
    If wasProtected And Not ws Is Nothing Then
        On Error Resume Next
        ApplyCalculatedColumnsProtection ws, GetLastDataRow(ws)
        On Error GoTo 0
    End If
    Err.Raise errNumber, "RefreshTaskTypeValidation", errDescription
End Sub

Public Function GetInvalidTaskType(ByVal ws As Worksheet, ByVal Target As Range) As String
    Dim changed As Range, item As Range, configured As Range, allowed As Object
    If CStr(ws.Range(COL_TYPE & HEADER_ROW).Value2) <> "타입" Or _
       CStr(ws.Range(COL_NO & HEADER_ROW).Value2) <> "No." Then Exit Function
    Set changed = Intersect(Target, ws.Range(COL_TYPE & DATA_START_ROW & ":" & COL_TYPE & ws.Rows.Count))
    If changed Is Nothing Then Exit Function
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbBinaryCompare
    For Each configured In ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).Range(TASK_TYPE_CONFIG_RANGE)
        If Not IsError(configured.Value2) Then allowed(CStr(configured.Value2)) = True
    Next configured
    For Each item In changed.Cells
        If IsError(item.Value2) Then
            GetInvalidTaskType = item.Address(False, False)
            Exit Function
        End If
        If Len(CStr(item.Value2)) > 0 Then
            If Not allowed.Exists(CStr(item.Value2)) Then
                GetInvalidTaskType = item.Address(False, False)
                Exit Function
            End If
        End If
    Next item
End Function
