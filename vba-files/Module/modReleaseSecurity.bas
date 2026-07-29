Attribute VB_Name = "modReleaseSecurity"
Option Explicit

Private Const SECURITY_SHEET_NAME As String = "__RELEASE_SECURITY"
Private Const GUIDE_SHEET_NAME As String = "사용안내"
Private Const INFO_SHEET_NAME As String = "배포정보"
Private Const SECURITY_MARKER As String = "RELEASE_SECURITY_V1"
Private Const SECURITY_ENABLED As String = "Y"
Private Const SHEET_NAME_SEPARATOR As String = vbLf
Private Const CODE_MODULUS As Long = 1679616       ' 36 ^ 4

Private mReleaseAuthenticated As Boolean

' 배포 스크립트에서 새 복사본에만 호출합니다.
' 개발 원본에는 보안 시트가 없으므로 Workbook_Open에서 아무 작업도 하지 않습니다.
Public Function ConfigureReleaseSecurity(ByVal releaseVersion As String, _
                                         ByVal releaseDateValue As Double, _
                                         ByVal expiryDateValue As Double, _
                                         ByVal renewalDays As Long, _
                                         ByVal renewalSecret As String, _
                                         ByVal releaseUser As String, _
                                         ByVal adminPassword As String) As String
    Dim previousEnableEvents As Boolean
    Dim securitySheet As Worksheet
    Dim guideSheet As Worksheet
    Dim infoSheet As Worksheet
    Dim visibleBusinessSheets As String

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False

    If IsReleaseSecurityEnabled() Then _
        Err.Raise 5, , "이미 배포 보안이 적용된 통합문서입니다. 개발 원본으로 다시 빌드하세요."
    If Len(Trim$(releaseVersion)) = 0 Then Err.Raise 5, , "배포 버전이 비어 있습니다."
    If renewalDays < 1 Then Err.Raise 5, , "연장 일수는 1일 이상이어야 합니다."
    If Len(renewalSecret) = 0 Then Err.Raise 5, , "기간 연장 비밀키가 비어 있습니다."
    If Len(Trim$(releaseUser)) = 0 Then Err.Raise 5, , "배포 대상 사용자가 비어 있습니다."
    If Len(adminPassword) = 0 Then Err.Raise 5, , "배포정보 확인 암호가 비어 있습니다."

    visibleBusinessSheets = CollectVisibleBusinessSheetNames()
    If Len(visibleBusinessSheets) = 0 Then _
        Err.Raise 5, , "배포할 표시 시트가 하나도 없습니다."

    Set guideSheet = GetOrCreateWorksheet(GUIDE_SHEET_NAME)
    Set securitySheet = GetOrCreateWorksheet(SECURITY_SHEET_NAME)
    Set infoSheet = GetOrCreateWorksheet(INFO_SHEET_NAME)

    With securitySheet
        .Range("A1").Value2 = SECURITY_MARKER
        .Range("B1").Value2 = SECURITY_ENABLED
        .Range("A2").Value2 = "release_version"
        .Range("B2").Value2 = releaseVersion
        .Range("A3").Value2 = "release_date"
        .Range("B3").Value2 = releaseDateValue
        .Range("A4").Value2 = "expiry_date"
        .Range("B4").Value2 = expiryDateValue
        .Range("A5").Value2 = "renewal_days"
        .Range("B5").Value2 = renewalDays
        .Range("A6").Value2 = "renewal_secret"
        .Range("B6").Value2 = renewalSecret
        .Range("A7").Value2 = "visible_sheets"
        .Range("B7").Value2 = visibleBusinessSheets
        .Range("A8").Value2 = "release_user"
        .Range("B8").Value2 = releaseUser
        .Range("A9").Value2 = "admin_password"
        .Range("B9").Value2 = adminPassword
        .Range("B3:B4").NumberFormat = "yyyy-mm-dd"
    End With

    WriteGuideSheet guideSheet, securitySheet
    WriteReleaseInfoSheet infoSheet, securitySheet, adminPassword
    HideBusinessSheets

    mReleaseAuthenticated = False
    ConfigureReleaseSecurity = BuildReleaseRenewalCode(releaseDateValue, renewalSecret)

SafeExit:
    Application.EnableEvents = previousEnableEvents
    Exit Function

EH:
    Application.EnableEvents = previousEnableEvents
    Err.Raise Err.Number, "ConfigureReleaseSecurity", Err.Description
End Function

Public Sub InitializeReleaseSecurity()
    Dim previousEnableEvents As Boolean
    Dim securitySheet As Worksheet
    Dim expiryDate As Date
    Dim renewalDays As Long
    Dim renewalSecret As String
    Dim expectedCode As String
    Dim enteredCode As Variant
    Dim attempt As Long
    Dim errorDescription As String

    If Not IsReleaseSecurityEnabled() Then Exit Sub

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False

    Set securitySheet = ThisWorkbook.Worksheets(SECURITY_SHEET_NAME)
    expiryDate = CDate(securitySheet.Range("B4").Value2)

    If Date <= expiryDate Then
        mReleaseAuthenticated = True
        ShowBusinessSheets
        ThisWorkbook.Saved = True
        GoTo SafeExit
    End If

    HideBusinessSheets
    renewalDays = CLng(securitySheet.Range("B5").Value2)
    renewalSecret = CStr(securitySheet.Range("B6").Value2)
    expectedCode = BuildReleaseRenewalCode(CDbl(Date), renewalSecret)

    For attempt = 1 To 3
        enteredCode = Application.InputBox( _
            Prompt:="사용 기간이 " & Format$(expiryDate, "yyyy-mm-dd") & _
                    "에 끝났습니다." & vbCrLf & _
                    "오늘 날짜용 기간 연장 코드를 입력하세요.", _
            Title:="사용 기간 연장", Type:=2)

        If VarType(enteredCode) = vbBoolean Then
            If enteredCode = False Then GoTo SafeExit
        End If

        If StrComp(Trim$(CStr(enteredCode)), expectedCode, vbTextCompare) = 0 Then
            securitySheet.Range("B4").Value2 = CDbl(Date + renewalDays)
            securitySheet.Range("B4").NumberFormat = "yyyy-mm-dd"
            WriteGuideSheet ThisWorkbook.Worksheets(GUIDE_SHEET_NAME), securitySheet

            ' 디스크에는 안내 시트만 보이는 잠금 상태로 만료일을 한 번 저장합니다.
            ThisWorkbook.Save

            mReleaseAuthenticated = True
            ShowBusinessSheets
            ThisWorkbook.Saved = True
            MsgBox "사용 기간이 " & Format$(Date + renewalDays, "yyyy-mm-dd") & _
                   "까지 연장되었습니다.", vbInformation, "기간 연장 완료"
            GoTo SafeExit
        End If

        MsgBox "연장 코드가 맞지 않습니다. (" & attempt & "/3)", _
               vbExclamation, "기간 연장 실패"
    Next attempt

SafeExit:
    Application.EnableEvents = previousEnableEvents
    Exit Sub

EH:
    errorDescription = Err.Description
    On Error Resume Next
    HideBusinessSheets
    Application.EnableEvents = previousEnableEvents
    MsgBox "배포본 보안을 초기화하지 못했습니다." & vbCrLf & _
           "원인: " & errorDescription, vbCritical, "배포본 보안 오류"
End Sub

Public Sub PrepareReleaseForSave(ByVal saveAsUI As Boolean, ByRef cancelSave As Boolean)
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim activeBusinessSheetName As String
    Dim saveSucceeded As Boolean
    Dim errorDescription As String

    If Not IsReleaseSecurityEnabled() Then Exit Sub
    If Not mReleaseAuthenticated Then Exit Sub

    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    If ActiveWorkbook Is ThisWorkbook Then
        If TypeOf ActiveSheet Is Worksheet Then
            activeBusinessSheetName = ActiveSheet.Name
        End If
    End If
    On Error GoTo EH
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    StoreVisibleBusinessSheetNames
    HideBusinessSheets

    If saveAsUI Then
        ' 기본 저장을 취소하고 이벤트가 꺼진 저장 대화상자에서 잠금 상태로 저장합니다.
        ' 사용자가 대화상자를 취소해도 현재 화면은 즉시 복원됩니다.
        cancelSave = True
        saveSucceeded = Application.Dialogs(xlDialogSaveAs).Show
        ShowBusinessSheets activeBusinessSheetName
        If saveSucceeded Then ThisWorkbook.Saved = True
        GoTo SafeExit
    End If

    ' 바깥 저장을 취소한 뒤 이벤트가 꺼진 내부 저장을 완전히 끝냅니다.
    ' 따라서 아래 화면 복원은 디스크 파일에 포함되지 않습니다.
    cancelSave = True
    ThisWorkbook.Save
    ShowBusinessSheets activeBusinessSheetName
    ThisWorkbook.Saved = True

SafeExit:
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

EH:
    errorDescription = Err.Description
    cancelSave = True
    On Error Resume Next
    ShowBusinessSheets activeBusinessSheetName
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "보안 저장 상태를 만들지 못해 저장을 취소했습니다." & vbCrLf & _
           "원인: " & errorDescription, vbExclamation, "저장 취소"
End Sub

Public Sub 배포정보확인()
    Dim previousEnableEvents As Boolean
    Dim wasSaved As Boolean
    Dim infoSheet As Worksheet
    Dim errorDescription As String

    If Not IsReleaseSecurityEnabled() Then
        MsgBox "개발 원본에는 배포정보가 없습니다.", vbInformation, "배포정보 확인"
        Exit Sub
    End If

    previousEnableEvents = Application.EnableEvents
    wasSaved = ThisWorkbook.Saved
    On Error GoTo EH
    Application.EnableEvents = False

    Set infoSheet = ThisWorkbook.Worksheets(INFO_SHEET_NAME)
    On Error Resume Next
    infoSheet.Unprotect
    On Error GoTo EH
    If infoSheet.ProtectContents Then GoTo SafeExit

    infoSheet.Protect Password:=GetReleaseAdminPassword(), DrawingObjects:=True, _
                      Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    infoSheet.Visible = xlSheetVisible
    infoSheet.Activate
    If wasSaved Then ThisWorkbook.Saved = True

SafeExit:
    Application.EnableEvents = previousEnableEvents
    Exit Sub

EH:
    errorDescription = Err.Description
    On Error Resume Next
    If Not infoSheet Is Nothing Then
        If Not infoSheet.ProtectContents Then
            infoSheet.Protect Password:=GetReleaseAdminPassword(), _
                              DrawingObjects:=True, Contents:=True, _
                              Scenarios:=True, UserInterfaceOnly:=True
        End If
    End If
    If wasSaved Then ThisWorkbook.Saved = True
    Application.EnableEvents = previousEnableEvents
    MsgBox "배포정보를 표시하지 못했습니다." & vbCrLf & _
           "원인: " & errorDescription, vbExclamation, "배포정보 확인"
End Sub

Public Sub HideReleaseInfoSheet()
    Dim previousEnableEvents As Boolean
    Dim wasSaved As Boolean
    Dim infoSheet As Worksheet

    If Not IsReleaseSecurityEnabled() Then Exit Sub

    Set infoSheet = ThisWorkbook.Worksheets(INFO_SHEET_NAME)
    If infoSheet.Visible <> xlSheetVisible Then Exit Sub

    previousEnableEvents = Application.EnableEvents
    wasSaved = ThisWorkbook.Saved
    On Error GoTo SafeExit
    Application.EnableEvents = False

    infoSheet.Visible = xlSheetVeryHidden
    If wasSaved Then ThisWorkbook.Saved = True

SafeExit:
    Application.EnableEvents = previousEnableEvents
End Sub

Private Function BuildReleaseRenewalCode(ByVal targetDateValue As Double, _
                                         ByVal renewalSecret As String) As String
    Dim targetDate As Date
    Dim payload As String
    Dim accumulatorA As Long
    Dim accumulatorB As Long
    Dim characterCode As Long
    Dim index As Long

    targetDate = CDate(targetDateValue)
    payload = RotateSecret(renewalSecret) & "|" & Format$(targetDate, "ddyyMM")
    accumulatorA = 7919
    accumulatorB = 104729

    For index = 1 To Len(payload)
        characterCode = AscW(Mid$(payload, index, 1)) And &HFFFF&
        accumulatorA = ((accumulatorA * 33) + characterCode + (index * 17)) Mod CODE_MODULUS
        accumulatorB = ((accumulatorB * 37) + (characterCode * 7) + (index * 13)) Mod CODE_MODULUS
    Next index

    BuildReleaseRenewalCode = ToBase36(accumulatorA, 4) & "-" & _
                              ToBase36(accumulatorB, 4)
End Function

Public Function IsReleaseSecurityEnabled() As Boolean
    Dim securitySheet As Worksheet

    On Error Resume Next
    Set securitySheet = ThisWorkbook.Worksheets(SECURITY_SHEET_NAME)
    On Error GoTo 0
    If securitySheet Is Nothing Then Exit Function

    IsReleaseSecurityEnabled = _
        (CStr(securitySheet.Range("A1").Value2) = SECURITY_MARKER And _
         UCase$(Trim$(CStr(securitySheet.Range("B1").Value2))) = SECURITY_ENABLED)
End Function

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Function CollectVisibleBusinessSheetNames() As String
    Dim worksheetItem As Worksheet
    Dim result As String

    For Each worksheetItem In ThisWorkbook.Worksheets
        If worksheetItem.Name <> SECURITY_SHEET_NAME And _
           worksheetItem.Name <> GUIDE_SHEET_NAME And _
           worksheetItem.Name <> INFO_SHEET_NAME And _
           worksheetItem.Visible = xlSheetVisible Then
            If Len(result) > 0 Then result = result & SHEET_NAME_SEPARATOR
            result = result & worksheetItem.Name
        End If
    Next worksheetItem

    CollectVisibleBusinessSheetNames = result
End Function

Private Sub StoreVisibleBusinessSheetNames()
    Dim securitySheet As Worksheet
    Dim visibleBusinessSheets As String

    Set securitySheet = ThisWorkbook.Worksheets(SECURITY_SHEET_NAME)
    visibleBusinessSheets = CollectVisibleBusinessSheetNames()
    If Len(visibleBusinessSheets) > 0 Then
        If CStr(securitySheet.Range("B7").Value2) <> visibleBusinessSheets Then
            securitySheet.Range("B7").Value2 = visibleBusinessSheets
        End If
    End If
End Sub

Private Sub HideBusinessSheets()
    Dim worksheetItem As Worksheet
    Dim guideSheet As Worksheet
    Dim securitySheet As Worksheet

    Set guideSheet = ThisWorkbook.Worksheets(GUIDE_SHEET_NAME)
    Set securitySheet = ThisWorkbook.Worksheets(SECURITY_SHEET_NAME)

    guideSheet.Visible = xlSheetVisible
    guideSheet.Activate

    For Each worksheetItem In ThisWorkbook.Worksheets
        If worksheetItem.Name <> GUIDE_SHEET_NAME Then
            worksheetItem.Visible = xlSheetVeryHidden
        End If
    Next worksheetItem

    securitySheet.Visible = xlSheetVeryHidden
End Sub

Private Sub ShowBusinessSheets(Optional ByVal preferredSheetName As String = "")
    Dim securitySheet As Worksheet
    Dim guideSheet As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim worksheetItem As Worksheet
    Dim visibleCount As Long
    Dim firstVisibleSheetName As String

    Set securitySheet = ThisWorkbook.Worksheets(SECURITY_SHEET_NAME)
    Set guideSheet = ThisWorkbook.Worksheets(GUIDE_SHEET_NAME)
    guideSheet.Visible = xlSheetVisible

    sheetNames = Split(CStr(securitySheet.Range("B7").Value2), SHEET_NAME_SEPARATOR)
    For Each sheetName In sheetNames
        If Len(CStr(sheetName)) > 0 Then
            Set worksheetItem = Nothing
            On Error Resume Next
            Set worksheetItem = ThisWorkbook.Worksheets(CStr(sheetName))
            On Error GoTo 0
            If Not worksheetItem Is Nothing Then
                worksheetItem.Visible = xlSheetVisible
                visibleCount = visibleCount + 1
                If Len(firstVisibleSheetName) = 0 Then firstVisibleSheetName = worksheetItem.Name
            End If
        End If
    Next sheetName

    securitySheet.Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(INFO_SHEET_NAME).Visible = xlSheetVeryHidden
    If visibleCount > 0 Then
        Set worksheetItem = Nothing
        If Len(preferredSheetName) > 0 Then
            On Error Resume Next
            Set worksheetItem = ThisWorkbook.Worksheets(preferredSheetName)
            On Error GoTo 0
        End If
        If worksheetItem Is Nothing Then
            Set worksheetItem = ThisWorkbook.Worksheets(firstVisibleSheetName)
        End If
        worksheetItem.Activate
        guideSheet.Visible = xlSheetVeryHidden
    End If
End Sub

Private Sub WriteGuideSheet(ByVal guideSheet As Worksheet, _
                            ByVal securitySheet As Worksheet)
    With guideSheet
        .Range("A1").Value2 = "업무 간트 배포본"
        .Range("A3").Value2 = "배포 버전"
        .Range("B3").Value2 = CStr(securitySheet.Range("B2").Value2)
        .Range("A4").Value2 = "배포일"
        .Range("B4").Value2 = securitySheet.Range("B3").Value2
        .Range("A5").Value2 = "사용 만료일"
        .Range("B5").Value2 = securitySheet.Range("B4").Value2
        .Range("A6:B6").ClearContents
        .Range("B4:B5").NumberFormat = "yyyy-mm-dd"
        .Range("A7").Value2 = "사용 기간이 끝나면 관리자에게 오늘 날짜용 연장 코드를 요청하세요."
        .Range("A8").Value2 = "이 시트만 보이는 경우 파일을 닫지 말고 연장 코드를 입력하세요."
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 18
        .Range("A3:A6").Font.Bold = True
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 22
        .Range("A7:A8").WrapText = True
    End With
End Sub

Private Sub WriteReleaseInfoSheet(ByVal infoSheet As Worksheet, _
                                  ByVal securitySheet As Worksheet, _
                                  ByVal adminPassword As String)
    With infoSheet
        .Cells.Clear
        .Range("A1").Value2 = "배포정보"
        .Range("A3").Value2 = "배포 대상 사용자"
        .Range("B3").Value2 = CStr(securitySheet.Range("B8").Value2)
        .Range("A4").Value2 = "배포 버전"
        .Range("B4").Value2 = CStr(securitySheet.Range("B2").Value2)
        .Range("A5").Value2 = "배포일"
        .Range("B5").Value2 = securitySheet.Range("B3").Value2
        .Range("A6").Value2 = "사용 만료일"
        .Range("B6").Value2 = securitySheet.Range("B4").Value2
        .Range("B5:B6").NumberFormat = "yyyy-mm-dd"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 18
        .Range("A3:A6").Font.Bold = True
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 22
        .Protect Password:=adminPassword, DrawingObjects:=True, _
                 Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        .Visible = xlSheetVeryHidden
    End With
End Sub

Private Function GetReleaseAdminPassword() As String
    GetReleaseAdminPassword = CStr( _
        ThisWorkbook.Worksheets(SECURITY_SHEET_NAME).Range("B9").Value2)
End Function

Private Function RotateSecret(ByVal value As String) As String
    If Len(value) <= 1 Then
        RotateSecret = value
    Else
        RotateSecret = Mid$(value, 2) & Left$(value, 1)
    End If
End Function

Private Function ToBase36(ByVal value As Long, ByVal width As Long) As String
    Const DIGITS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim result As String
    Dim digitValue As Long

    Do
        digitValue = value Mod 36
        result = Mid$(DIGITS, digitValue + 1, 1) & result
        value = value \ 36
    Loop While value > 0

    If Len(result) < width Then result = String$(width - Len(result), "0") & result
    ToBase36 = Right$(result, width)
End Function
