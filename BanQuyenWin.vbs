Option Explicit

'===========================================================
' BAN QUYEN WINDOWS V3
' PHAN 1 - KHOI TAO, ADMIN, WMI, KIEM TRA HE THONG
'===========================================================

Dim objShell
Dim objWshShell
Dim objFSO
Dim objWMIService

Dim colOS
Dim objOS

Dim colProducts
Dim objProduct

Dim colItems
Dim objItem

Dim windir
Dim sysFolder

Dim osName
Dim edition

Dim biosKey
Dim biosKeyPartial

Dim currentPartialKey
Dim licenseDesc
Dim licenseType

Dim isActivated
Dim isKMS
Dim isMAK
Dim isRetail
Dim isDigital

Dim hasInternet
Dim objPing
Dim objStatus

Dim promptMsg
Dim reportMsg
Dim info
Dim userChoice

'===========================================================
' KHOI TAO DOI TUONG
'===========================================================

Set objShell = CreateObject("Shell.Application")
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'===========================================================
' YEU CAU CHAY QUYEN ADMINISTRATOR
'===========================================================

If Not WScript.Arguments.Named.Exists("elevated") Then

    On Error Resume Next

    objShell.ShellExecute _
        "wscript.exe", _
        """" & WScript.ScriptFullName & """ /elevated", _
        "", _
        "runas", _
        1

    If Err.Number <> 0 Then

        MsgBox _
        "Ban can chay bang quyen Administrator.", _
        vbCritical, _
        "Quyen Administrator"

    End If

    On Error GoTo 0

    WScript.Quit

End If

'===========================================================
' KET NOI WMI
'===========================================================

On Error Resume Next

Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

If Err.Number <> 0 Then

    MsgBox _
    "Khong the ket noi dich vu WMI." & vbCrLf & _
    "Windows co the da bi loi hoac bi Ghost/Lite.", _
    vbCritical, _
    "Loi WMI"

    WScript.Quit

End If

On Error GoTo 0

'===========================================================
' XAC DINH THU MUC SYSTEM
'===========================================================

windir = objWshShell.ExpandEnvironmentStrings("%SystemRoot%")

sysFolder = windir & "\System32"

If objFSO.FolderExists(windir & "\Sysnative") Then
    sysFolder = windir & "\Sysnative"
End If

If Not objFSO.FileExists(sysFolder & "\slmgr.vbs") Then

    MsgBox _
    "Khong tim thay slmgr.vbs." & vbCrLf & _
    "Windows co the da bi rut gon.", _
    vbCritical, _
    "Loi He Thong"

    WScript.Quit

End If

'===========================================================
' DOC TEN WINDOWS
'===========================================================

osName = "Khong xac dinh"

Set colOS = objWMIService.ExecQuery( _
"SELECT Caption FROM Win32_OperatingSystem")

For Each objOS In colOS

    osName = Trim(objOS.Caption)

Next

edition = ""

If InStr(UCase(osName),"HOME") > 0 Then

    edition = "Home"

ElseIf InStr(UCase(osName),"PRO") > 0 Then

    edition = "Professional"

ElseIf InStr(UCase(osName),"ENTERPRISE") > 0 Then

    edition = "Enterprise"

ElseIf InStr(UCase(osName),"EDUCATION") > 0 Then

    edition = "Education"

Else

    edition = "Khong xac dinh"

End If

'===========================================================
' KHOI TAO CAC BIEN
'===========================================================

isActivated = False

isKMS = False
isMAK = False
isRetail = False
isDigital = False

licenseDesc = ""
licenseType = ""

currentPartialKey = "-"

biosKey = ""
biosKeyPartial = ""

hasInternet = False

promptMsg = ""
reportMsg = ""
info = ""

'===========================================================
' PHAN 2
' DOC LICENSE + BIOS KEY + KIEM TRA INTERNET
'===========================================================

Set colProducts = objWMIService.ExecQuery( _
"SELECT LicenseStatus,Description,PartialProductKey,LicenseFamily " & _
"FROM SoftwareLicensingProduct " & _
"WHERE PartialProductKey IS NOT NULL " & _
"AND ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'")

For Each objProduct In colProducts

    If InStr(1, objProduct.Description, "Windows", vbTextCompare) > 0 Then

        licenseDesc = objProduct.Description

        If Not IsNull(objProduct.PartialProductKey) Then
            currentPartialKey = Trim(objProduct.PartialProductKey)
        End If

        If objProduct.LicenseStatus = 1 Then
            isActivated = True
        End If

        Exit For

    End If

Next


'===========================================================
' PHAN LOAI LICENSE
'===========================================================

licenseType = "Khong xac dinh"

If InStr(UCase(licenseDesc),"KMS") > 0 Then

    licenseType = "Volume KMS"
    isKMS = True

ElseIf InStr(UCase(licenseDesc),"MAK") > 0 Then

    licenseType = "Volume MAK"
    isMAK = True

ElseIf InStr(UCase(licenseDesc),"OEM") > 0 Then

    licenseType = "OEM"

ElseIf InStr(UCase(licenseDesc),"RETAIL") > 0 Then

    isRetail = True

    '=======================================================
    ' RETAIL CO THE LA DIGITAL LICENSE
    '=======================================================

    licenseType = "Retail"

Else

    licenseType = licenseDesc

End If


'===========================================================
' DOC BIOS KEY
'===========================================================

Set colItems = objWMIService.ExecQuery( _
"SELECT OA3xOriginalProductKey FROM SoftwareLicensingService")

For Each objItem In colItems

    If Not IsNull(objItem.OA3xOriginalProductKey) Then

        biosKey = Trim(objItem.OA3xOriginalProductKey)

        Exit For

    End If

Next


If biosKey <> "" Then

    biosKeyPartial = Right(biosKey,5)

    On Error Resume Next

    objWshShell.Run _
    "cmd.exe /c <nul set /p=""" & biosKey & """ | clip", _
    0, _
    True

    On Error GoTo 0

End If


'===========================================================
' NHAN DIEN DIGITAL LICENSE
'===========================================================

If isRetail Then

    If isActivated Then

        If biosKey = "" Then

            licenseType = "Retail / Digital License"

            isDigital = True

        Else

            If UCase(currentPartialKey) <> UCase(biosKeyPartial) Then

                licenseType = "Retail / Digital License"

                isDigital = True

            End If

        End If

    End If

End If


'===========================================================
' KIEM TRA INTERNET
'===========================================================

Set objPing = GetObject( _
"winmgmts:{impersonationLevel=impersonate}") _
.ExecQuery( _
"SELECT StatusCode FROM Win32_PingStatus WHERE Address='8.8.8.8'")

For Each objStatus In objPing

    If Not IsNull(objStatus.StatusCode) Then

        If objStatus.StatusCode = 0 Then

            hasInternet = True

            Exit For

        End If

    End If

Next

'===========================================================
' PHAN 3
' HIEN THI THONG TIN VA HOI RESTORE
'===========================================================

info = ""
info = info & "=======================================" & vbCrLf
info = info & "KIEM TRA BAN QUYEN WINDOWS" & vbCrLf
info = info & "=======================================" & vbCrLf & vbCrLf

info = info & "He dieu hanh: " & osName & vbCrLf
info = info & "Edition: " & edition & vbCrLf

If isActivated Then
    info = info & "Trang thai: DA KICH HOAT" & vbCrLf
Else
    info = info & "Trang thai: CHUA KICH HOAT" & vbCrLf
End If

info = info & "Loai key: " & licenseType & vbCrLf
info = info & "Product Key: " & currentPartialKey & vbCrLf

If biosKey = "" Then
    info = info & "BIOS Key: KHONG CO" & vbCrLf
Else
    info = info & "BIOS Key: " & biosKeyPartial & vbCrLf
End If

info = info & vbCrLf
info = info & "Danh gia:" & vbCrLf

'===========================================================
' KHONG CO BIOS
'===========================================================

If biosKey = "" Then

    info = info & _
    "- May khong co BIOS Key." & vbCrLf & _
    "- Khong the khoi phuc OEM."

    MsgBox info, vbInformation, "Ban quyen Windows"

    WScript.Quit

End If

'===========================================================
' DANG DUNG BIOS KEY
'===========================================================

If UCase(currentPartialKey)=UCase(biosKeyPartial) Then

    info = info & _
    "- Windows dang kich hoat bang BIOS Key." & vbCrLf & _
    "- Khong can khoi phuc."

    MsgBox info, vbInformation, "Ban quyen Windows"

    WScript.Quit

End If

'===========================================================
' DANH GIA LICENSE
'===========================================================

If isKMS Then

    info = info & _
    "- Dang su dung Volume KMS." & vbCrLf & _
    "- Nen khoi phuc BIOS Key." & vbCrLf

ElseIf isMAK Then

    info = info & _
    "- Dang su dung Volume MAK." & vbCrLf & _
    "- Can can nhac truoc khi khoi phuc." & vbCrLf

ElseIf isDigital Then

    info = info & _
    "- Dang su dung Retail / Digital License." & vbCrLf & _
    "- Microsoft co the giu lai ban quyen hien tai." & vbCrLf

ElseIf isRetail Then

    info = info & _
    "- Dang su dung Retail." & vbCrLf & _
    "- Co the mat ban quyen neu khoi phuc." & vbCrLf

Else

    info = info & _
    "- Dang su dung OEM." & vbCrLf

End If

'===========================================================
' CANH BAO EDITION
'===========================================================

If edition = "Professional" Then

    info = info & vbCrLf
    info = info & "CANH BAO:" & vbCrLf
    info = info & _
    "- Windows hien tai la Professional." & vbCrLf
    info = info & _
    "- BIOS Key cua nha san xuat thuong la Home." & vbCrLf
    info = info & _
    "- Neu tiep tuc co the Microsoft se tu choi kich hoat." & vbCrLf

End If

If edition = "Enterprise" Then

    info = info & vbCrLf
    info = info & "CANH BAO:" & vbCrLf
    info = info & _
    "- Windows Enterprise khong kich hoat bang OEM Key." & vbCrLf
    info = info & _
    "- Nen cai lai Windows Professional/Home truoc." & vbCrLf

End If

info = info & vbCrLf
info = info & "Ban co muon khoi phuc BIOS Key khong?"

userChoice = MsgBox( _
info, _
vbYesNo + vbQuestion, _
"Khoi phuc BIOS Key")

If userChoice = vbNo Then

    WScript.Quit

End If

'===========================================================
' KIEM TRA INTERNET
'===========================================================

If Not hasInternet Then

    MsgBox _
    "May tinh khong co Internet." & vbCrLf & _
    "Khong the kich hoat voi Microsoft.", _
    vbExclamation, _
    "Khong co Internet"

    WScript.Quit

End If

'==========================
' HET PHAN 3
'==========================
'===========================================================
' PHAN 4
' KHOI PHUC BIOS KEY
'===========================================================

'-----------------------------------------------------------
' CANH BAO VOI ENTERPRISE / EDUCATION
'-----------------------------------------------------------

If edition = "Enterprise" Or edition = "Education" Then

    If MsgBox( _
        "Windows hien tai la " & edition & "." & vbCrLf & vbCrLf & _
        "OEM BIOS Key thuong khong kich hoat duoc phien ban nay." & vbCrLf & _
        "Ban van muon tiep tuc?", _
        vbYesNo + vbExclamation, _
        "Canh bao") = vbNo Then

        WScript.Quit

    End If

End If

'-----------------------------------------------------------
' DON DEP THONG TIN KMS
'-----------------------------------------------------------

On Error Resume Next

objWshShell.Run _
"cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /cpky", _
0, _
True

objWshShell.Run _
"cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /ckms", _
0, _
True

On Error GoTo 0

'-----------------------------------------------------------
' NAP BIOS KEY
'-----------------------------------------------------------

objWshShell.Run _
"cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /ipk " & biosKey, _
0, _
True

'-----------------------------------------------------------
' DOI PRODUCT KEY (NEU CO)
'-----------------------------------------------------------

If objFSO.FileExists(sysFolder & "\changepk.exe") Then

    objWshShell.Run _
    """" & sysFolder & "\changepk.exe"" /ProductKey " & biosKey, _
    0, _
    True

End If

WScript.Sleep 3000

'-----------------------------------------------------------
' KICH HOAT
'-----------------------------------------------------------

objWshShell.Run _
"cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /ato", _
0, _
True

WScript.Sleep 5000

'==========================
' HET PHAN 4
'==========================
'===========================================================
' PHAN 5
' KIEM TRA KET QUA
'===========================================================

Dim newActivated
Dim newPartialKey
Dim newLicenseDesc
Dim resultType

newActivated = False
newPartialKey = "-"
newLicenseDesc = ""
resultType = ""

Set colProducts = objWMIService.ExecQuery( _
"SELECT LicenseStatus,Description,PartialProductKey " & _
"FROM SoftwareLicensingProduct " & _
"WHERE PartialProductKey IS NOT NULL " & _
"AND ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'")

For Each objProduct In colProducts

    If InStr(1,objProduct.Description,"Windows",vbTextCompare)>0 Then

        newLicenseDesc = objProduct.Description

        If Not IsNull(objProduct.PartialProductKey) Then
            newPartialKey = Trim(objProduct.PartialProductKey)
        End If

        If objProduct.LicenseStatus = 1 Then
            newActivated = True
        End If

        Exit For

    End If

Next


'===========================================================
' PHAN LOAI LICENSE MOI
'===========================================================

If InStr(UCase(newLicenseDesc),"KMS")>0 Then

    resultType = "Volume KMS"

ElseIf InStr(UCase(newLicenseDesc),"MAK")>0 Then

    resultType = "Volume MAK"

ElseIf InStr(UCase(newLicenseDesc),"OEM")>0 Then

    resultType = "OEM"

ElseIf InStr(UCase(newLicenseDesc),"RETAIL")>0 Then

    resultType = "Retail / Digital License"

Else

    resultType = newLicenseDesc

End If


'===========================================================
' BAO CAO
'===========================================================

reportMsg = ""
reportMsg = reportMsg & "======================================" & vbCrLf
reportMsg = reportMsg & "KET QUA KHOI PHUC BAN QUYEN" & vbCrLf
reportMsg = reportMsg & "======================================" & vbCrLf & vbCrLf

reportMsg = reportMsg & "He dieu hanh : " & osName & vbCrLf
reportMsg = reportMsg & "Edition      : " & edition & vbCrLf

If newActivated Then
    reportMsg = reportMsg & "Trang thai   : DA KICH HOAT" & vbCrLf
Else
    reportMsg = reportMsg & "Trang thai   : CHUA KICH HOAT" & vbCrLf
End If

reportMsg = reportMsg & "Loai key     : " & resultType & vbCrLf
reportMsg = reportMsg & "Product Key  : " & newPartialKey & vbCrLf

If biosKey <> "" Then
    reportMsg = reportMsg & "BIOS Key     : " & biosKeyPartial & vbCrLf
End If

reportMsg = reportMsg & vbCrLf

'===========================================================
' KET LUAN
'===========================================================

If newActivated Then

    If UCase(newPartialKey)=UCase(biosKeyPartial) Then

        reportMsg = reportMsg & _
        "KET LUAN:" & vbCrLf & _
        "- Khoi phuc BIOS Key thanh cong." & vbCrLf & _
        "- Windows dang kich hoat bang BIOS Key."

        MsgBox reportMsg, vbInformation, "Thanh cong"

    Else

        reportMsg = reportMsg & _
        "KET LUAN:" & vbCrLf & _
        "- Windows da kich hoat." & vbCrLf & _
        "- Microsoft giu lai ban quyen hien tai." & vbCrLf & _
        "- Khong can tiep tuc khoi phuc."

        MsgBox reportMsg, vbInformation, "Hoan thanh"

    End If

Else

    reportMsg = reportMsg & _
    "KET LUAN:" & vbCrLf & _
    "- Khong kich hoat duoc BIOS Key." & vbCrLf & vbCrLf & _
    "Nguyen nhan co the:" & vbCrLf & _
    "- BIOS Key khong phu hop voi Edition Windows." & vbCrLf & _
    "- Windows bi loi sau khi crack." & vbCrLf & _
    "- BIOS Key khong hop le." & vbCrLf & _
    "- Can cai lai Windows dung phien ban OEM."

    MsgBox reportMsg, vbCritical, "That bai"

End If


'===========================================================
' GIAI PHONG BO NHO
'===========================================================

Set colOS = Nothing
Set colProducts = Nothing
Set colItems = Nothing

Set objOS = Nothing
Set objProduct = Nothing
Set objItem = Nothing

Set objPing = Nothing
Set objStatus = Nothing

Set objShell = Nothing
Set objWshShell = Nothing
Set objFSO = Nothing
Set objWMIService = Nothing

WScript.Quit