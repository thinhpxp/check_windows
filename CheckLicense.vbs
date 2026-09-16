'===========================================================
' CheckLicense.vbs (Bao cao ban quyen Windows & MS Office)
' Chay hoan toan silent (khong MsgBox / Khong cua so)
' Thu thap thong tin:
'   1. Ban quyen Windows: Installed Key (Full 25 ky tu), BIOS Key, Status, Type
'   2. Ban quyen MS Office: Phien ban, Build/Arch, Status, License Type, Partial Key
'   3. Thong tin may tinh: Ten may, Dong may, OS, Edition, User, IP
'   4. Thong tin chi tiet tu slmgr (/dli, /dlv) va ospp.vbs (/dstatus)
' Luu ket qua vao: \\10.43.4.103\doc\33_THINH\ketquakiemtra\<COMPUTERNAME>.txt
'===========================================================

Option Explicit

'-----------------------------------------------------------
' CAU HINH - CHINH SUA THEO MOI TRUONG CUA BAN
'-----------------------------------------------------------
Const RESULT_FOLDER = "\\10.43.4.103\doc\33_THINH\ketquakiemtra"

'-----------------------------------------------------------
' KHOI TAO DOI TUONG
'-----------------------------------------------------------
Dim objWMIService, objFSO, objWshShell, objNetwork
Dim colOS, colProducts
Dim objOS, objProduct

Set objFSO        = CreateObject("Scripting.FileSystemObject")
Set objWshShell   = CreateObject("WScript.Shell")
Set objNetwork    = CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'-----------------------------------------------------------
' LAY THONG TIN CO BAN MAY TINH
'-----------------------------------------------------------
Dim computerName, userName, osName, edition, productId, computerModel
Dim ipAddress, timestamp

computerName  = objNetwork.ComputerName
userName      = objWshShell.ExpandEnvironmentStrings("%USERNAME%")
timestamp     = Now()

' Lay thong tin Hang & Dong may tinh
computerModel = GetComputerModel()

' Lay IP Address chinh
ipAddress = GetIPAddress()

' Lay thong tin OS & Product ID qua WMI
Set colOS = objWMIService.ExecQuery("SELECT Caption, Version, BuildNumber, SerialNumber FROM Win32_OperatingSystem")
osName    = ""
edition   = ""
productId = "N/A"
For Each objOS In colOS
    osName    = Trim(objOS.Caption)
    If Not IsNull(objOS.SerialNumber) Then
        productId = Trim(objOS.SerialNumber)
    End If
    Exit For
Next

' Trich xuat Edition tu ten OS
edition = ExtractEdition(osName)

'-----------------------------------------------------------
' LAY THONG TIN BAN QUYEN WINDOWS (WMI SoftwareLicensing)
'-----------------------------------------------------------
Dim winLicenseStatus, winLicenseDesc, winPartialKey, winLicenseType, winProductChannel
Dim winIsActivated

winLicenseStatus  = -1
winLicenseDesc    = ""
winPartialKey     = "N/A"
winLicenseType    = "Unknown"
winProductChannel = ""
winIsActivated    = False

Set colProducts = objWMIService.ExecQuery( _
    "SELECT LicenseStatus, Description, PartialProductKey, ProductKeyChannel " & _
    "FROM SoftwareLicensingProduct " & _
    "WHERE PartialProductKey IS NOT NULL " & _
    "AND ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'")

For Each objProduct In colProducts
    If InStr(1, objProduct.Description, "Windows", vbTextCompare) > 0 Then
        winLicenseStatus = objProduct.LicenseStatus
        winLicenseDesc   = objProduct.Description
        If Not IsNull(objProduct.PartialProductKey) Then
            winPartialKey = Trim(objProduct.PartialProductKey)
        End If
        
        On Error Resume Next
        If Not IsNull(objProduct.ProductKeyChannel) Then
            winProductChannel = Trim(objProduct.ProductKeyChannel)
        End If
        On Error GoTo 0
        
        If winLicenseStatus = 1 Then winIsActivated = True
        Exit For
    End If
Next

' Phan loai Channel / License Type Windows
winLicenseType = ParseWindowsLicenseType(winProductChannel, winLicenseDesc)

' 1. Windows Installed Key (Full 25 ky tu tu Registry)
Dim fullWinInstalledKey, fullWinBiosKey
fullWinInstalledKey = GetInstalledProductKey()

' 2. Windows BIOS OEM Key (Full 25 ky tu tu Firmware Mainboard)
fullWinBiosKey = GetBIOSProductKey()

'-----------------------------------------------------------
' LAY THONG TIN MICROSOFT OFFICE & BAN QUYEN
'-----------------------------------------------------------
Dim officeName, officeVersion, officeBit, officeStatus, officeType, officePartialKey, officeRawDstatus
Dim hasOffice, osppPath

officeName       = "Khong cai dat Microsoft Office"
officeVersion    = "N/A"
officeBit        = "N/A"
officeStatus     = "N/A"
officeType       = "N/A"
officePartialKey = "N/A"
officeRawDstatus = ""
hasOffice        = False

' 1. Kiem tra phien ban Office cai tren may
Call DetectInstalledOffice(officeName, officeVersion, officeBit, hasOffice)

' 2. Tim script ospp.vbs cua Office de kiem tra ban quyen
osppPath = FindOsppPath()

If osppPath <> "" Then
    hasOffice = True
    officeRawDstatus = RunAndCapture("cscript.exe //nologo """ & osppPath & """ /dstatus")
    Call ParseOsppOutput(officeRawDstatus, officeStatus, officeType, officePartialKey)
Else
    ' Fallback qua WMI neu khong tim thay file ospp.vbs
    Call DetectOfficeWMI(officeStatus, officeType, officePartialKey, hasOffice)
End If

'-----------------------------------------------------------
' LAY THONG TIN THO TU SLMGR (Windows Chi tiet)
'-----------------------------------------------------------
Dim rawDli, rawDlv, sysFolder
sysFolder = objWshShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32"

rawDli = RunAndCapture("cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /dli")
rawDlv = RunAndCapture("cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /dlv")

'-----------------------------------------------------------
' TAO NOI DUNG BAO CAO TONG HOP
'-----------------------------------------------------------
Dim winActivatedStr
If winIsActivated Then
    winActivatedStr = "DA KICH HOAT (Licensed / Activated)"
Else
    winActivatedStr = "CHUA KICH HOAT - " & GetLicenseStatusDesc(winLicenseStatus)
End If

Dim output
output = ""
output = output & "==========================================================" & vbCrLf
output = output & "  BAO CAO BAN QUYEN WINDOWS & MICROSOFT OFFICE" & vbCrLf
output = output & "  Thoi gian quet       : " & timestamp & vbCrLf
output = output & "==========================================================" & vbCrLf
output = output & vbCrLf

output = output & "--- 1. THONG TIN BAN QUYEN WINDOWS ---" & vbCrLf
output = output & "Installed Key (Full) : " & fullWinInstalledKey & vbCrLf
output = output & "Trang thai kich hoat : " & winActivatedStr & vbCrLf
output = output & "Loai ban quyen       : " & winLicenseType & vbCrLf
output = output & "Partial Key (5 ky tu): " & winPartialKey & vbCrLf
output = output & "BIOS/OEM Key (Goc)   : " & fullWinBiosKey & vbCrLf
output = output & "Product ID           : " & productId & vbCrLf
output = output & vbCrLf

output = output & "--- 2. THONG TIN BAN QUYEN MICROSOFT OFFICE ---" & vbCrLf
output = output & "Phien ban Office     : " & officeName & vbCrLf
If hasOffice Then
    output = output & "Kien truc / Version  : " & officeBit & " (Build " & officeVersion & ")" & vbCrLf
    output = output & "Trang thai kich hoat : " & officeStatus & vbCrLf
    output = output & "Loai ban quyen       : " & officeType & vbCrLf
    output = output & "Partial Key (5 ky tu): " & officePartialKey & vbCrLf
Else
    output = output & "Trang thai           : Khong phat hien goi Office tren he thong" & vbCrLf
End If
output = output & vbCrLf

output = output & "--- 3. THONG TIN THIET BI & NGUOI DUNG ---" & vbCrLf
output = output & "Ten may              : " & computerName & vbCrLf
output = output & "Dong may             : " & computerModel & vbCrLf
output = output & "He dieu hanh (OS)    : " & osName & vbCrLf
output = output & "Edition              : " & edition & vbCrLf
output = output & "Username             : " & userName & vbCrLf
output = output & "Dia chi IP           : " & ipAddress & vbCrLf
output = output & vbCrLf

output = output & "--- 4. CHI TIET GIAY PHEP WINDOWS (slmgr /dli) ---" & vbCrLf
output = output & rawDli & vbCrLf

output = output & "--- 5. TOAN BO THONG TIN KICH HOAT WINDOWS (slmgr /dlv) ---" & vbCrLf
output = output & rawDlv & vbCrLf

output = output & "--- 6. CHI TIET BAN QUYEN OFFICE (ospp.vbs /dstatus) ---" & vbCrLf
If officeRawDstatus <> "" Then
    output = output & officeRawDstatus & vbCrLf
Else
    output = output & "(Khong co du lieu tu ospp.vbs hoac may chua cai Office)" & vbCrLf
End If
output = output & vbCrLf

output = output & "==========================================================" & vbCrLf

'-----------------------------------------------------------
' GHI FILE KET QUA
'-----------------------------------------------------------
Dim outputFile, outPath
outPath = RESULT_FOLDER & "\" & computerName & ".txt"

On Error Resume Next
If Not objFSO.FolderExists(RESULT_FOLDER) Then
    objFSO.CreateFolder(RESULT_FOLDER)
End If

Set outputFile = objFSO.CreateTextFile(outPath, True, False)
If Err.Number <> 0 Then
    Dim tempPath
    tempPath = objWshShell.ExpandEnvironmentStrings("%TEMP%") & "\LicenseCheck_" & computerName & ".txt"
    Err.Clear
    Set outputFile = objFSO.CreateTextFile(tempPath, True, False)
End If
On Error GoTo 0

outputFile.Write output
outputFile.Close

'-----------------------------------------------------------
' GIAI PHONG BO NHO
'-----------------------------------------------------------
Set colOS         = Nothing
Set colProducts   = Nothing
Set objOS         = Nothing
Set objProduct    = Nothing
Set objFSO        = Nothing
Set objWshShell   = Nothing
Set objNetwork    = Nothing
Set objWMIService = Nothing

WScript.Quit 0

'===========================================================
' CAC HAM LAY THONG TIN MICROSOFT OFFICE
'===========================================================

Sub DetectInstalledOffice(ByRef outName, ByRef outVer, ByRef outBit, ByRef outFound)
    On Error Resume Next
    Const HKLM = &H80000002
    Dim objReg, arrSubKeys, subkey, dispName, dispVer
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    ' 1. Kiem tra Office Click-to-Run (Office 2016, 2019, 2021, 2024, Microsoft 365)
    Dim c2rVer, c2rProd, c2rPlatform
    objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ClientVersionToReport", c2rVer
    If IsNull(c2rVer) Or c2rVer = "" Then
        objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "VersionToReport", c2rVer
    End If
    objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "ProductReleaseIDs", c2rProd
    objReg.GetStringValue HKLM, "SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "Platform", c2rPlatform
    
    If Not IsNull(c2rProd) And c2rProd <> "" Then
        outFound = True
        outVer   = c2rVer
        outBit   = c2rPlatform
        outName  = FriendlyOfficeName(c2rProd)
        Exit Sub
    End If
    
    ' 2. Kiem tra Registry Uninstall (Standard MSI & 32-bit/64-bit)
    Dim uninstallPaths, uPath
    uninstallPaths = Array( _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        
    For Each uPath In uninstallPaths
        objReg.EnumKey HKLM, uPath, arrSubKeys
        If Not IsNull(arrSubKeys) Then
            For Each subkey In arrSubKeys
                dispName = ""
                dispVer  = ""
                objReg.GetStringValue HKLM, uPath & "\" & subkey, "DisplayName", dispName
                If Not IsNull(dispName) And dispName <> "" Then
                    If (InStr(1, dispName, "Microsoft Office", vbTextCompare) > 0 Or InStr(1, dispName, "Microsoft 365", vbTextCompare) > 0) _
                       And InStr(1, dispName, "MUI", vbTextCompare) = 0 _
                       And InStr(1, dispName, "Language", vbTextCompare) = 0 _
                       And InStr(1, dispName, "Proof", vbTextCompare) = 0 _
                       And InStr(1, dispName, "Access Runtime", vbTextCompare) = 0 Then
                        
                        outFound = True
                        outName  = Trim(dispName)
                        objReg.GetStringValue HKLM, uPath & "\" & subkey, "DisplayVersion", dispVer
                        If Not IsNull(dispVer) Then outVer = Trim(dispVer)
                        If InStr(uPath, "WOW6432Node") > 0 Then
                            outBit = "32-bit (x86)"
                        Else
                            outBit = "64-bit (x64)"
                        End If
                        Exit Sub
                    End If
                End If
            Next
        End If
    Next
    On Error GoTo 0
End Sub

Function FriendlyOfficeName(prodId)
    Dim p
    p = UCase(prodId)
    If InStr(p, "O365PROPLUS") > 0 Or InStr(p, "M365") > 0 Then
        FriendlyOfficeName = "Microsoft 365 Apps for enterprise (" & prodId & ")"
    ElseIf InStr(p, "O365BUSINESS") > 0 Then
        FriendlyOfficeName = "Microsoft 365 Apps for business (" & prodId & ")"
    ElseIf InStr(p, "PROPLUS2024") > 0 Then
        FriendlyOfficeName = "Microsoft Office Professional Plus 2024"
    ElseIf InStr(p, "PROPLUS2021") > 0 Then
        FriendlyOfficeName = "Microsoft Office Professional Plus 2021"
    ElseIf InStr(p, "PROPLUS2019") > 0 Then
        FriendlyOfficeName = "Microsoft Office Professional Plus 2019"
    ElseIf InStr(p, "PROPLUS2016") > 0 Or InStr(p, "PROPLUSRETAIL") > 0 Then
        FriendlyOfficeName = "Microsoft Office Professional Plus 2016"
    ElseIf InStr(p, "STANDARD2021") > 0 Then
        FriendlyOfficeName = "Microsoft Office Standard 2021"
    ElseIf InStr(p, "STANDARD2019") > 0 Then
        FriendlyOfficeName = "Microsoft Office Standard 2019"
    ElseIf InStr(p, "HOMESTUDENT2021") > 0 Then
        FriendlyOfficeName = "Microsoft Office Home and Student 2021"
    ElseIf InStr(p, "HOMESTUDENT2019") > 0 Then
        FriendlyOfficeName = "Microsoft Office Home and Student 2019"
    ElseIf InStr(p, "HOMEBUSINESS2021") > 0 Then
        FriendlyOfficeName = "Microsoft Office Home and Business 2021"
    ElseIf InStr(p, "HOMEBUSINESS2019") > 0 Then
        FriendlyOfficeName = "Microsoft Office Home and Business 2019"
    Else
        FriendlyOfficeName = "Microsoft Office (" & prodId & ")"
    End If
End Function

Function FindOsppPath()
    On Error Resume Next
    Dim progFiles, progFilesX86, paths, p
    progFiles    = objWshShell.ExpandEnvironmentStrings("%ProgramFiles%")
    progFilesX86 = objWshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
    If InStr(progFilesX86, "%") > 0 Then progFilesX86 = progFiles
    
    paths = Array( _
        progFiles    & "\Microsoft Office\root\Office16\ospp.vbs", _
        progFilesX86 & "\Microsoft Office\root\Office16\ospp.vbs", _
        progFiles    & "\Microsoft Office\Office16\ospp.vbs", _
        progFilesX86 & "\Microsoft Office\Office16\ospp.vbs", _
        progFiles    & "\Microsoft Office\root\Office15\ospp.vbs", _
        progFilesX86 & "\Microsoft Office\root\Office15\ospp.vbs", _
        progFiles    & "\Microsoft Office\Office15\ospp.vbs", _
        progFilesX86 & "\Microsoft Office\Office15\ospp.vbs", _
        progFiles    & "\Microsoft Office\Office14\ospp.vbs", _
        progFilesX86 & "\Microsoft Office\Office14\ospp.vbs" _
    )
    
    FindOsppPath = ""
    For Each p In paths
        If objFSO.FileExists(p) Then
            FindOsppPath = p
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

Sub ParseOsppOutput(raw, ByRef outStatus, ByRef outType, ByRef outKey)
    On Error Resume Next
    Dim lines, line, trimmedLine
    outStatus = "Chua xac dinh"
    outType   = "N/A"
    outKey    = "N/A"
    
    If InStr(raw, "---LICENSED---") > 0 Then
        outStatus = "DA KICH HOAT (---LICENSED---)"
    ElseIf InStr(raw, "---NOTIFICATIONS---") > 0 Then
        outStatus = "CHUA KICH HOAT / THONG BAO (---NOTIFICATIONS---)"
    ElseIf InStr(raw, "---OOB_GRACE---") > 0 Then
        outStatus = "DANG DUNG THU / GRACE PERIOD (---OOB_GRACE---)"
    ElseIf InStr(raw, "---UNLICENSED---") > 0 Or InStr(raw, "---NOT LICENSED---") > 0 Then
        outStatus = "CHUA CAP PHEP (---UNLICENSED---)"
    End If
    
    lines = Split(raw, vbCrLf)
    For Each line In lines
        trimmedLine = Trim(line)
        If InStr(1, trimmedLine, "LICENSE DESCRIPTION:", vbTextCompare) > 0 Then
            outType = Trim(Mid(trimmedLine, InStr(trimmedLine, ":") + 1))
        ElseIf InStr(1, trimmedLine, "Last 5 characters of installed product key:", vbTextCompare) > 0 Then
            outKey = Trim(Mid(trimmedLine, InStr(trimmedLine, ":") + 1))
        End If
    Next
    On Error GoTo 0
End Sub

Sub DetectOfficeWMI(ByRef outStatus, ByRef outType, ByRef outKey, ByRef outFound)
    On Error Resume Next
    Dim colOff, objOff
    Set colOff = objWMIService.ExecQuery( _
        "SELECT LicenseStatus, Description, PartialProductKey " & _
        "FROM SoftwareLicensingProduct " & _
        "WHERE PartialProductKey IS NOT NULL " & _
        "AND ApplicationID='0ff1ce15-a989-479d-af46-f275c637705c'")
        
    For Each objOff In colOff
        outFound = True
        If objOff.LicenseStatus = 1 Then
            outStatus = "DA KICH HOAT (Licensed)"
        Else
            outStatus = "CHUA KICH HOAT (" & GetLicenseStatusDesc(objOff.LicenseStatus) & ")"
        End If
        outType = Trim(objOff.Description)
        If Not IsNull(objOff.PartialProductKey) Then
            outKey = Trim(objOff.PartialProductKey)
        End If
        Exit For
    Next
    On Error GoTo 0
End Sub

'===========================================================
' CAC HAM LAY THONG TIN WINDOWS & HARDWARE
'===========================================================

Function ParseWindowsLicenseType(channel, desc)
    Dim descUpper
    descUpper = UCase(desc)
    If channel <> "" Then
        If InStr(UCase(channel), "OEM") > 0 Then
            ParseWindowsLicenseType = "OEM (" & channel & ")"
        ElseIf InStr(UCase(channel), "RETAIL") > 0 Then
            ParseWindowsLicenseType = "Retail / Digital License (" & channel & ")"
        ElseIf InStr(UCase(channel), "VOLUME:MAK") > 0 Then
            ParseWindowsLicenseType = "Volume MAK"
        ElseIf InStr(UCase(channel), "VOLUME:GVLK") > 0 Or InStr(UCase(channel), "VOLUME") > 0 Then
            ParseWindowsLicenseType = "Volume KMS (GVLK)"
        Else
            ParseWindowsLicenseType = channel
        End If
    Else
        If InStr(descUpper, "KMS") > 0 Or InStr(descUpper, "GVLK") > 0 Then
            ParseWindowsLicenseType = "Volume KMS (GVLK)"
        ElseIf InStr(descUpper, "MAK") > 0 Then
            ParseWindowsLicenseType = "Volume MAK"
        ElseIf InStr(descUpper, "OEM") > 0 Then
            ParseWindowsLicenseType = "OEM"
        ElseIf InStr(descUpper, "RETAIL") > 0 Then
            ParseWindowsLicenseType = "Retail / Digital License"
        ElseIf desc <> "" Then
            ParseWindowsLicenseType = desc
        Else
            ParseWindowsLicenseType = "Chua xac dinh"
        End If
    End If
End Function

Function GetComputerModel()
    On Error Resume Next
    Dim colCS, objCS, colBB, objBB
    Dim mfg, mdl, fullStr
    mfg = ""
    mdl = ""
    
    Set colCS = objWMIService.ExecQuery("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
    For Each objCS In colCS
        If Not IsNull(objCS.Manufacturer) Then mfg = Trim(objCS.Manufacturer)
        If Not IsNull(objCS.Model) Then mdl = Trim(objCS.Model)
        Exit For
    Next
    
    If mfg = "" Or InStr(1, mfg, "To Be Filled", vbTextCompare) > 0 Or InStr(1, mfg, "System manufacturer", vbTextCompare) > 0 Then
        Set colBB = objWMIService.ExecQuery("SELECT Manufacturer, Product FROM Win32_BaseBoard")
        For Each objBB In colBB
            If Not IsNull(objBB.Manufacturer) Then mfg = Trim(objBB.Manufacturer)
            If Not IsNull(objBB.Product) Then mdl = Trim(objBB.Product)
            Exit For
        Next
    End If
    
    If mfg = "" And mdl = "" Then
        GetComputerModel = "N/A"
        Exit Function
    ElseIf mfg = "" Then
        GetComputerModel = mdl
        Exit Function
    ElseIf mdl = "" Then
        GetComputerModel = mfg
        Exit Function
    End If
    
    If InStr(1, mdl, mfg, vbTextCompare) = 1 Then
        fullStr = mdl
    Else
        fullStr = mfg & " " & mdl
    End If
    
    GetComputerModel = fullStr
    On Error GoTo 0
End Function

Function GetInstalledProductKey()
    On Error Resume Next
    Dim rawBytes, keyStr
    rawBytes = objWshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
    
    If Err.Number <> 0 Or IsNull(rawBytes) Or IsEmpty(rawBytes) Then
        GetInstalledProductKey = "N/A (Khong the doc tu Registry)"
        Err.Clear
        Exit Function
    End If
    
    keyStr = DecodeProductKey(rawBytes)
    If Trim(keyStr) = "" Or keyStr = "----" Then
        GetInstalledProductKey = "N/A (Khong the giai ma)"
    Else
        GetInstalledProductKey = keyStr
    End If
    On Error GoTo 0
End Function

Function DecodeProductKey(Key)
    On Error Resume Next
    Const KeyOffset = 52
    Dim isWin8, i, x, Cur, Last, KeyOutput, Chars
    Dim keypart1, insert, AKey, BKey, CKey, DKey, EKey
    
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    KeyOutput = ""
    
    isWin8 = (Key(66) \ 6) And 1
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
    
    i = 24
    Last = 0
    Do
        Cur = 0
        x = 14
        Do
            Cur = Cur * 256
            Cur = Key(x + KeyOffset) + Cur
            Key(x + KeyOffset) = (Cur \ 24)
            Cur = Cur Mod 24
            x = x - 1
        Loop While x >= 0
        i = i - 1
        KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
        Last = Cur
    Loop While i >= 0

    If (isWin8 = 1) Then
        keypart1 = Mid(KeyOutput, 2, Last)
        insert = "N"
        KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
        If Last = 0 Then KeyOutput = insert & KeyOutput
    End If

    If Len(KeyOutput) >= 25 Then
        AKey = Mid(KeyOutput, 1, 5)
        BKey = Mid(KeyOutput, 6, 5)
        CKey = Mid(KeyOutput, 11, 5)
        DKey = Mid(KeyOutput, 16, 5)
        EKey = Mid(KeyOutput, 21, 5)
        DecodeProductKey = AKey & "-" & BKey & "-" & CKey & "-" & DKey & "-" & EKey
    Else
        DecodeProductKey = "N/A"
    End If
    On Error GoTo 0
End Function

Function GetBIOSProductKey()
    On Error Resume Next
    Dim colSLS, objSLS, bKey
    bKey = ""
    
    Set colSLS = objWMIService.ExecQuery("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService")
    For Each objSLS In colSLS
        If Not IsNull(objSLS.OA3xOriginalProductKey) Then
            bKey = Trim(objSLS.OA3xOriginalProductKey)
            If bKey <> "" Then Exit For
        End If
    Next
    
    If bKey = "" Then
        GetBIOSProductKey = "N/A (May khong co OEM Key nhung trong BIOS/Mainboard)"
    Else
        GetBIOSProductKey = bKey
    End If
    On Error GoTo 0
End Function

Function GetIPAddress()
    Dim colAdapters, objAdapter
    Dim addresses, addr
    GetIPAddress = "N/A"
    Set colAdapters = objWMIService.ExecQuery( _
        "SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each objAdapter In colAdapters
        addresses = objAdapter.IPAddress
        If Not IsNull(addresses) Then
            For Each addr In addresses
                If InStr(addr, ".") > 0 And Left(addr, 3) <> "169" Then
                    GetIPAddress = addr
                    Exit For
                End If
            Next
        End If
        If GetIPAddress <> "N/A" Then Exit For
    Next
End Function

Function ExtractEdition(osCaption)
    Dim editions, ed
    editions = Array("Home Single Language", "Home", "Professional", "Pro", "Enterprise", "Education", _
                     "LTSC", "Server")
    ExtractEdition = "Unknown"
    For Each ed In editions
        If InStr(1, osCaption, ed, vbTextCompare) > 0 Then
            ExtractEdition = ed
            Exit For
        End If
    Next
End Function

Function RunAndCapture(cmd)
    Dim tmpFile, result, fso2, ts
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    tmpFile = objWshShell.ExpandEnvironmentStrings("%TEMP%") & "\cmd_out_" & Timer & ".tmp"

    On Error Resume Next
    objWshShell.Run "cmd.exe /c """ & cmd & """ > """ & tmpFile & """ 2>&1", 0, True

    result = ""
    If fso2.FileExists(tmpFile) Then
        Set ts = fso2.OpenTextFile(tmpFile, 1, False)
        result = ts.ReadAll()
        ts.Close
        fso2.DeleteFile tmpFile, True
    End If
    On Error GoTo 0

    Set fso2 = Nothing
    RunAndCapture = result
End Function

Function GetLicenseStatusDesc(statusCode)
    Select Case statusCode
        Case 0:  GetLicenseStatusDesc = "Chua duoc cap phep (Unlicensed)"
        Case 1:  GetLicenseStatusDesc = "Da duoc cap phep (Licensed)"
        Case 2:  GetLicenseStatusDesc = "Het han grace period ban dau"
        Case 3:  GetLicenseStatusDesc = "Grace period them"
        Case 4:  GetLicenseStatusDesc = "Grace period not yet"
        Case 5:  GetLicenseStatusDesc = "Thong bao"
        Case 6:  GetLicenseStatusDesc = "Extended Grace"
        Case Else: GetLicenseStatusDesc = "Ma trang thai: " & statusCode
    End Select
End Function
