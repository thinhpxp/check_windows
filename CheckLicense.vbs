'===========================================================
' CheckLicense.vbs (Bao cao ban quyen Windows & MS Office)
' Chay hoan toan SILENT 100% (Khong Popup / Khong MsgBox / Khong Window)
' Xuat ket qua dong thoi ra 2 dinh dang:
'   1. <COMPUTERNAME>.txt : Bao cao chi tiet day du
'   2. <COMPUTERNAME>.csv : Du lieu bang de import Excel / Tong hop
' Luu ket qua vao: \\10.43.4.103\doc\33_THINH\ketquakiemtra\
'===========================================================

Option Explicit

'-----------------------------------------------------------
' CHONG POPUP: TU DONG CHUYEN SANG CSCRIPT //B NEU DUNG WSCRIPT
'-----------------------------------------------------------
If InStr(1, WScript.FullName, "wscript.exe", vbTextCompare) > 0 Then
    Dim objSelfShell
    Set objSelfShell = CreateObject("WScript.Shell")
    ' Chay lai bang cscript.exe //nologo //b hoan toan an (0 = hidden)
    objSelfShell.Run "cscript.exe //nologo //b """ & WScript.ScriptFullName & """", 0, False
    Set objSelfShell = Nothing
    WScript.Quit 0
End If

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
    Dim officeDir, osppFile
    officeDir = objFSO.GetParentFolderName(osppPath)
    osppFile  = objFSO.GetFileName(osppPath)
    
    ' Chay ospp.vbs hoan toan an tu dung thu muc goc cua no
    officeRawDstatus = RunScriptInDir(officeDir, osppFile, "/dstatus")
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

rawDli = RunScriptInDir(sysFolder, "slmgr.vbs", "/dli")
rawDlv = RunScriptInDir(sysFolder, "slmgr.vbs", "/dlv")

'-----------------------------------------------------------
' TAO NOI DUNG BAO CAO TEXT (.TXT)
'-----------------------------------------------------------
Dim winActivatedStr
If winIsActivated Then
    winActivatedStr = "DA KICH HOAT (Licensed / Activated)"
Else
    winActivatedStr = "CHUA KICH HOAT - " & GetLicenseStatusDesc(winLicenseStatus)
End If

Dim outputTxt
outputTxt = ""
outputTxt = outputTxt & "==========================================================" & vbCrLf
outputTxt = outputTxt & "  BAO CAO BAN QUYEN WINDOWS & MICROSOFT OFFICE" & vbCrLf
outputTxt = outputTxt & "  Thoi gian quet       : " & timestamp & vbCrLf
outputTxt = outputTxt & "==========================================================" & vbCrLf
outputTxt = outputTxt & vbCrLf

outputTxt = outputTxt & "--- 1. THONG TIN THIET BI & NGUOI DUNG ---" & vbCrLf
outputTxt = outputTxt & "Ten may              : " & computerName & vbCrLf
outputTxt = outputTxt & "Dong may             : " & computerModel & vbCrLf
outputTxt = outputTxt & "He dieu hanh (OS)    : " & osName & vbCrLf
outputTxt = outputTxt & "Edition              : " & edition & vbCrLf
outputTxt = outputTxt & "Username             : " & userName & vbCrLf
outputTxt = outputTxt & "Dia chi IP           : " & ipAddress & vbCrLf
outputTxt = outputTxt & vbCrLf

outputTxt = outputTxt & "--- 2. TOM TAT THONG TIN BAN QUYEN WINDOWS ---" & vbCrLf
outputTxt = outputTxt & "Installed Key (Full) : " & fullWinInstalledKey & vbCrLf
outputTxt = outputTxt & "Trang thai kich hoat : " & winActivatedStr & vbCrLf
outputTxt = outputTxt & "Loai ban quyen       : " & winLicenseType & vbCrLf
outputTxt = outputTxt & "Partial Key (5 ky tu): " & winPartialKey & vbCrLf
outputTxt = outputTxt & "BIOS/OEM Key (Goc)   : " & fullWinBiosKey & vbCrLf
outputTxt = outputTxt & "Product ID           : " & productId & vbCrLf
outputTxt = outputTxt & vbCrLf

outputTxt = outputTxt & "--- 3. TOM TAT THONG TIN BAN QUYEN MICROSOFT OFFICE ---" & vbCrLf
outputTxt = outputTxt & "Phien ban Office     : " & officeName & vbCrLf
If hasOffice Then
    outputTxt = outputTxt & "Kien truc / Version  : " & officeBit & " (Build " & officeVersion & ")" & vbCrLf
    outputTxt = outputTxt & "Trang thai kich hoat : " & officeStatus & vbCrLf
    outputTxt = outputTxt & "Loai ban quyen       : " & officeType & vbCrLf
    outputTxt = outputTxt & "Partial Key (5 ky tu): " & officePartialKey & vbCrLf
Else
    outputTxt = outputTxt & "Trang thai           : Khong phat hien goi Office tren he thong" & vbCrLf
End If
outputTxt = outputTxt & vbCrLf

outputTxt = outputTxt & "--- 4. CHI TIET GIAY PHEP WINDOWS (slmgr /dli) ---" & vbCrLf
outputTxt = outputTxt & rawDli & vbCrLf

outputTxt = outputTxt & "--- 5. TOAN BO THONG TIN KICH HOAT WINDOWS (slmgr /dlv) ---" & vbCrLf
outputTxt = outputTxt & rawDlv & vbCrLf

outputTxt = outputTxt & "--- 6. CHI TIET BAN QUYEN OFFICE (ospp.vbs /dstatus) ---" & vbCrLf
If officeRawDstatus <> "" Then
    outputTxt = outputTxt & officeRawDstatus & vbCrLf
Else
    outputTxt = outputTxt & "(Khong co du lieu tu ospp.vbs hoac may chua cai Office)" & vbCrLf
End If
outputTxt = outputTxt & vbCrLf

outputTxt = outputTxt & "==========================================================" & vbCrLf

'-----------------------------------------------------------
' TAO NOI DUNG DANG BANG CSV (.CSV)
'-----------------------------------------------------------
Dim csvHeader, csvRow, outputCsv
csvHeader = "Ten may,Dong may,He dieu hanh,Edition,Product ID,Username,Dia chi IP,Win Installed Key,Win Trang thai,Win Loai key,Win Partial Key,Win BIOS Key,Office Phien ban,Office Kien truc,Office Build,Office Trang thai,Office Loai key,Office Partial Key,Thoi gian quet"

csvRow = EscapeCSV(computerName) & "," & _
         EscapeCSV(computerModel) & "," & _
         EscapeCSV(osName) & "," & _
         EscapeCSV(edition) & "," & _
         EscapeCSV(productId) & "," & _
         EscapeCSV(userName) & "," & _
         EscapeCSV(ipAddress) & "," & _
         EscapeCSV(fullWinInstalledKey) & "," & _
         EscapeCSV(winActivatedStr) & "," & _
         EscapeCSV(winLicenseType) & "," & _
         EscapeCSV(winPartialKey) & "," & _
         EscapeCSV(fullWinBiosKey) & "," & _
         EscapeCSV(officeName) & "," & _
         EscapeCSV(officeBit) & "," & _
         EscapeCSV(officeVersion) & "," & _
         EscapeCSV(officeStatus) & "," & _
         EscapeCSV(officeType) & "," & _
         EscapeCSV(officePartialKey) & "," & _
         EscapeCSV(timestamp)

outputCsv = csvHeader & vbCrLf & csvRow & vbCrLf

'-----------------------------------------------------------
' GHI FILE KET QUA (TXT & CSV)
'-----------------------------------------------------------
Dim txtPath, csvPath, targetFolder
targetFolder = RESULT_FOLDER

On Error Resume Next
If Not objFSO.FolderExists(targetFolder) Then
    objFSO.CreateFolder(targetFolder)
End If
If Err.Number <> 0 Then
    targetFolder = objWshShell.ExpandEnvironmentStrings("%TEMP%")
    Err.Clear
End If
On Error GoTo 0

txtPath = targetFolder & "\" & computerName & ".txt"
csvPath = targetFolder & "\" & computerName & ".csv"

' Ghi file TXT
Call SaveTextFile(txtPath, outputTxt)

' Ghi file CSV (UTF-8 de Excel hien thi tieng Viet khong loi font)
Call SaveTextFileUTF8(csvPath, outputCsv)

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
' CAC HAM HO TRO GHI FILE & ESCAPE CSV
'===========================================================

Function EscapeCSV(val)
    Dim s
    s = CStr(val)
    s = Replace(s, """", """""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    EscapeCSV = """" & s & """"
End Function

Sub SaveTextFile(filePath, contentText)
    On Error Resume Next
    Dim fOut, fsoLocal
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    Set fOut = fsoLocal.CreateTextFile(filePath, True, False)
    fOut.Write contentText
    fOut.Close
    Set fOut = Nothing
    Set fsoLocal = Nothing
    On Error GoTo 0
End Sub

Sub SaveTextFileUTF8(filePath, contentText)
    On Error Resume Next
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText contentText
    objStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    objStream.Close
    Set objStream = Nothing
    
    If Err.Number <> 0 Then
        Err.Clear
        Call SaveTextFile(filePath, contentText)
    End If
    On Error GoTo 0
End Sub

'===========================================================
' CAC HAM CHAY SCRIPT HE THONG (100% SILENT & KHONG LOI QUOTE)
'===========================================================

Function RunScriptInDir(dirPath, scriptName, switchParam)
    On Error Resume Next
    Dim tmpFile, batFile, result, fso2, ts, fBat
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    
    Randomize
    Dim rndId
    rndId = Int((999999 - 100000 + 1) * Rnd + 100000)
    
    tmpFile = objWshShell.ExpandEnvironmentStrings("%TEMP%") & "\out_" & rndId & ".tmp"
    batFile = objWshShell.ExpandEnvironmentStrings("%TEMP%") & "\run_" & rndId & ".bat"
    
    Set fBat = fso2.CreateTextFile(batFile, True, False)
    If dirPath <> "" Then
        fBat.WriteLine "@cd /d """ & dirPath & """"
    End If
    ' Su dung explicitly cscript.exe //nologo //b de chong tat ca popup/dialog
    fBat.WriteLine "@""%SystemRoot%\System32\cscript.exe"" //nologo //b """ & scriptName & """ " & switchParam & " > """ & tmpFile & """ 2>&1"
    fBat.Close
    
    ' Chay file bat hoan toan an (0 = an cua so, True = doi chay xong)
    objWshShell.Run "cmd.exe /c """"" & batFile & """""", 0, True
    
    result = ""
    If fso2.FileExists(tmpFile) Then
        Set ts = fso2.OpenTextFile(tmpFile, 1, False)
        result = ts.ReadAll()
        ts.Close
        fso2.DeleteFile tmpFile, True
    End If
    If fso2.FileExists(batFile) Then
        fso2.DeleteFile batFile, True
    End If
    
    Set fso2 = Nothing
    RunScriptInDir = result
    On Error GoTo 0
End Function

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
    If InStr(progFilesX86, "%") > 0 Or progFilesX86 = "" Then progFilesX86 = progFiles
    
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
    ElseIf InStr(raw, "ERROR CODE:") > 0 Then
        outStatus = "Loi xac thuc ban quyen"
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
