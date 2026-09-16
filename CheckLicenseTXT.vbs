'===========================================================
' CheckLicense.vbs (Bao cao ban quyen Windows - License Audit)
' Chay hoan toan silent (khong MsgBox / Khong cua so)
' Thu thap thong tin uu tien phuc vu lam bao cao ban quyen:
'   1. Installed Product Key (Full 25 ky tu), Trang thai & Loai Key
'   2. BIOS / OEM Key (Full 25 ky tu goc theo Mainboard)
'   3. Thong tin may tinh (Ten may, Dong may, OS, Edition, User, IP)
'   4. Thong tin chi tiet tu slmgr /dli va slmgr /dlv
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
' LAY THONG TIN CO BAN
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
' LAY THONG TIN BAN QUYEN (WMI SoftwareLicensing)
'-----------------------------------------------------------
Dim licenseStatus, licenseDesc, partialKey, licenseType, productChannel
Dim isActivated

licenseStatus  = -1
licenseDesc    = ""
partialKey     = "N/A"
licenseType    = "Unknown"
productChannel = ""
isActivated    = False

Set colProducts = objWMIService.ExecQuery( _
    "SELECT LicenseStatus, Description, PartialProductKey, ProductKeyChannel " & _
    "FROM SoftwareLicensingProduct " & _
    "WHERE PartialProductKey IS NOT NULL " & _
    "AND ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'")

For Each objProduct In colProducts
    If InStr(1, objProduct.Description, "Windows", vbTextCompare) > 0 Then
        licenseStatus = objProduct.LicenseStatus
        licenseDesc   = objProduct.Description
        If Not IsNull(objProduct.PartialProductKey) Then
            partialKey = Trim(objProduct.PartialProductKey)
        End If
        
        On Error Resume Next
        If Not IsNull(objProduct.ProductKeyChannel) Then
            productChannel = Trim(objProduct.ProductKeyChannel)
        End If
        On Error GoTo 0
        
        If licenseStatus = 1 Then isActivated = True
        Exit For
    End If
Next

'-----------------------------------------------------------
' PHAN LOAI LOAI BAN QUYEN / LICENSE CHANNEL
'-----------------------------------------------------------
Dim descUpper
descUpper = UCase(licenseDesc)

If productChannel <> "" Then
    If InStr(UCase(productChannel), "OEM") > 0 Then
        licenseType = "OEM (" & productChannel & ")"
    ElseIf InStr(UCase(productChannel), "RETAIL") > 0 Then
        licenseType = "Retail / Digital License (" & productChannel & ")"
    ElseIf InStr(UCase(productChannel), "VOLUME:MAK") > 0 Then
        licenseType = "Volume MAK"
    ElseIf InStr(UCase(productChannel), "VOLUME:GVLK") > 0 Or InStr(UCase(productChannel), "VOLUME") > 0 Then
        licenseType = "Volume KMS (GVLK)"
    Else
        licenseType = productChannel
    End If
Else
    If InStr(descUpper, "KMS") > 0 Or InStr(descUpper, "GVLK") > 0 Then
        licenseType = "Volume KMS (GVLK)"
    ElseIf InStr(descUpper, "MAK") > 0 Then
        licenseType = "Volume MAK"
    ElseIf InStr(descUpper, "OEM") > 0 Then
        licenseType = "OEM"
    ElseIf InStr(descUpper, "RETAIL") > 0 Then
        licenseType = "Retail / Digital License"
    ElseIf licenseDesc <> "" Then
        licenseType = licenseDesc
    Else
        licenseType = "Chua xac dinh"
    End If
End If

'-----------------------------------------------------------
' LAY FULL PRODUCT KEY (25 KY TU)
'-----------------------------------------------------------
Dim fullInstalledKey, fullBiosKey

' 1. Key dang cai dat trong Windows (giai ma tu Registry DigitalProductId)
fullInstalledKey = GetInstalledProductKey()

' 2. Key goc OEM nhung trong BIOS/UEFI firmware (Mainboard)
fullBiosKey = GetBIOSProductKey()

'-----------------------------------------------------------
' LAY THONG TIN THO TU SLMGR (phan chi tiet)
'-----------------------------------------------------------
Dim rawDli, rawDlv
Dim sysFolder
sysFolder = objWshShell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32"

rawDli = RunAndCapture("cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /dli")
rawDlv = RunAndCapture("cscript.exe //nologo """ & sysFolder & "\slmgr.vbs"" /dlv")

'-----------------------------------------------------------
' TAO NOI DUNG BAO CAO (UU TIEN INSTALLED KEY LEN DAU)
'-----------------------------------------------------------
Dim activatedStr
Dim output

If isActivated Then
    activatedStr = "DA KICH HOAT (Licensed / Activated)"
Else
    activatedStr = "CHUA KICH HOAT - " & GetLicenseStatusDesc(licenseStatus)
End If

output = ""
output = output & "==========================================================" & vbCrLf
output = output & "  BAO CAO BAN QUYEN WINDOWS (WINDOWS LICENSE AUDIT)" & vbCrLf
output = output & "  Thoi gian quet       : " & timestamp & vbCrLf
output = output & "==========================================================" & vbCrLf
output = output & vbCrLf

output = output & "--- 1. THONG TIN BAN QUYEN (INSTALLED LICENSE) ---" & vbCrLf
output = output & "Installed Key (Full) : " & fullInstalledKey & vbCrLf
output = output & "Trang thai kich hoat : " & activatedStr & vbCrLf
output = output & "Loai ban quyen       : " & licenseType & vbCrLf
output = output & "Partial Key (5 ky tu): " & partialKey & vbCrLf
output = output & "BIOS/OEM Key (Goc)   : " & fullBiosKey & vbCrLf
output = output & "Product ID           : " & productId & vbCrLf
output = output & vbCrLf

output = output & "--- 2. THONG TIN THIET BI & NGUOI DUNG ---" & vbCrLf
output = output & "Ten may              : " & computerName & vbCrLf
output = output & "Dong may             : " & computerModel & vbCrLf
output = output & "He dieu hanh (OS)    : " & osName & vbCrLf
output = output & "Edition              : " & edition & vbCrLf
output = output & "Username             : " & userName & vbCrLf
output = output & "Dia chi IP           : " & ipAddress & vbCrLf
output = output & vbCrLf

output = output & "--- 3. CHI TIET GIAY PHEP (slmgr /dli) ---" & vbCrLf
output = output & rawDli & vbCrLf

output = output & "--- 4. TOAN BO THONG TIN KICH HOAT (slmgr /dlv) ---" & vbCrLf
output = output & rawDlv & vbCrLf

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
    ' Fallback: luu vao %TEMP% neu khong ghi duoc vao network share
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
' CAC HAM LAY THONG TIN HE THONG & PRODUCT KEY
'===========================================================

' Ham lay Hang va Dong may tinh ghep thanh mot dong duy nhat
Function GetComputerModel()
    On Error Resume Next
    Dim colCS, objCS, colBB, objBB
    Dim mfg, mdl, fullStr
    mfg = ""
    mdl = ""
    
    ' 1. Lay tu Win32_ComputerSystem
    Set colCS = objWMIService.ExecQuery("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
    For Each objCS In colCS
        If Not IsNull(objCS.Manufacturer) Then mfg = Trim(objCS.Manufacturer)
        If Not IsNull(objCS.Model) Then mdl = Trim(objCS.Model)
        Exit For
    Next
    
    ' 2. Neu tra ve chuoi mac dinh/may lap rap (System manufacturer / To Be Filled By O.E.M.), thu qua Win32_BaseBoard
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
    
    ' Kiem tra tranh trung lap neu ten Model da bat dau bang ten Hang (VD: "Dell Optiplex 3060")
    If InStr(1, mdl, mfg, vbTextCompare) = 1 Then
        fullStr = mdl
    Else
        fullStr = mfg & " " & mdl
    End If
    
    GetComputerModel = fullStr
    On Error GoTo 0
End Function

' Ham doc va giai ma Installed Product Key tu Registry (Win 7/8/10/11)
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

' Thuat toan giai ma Base-24 DigitalProductId tuong thich Win 7, 8, 8.1, 10, 11
Function DecodeProductKey(Key)
    On Error Resume Next
    Const KeyOffset = 52
    Dim isWin8, i, x, Cur, Last, KeyOutput, Chars
    Dim keypart1, insert, AKey, BKey, CKey, DKey, EKey
    
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    KeyOutput = ""
    
    ' Kiem tra flag phien ban Win 8 / 10 / 11 tai byte 66
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

' Ham lay Full OEM Key nhung trong BIOS/UEFI qua WMI SoftwareLicensingService
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
    tmpFile = objWshShell.ExpandEnvironmentStrings("%TEMP%") & "\slmgr_out_" & Timer & ".tmp"

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
