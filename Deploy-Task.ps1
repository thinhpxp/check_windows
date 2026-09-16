#Requires -Version 3.0
<#
.SYNOPSIS
    Deploy-Task.ps1
    Tao Scheduled Task tren cac may Windows de chay CheckLicense.vbs tu dong.

.MO TA
    Script nay duoc chay MOT LAN tu may Fedora (thong qua PowerShell Remoting,
    hoac thu cong tren 1 may DC) de tao Task Scheduler entry tren cac may Windows.

    Sau khi deploy:
    - Task se chay CheckLicense.vbs luc KHOI DONG may (Startup)
    - Chay bang tai khoan SYSTEM (khong can user login)
    - Hoan toan an, khong hien thi cua so
    - Ket qua ghi vao: \\SERVER\share\Results\<COMPUTERNAME>.txt

.CACH SU DUNG
    Cach 1 - Chay tren tung may Windows qua PowerShell:
        .\Deploy-Task.ps1

    Cach 2 - Chay tu xa qua WinRM (neu da enable):
        $computers = @("PC-001","PC-002","PC-003")
        Invoke-Command -ComputerName $computers -FilePath .\Deploy-Task.ps1

    Cach 3 - Dung Group Policy (GPO) Startup Script:
        - Mo Group Policy Management Console
        - Tao GPO moi, link vao OU chua cac may can quet
        - Computer Configuration > Policies > Windows Settings > Scripts > Startup
        - Them Deploy-Task.ps1 vao danh sach script

.YEU CAU
    - PowerShell 3.0+
    - Quyen Administrator tren may dich
    - CheckLicense.vbs da co mat tren Network Drive

.CAU HINH - CHINH SUA TRUOC KHI CHAY
#>

# ============================================================
# CAU HINH - CHINH SUA THEO MOI TRUONG CUA BAN
# ============================================================

# Duong dan den script tren Network Drive
$ScriptPath   = "\\SERVER\share\Scripts\CheckLicense.vbs"

# Ten va mo ta Task
$TaskName     = "LicenseCheck_Silent"
$TaskDesc     = "Quet thong tin ban quyen Windows va luu vao Network Drive"

# Thoi diem chay: AtStartup (khi khoi dong may)
# Co the doi thanh: "Daily", "Weekly", "AtLogon"
$TriggerType  = "AtStartup"

# Tre them N giay sau khi startup truoc khi chay (tranh xung dot network)
$DelaySeconds = 60

# ============================================================
# KIEM TRA QUYEN ADMIN
# ============================================================
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Script can chay voi quyen Administrator!"
    exit 1
}

# ============================================================
# XOA TASK CU (neu ton tai)
# ============================================================
Write-Host "[*] Kiem tra Task Scheduler hien tai..." -ForegroundColor Cyan
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "[!] Xoa task cu: $TaskName" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# ============================================================
# TAO SCHEDULED TASK
# ============================================================
Write-Host "[*] Tao Scheduled Task: $TaskName" -ForegroundColor Cyan

# Action: chay cscript.exe //nologo //b de hoan toan an
$action = New-ScheduledTaskAction `
    -Execute "cscript.exe" `
    -Argument "//nologo //b ""$ScriptPath"""

# Trigger: khi khoi dong may, tre $DelaySeconds giay
$trigger = New-ScheduledTaskTrigger -AtStartup
$trigger.Delay = "PT${DelaySeconds}S"   # ISO 8601 duration

# Cai dat: chay voi SYSTEM, quyen cao nhat, an
$settings = New-ScheduledTaskSettingsSet `
    -Hidden `
    -RunOnlyIfNetworkAvailable `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew

# Principal: SYSTEM account
$principal = New-ScheduledTaskPrincipal `
    -UserId "NT AUTHORITY\SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

# Dang ky task
try {
    Register-ScheduledTask `
        -TaskName    $TaskName `
        -Description $TaskDesc `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Force | Out-Null

    Write-Host "[OK] Task da duoc tao thanh cong: $TaskName" -ForegroundColor Green
    Write-Host "     Script   : $ScriptPath"
    Write-Host "     Trigger  : $TriggerType + $DelaySeconds giay"
    Write-Host "     Tai khoan: NT AUTHORITY\SYSTEM"
}
catch {
    Write-Error "[FAIL] Khong the tao task: $_"
    exit 2
}

# ============================================================
# CHAY NGAY MOT LAN (khong doi reboot)
# ============================================================
Write-Host ""
Write-Host "[*] Chay task ngay bay gio de kiem tra..." -ForegroundColor Cyan
try {
    Start-ScheduledTask -TaskName $TaskName
    Start-Sleep -Seconds ($DelaySeconds + 10)

    $taskInfo = Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo
    if ($taskInfo.LastTaskResult -eq 0) {
        Write-Host "[OK] Task chay thanh cong! Ket qua da ghi vao Network Drive." -ForegroundColor Green
    } else {
        Write-Host "[WARN] Task chay xong voi ma loi: $($taskInfo.LastTaskResult)" -ForegroundColor Yellow
    }
}
catch {
    Write-Warning "Khong the kiem tra trang thai task: $_"
}

Write-Host ""
Write-Host "=== HOAN THANH ===" -ForegroundColor Green
Write-Host "Ket qua quet se co tai: \\SERVER\share\Results\$env:COMPUTERNAME.txt"
