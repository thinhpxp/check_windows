# Hướng dẫn triển khai License Check

## Kiến trúc tổng quan

```
Fedora Server (Samba)
│
├── /srv/samba/share/
│   ├── Scripts/
│   │   └── CheckLicense.vbs     ← Script chạy trên từng máy Windows
│   └── Results/
│       ├── PC-001.txt
│       ├── PC-002.txt
│       └── ...                  ← Kết quả từng máy
│
└── Deploy-Task.ps1              ← Chạy 1 lần để cài Scheduled Task
```

---

## Bước 1 — Chuẩn bị Samba Share trên Fedora

### 1.1 Tạo cấu trúc thư mục

```bash
sudo mkdir -p /srv/samba/share/Scripts
sudo mkdir -p /srv/samba/share/Results

# Sao chép script vào share
sudo cp /home/thinhpxp/Downloads/LicenseCheck/CheckLicense.vbs \
        /srv/samba/share/Scripts/

# Phân quyền:
# - Scripts: Domain Users chỉ đọc
# - Results: Domain Users có thể ghi (để máy tự lưu kết quả)
sudo chown -R root:sambashare /srv/samba/share/Scripts
sudo chown -R root:sambashare /srv/samba/share/Results

sudo chmod 755 /srv/samba/share/Scripts
sudo chmod 775 /srv/samba/share/Results   # group write
sudo chmod 644 /srv/samba/share/Scripts/CheckLicense.vbs
```

### 1.2 Cấu hình `/etc/samba/smb.conf`

```ini
[share]
   path = /srv/samba/share
   browseable = yes
   read only = no
   valid users = @"Domain Users"
   write list = @"Domain Users"   ; cần để máy ghi vào Results/
   create mask = 0664
   directory mask = 0775
   force group = sambashare
```

```bash
# Khởi động lại Samba
sudo systemctl restart smb nmb
```

### 1.3 Kiểm tra từ Windows client

```
# Mở Run (Win+R) và thử:
\\SERVER\share\Scripts\CheckLicense.vbs
```

---

## Bước 2 — Chỉnh sửa `CheckLicense.vbs`

Mở file và sửa dòng `RESULT_FOLDER` theo đường dẫn thực tế:

```vb
Const RESULT_FOLDER = "\\TÊNSERVER\share\Results"
' Hoặc dùng IP:
Const RESULT_FOLDER = "\\192.168.1.100\share\Results"
```

---

## Bước 3 — Triển khai qua GPO (khuyến nghị cho Domain)

> **Ưu điểm**: Tự động deploy tất cả máy trong domain, không cần chạy thủ công từng máy.

### 3.1 Tạo GPO Startup Script

1. Mở **Group Policy Management Console** (GPMC) trên Domain Controller
2. Tạo GPO mới: `LicenseCheck`
3. Link GPO vào OU chứa các máy Windows cần kiểm tra
4. Chỉnh sửa GPO:
   ```
   Computer Configuration
   └── Policies
       └── Windows Settings
           └── Scripts (Startup/Shutdown)
               └── Startup
   ```
5. Thêm script:
   - **Script Name**: `cscript.exe`
   - **Script Parameters**: `//nologo //b "\\SERVER\share\Scripts\CheckLicense.vbs"`

### 3.2 Hoặc dùng GPO PowerShell Script

1. Trong GPO chọn:
   ```
   Computer Configuration
   └── Policies
       └── Windows Settings
           └── Scripts (Startup/Shutdown)
               └── Startup → PowerShell Scripts tab
   ```
2. Thêm `Deploy-Task.ps1` (chạy 1 lần để cài Task Scheduler, sau đó task tự chạy lúc startup)

---

## Bước 4 — Triển khai thủ công (không có WinRM)

Nếu không muốn dùng GPO, chạy `Deploy-Task.ps1` thủ công trên từng máy:

```powershell
# Mở PowerShell với quyền Admin trên máy Windows, rồi chạy:
\\SERVER\share\Scripts\Deploy-Task.ps1

# Hoặc từ Fedora qua net rpc (nếu đã cài samba-client):
net rpc service status "Schedule" -U Administrator%Password -S PC-001
```

---

## Bước 5 (Tùy chọn) — Deploy song song từ Fedora qua WinRM

Nếu muốn enable WinRM để deploy nhanh cho 10–50 máy:

### 5.1 Enable WinRM trên Windows clients (qua GPO)

```
Computer Configuration
└── Policies
    └── Administrative Templates
        └── Windows Components
            └── Windows Remote Management (WinRM)
                └── WinRM Service → Allow remote server management → Enabled
```

### 5.2 Chạy từ Fedora

```bash
# Cài winrm client trên Fedora
pip3 install pywinrm requests-kerberos

# Hoặc dùng Ansible (khuyến nghị cho 50+ máy)
sudo dnf install ansible

# Tạo inventory
cat > inventory.ini << 'EOF'
[windows]
PC-001.domain.local
PC-002.domain.local
PC-003.domain.local

[windows:vars]
ansible_user=Administrator
ansible_password=YourPassword
ansible_connection=winrm
ansible_winrm_transport=kerberos
ansible_winrm_server_cert_validation=ignore
EOF

# Chạy script trên tất cả máy cùng lúc
ansible windows -i inventory.ini -m win_shell \
  -a 'cscript.exe //nologo //b "\\SERVER\share\Scripts\CheckLicense.vbs"'
```

---

## Bước 6 — Xem kết quả từ Fedora

```bash
# Liệt kê tất cả kết quả
ls -la /srv/samba/share/Results/

# Xem kết quả của một máy cụ thể
cat /srv/samba/share/Results/PC-001.txt

# Tổng hợp nhanh trạng thái tất cả máy
grep -h "Trang thai" /srv/samba/share/Results/*.txt | sort | uniq -c | sort -rn

# Tìm máy chưa kích hoạt
grep -l "CHUA KICH HOAT" /srv/samba/share/Results/*.txt

# Export tổng hợp ra CSV
echo "Ten may,Trang thai,Loai key,Partial Key" > /tmp/license_summary.csv
for f in /srv/samba/share/Results/*.txt; do
  pc=$(basename "$f" .txt)
  status=$(grep "Trang thai" "$f" | cut -d: -f2- | xargs)
  type=$(grep "Loai key" "$f" | cut -d: -f2- | xargs)
  key=$(grep "Partial Key" "$f" | cut -d: -f2- | xargs)
  echo "$pc,$status,$type,$key" >> /tmp/license_summary.csv
done
cat /tmp/license_summary.csv
```

---

## Thứ tự ưu tiên triển khai

| Phương án | Độ phức tạp | Phù hợp với |
|-----------|-------------|-------------|
| **GPO Startup Script** | ⭐ Thấp | Domain, không cần cấu hình thêm |
| **Task Scheduler thủ công** | ⭐⭐ Trung bình | Vài máy, không có GPO |
| **Ansible + WinRM** | ⭐⭐⭐ Cao | 50+ máy, cần tốc độ |

> [!TIP]
> Với domain environment, **GPO Startup Script** là cách đơn giản nhất — chỉ cần cấu hình 1 lần, tất cả máy trong OU tự áp dụng.

> [!IMPORTANT]
> Thư mục `Results/` trên Samba phải có quyền **write cho Domain Users** để các máy Windows ghi file kết quả vào được.
