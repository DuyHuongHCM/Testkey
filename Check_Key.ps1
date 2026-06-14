<#
.SYNOPSIS
    Chương trình kiểm tra chi tiết thông tin phần cứng máy tính và phân tích tính hợp lệ 
    (Chính hãng vs. Bẻ khóa/Crack) của Windows và Microsoft Office.
.DESCRIPTION
    Script thực hiện quét sâu thông tin hệ thống, kiểm tra Registry, tệp tin hệ thống,
    phân tích phương thức kích hoạt hệ điều hành và Office để đưa ra kết luận chi tiết.
.AUTHOR
    Thư ký AI của bạn
#>

# 1. TỰ ĐỘNG KIỂM TRA VÀ YÊU CẦU QUYỀN ADMINISTRATOR CHUẨN XÁC
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Trường hợp người dùng copy-paste trực tiếp toàn bộ code vào cửa sổ PowerShell thường thay vì chạy file .ps1
    if ([string]::IsNullOrEmpty($PSCommandPath)) {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        Write-Host "==========================================================" -ForegroundColor Red
        Write-Host " LỖI: BẠN ĐANG DÁN TRỰC TIẾP CODE VÀO CỬA SỔ POWERSHELL THƯỜNG!" -ForegroundColor Yellow
        Write-Host " Vui lòng làm theo một trong hai cách sau:" -ForegroundColor White
        Write-Host " Cách 1: Tìm 'PowerShell', chuột phải chọn 'Run as Administrator', rồi dán lại code." -ForegroundColor Green
        Write-Host " Cách 2: Lưu đoạn code này thành file '.ps1' (VD: Check.ps1) rồi chạy file đó." -ForegroundColor Green
        Write-Host "==========================================================" -ForegroundColor Red
        Read-Host "Nhấn Enter để thoát..."
        Exit
    }
    
    Write-Host "==========================================================" -ForegroundColor Yellow
    Write-Host " ĐANG TỰ ĐỘNG YÊU CẦU CẤP QUYỀN ADMINISTRATOR (QUYỀN QUẢN TRỊ)" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Yellow
    
    try {
        # Sử dụng danh sách tham số dạng mảng để xử lý hoàn hảo đường dẫn chứa khoảng trắng
        Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath -Verb RunAs
    } catch {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        Write-Host "Không thể tự động cấp quyền Administrator: $_" -ForegroundColor Red
        Write-Host "Vui lòng click chuột phải vào file script và chọn 'Run with PowerShell' hoặc chạy PowerShell dưới quyền Admin trước." -ForegroundColor Yellow
        Read-Host "Nhấn Enter để thoát..."
    }
    Exit
}

# Khởi tạo bảng mã UTF-8 để hiển thị tiếng Việt không bị lỗi font
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "    HỆ THỐNG KIỂM TRA BẢN QUYỀN WINDOWS & OFFICE" -ForegroundColor Green
Write-Host "                 Đang quét dữ liệu... " -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Green

# 2. THU THẬP THÔNG TIN PHẦN CỨNG (PC/LAPTOP INFO)
Write-Host "[1/4] Đang thu thập thông tin thiết bị..." -ForegroundColor Cyan
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $computerSystem = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_Bios
    $processor = Get-CimInstance Win32_Processor
    $motherboard = Get-CimInstance Win32_BaseBoard
    $ramTotal = [Math]::Round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 2)
    $disks = Get-CimInstance Win32_DiskDrive | ForEach-Object {
        $sizeGB = [Math]::Round($_.Size / 1GB, 2)
        "$($_.Model) ($sizeGB GB)"
    } -join ", "

    $pcInfo = @{
        "DeviceName"    = $computerSystem.Name
        "Manufacturer"  = $computerSystem.Manufacturer
        "Model"         = $computerSystem.Model
        "SerialNumber"  = $bios.SerialNumber
        "CPU"           = $processor.Name
        "RAM"           = "$ramTotal GB"
        "Storage"       = $disks
        "Motherboard"   = "$($motherboard.Manufacturer) $($motherboard.Product)"
        "OSName"        = $os.Caption
        "OSVersion"     = $os.Version
        "OSBuild"       = $os.BuildNumber
        "DomainJoined"  = $computerSystem.PartOfDomain
    }
} catch {
    Write-Host "Lỗi khi thu thập thông tin phần cứng: $_" -ForegroundColor Red
}

# 3. KIỂM TRA BẢN QUYỀN WINDOWS
Write-Host "[2/4] Đang kiểm tra trạng thái bản quyền Windows..." -ForegroundColor Cyan

$winLicenseStatus = "Không xác định"
$winChannel = "Không xác định"
$winKmsServer = ""
$winIsGenuine = "Không rõ"
$winDetailReason = ""

try {
    # Truy vấn thông tin kích hoạt Windows từ SoftwareLicensingProduct
    # ApplicationID của Windows: 55c92734-d682-4d71-983e-d6ec3f16059f
    $winProducts = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey != null"
    
    if ($winProducts) {
        $winProduct = $winProducts[0]
        
        # Trạng thái License
        # 1 = Licensed (Đã kích hoạt hợp lệ)
        switch ($winProduct.LicenseStatus) {
            0 { $winLicenseStatus = "Chưa kích hoạt (Unlicensed)" }
            1 { $winLicenseStatus = "Đã kích hoạt (Licensed)" }
            2 { $winLicenseStatus = "Thời gian chờ gia hạn (OOB Grace)" }
            3 { $winLicenseStatus = "Thời gian chờ kích hoạt lại (OOT Grace)" }
            4 { $winLicenseStatus = "Không chính hãng tạm thời (Non-Genuine Grace)" }
            5 { $winLicenseStatus = "Thông báo hết hạn (Notification)" }
            default { $winLicenseStatus = "Khác" }
        }

        # Kênh bản quyền
        $description = $winProduct.Description
        if ($description -like "*RETAIL*") { $winChannel = "Retail (Bán lẻ cá nhân)" }
        elseif ($description -like "*OEM*") { $winChannel = "OEM (Nhà sản xuất máy tính cài sẵn)" }
        elseif ($description -like "*VOLUME_KMS*") { $winChannel = "Volume:KMS (Máy chủ kích hoạt doanh nghiệp)" }
        elseif ($description -like "*VOLUME_MAK*") { $winChannel = "Volume:MAK (Khóa kích hoạt nhiều máy)" }
        else { $winChannel = "Volume (Doanh nghiệp)" }

        # KMS Server sử dụng (nếu có)
        $winKmsServer = $winProduct.KeyManagementServiceMachine

        # Đánh giá tính chính hãng của Windows
        if ($winProduct.LicenseStatus -eq 1) {
            if ($winChannel -like "*Retail*" -or $winChannel -like "*OEM*" -or $winChannel -like "*MAK*") {
                $winIsGenuine = "CHÍNH HÃNG THỰC SỰ"
                $winDetailReason = "Thiết bị sử dụng giấy phép vĩnh viễn chính thức (Retail/OEM). Không có dấu hiệu can thiệp của phần mềm crack."
            } elseif ($winChannel -like "*KMS*") {
                # Kiểm tra xem máy có thuộc Domain doanh nghiệp không
                if ($pcInfo.DomainJoined -eq $true) {
                    $winIsGenuine = "HỢP LỆ (DOANH NGHIỆP)"
                    $winDetailReason = "Hệ thống được kích hoạt qua máy chủ KMS nội bộ của doanh nghiệp ($winKmsServer) khi tham gia vào mạng nội bộ."
                } else {
                    # KMS trên máy cá nhân không thuộc Domain thường là bẻ khóa lậu (KMSAuto, KMSpico, MAS)
                    $winIsGenuine = "CẢNH BÁO: CÓ THỂ LÀ BẺ KHÓA LẬU (CRACK/KMS BYPASS)"
                    $winDetailReason = "Máy tính cá nhân không thuộc mạng doanh nghiệp nhưng lại kích hoạt bằng giao thức KMS thông qua máy chủ lậu hoặc máy chủ giả lập tại chỗ ($winKmsServer)."
                }
            }
        } else {
            $winIsGenuine = "CHƯA KÍCH HOẠT"
            $winDetailReason = "Windows hiện tại chưa được kích hoạt bản quyền."
        }
    } else {
        $winIsGenuine = "KHÔNG TÌM THẤY THÔNG TIN"
        $winDetailReason = "Không tìm thấy thông tin sản phẩm Windows hợp lệ trong hệ thống."
    }
} catch {
    $winIsGenuine = "LỖI KIỂM TRA"
    $winDetailReason = "Gặp lỗi trong quá trình phân tích: $_"
}

# 4. KIỂM TRA BẢN QUYỀN OFFICE
Write-Host "[3/4] Đang kiểm tra trạng thái bản quyền Microsoft Office..." -ForegroundColor Cyan

$officeStatusList = @()
$officeIsGenuine = "Không cài đặt"
$officeDetailReason = "Không phát hiện thấy bản cài đặt Microsoft Office cổ điển (Click-to-Run) hoặc Office chưa được kích hoạt."

# Thư mục chứa tệp quản lý bản quyền Office (ospp.vbs) thông dụng
$officePaths = @(
    "C:\Program Files\Microsoft Office\Office16",
    "C:\Program Files (x86)\Microsoft Office\Office16",
    "C:\Program Files\Microsoft Office\Office15",
    "C:\Program Files (x86)\Microsoft Office\Office15"
)

$osppPath = ""
foreach ($path in $officePaths) {
    if (Test-Path "$path\ospp.vbs") {
        $osppPath = "$path\ospp.vbs"
        break
    }
}

# Kiểm tra công nghệ bẻ khóa "Ohook" phổ biến hiện nay
$ohookDetected = $false
$ohookPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\sppc.dll",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\sppc.dll"
)
foreach ($oPath in $ohookPaths) {
    if (Test-Path $oPath) {
        $ohookDetected = $true
        break
    }
}

if ($osppPath) {
    try {
        # Chạy file quản lý bản quyền Office để lấy trạng thái chi tiết
        $osppOutput = cscript.exe //NoLogo "$osppPath" /dstatus
        
        $currentOffice = @{}
        $licenseName = ""
        $licenseStatus = ""
        $kmsHost = ""

        foreach ($line in $osppOutput) {
            if ($line -like "*LICENSE NAME:*") {
                $licenseName = ($line -split "LICENSE NAME:")[1].Trim()
            }
            if ($line -like "*LICENSE STATUS:*") {
                $licenseStatus = ($line -split "LICENSE STATUS:")[1].Trim()
            }
            if ($line -like "*KMS machine name:*") {
                $kmsHost = ($line -split "KMS machine name:")[1].Trim()
            }

            # Gom nhóm thông tin khi kết thúc bản ghi một sản phẩm
            if ($line -like "*---*") {
                if ($licenseName) {
                    $officeStatusList += [PSCustomObject]@{
                        Name   = $licenseName
                        Status = $licenseStatus
                        KMS    = $kmsHost
                    }
                    # Reset biến tạm
                    $licenseName = ""
                    $licenseStatus = ""
                    $kmsHost = ""
                }
            }
        }
        # Thêm sản phẩm cuối cùng nếu có
        if ($licenseName) {
            $officeStatusList += [PSCustomObject]@{
                Name   = $licenseName
                Status = $licenseStatus
                KMS    = $kmsHost
            }
        }

        # Đánh giá tính chính hãng của Office
        if ($officeStatusList.Count -gt 0) {
            $licensedProducts = $officeStatusList | Where-Object { $_.Status -like "*LICENSED*" }
            
            if ($ohookDetected) {
                $officeIsGenuine = "PHÁT HIỆN BẺ KHÓA LẬU (CRACK BẰNG OHOOK BYPASS)"
                $officeDetailReason = "Phát hiện tệp tin 'sppc.dll' lạ được cài cắm trực tiếp trong thư mục Microsoft Office. Đây là dấu hiệu rõ ràng của phương pháp bẻ khóa Ohook bypass tinh vi nhằm lừa gạt hệ thống bản quyền."
            } elseif ($licensedProducts) {
                # Kiểm tra KMS lậu
                $kmsActivated = $licensedProducts | Where-Object { $_.Name -like "*KMS*" -or $_.KMS -ne "" }
                if ($kmsActivated) {
                    if ($pcInfo.DomainJoined -eq $true) {
                        $officeIsGenuine = "HỢP LỆ (DOANH NGHIỆP)"
                        $officeDetailReason = "Office được kích hoạt thông qua hạ tầng KMS nội bộ thuộc mạng lưới doanh nghiệp của bạn."
                    } else {
                        $officeIsGenuine = "CẢNH BÁO: BẺ KHÓA LẬU (CRACK BẰNG KMS)"
                        $officeDetailReason = "Phát hiện Office được đăng ký kích hoạt dưới dạng Volume cấp phép lớn (KMS) nhưng máy tính cá nhân không thuộc quản lý của bất kỳ Active Directory Domain doanh nghiệp nào. Khả năng cao đây là crack."
                    }
                } else {
                    $officeIsGenuine = "CHÍNH HÃNG THỰC SỰ"
                    $officeDetailReason = "Microsoft Office được kích hoạt hợp lệ bằng giấy phép bán lẻ vĩnh viễn (Retail/OEM/Subscription) hoặc liên kết trực tiếp với tài khoản Microsoft 365 chính chủ."
                }
            } else {
                $officeIsGenuine = "CHƯA KÍCH HOẠT"
                $officeDetailReason = "Đã tìm thấy Office trên máy nhưng trạng thái hiện tại là dùng thử hoặc chưa được kích hoạt bản quyền."
            }
        }
    } catch {
        $officeIsGenuine = "LỖI KIỂM TRA"
        $officeDetailReason = "Gặp lỗi khi phân tích giấy phép Office thông qua hệ thống script: $_"
    }
} else {
    # Kiểm tra thêm phiên bản Office dạng Store App (UWP) hoặc Microsoft 365 hiện đại nếu không có OSPP.vbs
    $uwpOffice = Get-AppxPackage -Name "*Microsoft.Office.Desktop*"
    if ($uwpOffice) {
        $officeIsGenuine = "CẦN KIỂM TRA TRONG APP (MICROSOFT 365)"
        $officeDetailReason = "Hệ thống phát hiện phiên bản Microsoft Office Store App (UWP) được cài đặt. Đối với phiên bản này, vui lòng mở trực tiếp ứng dụng Word -> File -> Account để xem trạng thái kích hoạt tài khoản của bạn."
    }
}

# 5. QUÉT CÁC DẤU VẾT HACK/CRACK TRÊN HỆ THỐNG
Write-Host "[4/4] Đang quét hệ thống tìm công cụ kích hoạt lậu..." -ForegroundColor Cyan

$detectedCracks = @()
$commonCrackPaths = @(
    "C:\Windows\KMS",
    "C:\Program Files\KMSpico",
    "C:\Program Files (x86)\KMSpico",
    "C:\Windows\SECOH-QAD.exe",
    "C:\Windows\SECOH-QAD.dll",
    "C:\Windows\KMSeldi.exe",
    "C:\Windows\Setup\Scripts\MAS.cmd"
)

foreach ($path in $commonCrackPaths) {
    if (Test-Path $path) {
        $detectedCracks += [PSCustomObject]@{
            "Path" = $path
            "Type" = "Phát hiện tệp/thư mục của công cụ Crack (KMSpico/KMSAuto/MAS)"
        }
    }
}

# Kiểm tra các máy chủ KMS lậu phổ biến trong Registry
$kmsRegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
)
$suspiciousKmsServers = @("kms.digiboy.ir", "kms.lotro.cc", "kms8.msguides.com", "msguides.com", "127.0.0.1", "localhost", "::1")

foreach ($rPath in $kmsRegPaths) {
    if (Test-Path $rPath) {
        $val = Get-ItemProperty -Path $rPath -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
        if ($val -and $val.KeyManagementServiceMachine) {
            $kmsMachine = $val.KeyManagementServiceMachine
            foreach ($sus in $suspiciousKmsServers) {
                if ($kmsMachine -like "*$sus*") {
                    $detectedCracks += [PSCustomObject]@{
                        "Path" = "$rPath \ KeyManagementServiceMachine"
                        "Type" = "Trỏ đến máy chủ KMS lậu hoặc nội bộ ảo: $kmsMachine"
                    }
                    break
                }
            }
        }
    }
}

# 6. HIỂN THỊ KẾT QUẢ RA MÀN HÌNH CONSOLE (COLORFUL PRINT)
Clear-Host
Write-Host "==========================================================================" -ForegroundColor Green
Write-Host "                KẾT QUẢ PHÂN TÍCH HỆ THỐNG BẢN QUYỀN" -ForegroundColor Green
Write-Host "==========================================================================" -ForegroundColor Green

Write-Host "`n[+] THÔNG TIN PHẦN CỨNG MÁY TÍNH:" -ForegroundColor Blue
$pcInfo.Keys | ForEach-Object {
    Write-Host "   - $_ : " -NoNewline -ForegroundColor Gray
    Write-Host "$($pcInfo[$_])" -ForegroundColor White
}

Write-Host "`n[+] TRẠNG THÁI BẢN QUYỀN WINDOWS:" -ForegroundColor Blue
Write-Host "   - Phiên bản: " -NoNewline -ForegroundColor Gray; Write-Host $pcInfo.OSName -ForegroundColor White
Write-Host "   - Kênh phân phối: " -NoNewline -ForegroundColor Gray; Write-Host $winChannel -ForegroundColor White
Write-Host "   - Trạng thái giấy phép: " -NoNewline -ForegroundColor Gray; Write-Host $winLicenseStatus -ForegroundColor White
Write-Host "   - Máy chủ KMS: " -NoNewline -ForegroundColor Gray; Write-Host (If ($winKmsServer) { $winKmsServer } else { "Không có (Sử dụng mã vĩnh viễn)" }) -ForegroundColor White
Write-Host "   - Đánh giá: " -NoNewline -ForegroundColor Gray
if ($winIsGenuine -eq "CHÍNH HÃNG THỰC SỰ" -or $winIsGenuine -eq "HỢP LỆ (DOANH NGHIỆP)") {
    Write-Host $winIsGenuine -ForegroundColor Green
} else {
    Write-Host $winIsGenuine -ForegroundColor Red
}
Write-Host "   - Chi tiết: " -NoNewline -ForegroundColor Gray; Write-Host $winDetailReason -ForegroundColor Yellow

Write-Host "`n[+] TRẠNG THÁI BẢN QUYỀN MICROSOFT OFFICE:" -ForegroundColor Blue
Write-Host "   - Đánh giá: " -NoNewline -ForegroundColor Gray
if ($officeIsGenuine -eq "CHÍNH HÃNG THỰC SỰ" -or $officeIsGenuine -eq "HỢP LỆ (DOANH NGHIỆP)") {
    Write-Host $officeIsGenuine -ForegroundColor Green
} elseif ($officeIsGenuine -eq "Không cài đặt") {
    Write-Host $officeIsGenuine -ForegroundColor Gray
} else {
    Write-Host $officeIsGenuine -ForegroundColor Red
}
Write-Host "   - Chi tiết: " -NoNewline -ForegroundColor Gray; Write-Host $officeDetailReason -ForegroundColor Yellow

if ($officeStatusList.Count -gt 0) {
    Write-Host "   - Chi tiết các gói Office phát hiện:" -ForegroundColor Gray
    foreach ($o in $officeStatusList) {
        Write-Host "     * $($o.Name) -> Trạng thái: " -NoNewline -ForegroundColor DarkGray
        if ($o.Status -like "*LICENSED*") {
            Write-Host "LICENSED (Đã kích hoạt)" -ForegroundColor Green
        } else {
            Write-Host $o.Status -ForegroundColor Red
        }
    }
}

if ($detectedCracks.Count -gt 0) {
    Write-Host "`n[!] CẢNH BÁO MỐI ĐE DỌA PHÁT HIỆN TRÊN MÁY:" -ForegroundColor Red
    foreach ($c in $detectedCracks) {
        Write-Host "   - Loại: $($c.Type)" -ForegroundColor Yellow
        Write-Host "     Vị trí/Thông số: $($c.Path)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "`n[+] KHÔNG PHÁT HIỆN TIẾN TRÌNH HOẶC TỆP TIN BẺ KHÓA LẠ NẰM TRONG HỆ THỐNG" -ForegroundColor Green
}

# 7. XUẤT BÁO CÁO HTML ĐẸP MẮT RA DESKTOP
$reportPath = "$env:USERPROFILE\Desktop\BaoCao_BanQuyen.html"

# CSS & HTML Template
$htmlContent = @"
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Báo Cáo Tình Trạng Bản Quyền Hệ Thống</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f3f4f6; color: #1f2937; margin: 0; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h1 { text-align: center; color: #1e3a8a; margin-bottom: 5px; }
        .subtitle { text-align: center; color: #6b7280; font-size: 0.95rem; margin-bottom: 30px; }
        .section { margin-bottom: 25px; border-bottom: 1px solid #e5e7eb; padding-bottom: 20px; }
        .section:last-child { border-bottom: none; }
        .section-title { font-size: 1.2rem; color: #1e3a8a; font-weight: 600; margin-bottom: 15px; display: flex; align-items: center; border-left: 4px solid #3b82f6; padding-left: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        td, th { padding: 10px 12px; text-align: left; font-size: 0.95rem; }
        tr:nth-child(even) { background-color: #f9fafb; }
        td.label { font-weight: 600; color: #4b5563; width: 30%; }
        .badge { display: inline-block; padding: 6px 12px; border-radius: 50px; font-weight: bold; font-size: 0.85rem; }
        .badge-success { background-color: #d1fae5; color: #065f46; }
        .badge-danger { background-color: #fee2e2; color: #991b1b; }
        .badge-warning { background-color: #fef3c7; color: #92400e; }
        .badge-gray { background-color: #f3f4f6; color: #374151; }
        .warning-box { background-color: #fffbeb; border: 1px solid #fef3c7; border-left: 4px solid #d97706; padding: 15px; border-radius: 6px; margin-top: 10px; }
        .warning-box-title { font-weight: bold; color: #b45309; margin-bottom: 5px; }
        .footer { text-align: center; margin-top: 30px; font-size: 0.85rem; color: #9ca3af; }
    </style>
</head>
<body>
    <div class="container">
        <h1>BÁO CÁO BẢN QUYỀN HỆ THỐNG</h1>
        <div class="subtitle">Báo cáo được khởi tạo tự động bởi Thư ký AI của bạn vào lúc $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</div>

        <!-- PHẦN 1: THÔNG TIN THIẾT BỊ -->
        <div class="section">
            <div class="section-title">1. Thông Tin Phần Cứng Máy Tính</div>
            <table>
                <tr><td class="label">Tên thiết bị</td><td>$($pcInfo.DeviceName)</td></tr>
                <tr><td class="label">Hãng sản xuất & Model</td><td>$($pcInfo.Manufacturer) - $($pcInfo.Model)</td></tr>
                <tr><td class="label">Mã Serial Number</td><td>$($pcInfo.SerialNumber)</td></tr>
                <tr><td class="label">Bộ vi xử lý (CPU)</td><td>$($pcInfo.CPU)</td></tr>
                <tr><td class="label">Bộ nhớ trong (RAM)</td><td>$($pcInfo.RAM)</td></tr>
                <tr><td class="label">Ổ cứng lưu trữ</td><td>$($pcInfo.Storage)</td></tr>
                <tr><td class="label">Bo mạch chủ (Motherboard)</td><td>$($pcInfo.Motherboard)</td></tr>
            </table>
        </div>

        <!-- PHẦN 2: BẢN QUYỀN WINDOWS -->
        <div class="section">
            <div class="section-title">2. Phân Tích Bản Quyền Windows</div>
            <table>
                <tr><td class="label">Hệ điều hành</td><td>$($pcInfo.OSName)</td></tr>
                <tr><td class="label">Phiên bản Build</td><td>$($pcInfo.OSVersion) (Build $($pcInfo.OSBuild))</td></tr>
                <tr><td class="label">Kênh phân phối</td><td>$winChannel</td></tr>
                <tr><td class="label">Trạng thái kích hoạt</td><td>$winLicenseStatus</td></tr>
                <tr>
                    <td class="label">Đánh giá độ chính hãng</td>
                    <td>
                        $(if ($winIsGenuine -eq "CHÍNH HÃNG THỰC SỰ" -or $winIsGenuine -eq "HỢP LỆ (DOANH NGHIỆP)") {
                            "<span class='badge badge-success'>CHÍNH HÃNG THỰC SỰ</span>"
                        } else {
                            "<span class='badge badge-danger'>CÓ DẤU HIỆU BẺ KHÓA / CRACK</span>"
                        })
                    </td>
                </tr>
                <tr><td class="label">Phân tích kỹ thuật</td><td style="color: #6b7280; font-style: italic;">$winDetailReason</td></tr>
            </table>
        </div>

        <!-- PHẦN 3: BẢN QUYỀN MICROSOFT OFFICE -->
        <div class="section">
            <div class="section-title">3. Phân Tích Bản Quyền Microsoft Office</div>
            <table>
                <tr>
                    <td class="label">Đánh giá bản quyền</td>
                    <td>
                        $(if ($officeIsGenuine -eq "CHÍNH HÃNG THỰC SỰ" -or $officeIsGenuine -eq "HỢP LỆ (DOANH NGHIỆP)") {
                            "<span class='badge badge-success'>CHÍNH HÃNG THỰC SỰ</span>"
                        } elseif ($officeIsGenuine -eq "Không cài đặt") {
                            "<span class='badge badge-gray'>CHƯA CÀI ĐẶT HOẶC CẦN KIỂM TRA TỰ TAY</span>"
                        } else {
                            "<span class='badge badge-danger'>CÓ DẤU HIỆU BẺ KHÓA / CRACK</span>"
                        })
                    </td>
                </tr>
                <tr><td class="label">Phân tích kỹ thuật</td><td style="color: #6b7280; font-style: italic;">$officeDetailReason</td></tr>
            </table>

            $(if ($officeStatusList.Count -gt 0) {
                $subTable = "<h4>Các ứng dụng Office phát hiện được:</h4><table>"
                foreach ($o in $officeStatusList) {
                    $statusBadge = if ($o.Status -like "*LICENSED*") { "<span class='badge badge-success'>LICENSED</span>" } else { "<span class='badge badge-danger'>$($o.Status)</span>" }
                    $subTable += "<tr><td class='label'>$($o.Name)</td><td>$statusBadge</td></tr>"
                }
                $subTable += "</table>"
                $subTable
            })
        </div>

        <!-- PHẦN 4: DẤU VẾT CRACK ĐỘC HẠI -->
        <div class="section">
            <div class="section-title">4. Quét Dấu Vết Phần Mềm Crack Hệ Thống</div>
            $(if ($detectedCracks.Count -gt 0) {
                $warningHtml = "<div class='warning-box'><div class='warning-box-title'>CẢNH BÁO: PHÁT HIỆN CÔNG CỤ HOẶC KHÓA REGISTRY CRACK LẬU</div><ul>"
                foreach ($c in $detectedCracks) {
                    $warningHtml += "<li><strong>$($c.Type)</strong> tại: <code>$($c.Path)</code></li>"
                }
                $warningHtml += "</ul><p style='font-size: 0.85rem; color: #b45309; margin-top: 10px;'><em>Lưu ý: Các công cụ crack (như KMSpico, KMSAuto, MAS...) thường can thiệp sâu vào nhân hệ điều hành, tắt Windows Defender, mở cổng mạng và có nguy cơ cao đính kèm virus, mã độc tống tiền (Ransomware), Trojan đánh cắp mật khẩu ngân hàng.</em></p></div>"
                $warningHtml
            } else {
                "<div style='background-color: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; padding: 15px; border-radius: 6px;'>Hệ thống sạch! Không phát hiện thấy tệp bẻ khóa lậu hoặc cấu hình máy chủ KMS bất hợp pháp lưu trữ cục bộ.</div>"
            })
        </div>

        <div class="footer">
            Báo cáo được thực hiện trên thiết bị: $($pcInfo.DeviceName) | Kỹ thuật viên: AI Secretary <br>
            Cám ơn anh/chị đã sử dụng dịch vụ của em!
        </div>
    </div>
</body>
</html>
"@

try {
    $htmlContent | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "`n==========================================================================" -ForegroundColor Green
    Write-Host "[THÀNH CÔNG] Đã xuất báo cáo chi tiết ra màn hình Desktop!" -ForegroundColor Cyan
    Write-Host "Vị trí file: $reportPath" -ForegroundColor Yellow
    Write-Host "==========================================================================" -ForegroundColor Green
} catch {
    Write-Host "Không thể ghi file báo cáo ra Desktop: $_" -ForegroundColor Red
}

# 8. TỰ ĐỘNG MỞ FILE BÁO CÁO VÀ CHỜ NGƯỜI DÙNG ĐỌC KẾT QUẢ TRÊN CONSOLE
try {
    Start-Process -FilePath $reportPath
} catch {
    Write-Host "Không thể tự động mở báo cáo bằng trình duyệt: $_" -ForegroundColor Yellow
    Write-Host "Anh/chị vui lòng tự mở file 'BaoCao_BanQuyen.html' thủ công trên Desktop nhé." -ForegroundColor White
}

Write-Host "`n[HOÀN THÀNH] Toàn bộ quy trình quét đã hoàn tất." -ForegroundColor Green
Write-Host "Màn hình sẽ không tự tắt để anh/chị xem chi tiết kết quả quét bên trên." -ForegroundColor Gray
Read-Host "Nhấn phím Enter để đóng cửa sổ này..."