<#
.SYNOPSIS
    Kịch bản tự động đồng bộ các file báo cáo Excel sang thư mục OneDrive Eureka
    Thời gian chạy khuyến nghị: Hàng ngày lúc 11:00 AM
#>

$SourceDir = "D:\AI AGENT THUY\AI branding\reports"
$TargetDir = "C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh"
$LogFile   = "D:\AI AGENT THUY\AI branding\reports\sync_log.txt"

$FilesToSync = @(
    "Danh_sach_Van_ban_Phap_luat_XNK.xlsx",
    "SEO_Keyword_Trends_Eureka.xlsx"
)

function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogLine = "[$Timestamp] $Message"
    Write-Host $LogLine
    Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue
}

Write-Log "=== BẮT ĐẦU ĐỒNG BỘ REPORT SANG ONEDRIVE ==="

if (-not (Test-Path -Path $TargetDir)) {
    Write-Log "[ERROR] Thư mục đích OneDrive không tồn tại: $TargetDir"
    exit 1
}

$SuccessCount = 0

foreach ($FileName in $FilesToSync) {
    $SourcePath = Join-Path -Path $SourceDir -ChildPath $FileName
    $TargetPath = Join-Path -Path $TargetDir -ChildPath $FileName

    if (Test-Path -Path $SourcePath) {
        try {
            Copy-Item -Path $SourcePath -Destination $TargetPath -Force -ErrorAction Stop
            $SourceHash = (Get-FileHash -Path $SourcePath -Algorithm SHA256).Hash
            $TargetHash = (Get-FileHash -Path $TargetPath -Algorithm SHA256).Hash
            $FileInfo = Get-Item $TargetPath
            if ($SourceHash -eq $TargetHash) {
                attrib +p "$TargetPath" | Out-Null
                Write-Log "[SUCCESS] Đã đồng bộ '$FileName' ($($FileInfo.Length) bytes, SHA256: $SourceHash) -> OneDrive"
                $SuccessCount++
            } else {
                Write-Log "[ERROR] File '$FileName' copy xong nhưng hash không khớp!"
            }
        }
        catch {
            Write-Log "[ERROR] Lỗi khi đồng bộ '$FileName': $_"
        }
    }
    else {
        Write-Log "[WARNING] Không tìm thấy file nguồn: $SourcePath"
    }
}

Write-Log "=== HOÀN THÀNH: $SuccessCount/$($FilesToSync.Count) file đã đồng bộ thành công ==="
