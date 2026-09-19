# ================================================================
# SCRIPT CHUYEN DU LIEU ZALO VA WECHAT SANG O D
# Tao Junction Link tu C sang D
# ================================================================

function Stop-ChatApps {
    Write-Host "Dang tat cac ung dung Zalo va WeChat/Weixin..." -ForegroundColor Yellow
    $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Path -match 'Tencent|Weixin|Zalo|WeChat' -or
        $_.ProcessName -match 'Weixin|crashpad_handler|Zalo|WeChat|Tencent|xwechat'
    }
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
}

Stop-ChatApps

function Move-And-Junction($src, $dst, $label) {
    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Write-Host "--- CHUYEN [$label] ---" -ForegroundColor Cyan
    Write-Host "Nguon (C): $src"
    Write-Host "Dich  (D): $dst"

    Stop-ChatApps

    if (-not (Test-Path $src)) {
        Write-Host "Thu muc nguon $src khong ton tai." -ForegroundColor Yellow
        if (Test-Path $dst) {
            Write-Host "Tao Junction cho thu muc rong..." -ForegroundColor Yellow
            cmd /c mklink /J "$src" "$dst"
        }
        return
    }

    # Kiem tra neu src da la Junction roi
    $item = Get-Item $src -Force -ErrorAction SilentlyContinue
    if ($item.LinkType -eq "Junction") {
        Write-Host "SUCCESS: Thu muc $src DA LA JUNCTION sang $($item.Target)." -ForegroundColor Green
        return
    }

    if (-not (Test-Path $dst)) {
        New-Item -ItemType Directory -Path $dst -Force | Out-Null
    }

    Write-Host "Dang di chuyen du lieu tu C sang D..." -ForegroundColor Yellow
    robocopy "$src" "$dst" /E /MOVE /R:2 /W:1 /NJH /NJS /NDL /NC /NS /NP | Out-Null

    if (Test-Path $src) {
        Remove-Item -Path $src -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Dang tao Junction Link..." -ForegroundColor Yellow
    $cmdOutput = cmd /c mklink /J "$src" "$dst"
    Write-Host $cmdOutput

    $newItem = Get-Item $src -Force -ErrorAction SilentlyContinue
    if ($newItem.LinkType -eq "Junction") {
        Write-Host "SUCCESS: Da tao Junction thanh cong: $label -> $($newItem.Target)" -ForegroundColor Green
    } else {
        Write-Host "ERROR: Khong the tao Junction cho $label" -ForegroundColor Red
    }
}

Move-And-Junction "C:\Users\thuyn\AppData\Roaming\ZaloData" "D:\AppData\Roaming\ZaloData" "ZaloData (AppData)"
Move-And-Junction "C:\Users\thuyn\AppData\Roaming\Tencent"  "D:\AppData\Roaming\Tencent"  "Tencent / WeChat (AppData)"
Move-And-Junction "C:\Users\thuyn\Documents\WeChat Files"   "D:\Documents\WeChat Files"   "WeChat Files (Documents)"
Move-And-Junction "C:\Users\thuyn\Documents\xwechat_files"  "D:\Documents\xwechat_files"  "xwechat_files (Documents)"

Write-Host ""
Write-Host "=================================================" -ForegroundColor Green
Write-Host "HOAN TAT CHUYEN DONG BO ZALO VA WECHAT SANG O D!" -ForegroundColor Green
Write-Host "=================================================" -ForegroundColor Green
