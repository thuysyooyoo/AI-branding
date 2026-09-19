# ================================================================
# WEEKLY CLEANUP SCRIPT
# Chay vao luc 11:00 sang thu hai hang tuan (tu dong qua Task Scheduler)
# Roaming -> D:\AppData\Roaming (junction tu C)
# Log luu tai: $env:USERPROFILE\CleanupLogs\
# ================================================================

$logDir  = "$env:USERPROFILE\CleanupLogs"
$logFile = "$logDir\cleanup_$(Get-Date -Format 'yyyy-MM').log"
$totalFreed = 0

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Add-Content $logFile $line
    Write-Host $line
}

function Get-FolderSize($path) {
    if (-not (Test-Path $path)) { return 0 }
    (Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
}

function Clear-Folder($path, $label) {
    if (-not (Test-Path $path)) { Log "  SKIP: $label (khong ton tai)"; return 0 }
    $before = Get-FolderSize $path
    Get-ChildItem $path -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $freed = $before - (Get-FolderSize $path)
    Log "  OK [$label]: Giai phong $([math]::Round($freed/1MB,1)) MB"
    return $freed
}

Log "================================================"
Log "  BAT DAU DON DEP - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Log "================================================"

# ----------------------------------------------------------------
# [1] Kiem tra Junction Roaming (C -> D)
# ----------------------------------------------------------------
Log ""
Log "[1] KIEM TRA JUNCTION ROAMING..."
$r = Get-Item "$env:USERPROFILE\AppData\Roaming" -Force -ErrorAction SilentlyContinue
if ($r.LinkType -eq "Junction" -and $r.Target -eq "D:\AppData\Roaming") {
    Log "  OK: Junction hop le -> D:\AppData\Roaming"
} else {
    Log "  CANH BAO: Junction co van de hoac khong dung (LinkType=$($r.LinkType) | Target=$($r.Target))"
}

# ----------------------------------------------------------------
# [2] Temp Files
# ----------------------------------------------------------------
Log ""
Log "[2] XOA TEMP FILES..."
$totalFreed += Clear-Folder $env:TEMP             "User Temp (%TEMP%)"
$totalFreed += Clear-Folder "C:\Windows\Temp"     "Windows Temp"
$totalFreed += Clear-Folder "C:\Windows\Prefetch" "Prefetch"

# ----------------------------------------------------------------
# [3] Windows Update Cache
# ----------------------------------------------------------------
Log ""
Log "[3] WINDOWS UPDATE CACHE..."
Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
$totalFreed += Clear-Folder "C:\Windows\SoftwareDistribution\Download" "WU Download Cache"
Start-Service wuauserv -ErrorAction SilentlyContinue

# ----------------------------------------------------------------
# [4] File cai dat cu trong Downloads (>30 ngay)
# ----------------------------------------------------------------
Log ""
Log "[4] FILE CAI DAT CU TRONG DOWNLOADS (>30 ngay)..."
$cutoff   = (Get-Date).AddDays(-30)
$downloadsPath = "$env:USERPROFILE\Downloads"
$oldFiles = Get-ChildItem $downloadsPath -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $cutoff -and
                   $_.Extension -in @('.exe','.msi','.msp','.msix','.appx','.zip') }
$oldSize  = ($oldFiles | Measure-Object Length -Sum).Sum
$oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
Log "  OK: Xoa $($oldFiles.Count) file ($([math]::Round($oldSize/1MB,1)) MB)"
$totalFreed += $oldSize

# ----------------------------------------------------------------
# [5] Browser Cache
# ----------------------------------------------------------------
Log ""
Log "[5] BROWSER CACHE..."
$caches = @{
    "CocCoc" = "$env:APPDATA\CocCoc\Browser\User Data\Default\Cache"
    "Chrome" = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
    "Edge"   = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
}
foreach ($b in $caches.GetEnumerator()) {
    $totalFreed += Clear-Folder $b.Value "$($b.Key) Cache"
}

# ----------------------------------------------------------------
# [6] Recycle Bin
# ----------------------------------------------------------------
Log ""
Log "[6] RECYCLE BIN..."
Clear-RecycleBin -Force -ErrorAction SilentlyContinue
Log "  OK: Da don sach Recycle Bin"

# ----------------------------------------------------------------
# [7] Thumbnail Cache
# ----------------------------------------------------------------
Log ""
Log "[7] THUMBNAIL CACHE..."
$thumbs    = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" `
               -Filter "thumbcache_*.db" -File -ErrorAction SilentlyContinue
$thumbSize = ($thumbs | Measure-Object Length -Sum).Sum
$thumbs | Remove-Item -Force -ErrorAction SilentlyContinue
Log "  OK: $($thumbs.Count) files ($([math]::Round($thumbSize/1MB,1)) MB)"
$totalFreed += $thumbSize

# ----------------------------------------------------------------
# [8] Antigravity Cache (Conversations with no update > 2 days)
# ----------------------------------------------------------------
Log ""
Log "[8] ANTIGRAVITY CACHE..."
$antigravityPath = "$env:USERPROFILE\.gemini\antigravity"
$brainPath = Join-Path $antigravityPath "brain"
$convPath = Join-Path $antigravityPath "conversations"
$freedAntigravity = 0
$deletedCount = 0

if (Test-Path $antigravityPath) {
    # Get all GUIDs from brain and conversations directories
    $guids = @()
    if (Test-Path $brainPath) {
        $guids += Get-ChildItem $brainPath -Directory | Where-Object { $_.Name -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' } | Select-Object -ExpandProperty Name
    }
    if (Test-Path $convPath) {
        $guids += Get-ChildItem $convPath -File | Where-Object { $_.BaseName -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$' } | Select-Object -ExpandProperty BaseName
    }
    $guids = $guids | Select-Object -Unique
    
    $cutoff = (Get-Date).AddDays(-2)
    
    foreach ($guid in $guids) {
        $files = @()
        $bDir = Join-Path $brainPath $guid
        if (Test-Path $bDir) {
            $files += Get-Item $bDir
            $files += Get-ChildItem $bDir -Recurse -ErrorAction SilentlyContinue
        }
        $cFiles = Get-ChildItem $convPath -Filter "$guid*" -ErrorAction SilentlyContinue
        if ($cFiles) {
            $files += $cFiles
        }
        
        if ($files) {
            $lastWrite = ($files | Measure-Object -Property LastWriteTime -Maximum).Maximum
            if ($lastWrite -lt $cutoff) {
                # Calculate size to be freed
                $size = 0
                foreach ($f in $files) {
                    if (-not $f.PSIsContainer) {
                        $size += $f.Length
                    }
                }
                
                # Delete files and directories
                if (Test-Path $bDir) {
                    Remove-Item -Path $bDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                foreach ($cf in $cFiles) {
                    Remove-Item -Path $cf.FullName -Force -ErrorAction SilentlyContinue
                }
                
                $freedAntigravity += $size
                $deletedCount++
                Log "  Xoa cuoc tro chuyen: $guid (Cap nhat lan cuoi: $($lastWrite.ToString('yyyy-MM-dd HH:mm:ss')) | Giai phong $([math]::Round($size/1MB,2)) MB)"
            }
        }
    }
}
Log "  OK: Xoa $deletedCount cuoc tro chuyen cu ($([math]::Round($freedAntigravity/1MB,1)) MB)"
$totalFreed += $freedAntigravity

# ----------------------------------------------------------------
# [9] Ket qua
# ----------------------------------------------------------------
Log ""
Log "================================================"
Log "  TONG GIAI PHONG: $([math]::Round($totalFreed/1GB,3)) GB ($([math]::Round($totalFreed/1MB,1)) MB)"
Log ""
Get-PSDrive C,D -ErrorAction SilentlyContinue | ForEach-Object {
    Log "  O $($_.Name): Da dung $([math]::Round($_.Used/1GB,2)) GB | Con trong $([math]::Round($_.Free/1GB,2)) GB"
}
Log ""
Log "  HOAN TAT: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log "================================================"
