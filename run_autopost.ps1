# =============================================================
# run_autopost.ps1
# Jalankan tiktok_autopost.py langsung dari PowerShell.
# Cara pakai:
#   .\run_autopost.ps1                      # langsung sekali jalan
#   .\run_autopost.ps1 -Jadwal "21:15"      # tunggu sampai jam 21:15 lalu jalan
#   .\run_autopost.ps1 -Loop -Interval 4    # loop tiap 4 jam
#   .\run_autopost.ps1 -Jadwal "21:15" -Loop -Interval 4  # jadwal + loop
# =============================================================

param(
    [string] $Jadwal   = "",   # waktu mulai, format "HH:mm"  contoh: "21:15"
    [switch] $Loop,            # ulangi terus setelah selesai
    [int]    $Interval = 4     # jeda antar run (jam) kalau -Loop aktif
)

# --- Sesuaikan path Python dan script di sini ---
$PythonExe  = "python"
$ScriptPath = Join-Path $PSScriptRoot "tiktok_autopost.py"
# ------------------------------------------------

function Run-Autopost {
    $waktu = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host ""
    Write-Host "[$waktu] === Mulai autopost ===" -ForegroundColor Cyan

    if (-not (Test-Path $ScriptPath)) {
        Write-Host "ERROR: $ScriptPath tidak ditemukan." -ForegroundColor Red
        return
    }

    & $PythonExe $ScriptPath

    if ($LASTEXITCODE -eq 0) {
        Write-Host "[$( Get-Date -Format 'HH:mm:ss' )] Selesai (sukses)" -ForegroundColor Green
    } else {
        Write-Host "[$( Get-Date -Format 'HH:mm:ss' )] Selesai (exit code $LASTEXITCODE)" -ForegroundColor Yellow
    }
}

function Wait-UntilJadwal {
    param([string] $WaktuTarget)

    $now    = Get-Date
    $target = [datetime]::ParseExact(
                  "$( $now.ToString('yyyy-MM-dd') ) $WaktuTarget",
                  "yyyy-MM-dd HH:mm", $null)

    # Kalau jam target sudah lewat hari ini, jadwalkan besok
    if ($target -le $now) {
        $target = $target.AddDays(1)
    }

    $selisih = ($target - (Get-Date)).TotalSeconds
    Write-Host "Menunggu jadwal jam $WaktuTarget (sekitar $( [math]::Round($selisih/60) ) menit lagi)..." -ForegroundColor DarkYellow

    while ((Get-Date) -lt $target) {
        $sisa = ($target - (Get-Date)).TotalSeconds
        Write-Host "`r  Mulai dalam $( [math]::Round($sisa) ) detik...  " -NoNewline -ForegroundColor DarkGray
        Start-Sleep -Seconds 10
    }
    Write-Host ""
}

# --- Jalankan ---
if ($Jadwal -ne "") {
    Wait-UntilJadwal -WaktuTarget $Jadwal
}

if ($Loop) {
    $intervalDetik = $Interval * 3600
    Write-Host "Mode LOOP aktif — interval $Interval jam. Tekan Ctrl+C untuk berhenti." -ForegroundColor Magenta
    while ($true) {
        Run-Autopost
        $berikutnya = (Get-Date).AddSeconds($intervalDetik).ToString("HH:mm:ss")
        Write-Host "Run berikutnya jam $berikutnya. Menunggu..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $intervalDetik
    }
} else {
    Run-Autopost
}
