# =============================================================
# run_autopost.ps1
# Jalankan tiktok_autopost.py langsung dari PowerShell.
# Cara pakai:
#   .\run_autopost.ps1              # langsung sekali jalan
#   .\run_autopost.ps1 -Loop        # jalan terus tiap JAM_INTERVAL jam
#   .\run_autopost.ps1 -Loop -Jam 3 # jalan terus tiap 3 jam
# =============================================================

param(
    [switch] $Loop,           # aktifkan mode loop
    [int]    $Jam = 4         # interval antar run (jam), default 4 jam
)

# --- Sesuaikan path Python dan script di sini ---
$PythonExe  = "python"                        # ganti ke "python3" kalau perlu
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

# --- Jalankan ---
if ($Loop) {
    $intervalDetik = $Jam * 3600
    Write-Host "Mode LOOP aktif — interval $Jam jam. Tekan Ctrl+C untuk berhenti." -ForegroundColor Magenta
    while ($true) {
        Run-Autopost
        $berikutnya = (Get-Date).AddSeconds($intervalDetik).ToString("HH:mm:ss")
        Write-Host "Run berikutnya jam $berikutnya. Menunggu..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $intervalDetik
    }
} else {
    Run-Autopost
}
