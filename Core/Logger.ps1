# ============================================================
# ESTAR Admin Toolkit
# File: Core\Logger.ps1
# Version: 2.0.0
# ============================================================

function Write-Log {

    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR","DEBUG")]
        [string]$Level = "INFO"
    )

    $time = Get-Date -Format $DateFormat
    $line = "[{0}] [{1}] {2}" -f $time,$Level,$Message

    if($EnableLog){
        try {
            Add-Content -Path $LogFile -Value $line -ErrorAction Stop
        }
        catch {
            Write-Host "[ATTENZIONE] Impossibile scrivere il file di log ('$LogFile'): $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "[ATTENZIONE] Il logging su file viene disabilitato per questa sessione. Verifica i permessi della cartella Logs." -ForegroundColor Yellow
            $Global:EnableLog = $false
        }
    }

    switch($Level){
        "INFO"    { Write-Host $line -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        "WARNING" { Write-Host $line -ForegroundColor Yellow }
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "DEBUG"   { Write-Host $line -ForegroundColor DarkGray }
    }
}

function Write-Info       { param([string]$Message) Write-Log -Message $Message -Level INFO }
function Write-Success    { param([string]$Message) Write-Log -Message $Message -Level SUCCESS }
function Write-WarningMsg { param([string]$Message) Write-Log -Message $Message -Level WARNING }
function Write-ErrorMsg   { param([string]$Message) Write-Log -Message $Message -Level ERROR }
function Write-DebugLog   { param([string]$Message) Write-Log -Message $Message -Level DEBUG }
