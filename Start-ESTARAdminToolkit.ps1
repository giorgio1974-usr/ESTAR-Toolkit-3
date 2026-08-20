# ============================================================
# ESTAR Admin Toolkit
# File: Start-ESTARAdminToolkit.ps1
# Version: 2.0.0
# ============================================================

#============================================================
# Caricamento configurazione
#============================================================

. "$PSScriptRoot\Config.ps1"

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
#============================================================
# Caricamento Core
#============================================================

$CoreFiles = @(
    "Common.ps1"
    "Logger.ps1"
    "Version.ps1"
    "Initialize.ps1"
    "Network.ps1"
    "Results.ps1"
    "Runspace.ps1"
    "Export.ps1"
)

foreach ($File in $CoreFiles) {

    $Path = Join-Path $PSScriptRoot "Core\$File"

    if (-not (Test-Path $Path)) {
        throw "File Core mancante: $Path"
    }

    . $Path
}

#============================================================
# Inizializzazione Toolkit
#============================================================

Initialize-Toolkit

#============================================================
# Caricamento Moduli
#============================================================

Get-ChildItem (Join-Path $PSScriptRoot "Modules\*.ps1") |
    Sort-Object Name |
    ForEach-Object {

        try {

            . $_.FullName

        }
        catch {

            Write-Error "Errore caricando $($_.Name): $($_.Exception.Message)"
            throw

        }

    }

#============================================================
# Avvio menu
#============================================================

$ExitToolkit = $false

do {

    Show-Header

    Write-Host "1  - Ping Massivo"
    Write-Host "2  - Utenti Locali"
    Write-Host "3  - Cerca utente/gruppo"
    Write-Host ""
    Write-Host "--- OPERAZIONI SU GRUPPI ---" -ForegroundColor DarkGray
    Write-Host "4  - Aggiungi membro gruppo"
    Write-Host "5  - Rimuovi membro gruppo"
    Write-Host ""
    Write-Host "--- Utility ---" -ForegroundColor DarkGray
    Write-Host "6  - Test WinRM"
    Write-Host "7  - Analizza report Excel (utenti unici)"
    Write-Host "8  - Pulizia Disco (PC remoto)"
    Write-Host ""
    Write-Host "99 - Informazioni"
    Write-Host "0  - Esci"
    Write-Host ""

    $Choice = Read-Host "Selezione"

    switch ($Choice) {

        "1"  { if(Get-Command Invoke-PingMassivo -EA SilentlyContinue){Invoke-PingMassivo}; Pause-Toolkit }
        "2"  { if(Get-Command Invoke-GetLocalUsers -EA SilentlyContinue){Invoke-GetLocalUsers}; Pause-Toolkit }
        "3"  { if(Get-Command Invoke-SearchGroupMember -EA SilentlyContinue){Invoke-SearchGroupMember}; Pause-Toolkit }
        "4"  { if(Get-Command Invoke-AddGroupMember -EA SilentlyContinue){Invoke-AddGroupMember}; Pause-Toolkit }
        "5"  { if(Get-Command Invoke-RemoveGroupMember -EA SilentlyContinue){Invoke-RemoveGroupMember}; Pause-Toolkit }
        "6"  { if(Get-Command Invoke-TestWinRM -EA SilentlyContinue){Invoke-TestWinRM}; Pause-Toolkit }
        "7"  { if(Get-Command Start-AnalyzeLocalUsers -EA SilentlyContinue){Start-AnalyzeLocalUsers} }

        "8"  { if(Get-Command Invoke-PuliziaDisco -EA SilentlyContinue){Invoke-PuliziaDisco}; Pause-Toolkit }

        "99" {
            Write-Host ""
            Write-Host "ESTAR Admin Toolkit $ToolkitVersion"
            Write-Host "Build $ToolkitBuild"
            Pause-Toolkit
        }

        "0" {
            $ExitToolkit = $true
        }

        default {
            Write-WarningMsg "Scelta non valida."
            Pause-Toolkit
        }
    }

} until ($ExitToolkit)

Write-Host ""
Write-Host "Chiusura Toolkit..." -ForegroundColor Green
