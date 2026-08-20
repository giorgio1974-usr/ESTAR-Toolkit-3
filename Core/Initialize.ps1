# ============================================================
# ESTAR Admin Toolkit
# File: Core\Initialize.ps1
# Version: 2.0.0
# ============================================================

function Initialize-Toolkit {

    $RequiredFolders = @(
        $CoreFolder,
        $ModulesFolder,
        $ReportsFolder,
        $LogsFolder,
        $TempFolder,
        $TestsFolder
    )

    foreach($Folder in $RequiredFolders){

        if([string]::IsNullOrWhiteSpace($Folder)){
            throw "Variabile cartella non valorizzata."
        }

        if(-not (Test-Path $Folder)){
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
            Write-Info "Creata cartella: $Folder"
        }

    }

    if(-not (Test-Administrator)){
        Write-WarningMsg "Toolkit avviato senza privilegi amministrativi."
    }

    Write-Success "Toolkit inizializzato."

    return $true
}
