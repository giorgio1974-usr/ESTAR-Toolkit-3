# ============================================================
# ESTAR Admin Toolkit
# File: Core\Common.ps1
# Version: 2.0.0
# ============================================================

function Show-Header {

    Clear-Host

    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "           ESTAR Admin Toolkit" -ForegroundColor White
    Write-Host ("             Version {0}" -f $ToolkitVersion) -ForegroundColor DarkGray
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
}

function Pause-Toolkit {

    Write-Host ""
    [void](Read-Host "Premi INVIO per continuare")

}

function Confirm-YesNo {

    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    do{
        $r = (Read-Host "$Message (S/N)").Trim().ToUpper()
    }until($r -in @("S","N"))

    return ($r -eq "S")
}

function Read-Choice {

    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    return (Read-Host $Prompt).Trim()
}

function Test-Administrator {

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)

    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
