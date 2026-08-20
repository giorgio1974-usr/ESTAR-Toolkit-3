# ============================================================
# ESTAR Admin Toolkit
# File: Core\Version.ps1
# Version: 2.0.0
# ============================================================

$script:ToolkitInfo = [PSCustomObject]@{
    Name        = "ESTAR Admin Toolkit"
    Version     = "2.0.0"
    Build       = "001"
    ReleaseDate = "2026-07-24"
    Author      = "OpenAI + Giorgio"
}

function Get-ToolkitVersion {
    return $script:ToolkitInfo
}

function Show-ToolkitInfo {

    $i = Get-ToolkitVersion

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host (" {0}" -f $i.Name) -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host (" Version : {0}" -f $i.Version)
    Write-Host (" Build   : {0}" -f $i.Build)
    Write-Host (" Release : {0}" -f $i.ReleaseDate)
    Write-Host ""
}
