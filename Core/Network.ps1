# ============================================================
# ESTAR Admin Toolkit
# File: Core\Network.ps1
# Version: 2.0.0
# ============================================================



function Get-IPRange {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Network,
        [Parameter(Mandatory)][ValidateRange(0,255)][int]$StartVlan,
        [Parameter(Mandatory)][ValidateRange(0,255)][int]$EndVlan,
        [Parameter(Mandatory)][ValidateRange(1,254)][int]$StartHost,
        [Parameter(Mandatory)][ValidateRange(1,254)][int]$EndHost
    )

    foreach($Vlan in $StartVlan..$EndVlan){
        foreach($HostNumber in $StartHost..$EndHost){
            "{0}.{1}.{2}" -f $Network,$Vlan,$HostNumber
        }
    }
}

function Read-NetworkRange {

    Write-Host ""
    Write-Host "Intervallo da analizzare" -ForegroundColor Cyan
    Write-Host ""

    do { $StartVlan = [int](Read-Host "VLAN iniziale") }
    until($StartVlan -ge 0 -and $StartVlan -le 255)

    do { $EndVlan = [int](Read-Host "VLAN finale") }
    until($EndVlan -ge $StartVlan -and $EndVlan -le 255)

    do { $StartHost = [int](Read-Host "Host iniziale") }
    until($StartHost -ge 1 -and $StartHost -le 254)

    do { $EndHost = [int](Read-Host "Host finale") }
    until($EndHost -ge $StartHost -and $EndHost -le 254)

    Write-Info ("Generazione IP da {0}.{1}.x a {0}.{2}.x" -f $DefaultNetwork,$StartVlan,$EndVlan)

    Get-IPRange `
        -Network $DefaultNetwork `
        -StartVlan $StartVlan `
        -EndVlan $EndVlan `
        -StartHost $StartHost `
        -EndHost $EndHost
}

function Test-PCOnline {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Computer,
        [int]$Timeout = $Global:PingTimeout
    )

    try{
        $Ping = [System.Net.NetworkInformation.Ping]::new()
        $Reply = $Ping.Send($Computer,$Timeout)
        return ($Reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
    }
    catch{
        return $false
    }
    finally{
        if($Ping){ $Ping.Dispose() }
    }
}
function Get-ComputerList {

    [CmdletBinding()]
    param()

    return Read-NetworkRange

}
