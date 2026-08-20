# ============================================================
# ESTAR Admin Toolkit
# File: Core\Results.ps1
# Version: 2.0.0
# ============================================================

$script:ToolkitResults = New-Object System.Collections.Generic.List[object]

function Clear-ToolkitResults {
    $script:ToolkitResults.Clear()
}

function Add-ToolkitResult {

    [CmdletBinding()]
    param(

    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        Position = 0
    )]
    $InputObject,

    [string]$Computer,

    [string]$Operation = "",

    [bool]$Success = $true,

    [string]$Message = "",

    $Data = $null

)

if ($PSBoundParameters.ContainsKey("InputObject")) {

    if ($null -eq $script:ToolkitResults) {

        $script:ToolkitResults = New-Object System.Collections.Generic.List[object]

    }

    [void]$script:ToolkitResults.Add($InputObject)

    return $InputObject

}
    $obj = [pscustomobject]@{
        Timestamp = Get-Date -Format $DateFormat
        Computer  = $Computer
        Operation = $Operation
        Success   = $Success
        Message   = $Message
        Data      = $Data
    }

    [void]$script:ToolkitResults.Add($obj)
    return $obj
}

function Get-ToolkitResults {
    return $script:ToolkitResults
}

function Show-ToolkitResults {

    if($script:ToolkitResults.Count -eq 0){
        Write-WarningMsg "Nessun risultato disponibile."
        return
    }

    $script:ToolkitResults |
        Sort-Object Computer |
        Format-Table Timestamp,Computer,Operation,Success,Message -AutoSize
}
