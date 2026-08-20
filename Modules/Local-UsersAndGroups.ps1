# ============================================================
# ESTAR Admin Toolkit
# File: Modules\Local-UsersAndGroups.ps1
# Gestione remota di utenti e gruppi locali tramite provider ADSI WinNT.
# ============================================================

function Read-ClientName {

    [CmdletBinding()]
    param(
        [string]$Prompt = 'Nome DNS o indirizzo IP del client'
    )

    do {
        $computer = (Read-Host $Prompt).Trim()
    } until (-not [string]::IsNullOrWhiteSpace($computer))

    return $computer
}

function Get-ClientNetBIOSName {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer
    )

    try {
        $computerDirectory = [ADSI]"WinNT://$Computer,computer"
        return [string]$computerDirectory.Name
    }
    catch {
        return $Computer
    }
}

function Get-WinNTOrigin {

    # Determina se un oggetto ADSI (utente/gruppo) e' Locale o di Dominio
    # confrontando il segmento iniziale dell'AdsPath col nome NetBIOS del client.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LocalName,

        [Parameter(Mandatory)]
        [string]$AdsPath
    )

    $stripped = ($AdsPath -replace '^WinNT://', '').TrimEnd('/')
    $origin = ($stripped -split '/')[0]

    if ($origin -ieq $LocalName) {
        return 'Locale'
    }

    return "Dominio ($origin)"
}

function ConvertTo-WinNTAdsPath {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [Parameter(Mandatory)]
        [string]$Identity
    )

    $identity = $Identity.Trim()

    if ($identity.StartsWith('.\')) {
        $identity = "$Computer\$($identity.Substring(2))"
    }
    elseif ($identity -notmatch '[\\/]') {
        $identity = "$Computer\$identity"
    }

    $identity = $identity -replace '\\', '/'
    return "WinNT://$identity"
}

function Test-ClientLocalGroupMembership {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$ExistingMembers,

        [Parameter(Mandatory)]
        [string]$MemberPath
    )

    $expected = ($MemberPath -replace '^WinNT://', '').TrimEnd('/')

    return @(
        $ExistingMembers | Where-Object {
            $actual = ($_.AdsPath -replace '^WinNT://', '').TrimEnd('/')
            $actual -ieq $expected -or $actual.EndsWith("/$expected", [System.StringComparison]::OrdinalIgnoreCase)
        }
    ).Count -gt 0
}

function Get-ClientLocalGroups {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer
    )

    try {
        $computerDirectory = [ADSI]"WinNT://$Computer,computer"

        foreach ($child in $computerDirectory.psbase.Children) {
            if ($child.SchemaClassName -eq 'group') {
                [pscustomobject]@{
                    Computer    = $Computer
                    Name        = [string]$child.Name
                    Description = [string]$child.Description
                }
            }
        }
    }
    catch {
        throw "Impossibile leggere i gruppi locali di '$Computer': $($_.Exception.Message)"
    }
}

function Get-ClientLocalUsers {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer
    )

    try {
        $computerDirectory = [ADSI]"WinNT://$Computer,computer"

        foreach ($child in $computerDirectory.psbase.Children) {
            if ($child.SchemaClassName -eq 'user') {
                [pscustomobject]@{
                    Computer    = $Computer
                    Name        = [string]$child.Name
                    FullName    = [string]$child.FullName
                    Description = [string]$child.Description
                    Disabled    = [bool]$child.AccountDisabled
                    PasswordAge = [int]$child.psbase.Properties['PasswordAge'].Value
                }
            }
        }
    }
    catch {
        throw "Impossibile leggere gli utenti locali di '$Computer': $($_.Exception.Message)"
    }
}

function Get-ClientLocalGroupMembers {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [Parameter(Mandatory)]
        [string]$Group
    )

    try {
        $groupPath = "WinNT://$Computer/$Group,group"
        $groupObject = [ADSI]$groupPath
        $null = $groupObject.Name

        foreach ($member in @($groupObject.psbase.Invoke('Members'))) {
            $memberType = $member.GetType()
            $adsPath = [string]$memberType.InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)
            $name = [string]$memberType.InvokeMember('Name', 'GetProperty', $null, $member, $null)
            $class = [string]$memberType.InvokeMember('Class', 'GetProperty', $null, $member, $null)

            [pscustomobject]@{
                Computer = $Computer
                Group    = $Group
                Name     = $name
                Type     = $class
                AdsPath  = $adsPath
            }
        }
    }
    catch {
        throw "Impossibile leggere i membri del gruppo '$Group' su '$Computer': $($_.Exception.Message)"
    }
}

function Find-ClientLocalGroupMember {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [Parameter(Mandatory)]
        [string]$Identity
    )

    $matches = foreach ($group in Get-ClientLocalGroups -Computer $Computer) {
        Get-ClientLocalGroupMembers -Computer $Computer -Group $group.Name |
            Where-Object {
                $_.Name -like "*$Identity*" -or $_.AdsPath -like "*$Identity*"
            }
    }

    return @($matches | Sort-Object Group, Name)
}

function Add-ClientLocalGroupMember {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [Parameter(Mandatory)]
        [string]$Group,

        [Parameter(Mandatory)]
        [string]$Member
    )

    $memberPath = ConvertTo-WinNTAdsPath -Computer $Computer -Identity $Member

    try {
        $groupObject = [ADSI]"WinNT://$Computer/$Group,group"
        $null = $groupObject.Name

        $alreadyMember = Test-ClientLocalGroupMembership `
            -ExistingMembers @(Get-ClientLocalGroupMembers -Computer $Computer -Group $Group) `
            -MemberPath $memberPath

        if ($alreadyMember) {
            throw "'$Member' è già membro del gruppo '$Group' su '$Computer'."
        }

        $groupObject.Add($memberPath)
        Write-Success "Aggiunto '$Member' al gruppo '$Group' su '$Computer'."
    }
    catch {
        throw "Impossibile aggiungere '$Member' al gruppo '$Group' su '$Computer': $($_.Exception.Message)"
    }
}

function Remove-ClientLocalGroupMember {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [Parameter(Mandatory)]
        [string]$Group,

        [Parameter(Mandatory)]
        [string]$Member
    )

    $memberPath = ConvertTo-WinNTAdsPath -Computer $Computer -Identity $Member

    try {
        $groupObject = [ADSI]"WinNT://$Computer/$Group,group"
        $null = $groupObject.Name

        $existingMember = Test-ClientLocalGroupMembership `
            -ExistingMembers @(Get-ClientLocalGroupMembers -Computer $Computer -Group $Group) `
            -MemberPath $memberPath

        if (-not $existingMember) {
            throw "'$Member' non è membro del gruppo '$Group' su '$Computer'."
        }

        $groupObject.Remove($memberPath)
        Write-Success "Rimosso '$Member' dal gruppo '$Group' su '$Computer'."
    }
    catch {
        throw "Impossibile rimuovere '$Member' dal gruppo '$Group' su '$Computer': $($_.Exception.Message)"
    }
}

function Invoke-LocalUsersCensus {

    [CmdletBinding()]
    param(
        [string[]]$Computer,

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = $Global:MaxThreads,

        [switch]$SkipReport,

        [switch]$SkipOpenPrompt
    )

    if (-not $Computer) {
        $Computer = @(Get-ComputerList)
    }

    if ($Computer.Count -eq 0) {
        Write-WarningMsg 'Nessun indirizzo IP da analizzare.'
        return @()
    }

    Write-Info ("Inventario utenti locali su {0} indirizzi IP..." -f $Computer.Count)

    $pingScript = {
        param($Target, $PingTimeout, $DnsTimeout)

        $ping = $null

        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($Target, $PingTimeout)

            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $dnsName = ''

                try {
                    $dnsTask = [System.Net.Dns]::GetHostEntryAsync($Target)

                    if ($dnsTask.Wait($DnsTimeout)) {
                        $dnsName = $dnsTask.Result.HostName
                    }
                }
                catch {
                    $dnsName = ''
                }

                [pscustomobject]@{
                    Computer = $Target
                    DnsName  = $dnsName
                }
            }
        }
        catch {
        }
        finally {
            if ($null -ne $ping) {
                $ping.Dispose()
            }
        }
    }

    $onlineClients = @(Invoke-Parallel `
        -InputObject $Computer `
        -ScriptBlock $pingScript `
        -ArgumentList @($Global:PingTimeout, $Global:DnsTimeout) `
        -ThrottleLimit $ThrottleLimit)

    if ($onlineClients.Count -eq 0) {
        Write-WarningMsg 'Nessun client raggiungibile nell''intervallo selezionato.'
        return @()
    }

    Write-Info ("Client raggiungibili: {0}. Lettura degli utenti locali..." -f $onlineClients.Count)

    $inventoryScript = {
        param($Client)

        try {
            $computerDirectory = [ADSI]("WinNT://{0},computer" -f $Client.Computer)
            $localUsers = @(
                $computerDirectory.psbase.Children |
                    Where-Object SchemaClassName -eq 'user' |
                    ForEach-Object { [string]$_.Name } |
                    Sort-Object
            )

            $administrators = @()
            $adminGroup = @(
                $computerDirectory.psbase.Children |
                    Where-Object { $_.SchemaClassName -eq 'group' -and $_.Name -eq 'Administrators' }
            ) | Select-Object -First 1

            if ($adminGroup) {
                foreach ($member in @($adminGroup.psbase.Invoke('Members'))) {
                    $memberType = $member.GetType()
                    $administrators += [string]$memberType.InvokeMember('Name', 'GetProperty', $null, $member, $null)
                }
            }

            $normalUsers = @(
                $localUsers | Where-Object { $_ -notin $administrators }
            )

            [pscustomobject]@{
                'Indirizzo IP'                 = $Client.Computer
                'Nome DNS'                     = $Client.DnsName
                'Utenti gruppo Administrators' = (@($administrators | Sort-Object -Unique) -join '; ')
                'Utenti normali'               = ($normalUsers -join '; ')
                Error                          = ''
            }
        }
        catch {
            [pscustomobject]@{
                'Indirizzo IP'                 = $Client.Computer
                'Nome DNS'                     = $Client.DnsName
                'Utenti gruppo Administrators' = ''
                'Utenti normali'               = ''
                Error                          = $_.Exception.Message
            }
        }
    }

    $scanResults = @(Invoke-Parallel `
        -InputObject $onlineClients `
        -ScriptBlock $inventoryScript `
        -ThrottleLimit $ThrottleLimit)

    $errors = @($scanResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Error) })
    $results = @(
        $scanResults |
            Where-Object { [string]::IsNullOrWhiteSpace($_.Error) } |
            Sort-Object {
                if ($_.'Indirizzo IP' -match '^\d{1,3}(\.\d{1,3}){3}$') {
                    (($_.'Indirizzo IP' -split '\.') | ForEach-Object { '{0:D3}' -f [int]$_ }) -join '.'
                }
                else {
                    $_.'Indirizzo IP'
                }
            }
    )

    $reportData = @(
        $results |
            Select-Object 'Indirizzo IP', 'Nome DNS', 'Utenti gruppo Administrators', 'Utenti normali'
    )

    Clear-ToolkitResults
    $reportData | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    if ($reportData.Count -gt 0) {
        $reportData |
            Format-Table -AutoSize |
            Out-Host

        Write-Success ("Client inventariati: {0}" -f $reportData.Count)
    }

    if ($errors.Count -gt 0) {
        Write-WarningMsg ("Client raggiungibili ma non interrogabili: {0}" -f $errors.Count)
        $errors | Select-Object 'Indirizzo IP', Error | Format-Table -AutoSize | Out-Host
    }

    $report = $null

    if ($reportData.Count -gt 0 -and -not $SkipReport) {
        $report = Export-ToolkitExcel `
            -Module 'Utenti-Locali' `
            -Worksheet 'Dati' `
            -Data $reportData `
            -ReportName 'UtentiLocali' `
            -Metadata @{
                Subnet = $DefaultNetwork
            }

        if ($report) {
            Write-Success "Report creato: $report"
        }
    }

    if ($report -and -not $SkipOpenPrompt) {
        if (Confirm-YesNo -Message 'Aprire il report?') {
            Open-ToolkitReport -Path $report
        }
    }

    return $reportData
}

function Invoke-LocalAdministratorsAnomalySearch {

    [CmdletBinding()]
    param(
        [string[]]$Computer,

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = $Global:MaxThreads,

        [switch]$ExcludeStandardMembers,

        [switch]$SkipReport,

        [switch]$SkipOpenPrompt
    )

    if (-not $Computer) {
        $Computer = @(Get-ComputerList)
    }

    if ($Computer.Count -eq 0) {
        Write-WarningMsg 'Nessun indirizzo IP da analizzare.'
        return @()
    }

    $scanLabel = if ($ExcludeStandardMembers) { 'Ricerca amministratori locali non standard' } else { 'Ricerca membri di Administrators' }
    Write-Info ("{0} su {1} indirizzi IP..." -f $scanLabel, $Computer.Count)

    $pingScript = {
        param($Target, $PingTimeout, $DnsTimeout)

        $ping = $null

        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($Target, $PingTimeout)

            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $dnsName = ''

                try {
                    $dnsTask = [System.Net.Dns]::GetHostEntryAsync($Target)

                    if ($dnsTask.Wait($DnsTimeout)) {
                        $dnsName = $dnsTask.Result.HostName
                    }
                }
                catch {
                    $dnsName = ''
                }

                [pscustomobject]@{
                    Computer = $Target
                    DnsName  = $dnsName
                }
            }
        }
        catch {
        }
        finally {
            if ($null -ne $ping) {
                $ping.Dispose()
            }
        }
    }

    $onlineClients = @(Invoke-Parallel `
        -InputObject $Computer `
        -ScriptBlock $pingScript `
        -ArgumentList @($Global:PingTimeout, $Global:DnsTimeout) `
        -ThrottleLimit $ThrottleLimit)

    if ($onlineClients.Count -eq 0) {
        Write-WarningMsg 'Nessun client raggiungibile nell''intervallo selezionato.'
        return @()
    }

    Write-Info ("Client raggiungibili: {0}. Lettura del gruppo Administrators..." -f $onlineClients.Count)

    $administratorsScript = {
        param($Client, $ExcludeStandard)

        try {
            $computerDirectory = [ADSI]("WinNT://{0},computer" -f $Client.Computer)
            $adminGroup = @(
                $computerDirectory.psbase.Children |
                    Where-Object { $_.SchemaClassName -eq 'group' -and $_.Name -eq 'Administrators' }
            ) | Select-Object -First 1

            $members = @()

            if ($adminGroup) {
                foreach ($member in @($adminGroup.psbase.Invoke('Members'))) {
                    $memberType = $member.GetType()
                    $memberName = [string]$memberType.InvokeMember('Name', 'GetProperty', $null, $member, $null)
                    $memberSid = ''

                    try {
                        $sidBytes = [byte[]]$memberType.InvokeMember('ObjectSID', 'GetProperty', $null, $member, $null)
                        $memberSid = [System.Security.Principal.SecurityIdentifier]::new($sidBytes, 0).Value
                    }
                    catch {
                    }

                    $isStandard = $memberSid -match '^S-1-5-32-' -or
                        $memberSid -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20') -or
                        $memberSid -match '-(500|512|513|514|515|516|518|519|520|525|526|527)$'

                    if (-not $ExcludeStandard -or -not $isStandard) {
                        $members += $memberName
                    }
                }
            }

            [pscustomobject]@{
                Computer       = $Client.Computer
                DnsName        = $Client.DnsName
                Administrators = @($members | Sort-Object -Unique)
                Error          = ''
            }
        }
        catch {
            [pscustomobject]@{
                Computer       = $Client.Computer
                DnsName        = $Client.DnsName
                Administrators = @()
                Error          = $_.Exception.Message
            }
        }
    }

    $scanResults = @(Invoke-Parallel `
        -InputObject $onlineClients `
        -ScriptBlock $administratorsScript `
        -ArgumentList @($ExcludeStandardMembers.IsPresent) `
        -ThrottleLimit $ThrottleLimit)

    $errors = @($scanResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Error) })
    $validResults = @(
        $scanResults |
            Where-Object { [string]::IsNullOrWhiteSpace($_.Error) } |
            Sort-Object {
                if ($_.Computer -match '^\d{1,3}(\.\d{1,3}){3}$') {
                    (($_.Computer -split '\.') | ForEach-Object { '{0:D3}' -f [int]$_ }) -join '.'
                }
                else {
                    $_.Computer
                }
            }
    )

    $memberFrequency = @{}

    foreach ($client in $validResults) {
        foreach ($member in @($client.Administrators | Sort-Object -Unique)) {
            if ($memberFrequency.ContainsKey($member)) {
                $memberFrequency[$member]++
            }
            else {
                $memberFrequency[$member] = 1
            }
        }
    }

    $orderedMembers = @(
        $memberFrequency.Keys |
            Sort-Object `
                @{ Expression = { $memberFrequency[$_] }; Descending = $true }, `
                @{ Expression = { $_ }; Descending = $false }
    )

    $reportData = foreach ($client in $validResults) {
        $row = [ordered]@{
            'Indirizzo IP' = $client.Computer
            'Nome DNS'     = $client.DnsName
        }

        foreach ($member in $orderedMembers) {
            $columnName = 'Utente - {0}' -f $member
            $row[$columnName] = if ($client.Administrators -contains $member) { $member } else { '' }
        }

        [pscustomobject]$row
    }

    Clear-ToolkitResults
    $reportData | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    if ($reportData.Count -gt 0) {
        $reportData | Format-Table -AutoSize | Out-Host
        Write-Success ("Client analizzati: {0}" -f $reportData.Count)
    }

    if ($errors.Count -gt 0) {
        Write-WarningMsg ("Client raggiungibili ma non interrogabili: {0}" -f $errors.Count)
        $errors | Select-Object Computer, Error | Format-Table -AutoSize | Out-Host
    }

    $report = $null

    if ($reportData.Count -gt 0 -and -not $SkipReport) {
        $reportModule = if ($ExcludeStandardMembers) { 'Amministratori-Locali-Non-Standard' } else { 'Ricerca-Anomalie-Administrators' }
        $reportName = if ($ExcludeStandardMembers) { 'AmministratoriLocali' } else { 'RicercaAnomalieAdministrators' }

        $report = Export-ToolkitExcel `
            -Module $reportModule `
            -Worksheet 'Dati' `
            -Data $reportData `
            -ReportName $reportName `
            -Metadata @{
                Subnet = $DefaultNetwork
            }

        if ($report) {
            Write-Success "Report creato: $report"
        }
    }

    if ($report -and -not $SkipOpenPrompt) {
        if (Confirm-YesNo -Message 'Aprire il report?') {
            Open-ToolkitReport -Path $report
        }
    }

    return @($reportData)
}

function Invoke-GetLocalUsers {

    [CmdletBinding()]
    param()

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '               Utenti locali' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '1 - Censimento'
    Write-Host '2 - Ricerca anomalie'
    Write-Host '3 - Amministratori locali'
    Write-Host ''

    $choice = Read-Host 'Selezione'

    switch ($choice) {
        '1' { Invoke-LocalUsersCensus }
        '2' { Invoke-LocalAdministratorsAnomalySearch }
        '3' { Invoke-LocalAdministratorsAnomalySearch -ExcludeStandardMembers }
        default { Write-WarningMsg 'Scelta non valida.' }
    }
}

function Invoke-GroupMembersLegacySearch {

    [CmdletBinding()]
    param(
        [string]$Identity,

        [string[]]$Computer,

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = $Global:MaxThreads
    )

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        $Identity = (Read-Host 'Utente o gruppo da cercare (nome o dominio\\nome)').Trim()
    }

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        Write-WarningMsg 'Ricerca annullata: identità non specificata.'
        return
    }

    if (-not $Computer) {
        $Computer = @(Get-ComputerList)
    }

    if ($Computer.Count -eq 0) {
        Write-WarningMsg 'Nessun indirizzo IP da analizzare.'
        return
    }

    Write-Info ("Ricerca di '$Identity' su {0} indirizzi IP..." -f $Computer.Count)

    $pingScript = {
        param($Target, $PingTimeout)

        $ping = $null

        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($Target, $PingTimeout)

            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                [pscustomobject]@{ Computer = $Target; Online = $true }
            }
        }
        catch {
        }
        finally {
            if ($null -ne $ping) {
                $ping.Dispose()
            }
        }
    }

    $onlineComputers = @(
        Invoke-Parallel `
            -InputObject $Computer `
            -ScriptBlock $pingScript `
            -ArgumentList @($Global:PingTimeout) `
            -ThrottleLimit $ThrottleLimit |
            Where-Object Online |
            Select-Object -ExpandProperty Computer
    )

    if ($onlineComputers.Count -eq 0) {
        Write-WarningMsg 'Nessun client raggiungibile nell''intervallo selezionato.'
        return
    }

    Write-Info ("Client raggiungibili: {0}. Analisi dei gruppi locali..." -f $onlineComputers.Count)

    $searchScript = {
        param($Target, $SearchIdentity)

        $identityForPath = $SearchIdentity -replace '\\', '/'

        try {
            $computerDirectory = [ADSI]"WinNT://$Target,computer"

            foreach ($group in $computerDirectory.psbase.Children) {
                if ($group.SchemaClassName -ne 'group') {
                    continue
                }

                foreach ($member in @($group.psbase.Invoke('Members'))) {
                    $memberType = $member.GetType()
                    $memberName = [string]$memberType.InvokeMember('Name', 'GetProperty', $null, $member, $null)
                    $memberClass = [string]$memberType.InvokeMember('Class', 'GetProperty', $null, $member, $null)
                    $memberPath = [string]$memberType.InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)

                    if (
                        $memberName.IndexOf($SearchIdentity, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                        $memberPath.IndexOf($identityForPath, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                    ) {
                        [pscustomobject]@{
                            Computer = $Target
                            Group    = [string]$group.Name
                            Name     = $memberName
                            Type     = $memberClass
                            AdsPath  = $memberPath
                            Error    = ''
                        }
                    }
                }
            }
        }
        catch {
            [pscustomobject]@{
                Computer = $Target
                Group    = ''
                Name     = ''
                Type     = 'Error'
                AdsPath  = ''
                Error    = $_.Exception.Message
            }
        }
    }

    $scanResults = @(Invoke-Parallel `
        -InputObject $onlineComputers `
        -ScriptBlock $searchScript `
        -ArgumentList @($Identity) `
        -ThrottleLimit $ThrottleLimit)

    $errors = @($scanResults | Where-Object Type -eq 'Error')
    $results = @($scanResults | Where-Object Type -ne 'Error' | Sort-Object Computer, Group, Name)

    Clear-ToolkitResults
    $results | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    if ($results.Count -eq 0) {
        Write-WarningMsg "Nessun membro corrispondente a '$Identity' nei client raggiungibili."
    }
    else {
        $results | Format-Table Computer, Group, Name, Type, AdsPath -AutoSize | Out-Host
        Write-Success ("Corrispondenze trovate: {0}" -f $results.Count)
    }

    if ($errors.Count -gt 0) {
        Write-WarningMsg ("Client non interrogabili dopo il ping: {0}" -f $errors.Count)
        $errors | Select-Object Computer, Error | Format-Table -AutoSize | Out-Host
    }

    return $results
}

function Invoke-NonStandardLocalGroupsCensus {

    [CmdletBinding()]
    param(
        [string[]]$Computer,
        [ValidateRange(1, 256)][int]$ThrottleLimit = $Global:MaxThreads,
        [switch]$SkipReport,
        [switch]$SkipOpenPrompt
    )

    if (-not $Computer) { $Computer = @(Get-ComputerList) }
    if ($Computer.Count -eq 0) { Write-WarningMsg 'Nessun indirizzo IP da analizzare.'; return @() }

    $pingScript = {
        param($Target, $PingTimeout, $DnsTimeout)
        $ping = $null
        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            if ($ping.Send($Target, $PingTimeout).Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $dnsName = ''
                try {
                    $dnsTask = [System.Net.Dns]::GetHostEntryAsync($Target)
                    if ($dnsTask.Wait($DnsTimeout)) { $dnsName = $dnsTask.Result.HostName }
                } catch {}
                [pscustomobject]@{ Computer = $Target; DnsName = $dnsName }
            }
        }
        catch {}
        finally { if ($null -ne $ping) { $ping.Dispose() } }
    }

    Write-Info ("Censimento gruppi non standard su {0} indirizzi IP..." -f $Computer.Count)
    $onlineClients = @(Invoke-Parallel -InputObject $Computer -ScriptBlock $pingScript -ArgumentList @($Global:PingTimeout, $Global:DnsTimeout) -ThrottleLimit $ThrottleLimit)
    if ($onlineClients.Count -eq 0) { Write-WarningMsg 'Nessun client raggiungibile nell''intervallo selezionato.'; return @() }

    $censusScript = {
        param($Client)
        try {
            $computerDirectory = [ADSI]("WinNT://{0},computer" -f $Client.Computer)
            foreach ($group in $computerDirectory.psbase.Children | Where-Object SchemaClassName -eq 'group') {
                $sidBytes = [byte[]]$group.psbase.Properties['ObjectSID'].Value
                $sid = [System.Security.Principal.SecurityIdentifier]::new($sidBytes, 0).Value

                if ($sid -notlike 'S-1-5-32-*') {
                    [pscustomobject]@{
                        'Indirizzo IP' = $Client.Computer
                        'Nome DNS' = $Client.DnsName
                        'Gruppo locale non standard' = [string]$group.Name
                        'Descrizione' = [string]$group.Description
                        Error = ''
                    }
                }
            }
        }
        catch {
            [pscustomobject]@{
                'Indirizzo IP' = $Client.Computer
                'Nome DNS' = $Client.DnsName
                'Gruppo locale non standard' = ''
                'Descrizione' = ''
                Error = $_.Exception.Message
            }
        }
    }

    $scanResults = @(Invoke-Parallel -InputObject $onlineClients -ScriptBlock $censusScript -ThrottleLimit $ThrottleLimit)
    $errors = @($scanResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Error) })
    $reportData = @($scanResults | Where-Object { [string]::IsNullOrWhiteSpace($_.Error) } | Select-Object 'Indirizzo IP', 'Nome DNS', 'Gruppo locale non standard', 'Descrizione' | Sort-Object 'Indirizzo IP', 'Gruppo locale non standard')

    Clear-ToolkitResults
    $reportData | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    if ($reportData.Count -eq 0) { Write-Info 'Nessun gruppo locale non standard trovato.' }
    else { $reportData | Format-Table -AutoSize | Out-Host; Write-Success ("Gruppi non standard trovati: {0}" -f $reportData.Count) }
    if ($errors.Count -gt 0) { Write-WarningMsg ("Client raggiungibili ma non interrogabili: {0}" -f $errors.Count); $errors | Select-Object 'Indirizzo IP', Error | Format-Table -AutoSize | Out-Host }

    $report = $null
    if ($reportData.Count -gt 0 -and -not $SkipReport) {
        $report = Export-ToolkitExcel -Module 'Censimento-Gruppi-Non-Standard' -Worksheet 'Dati' -Data $reportData -ReportName 'GruppiNonStandard' -Metadata @{ Subnet = $DefaultNetwork }
        if ($report) { Write-Success "Report creato: $report" }
    }
    if ($report -and -not $SkipOpenPrompt -and (Confirm-YesNo -Message 'Aprire il report?')) { Open-ToolkitReport -Path $report }
    return $reportData
}

function Invoke-DomainGroupLocalMembershipAnomalies {

    [CmdletBinding()]
    param(
        [string[]]$Computer,
        [ValidateRange(1, 256)][int]$ThrottleLimit = $Global:MaxThreads,
        [switch]$SkipReport,
        [switch]$SkipOpenPrompt
    )

    if (-not $Computer) { $Computer = @(Get-ComputerList) }
    if ($Computer.Count -eq 0) { Write-WarningMsg 'Nessun indirizzo IP da analizzare.'; return @() }

    $pingScript = {
        param($Target, $PingTimeout, $DnsTimeout)
        $ping = $null
        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            if ($ping.Send($Target, $PingTimeout).Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $dnsName = ''
                try {
                    $dnsTask = [System.Net.Dns]::GetHostEntryAsync($Target)
                    if ($dnsTask.Wait($DnsTimeout)) { $dnsName = $dnsTask.Result.HostName }
                } catch {}
                [pscustomobject]@{ Computer = $Target; DnsName = $dnsName }
            }
        }
        catch {}
        finally { if ($null -ne $ping) { $ping.Dispose() } }
    }

    Write-Info ("Ricerca gruppi di dominio su {0} indirizzi IP..." -f $Computer.Count)
    $onlineClients = @(Invoke-Parallel -InputObject $Computer -ScriptBlock $pingScript -ArgumentList @($Global:PingTimeout, $Global:DnsTimeout) -ThrottleLimit $ThrottleLimit)
    if ($onlineClients.Count -eq 0) { Write-WarningMsg 'Nessun client raggiungibile nell''intervallo selezionato.'; return @() }

    $anomalyScript = {
        param($Client)
        try {
            $computerDirectory = [ADSI]("WinNT://{0},computer" -f $Client.Computer)
            foreach ($localGroup in $computerDirectory.psbase.Children | Where-Object SchemaClassName -eq 'group') {
                foreach ($member in @($localGroup.psbase.Invoke('Members'))) {
                    $memberType = $member.GetType()
                    $memberClass = [string]$memberType.InvokeMember('Class', 'GetProperty', $null, $member, $null)
                    $memberPath = [string]$memberType.InvokeMember('AdsPath', 'GetProperty', $null, $member, $null)

                    # I gruppi di dominio sono esposti come WinNT://DOMINIO/Gruppo;
                    # un gruppo locale ha anche il nome del computer nel percorso.
                    $pathParts = @((($memberPath -replace '^WinNT://', '') -split '/') | Where-Object { $_ })

                    if (
                        $memberClass -eq 'Group' -and
                        $pathParts.Count -eq 2 -and
                        $pathParts[0] -notin @('NT AUTHORITY', 'NT SERVICE', 'BUILTIN')
                    ) {
                        [pscustomobject]@{
                            'Indirizzo IP' = $Client.Computer
                            'Nome DNS' = $Client.DnsName
                            'Gruppo dominio' = [string]$memberType.InvokeMember('Name', 'GetProperty', $null, $member, $null)
                            'Gruppo locale' = [string]$localGroup.Name
                            Error = ''
                        }
                    }
                }
            }
        }
        catch {
            [pscustomobject]@{
                'Indirizzo IP' = $Client.Computer
                'Nome DNS' = $Client.DnsName
                'Gruppo dominio' = ''
                'Gruppo locale' = ''
                Error = $_.Exception.Message
            }
        }
    }

    $scanResults = @(Invoke-Parallel -InputObject $onlineClients -ScriptBlock $anomalyScript -ThrottleLimit $ThrottleLimit)
    $errors = @($scanResults | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Error) })
    $reportData = @($scanResults | Where-Object { [string]::IsNullOrWhiteSpace($_.Error) } | Select-Object 'Indirizzo IP', 'Nome DNS', 'Gruppo dominio', 'Gruppo locale' | Sort-Object 'Indirizzo IP', 'Gruppo dominio', 'Gruppo locale')

    Clear-ToolkitResults
    $reportData | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    if ($reportData.Count -eq 0) { Write-Info 'Nessun gruppo di dominio presente nei gruppi locali.' }
    else { $reportData | Format-Table -AutoSize | Out-Host; Write-Success ("Appartenenze di dominio trovate: {0}" -f $reportData.Count) }
    if ($errors.Count -gt 0) { Write-WarningMsg ("Client raggiungibili ma non interrogabili: {0}" -f $errors.Count); $errors | Select-Object 'Indirizzo IP', Error | Format-Table -AutoSize | Out-Host }

    $report = $null
    if ($reportData.Count -gt 0 -and -not $SkipReport) {
        $report = Export-ToolkitExcel -Module 'Anomalie-Gruppi-Dominio' -Worksheet 'Dati' -Data $reportData -ReportName 'GruppiDominioLocali' -Metadata @{ Subnet = $DefaultNetwork }
        if ($report) { Write-Success "Report creato: $report" }
    }
    if ($report -and -not $SkipOpenPrompt -and (Confirm-YesNo -Message 'Aprire il report?')) { Open-ToolkitReport -Path $report }
    return $reportData
}

function Invoke-SearchGroupMember {

    [CmdletBinding()]
    param()

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '            Utenti e gruppi locali' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '1 - Censimento'
    Write-Host '2 - Anomalie'
    Write-Host ''

    switch (Read-Host 'Selezione') {
        '1' { Invoke-NonStandardLocalGroupsCensus }
        '2' { Invoke-DomainGroupLocalMembershipAnomalies }
        default { Write-WarningMsg 'Scelta non valida.' }
    }
}

function Invoke-AddGroupMember {

    [CmdletBinding()]
    param()

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '          Aggiungi membro a un gruppo locale' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan

    $computer = Read-ClientName
    $localName = Get-ClientNetBIOSName -Computer $computer

    Write-Info "Recupero elenco gruppi locali di '$computer'..."

    try {
        $groups = @(Get-ClientLocalGroups -Computer $computer | Sort-Object Name)
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return
    }

    if ($groups.Count -eq 0) {
        Write-WarningMsg "Nessun gruppo locale trovato su '$computer'."
        return
    }

    Write-Host ''
    Write-Host "Gruppi locali su '$computer':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $groups.Count; $i++) {
        $desc = if ($groups[$i].Description) { " - $($groups[$i].Description)" } else { '' }
        Write-Host ("  {0,3}) {1}  [Locale]{2}" -f ($i + 1), $groups[$i].Name, $desc)
    }
    Write-Host '    0) Annulla'
    Write-Host ''

    $groupChoice = Read-Host 'Seleziona il numero del gruppo'

    if ($groupChoice -eq '0' -or [string]::IsNullOrWhiteSpace($groupChoice)) {
        Write-Info 'Operazione annullata.'
        return
    }

    $groupIndex = 0
    if (-not [int]::TryParse($groupChoice, [ref]$groupIndex) -or $groupIndex -lt 1 -or $groupIndex -gt $groups.Count) {
        Write-WarningMsg 'Selezione non valida.'
        return
    }

    $selectedGroup = $groups[$groupIndex - 1]

    Write-Info "Recupero utenti locali di '$computer'..."

    try {
        $users = @(Get-ClientLocalUsers -Computer $computer | Sort-Object Name)
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return
    }

    $manualOption = $users.Count + 1

    Write-Host ''
    Write-Host "Utenti locali su '$computer':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $users.Count; $i++) {
        $stato = if ($users[$i].Disabled) { ' (disabilitato)' } else { '' }
        Write-Host ("  {0,3}) {1}  [Locale]{2}" -f ($i + 1), $users[$i].Name, $stato)
    }
    Write-Host ("  {0,3}) Altro (inserisci manualmente DOMINIO\utente)" -f $manualOption)
    Write-Host '    0) Annulla'
    Write-Host ''

    $memberChoice = Read-Host 'Seleziona il numero dell''utente da aggiungere'

    if ($memberChoice -eq '0' -or [string]::IsNullOrWhiteSpace($memberChoice)) {
        Write-Info 'Operazione annullata.'
        return
    }

    $memberIndex = 0
    if (-not [int]::TryParse($memberChoice, [ref]$memberIndex) -or $memberIndex -lt 1 -or $memberIndex -gt $manualOption) {
        Write-WarningMsg 'Selezione non valida.'
        return
    }

    if ($memberIndex -eq $manualOption) {
        $member = (Read-Host 'Membro da aggiungere (DOMINIO\utente, CLIENT\utente o .\utente)').Trim()

        if ([string]::IsNullOrWhiteSpace($member)) {
            Write-WarningMsg 'Operazione annullata: nessun membro specificato.'
            return
        }
    }
    else {
        $member = $users[$memberIndex - 1].Name
    }

    if (-not (Confirm-YesNo -Message "Aggiungere '$member' al gruppo '$($selectedGroup.Name)' su '$computer'?")) {
        Write-Info 'Operazione annullata dall''utente.'
        return
    }

    try {
        Add-ClientLocalGroupMember -Computer $computer -Group $selectedGroup.Name -Member $member
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
    }
}

function Invoke-RemoveGroupMember {

    [CmdletBinding()]
    param()

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '          Rimuovi membro da un gruppo locale' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan

    $computer = Read-ClientName
    $localName = Get-ClientNetBIOSName -Computer $computer

    Write-Info "Recupero elenco gruppi locali di '$computer'..."

    try {
        $groups = @(Get-ClientLocalGroups -Computer $computer | Sort-Object Name)
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return
    }

    if ($groups.Count -eq 0) {
        Write-WarningMsg "Nessun gruppo locale trovato su '$computer'."
        return
    }

    Write-Host ''
    Write-Host "Gruppi locali su '$computer':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $groups.Count; $i++) {
        $desc = if ($groups[$i].Description) { " - $($groups[$i].Description)" } else { '' }
        Write-Host ("  {0,3}) {1}  [Locale]{2}" -f ($i + 1), $groups[$i].Name, $desc)
    }
    Write-Host '    0) Annulla'
    Write-Host ''

    $groupChoice = Read-Host 'Seleziona il numero del gruppo'

    if ($groupChoice -eq '0' -or [string]::IsNullOrWhiteSpace($groupChoice)) {
        Write-Info 'Operazione annullata.'
        return
    }

    $groupIndex = 0
    if (-not [int]::TryParse($groupChoice, [ref]$groupIndex) -or $groupIndex -lt 1 -or $groupIndex -gt $groups.Count) {
        Write-WarningMsg 'Selezione non valida.'
        return
    }

    $selectedGroup = $groups[$groupIndex - 1]

    Write-Info "Recupero membri del gruppo '$($selectedGroup.Name)'..."

    try {
        $members = @(Get-ClientLocalGroupMembers -Computer $computer -Group $selectedGroup.Name | Sort-Object Name)
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return
    }

    if ($members.Count -eq 0) {
        Write-WarningMsg "Il gruppo '$($selectedGroup.Name)' non ha membri."
        return
    }

    Write-Host ''
    Write-Host "Membri del gruppo '$($selectedGroup.Name)' su '$computer':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $members.Count; $i++) {
        $origin = Get-WinNTOrigin -LocalName $localName -AdsPath $members[$i].AdsPath
        Write-Host ("  {0,3}) {1}  [{2}] [{3}]" -f ($i + 1), $members[$i].Name, $members[$i].Type, $origin)
    }
    Write-Host '    0) Annulla'
    Write-Host ''

    $memberChoice = Read-Host 'Seleziona il numero del membro da rimuovere'

    if ($memberChoice -eq '0' -or [string]::IsNullOrWhiteSpace($memberChoice)) {
        Write-Info 'Operazione annullata.'
        return
    }

    $memberIndex = 0
    if (-not [int]::TryParse($memberChoice, [ref]$memberIndex) -or $memberIndex -lt 1 -or $memberIndex -gt $members.Count) {
        Write-WarningMsg 'Selezione non valida.'
        return
    }

    $selectedMember = $members[$memberIndex - 1]

    # Ricostruisco l'identita' a partire dall'AdsPath gia' noto (Dominio/Computer + Nome),
    # cosi' evito ambiguita' tra membri locali e membri di dominio nello stesso gruppo.
    $memberIdentity = ($selectedMember.AdsPath -replace '^WinNT://', '').TrimEnd('/')

    if (-not (Confirm-YesNo -Message "Rimuovere '$($selectedMember.Name)' dal gruppo '$($selectedGroup.Name)' su '$computer'?")) {
        Write-Info 'Operazione annullata dall''utente.'
        return
    }

    try {
        Remove-ClientLocalGroupMember -Computer $computer -Group $selectedGroup.Name -Member $memberIdentity
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
    }
}

function Invoke-TestWinRM {

    [CmdletBinding()]
    param()

    $computer = Read-ClientName

    try {
        $result = Test-WSMan -ComputerName $computer -ErrorAction Stop
        Write-Success "WinRM disponibile su '$computer' (protocollo $($result.ProtocolVersion))."
    }
    catch {
        Write-ErrorMsg "WinRM non disponibile su '$computer': $($_.Exception.Message)"
    }
}
