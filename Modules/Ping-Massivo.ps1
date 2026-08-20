# ============================================================
# ESTAR Admin Toolkit
# File: Modules\Ping-Massivo.ps1
# ============================================================

function Invoke-PingMassivo {

    [CmdletBinding()]
    param(
        [string[]]$Computer,

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = $Global:MaxThreads,

        [switch]$SkipReport,

        [switch]$SkipOpenPrompt
    )

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '                Ping Massivo' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan

    if (-not $Computer) {
        $Computer = @(Get-ComputerList)
    }

    if ($Computer.Count -eq 0) {
        Write-WarningMsg 'Nessun host da analizzare.'
        return @()
    }

    Write-Info ("Host da analizzare: {0}" -f $Computer.Count)

    $scanScript = {
        param($Target, $PingTimeout, $DnsTimeout)

        $ping = $null
        $online = $false
        $roundTripTime = $null
        $hostName = ''
        $errorMessage = ''

        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($Target, $PingTimeout)
            $online = $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success

            if ($online) {
                $roundTripTime = $reply.RoundtripTime

                try {
                    $dnsTask = [System.Net.Dns]::GetHostEntryAsync($Target)

                    if ($dnsTask.Wait($DnsTimeout)) {
                        $hostName = $dnsTask.Result.HostName
                    }
                }
                catch {
                    $hostName = ''
                }
            }
            else {
                $errorMessage = [string]$reply.Status
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
        }
        finally {
            if ($null -ne $ping) {
                $ping.Dispose()
            }
        }

        [pscustomobject]@{
            Computer = $Target
            Hostname = $hostName
            Online   = if ($online) { 'SI' } else { 'NO' }
            RTT      = $roundTripTime
            Note     = $errorMessage
        }
    }

    $startTime = Get-Date

    $results = @(Invoke-Parallel `
        -InputObject $Computer `
        -ScriptBlock $scanScript `
        -ArgumentList @($Global:PingTimeout, $Global:DnsTimeout) `
        -ThrottleLimit $ThrottleLimit)

    $elapsed = (Get-Date) - $startTime
    $total = $results.Count
    $online = @($results | Where-Object Online -eq 'SI').Count
    $offline = $total - $online

    Clear-ToolkitResults
    $results | ForEach-Object { Add-ToolkitResult -InputObject $_ | Out-Null }

    Write-Host ''
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host '              RISULTATI' -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host ("Host analizzati : {0}" -f $total)
    Write-Host ("Online          : {0}" -f $online) -ForegroundColor Green
    Write-Host ("Offline         : {0}" -f $offline) -ForegroundColor Yellow
    Write-Host ("Tempo           : {0:N1} sec." -f $elapsed.TotalSeconds)

    $report = $null

    if (-not $SkipReport) {
        $report = Export-ToolkitExcel `
            -Module 'Ping-Massivo' `
            -Worksheet 'Dati' `
            -Data $results `
            -ReportName 'PingMassivo' `
            -Metadata @{
                ScanTime = '{0:N1} sec.' -f $elapsed.TotalSeconds
                Subnet   = $DefaultNetwork
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

    return $results
}

function Test-PingMassivo {

    [CmdletBinding()]
    param()

    try {
        if (-not (Test-ExcelEngine)) {
            throw 'Excel Engine non disponibile.'
        }

        if (-not (Test-RunspaceEngine)) {
            throw 'Runspace Engine non disponibile.'
        }

        Write-Success 'Modulo Ping Massivo pronto.'
        return $true
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return $false
    }
}

Set-Alias PingMassivo Invoke-PingMassivo
