# ============================================================
# ESTAR Admin Toolkit
# File: Core\Runspace.ps1
# ============================================================

Set-StrictMode -Version Latest

function Invoke-Parallel {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$InputObject,

        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,

        [object[]]$ArgumentList = @(),

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = 30
    )

    $items = @($InputObject)

    if ($items.Count -eq 0) {
        return @()
    }

    Write-Info ("Avvio RunspacePool ({0} thread, {1} elementi)..." -f $ThrottleLimit, $items.Count)

    $pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.ApartmentState = 'MTA'
    $pool.ThreadOptions = 'ReuseThread'

    $pending = [System.Collections.Queue]::new()
    foreach ($item in $items) {
        $pending.Enqueue($item)
    }

    $running = [System.Collections.Generic.List[object]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $completed = 0

    try {
        $pool.Open()

        while ($pending.Count -gt 0 -or $running.Count -gt 0) {
            while ($pending.Count -gt 0 -and $running.Count -lt $ThrottleLimit) {
                $powerShell = [PowerShell]::Create()
                $powerShell.RunspacePool = $pool
                $null = $powerShell.AddScript($ScriptBlock)
                $null = $powerShell.AddArgument($pending.Dequeue())

                foreach ($argument in $ArgumentList) {
                    $null = $powerShell.AddArgument($argument)
                }

                $running.Add([pscustomobject]@{
                    PowerShell = $powerShell
                    Handle     = $powerShell.BeginInvoke()
                })
            }

            $finished = @($running | Where-Object { $_.Handle.IsCompleted })

            if ($finished.Count -eq 0) {
                Start-Sleep -Milliseconds 40
                continue
            }

            foreach ($job in $finished) {
                try {
                    foreach ($output in @($job.PowerShell.EndInvoke($job.Handle))) {
                        if ($null -ne $output) {
                            $results.Add($output)
                        }
                    }
                }
                catch {
                    Write-ErrorMsg $_.Exception.Message
                }
                finally {
                    $job.PowerShell.Dispose()
                    $null = $running.Remove($job)
                }

                $completed++

                if ($ShowProgressBar) {
                    $percent = [math]::Round(($completed / $items.Count) * 100, 0)
                    Write-Progress `
                        -Activity 'Scansione in corso...' `
                        -Status ("{0}/{1}" -f $completed, $items.Count) `
                        -PercentComplete $percent
                }
            }
        }
    }
    finally {
        Write-Progress -Activity 'Scansione in corso...' -Completed

        foreach ($job in @($running)) {
            $job.PowerShell.Dispose()
        }

        if ($null -ne $pool) {
            $pool.Close()
            $pool.Dispose()
        }
    }

    if ($results.Count -gt 0 -and $results[0].PSObject.Properties.Name -contains 'Computer') {
        if ($results[0].Computer -match '^\d{1,3}(\.\d{1,3}){3}$') {
            return @($results | Sort-Object {
                $parts = $_.Computer -split '\.'
                ($parts | ForEach-Object { '{0:D3}' -f [int]$_ }) -join '.'
            })
        }

        return @($results | Sort-Object Computer)
    }

    Write-Info ("Runspace completati: {0}" -f $results.Count)
    return @($results)
}

function Test-RunspaceEngine {

    [CmdletBinding()]
    param()

    try {
        $result = Invoke-Parallel `
            -InputObject @(1, 2, 3) `
            -ThrottleLimit 3 `
            -ArgumentList 2 `
            -ScriptBlock {
                param($Value, $Multiplier)

                [pscustomobject]@{
                    Number = $Value
                    Value  = $Value * $Multiplier
                }
            }

        if ($result.Count -ne 3 -or ($result | Measure-Object -Property Value -Sum).Sum -ne 12) {
            throw 'Numero o valore dei risultati non corretto.'
        }

        Write-Success 'Runspace Engine pronto.'
        return $true
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return $false
    }
}
