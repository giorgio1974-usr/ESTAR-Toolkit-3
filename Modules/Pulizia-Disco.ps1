# ============================================================
# ESTAR Admin Toolkit
# File: Modules\Pulizia-Disco.ps1
# Pulizia remota di file temporanei, cache Windows Update,
# cache browser e Cestino su un client via WinRM.
# ============================================================

function Get-RemoteFolderSizeMB {

    # Eseguito DENTRO la sessione remota (Invoke-Command),
    # non richiamare direttamente dalla console locale.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return 0
    }

    $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue

    if (-not $items) {
        return 0
    }

    $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum

    if (-not $size) {
        return 0
    }

    return [math]::Round($size / 1MB, 2)
}

function Invoke-RemoteDiskCleanup {

    # Eseguito DENTRO la sessione remota (Invoke-Command).

    [CmdletBinding()]
    param()

    function Get-FolderSizeMB($Path) {
        if (-not (Test-Path $Path)) { return 0 }
        $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        if (-not $items) { return 0 }
        $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if (-not $size) { return 0 }
        return [math]::Round($size / 1MB, 2)
    }

    $detail = [ordered]@{}

    # --- Temp utenti (tutti i profili) ---
    $userTempPaths = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName 'AppData\Local\Temp' } |
        Where-Object { Test-Path $_ }

    $beforeUserTemp = 0
    foreach ($p in $userTempPaths) { $beforeUserTemp += (Get-FolderSizeMB $p) }
    foreach ($p in $userTempPaths) {
        Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue |
            Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    }
    $afterUserTemp = 0
    foreach ($p in $userTempPaths) { $afterUserTemp += (Get-FolderSizeMB $p) }
    $detail['TempUtente_MB'] = [math]::Round($beforeUserTemp - $afterUserTemp, 2)

    # --- Temp di Windows ---
    $winTemp = 'C:\Windows\Temp'
    $before = Get-FolderSizeMB $winTemp
    Get-ChildItem -Path $winTemp -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    $after = Get-FolderSizeMB $winTemp
    $detail['TempWindows_MB'] = [math]::Round($before - $after, 2)

    # --- Cache Windows Update ---
    $wuCache = 'C:\Windows\SoftwareDistribution\Download'
    $before = Get-FolderSizeMB $wuCache
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $wuCache -Recurse -Force -ErrorAction SilentlyContinue |
            Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    } catch {}
    $after = Get-FolderSizeMB $wuCache
    $detail['CacheWindowsUpdate_MB'] = [math]::Round($before - $after, 2)

    # --- Cache browser (Chrome/Edge, tutti i profili utente) ---
    $browserCaches = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        @(
            (Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data\Default\Cache'),
            (Join-Path $_.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Cache')
        )
    } | Where-Object { Test-Path $_ }

    $beforeBrowser = 0
    foreach ($c in $browserCaches) { $beforeBrowser += (Get-FolderSizeMB $c) }
    foreach ($c in $browserCaches) {
        Get-ChildItem -Path $c -Recurse -Force -ErrorAction SilentlyContinue |
            Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    }
    $afterBrowser = 0
    foreach ($c in $browserCaches) { $afterBrowser += (Get-FolderSizeMB $c) }
    $detail['CacheBrowser_MB'] = [math]::Round($beforeBrowser - $afterBrowser, 2)

    # --- Cestino ---
    $cestinoOk = $true
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
    } catch {
        $cestinoOk = $false
    }
    $detail['CestinoSvuotato'] = $cestinoOk

    $disk = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    $freeGB = if ($disk) { [math]::Round($disk.Free / 1GB, 2) } else { $null }

    $totaleLiberatoMB = [math]::Round(
        $detail['TempUtente_MB'] +
        $detail['TempWindows_MB'] +
        $detail['CacheWindowsUpdate_MB'] +
        $detail['CacheBrowser_MB'], 2)

    [pscustomobject]@{
        Computer            = $env:COMPUTERNAME
        TempUtente_MB       = $detail['TempUtente_MB']
        TempWindows_MB      = $detail['TempWindows_MB']
        CacheWU_MB          = $detail['CacheWindowsUpdate_MB']
        CacheBrowser_MB     = $detail['CacheBrowser_MB']
        CestinoSvuotato     = $detail['CestinoSvuotato']
        TotaleLiberato_MB   = $totaleLiberatoMB
        SpazioLiberoC_GB    = $freeGB
    }
}

function New-DcomCimSession {

    # Crea una sessione CIM forzando il protocollo DCOM (porta 135/RPC),
    # utilizzabile anche quando WinRM non e' configurato sul target.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer
    )

    $sessionOption = New-CimSessionOption -Protocol Dcom
    return New-CimSession -ComputerName $Computer -SessionOption $sessionOption -ErrorAction Stop
}

function Enable-RemoteWinRM {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer
    )

    $cimSession = $null

    try {
        $cimSession = New-DcomCimSession -Computer $Computer

        $proc = Invoke-CimMethod `
            -CimSession $cimSession `
            -ClassName Win32_Process `
            -MethodName Create `
            -Arguments @{ CommandLine = 'winrm quickconfig -quiet' } `
            -ErrorAction Stop

        if ($proc.ReturnValue -ne 0) {
            Write-ErrorMsg "'winrm quickconfig' su '$Computer' ha restituito codice $($proc.ReturnValue)."
            return $false
        }

        return $true
    }
    catch {
        Write-ErrorMsg "Impossibile contattare '$Computer' via WMI/DCOM per abilitare WinRM: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($cimSession) {
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-DiskCleanupViaWmi {

    # Fallback quando WinRM non e' disponibile: copia uno script sulla
    # condivisione amministrativa C$ del target, lo esegue via WMI/DCOM
    # (Win32_Process, porta 135), poi legge il risultato da un file JSON
    # scritto dal PC remoto sulla stessa condivisione.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Computer,

        [int]$TimeoutSeconds = 300
    )

    $jobId          = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $remoteScript   = "C:\Windows\Temp\ESTAR-Cleanup-$jobId.ps1"
    $remoteResult   = "C:\Windows\Temp\ESTAR-Cleanup-$jobId-result.json"
    $shareScript    = "\\$Computer\C$\Windows\Temp\ESTAR-Cleanup-$jobId.ps1"
    $shareResult    = "\\$Computer\C$\Windows\Temp\ESTAR-Cleanup-$jobId-result.json"
    $localTempScript = Join-Path $env:TEMP "ESTAR-Cleanup-$jobId.ps1"

    $cimSession = $null

    $scriptBody = @'
function Get-FolderSizeMB($Path) {
    if (-not (Test-Path $Path)) { return 0 }
    $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $items) { return 0 }
    $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    if (-not $size) { return 0 }
    return [math]::Round($size / 1MB, 2)
}

$userTempPaths = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
    ForEach-Object { Join-Path $_.FullName 'AppData\Local\Temp' } |
    Where-Object { Test-Path $_ }

$beforeUserTemp = 0
foreach ($p in $userTempPaths) { $beforeUserTemp += (Get-FolderSizeMB $p) }
foreach ($p in $userTempPaths) {
    Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}
$afterUserTemp = 0
foreach ($p in $userTempPaths) { $afterUserTemp += (Get-FolderSizeMB $p) }

$winTemp = 'C:\Windows\Temp'
$beforeWinTemp = Get-FolderSizeMB $winTemp
Get-ChildItem -Path $winTemp -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike 'ESTAR-Cleanup-*' } |
    Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
$afterWinTemp = Get-FolderSizeMB $winTemp

$wuCache = 'C:\Windows\SoftwareDistribution\Download'
$beforeWu = Get-FolderSizeMB $wuCache
try {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path $wuCache -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
} catch {}
$afterWu = Get-FolderSizeMB $wuCache

$browserCaches = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    @(
        (Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data\Default\Cache'),
        (Join-Path $_.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Cache')
    )
} | Where-Object { Test-Path $_ }

$beforeBrowser = 0
foreach ($c in $browserCaches) { $beforeBrowser += (Get-FolderSizeMB $c) }
foreach ($c in $browserCaches) {
    Get-ChildItem -Path $c -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}
$afterBrowser = 0
foreach ($c in $browserCaches) { $afterBrowser += (Get-FolderSizeMB $c) }

$cestinoOk = $true
try { Clear-RecycleBin -Force -ErrorAction Stop } catch { $cestinoOk = $false }

$disk = Get-PSDrive -Name C -ErrorAction SilentlyContinue
$freeGB = if ($disk) { [math]::Round($disk.Free / 1GB, 2) } else { $null }

$tempUtenteMB   = [math]::Round($beforeUserTemp - $afterUserTemp, 2)
$tempWindowsMB  = [math]::Round($beforeWinTemp - $afterWinTemp, 2)
$cacheWuMB      = [math]::Round($beforeWu - $afterWu, 2)
$cacheBrowserMB = [math]::Round($beforeBrowser - $afterBrowser, 2)
$totaleMB       = [math]::Round($tempUtenteMB + $tempWindowsMB + $cacheWuMB + $cacheBrowserMB, 2)

[pscustomobject]@{
    Computer          = $env:COMPUTERNAME
    TempUtente_MB     = $tempUtenteMB
    TempWindows_MB    = $tempWindowsMB
    CacheWU_MB        = $cacheWuMB
    CacheBrowser_MB   = $cacheBrowserMB
    CestinoSvuotato   = $cestinoOk
    TotaleLiberato_MB = $totaleMB
    SpazioLiberoC_GB  = $freeGB
} | ConvertTo-Json | Out-File -FilePath '__RESULT_PATH__' -Encoding UTF8 -Force
'@

    $scriptBody = $scriptBody.Replace('__RESULT_PATH__', $remoteResult)

    try {
        Write-Info "Copia dello script sulla condivisione amministrativa di '$Computer'..."
        Set-Content -Path $localTempScript -Value $scriptBody -Encoding UTF8 -Force
        Copy-Item -Path $localTempScript -Destination $shareScript -Force -ErrorAction Stop
    }
    catch {
        Write-ErrorMsg "Impossibile copiare lo script su '\\$Computer\C$': $($_.Exception.Message)"
        Write-Info "Verifica che la condivisione amministrativa C$ sia raggiungibile e che tu abbia diritti di amministratore su '$Computer'."
        Remove-Item -Path $localTempScript -Force -ErrorAction SilentlyContinue
        return $null
    }

    try {
        Write-Info "Avvio pulizia remota su '$Computer' via WMI/DCOM..."
        $cimSession = New-DcomCimSession -Computer $Computer

        $commandLine = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$remoteScript`""

        $proc = Invoke-CimMethod `
            -CimSession $cimSession `
            -ClassName Win32_Process `
            -MethodName Create `
            -Arguments @{ CommandLine = $commandLine } `
            -ErrorAction Stop

        if ($proc.ReturnValue -ne 0) {
            Write-ErrorMsg "Impossibile avviare la pulizia remota su '$Computer' (codice ritorno $($proc.ReturnValue))."
            return $null
        }

        $remotePid = $proc.ProcessId
        Write-Info "Processo remoto avviato (PID $remotePid). Attesa completamento..."

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        while ($true) {
            $running = Get-CimInstance -CimSession $cimSession -ClassName Win32_Process -Filter "ProcessId=$remotePid" -ErrorAction SilentlyContinue

            if (-not $running) {
                break
            }

            if ($stopwatch.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                Write-ErrorMsg "Timeout (${TimeoutSeconds}s) in attesa del completamento su '$Computer'."
                return $null
            }

            Start-Sleep -Seconds 3
        }

        if (-not (Test-Path $shareResult)) {
            Write-ErrorMsg "Pulizia remota terminata ma nessun file di risultato trovato su '$Computer'."
            return $null
        }

        $json = Get-Content -Path $shareResult -Raw -ErrorAction Stop
        $result = $json | ConvertFrom-Json

        return $result
    }
    catch {
        Write-ErrorMsg "Errore durante la pulizia remota via WMI su '$Computer': $($_.Exception.Message)"
        return $null
    }
    finally {
        Remove-Item -Path $localTempScript -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $shareScript -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $shareResult -Force -ErrorAction SilentlyContinue

        if ($cimSession) {
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-PuliziaDisco {

    [CmdletBinding()]
    param(
        [string]$Computer,

        [switch]$SkipReport
    )

    Clear-Host -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host '           Pulizia Disco (PC remoto)' -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan

    if (-not $Computer) {
        $Computer = Read-ClientName -Prompt 'Nome DNS o indirizzo IP del PC su cui eseguire la pulizia'
    }

    Write-Info "Verifica raggiungibilita' di '$Computer'..."

    if (-not (Test-PCOnline -Computer $Computer)) {
        Write-ErrorMsg "'$Computer' non risponde al ping. Operazione annullata."
        return $null
    }

    $useWmiFallback = $false

    try {
        Test-WSMan -ComputerName $Computer -ErrorAction Stop | Out-Null
    }
    catch {
        Write-WarningMsg "WinRM non disponibile su '$Computer': $($_.Exception.Message)"

        if (-not (Confirm-YesNo -Message "Vuoi provare ad abilitare WinRM da remoto su '$Computer' adesso?")) {
            Write-Info 'Operazione annullata.'
            return $null
        }

        Write-Info "Tentativo di abilitazione remota di WinRM su '$Computer' (via WMI/DCOM)..."

        if (-not (Enable-RemoteWinRM -Computer $Computer)) {
            Write-ErrorMsg "Non e' stato possibile abilitare WinRM su '$Computer'. Verifica manualmente (accesso diretto alla macchina o GPO dedicata)."
            return $null
        }

        Write-Info 'Attesa avvio del servizio WinRM (5 secondi)...'
        Start-Sleep -Seconds 5

        try {
            Test-WSMan -ComputerName $Computer -ErrorAction Stop | Out-Null
            Write-Success "WinRM ora disponibile su '$Computer'."
        }
        catch {
            Write-WarningMsg "WinRM ancora non disponibile su '$Computer' dopo il tentativo di abilitazione."
            Write-Info "Provo un metodo alternativo che non richiede WinRM (esecuzione via WMI/DCOM su condivisione amministrativa C$)."
            $useWmiFallback = $true
        }
    }

    Write-WarningMsg "L'operazione cancellera' file temporanei, cache Windows Update, cache browser e svuotera' il Cestino su '$Computer'."

    if (-not (Confirm-YesNo -Message "Procedere con la pulizia disco su '$Computer'?")) {
        Write-Info 'Operazione annullata dall''utente.'
        return $null
    }

    Write-Info "Pulizia in corso su '$Computer' (potrebbe richiedere qualche minuto)..."

    $result = $null

    if ($useWmiFallback) {
        $result = Invoke-DiskCleanupViaWmi -Computer $Computer

        if (-not $result) {
            Write-ErrorMsg "Impossibile completare la pulizia su '$Computer' ne' via WinRM ne' via WMI. Verifica connettivita' di rete (porte 445/135) e permessi amministrativi."
            return $null
        }
    }
    else {
        try {
            $result = Invoke-Command `
                -ComputerName $Computer `
                -ScriptBlock ${function:Invoke-RemoteDiskCleanup} `
                -ErrorAction Stop
        }
        catch {
            Write-ErrorMsg "Errore durante la pulizia su '$Computer': $($_.Exception.Message)"
            return $null
        }
    }

    if (-not $result) {
        Write-WarningMsg 'Nessun risultato ricevuto dal client.'
        return $null
    }

    Clear-ToolkitResults
    Add-ToolkitResult -InputObject $result | Out-Null


    Write-Host ''
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host '                RISULTATI' -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host ("Computer               : {0}" -f $result.Computer)
    Write-Host ("Temp utente             : {0} MB" -f $result.TempUtente_MB)
    Write-Host ("Temp Windows            : {0} MB" -f $result.TempWindows_MB)
    Write-Host ("Cache Windows Update    : {0} MB" -f $result.CacheWU_MB)
    Write-Host ("Cache browser           : {0} MB" -f $result.CacheBrowser_MB)
    Write-Host ("Cestino svuotato        : {0}" -f $(if ($result.CestinoSvuotato) { 'SI' } else { 'NO' }))
    Write-Host ("Totale liberato         : {0} MB" -f $result.TotaleLiberato_MB) -ForegroundColor Green
    Write-Host ("Spazio libero su C:     : {0} GB" -f $result.SpazioLiberoC_GB) -ForegroundColor Magenta

    Write-Success "Pulizia completata su '$Computer'."

    $report = $null

    if (-not $SkipReport) {
        $report = Export-ToolkitExcel `
            -Module 'Pulizia-Disco' `
            -Worksheet 'Dati' `
            -Data @($result) `
            -ReportName 'PuliziaDisco'

        if ($report) {
            Write-Success "Report creato: $report"
        }
    }

    if ($report) {
        if (Confirm-YesNo -Message 'Aprire il report?') {
            Open-ToolkitReport -Path $report
        }
    }

    return $result
}

function Test-PuliziaDisco {

    [CmdletBinding()]
    param()

    try {
        if (-not (Get-Command Invoke-Command -EA SilentlyContinue)) {
            throw 'Invoke-Command non disponibile.'
        }

        Write-Success 'Modulo Pulizia Disco pronto.'
        return $true
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return $false
    }
}

Set-Alias PuliziaDisco Invoke-PuliziaDisco
