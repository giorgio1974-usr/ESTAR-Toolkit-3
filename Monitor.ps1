<#
.SYNOPSIS
    Monitor.ps1 - Monitoraggio "stile terminale hacker" della macchina.

.DESCRIZIONE
    Mostra a schermo, con testo verde su sfondo nero, informazioni aggiornate
    periodicamente su:
      - Data/ora
      - Utilizzo CPU
      - Utilizzo memoria (RAM)
      - Processo che consuma più risorse
      - Spazio disco (unità C:)
      - Arrivo di nuove email (via IMAP, stessa casella usata da Thunderbird)
        con avviso grande e lampeggiante quando arriva posta nuova

.USO
    Apri PowerShell ed esegui:
        .\Monitor.ps1
    Se l'esecuzione script è bloccata dalle policy di sistema, esegui prima:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    Premi CTRL+C per interrompere.

.NOTE SU POSTA (Thunderbird)
    Thunderbird non offre un'API di automazione esterna come Outlook, quindi
    lo script si collega DIRETTAMENTE via IMAP allo stesso server a cui è
    collegato Thunderbird, per leggere il numero di email non lette.
    Configura i parametri qui sotto (host, porta, utente). La password viene
    chiesta in modo sicuro all'avvio (non viene mai scritta nel file).

    Esempi di server comuni:
      Gmail        -> imap.gmail.com   porta 993  (serve una "App Password")
      Outlook.com  -> outlook.office365.com  porta 993
      Altro        -> guarda in Thunderbird: Impostazioni account > Server > Nome server
#>

# ============== CONFIGURAZIONE IMAP (da modificare) ==============
$ImapServer = "imaps.sanita.toscana.it"      # <-- metti qui il tuo server IMAP
$ImapPort   = 993
$ImapUser   = "giorgio.perini@estar.toscana.it"   # <-- il tuo indirizzo email
# ====================================================================

# --- Aspetto console: sfondo nero, testo verde ---
$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "Green"
Clear-Host

function Write-Line {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Green
}

# --- Banner grande e lampeggiante per gli avvisi importanti ---
function Show-BigFlashAlert {
    param(
        [string[]]$Lines,
        [int]$FlashCount = 6
    )
    $width = 60
    $border = "*" * $width

    for ($i = 0; $i -lt $FlashCount; $i++) {
        Clear-Host
        if ($i % 2 -eq 0) {
            $fg = "Black"; $bg = "Green"
        } else {
            $fg = "Green"; $bg = "Black"
        }
        $Host.UI.RawUI.ForegroundColor = $fg
        $Host.UI.RawUI.BackgroundColor = $bg

        Write-Host ""
        Write-Host $border -ForegroundColor $fg -BackgroundColor $bg
        Write-Host ""
        foreach ($line in $Lines) {
            $pad = [math]::Max(0, [math]::Floor(($width - $line.Length) / 2))
            Write-Host ((" " * $pad) + $line.ToUpper()) -ForegroundColor $fg -BackgroundColor $bg
        }
        Write-Host ""
        Write-Host $border -ForegroundColor $fg -BackgroundColor $bg
        Write-Host ""

        [console]::Beep(1000, 150)
        Start-Sleep -Milliseconds 300
    }

    # ripristina aspetto normale
    $Host.UI.RawUI.ForegroundColor = "Green"
    $Host.UI.RawUI.BackgroundColor = "Black"
    Clear-Host
}

# --- Credenziali IMAP (password richiesta in modo sicuro, non salvata) ---
Write-Line "Inserisci la password per $ImapUser (server $ImapServer):"
$ImapCred = Get-Credential -UserName $ImapUser -Message "Password casella email (Thunderbird/IMAP)"

function Get-UnreadCountIMAP {
    param($Server, $Port, $Cred)

    try {
        $tcp = New-Object System.Net.Sockets.TcpClient($Server, $Port)
        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, ({ $true }))
        $ssl.AuthenticateAsClient($Server)

        $reader = New-Object System.IO.StreamReader($ssl)
        $writer = New-Object System.IO.StreamWriter($ssl)
        $writer.NewLine = "`r`n"
        $writer.AutoFlush = $true

        # legge il saluto iniziale del server
        $null = $reader.ReadLine()

        $user = $Cred.UserName
        $pass = $Cred.GetNetworkCredential().Password

        $writer.WriteLine("a1 LOGIN $user $pass")
        do { $resp = $reader.ReadLine() } while ($resp -notmatch "^a1 ")

        if ($resp -notmatch "^a1 OK") {
            $writer.WriteLine("a4 LOGOUT")
            $tcp.Close()
            return $null   # login fallito
        }

        $writer.WriteLine("a2 SELECT INBOX")
        do { $resp = $reader.ReadLine() } while ($resp -notmatch "^a2 ")

        $writer.WriteLine("a3 SEARCH UNSEEN")
        $searchLine = ""
        do {
            $line = $reader.ReadLine()
            if ($line -match "^\* SEARCH") { $searchLine = $line }
        } while ($line -notmatch "^a3 ")

        $writer.WriteLine("a4 LOGOUT")
        $tcp.Close()

        if ($searchLine -match "^\* SEARCH\s*(.*)$") {
            $ids = $matches[1].Trim()
            if ([string]::IsNullOrWhiteSpace($ids)) { return 0 }
            return ($ids -split "\s+").Count
        }
        return 0
    } catch {
        return $null   # errore di connessione: verrà ritentato al giro successivo
    }
}

$lastUnreadCount = -1

Write-Line "=================================================="
Write-Line "   MONITOR DI SISTEMA - terminale verde"
Write-Line "   Premi CTRL+C per uscire"
Write-Line "=================================================="
Start-Sleep -Seconds 1

while ($true) {

    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # --- CPU ---
    $cpu = (Get-CimInstance Win32_Processor |
            Measure-Object -Property LoadPercentage -Average).Average

    # --- Memoria ---
    $os       = Get-CimInstance Win32_OperatingSystem
    $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $memFree  = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $memUsed  = [math]::Round($memTotal - $memFree, 2)
    $memPct   = [math]::Round(($memUsed / $memTotal) * 100, 1)

    # --- Processo più pesante (per uso CPU cumulativo) ---
    $topProc = Get-Process | Sort-Object CPU -Descending | Select-Object -First 1

    # --- Disco C: ---
    $disk = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    if ($disk) {
        $diskFreeGB = [math]::Round($disk.Free / 1GB, 1)
        $diskUsedGB = [math]::Round($disk.Used / 1GB, 1)
    }

    Write-Line "[$now] CPU: $cpu`%  |  RAM: $memUsed GB / $memTotal GB ($memPct`%)"

    if ($topProc) {
        Write-Line "   Processo top: $($topProc.ProcessName) (PID $($topProc.Id)) - CPU tot: $([math]::Round($topProc.CPU,1))s"
    }

    if ($disk) {
        Write-Line "   Disco C: usati $diskUsedGB GB / liberi $diskFreeGB GB"
    }

    # --- Controllo nuove email (IMAP / Thunderbird) ---
    $unread = Get-UnreadCountIMAP -Server $ImapServer -Port $ImapPort -Cred $ImapCred

    if ($null -ne $unread) {
        Write-Line "   Posta non letta: $unread"

        if ($lastUnreadCount -ge 0 -and $unread -gt $lastUnreadCount) {
            $nuove = $unread - $lastUnreadCount
            Show-BigFlashAlert -Lines @(
                "NUOVA EMAIL!",
                "$nuove nuovo/i messaggio/i",
                "Non letti totali: $unread"
            )
        }
        $lastUnreadCount = $unread
    } else {
        Write-Line "   (impossibile controllare la posta - verifica server/credenziali)"
    }

    Write-Line "--------------------------------------------------"
    Start-Sleep -Seconds 5
}
