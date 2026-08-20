# =====================================================================
# Analyze-LocalUsers.ps1
# Analizza un file Excel di utenti locali
# Crea il foglio "Amministratori locali" con l'elenco degli utenti locali univoci
#
# NOTA: questo modulo viene caricato tramite dot-source insieme a tutti
# gli altri in Modules\*.ps1 all'avvio del toolkit. Per questo motivo
# NON deve contenere codice eseguito a livello di file (top-level): tutto
# ciò che serve solo quando l'utente sceglie questa funzione dal menu va
# dentro Start-AnalyzeLocalUsers, altrimenti verrebbe eseguito ad ogni
# avvio del toolkit, prima ancora di mostrare il menu.
# =====================================================================

function Start-AnalyzeLocalUsers
{
    Clear-Host

    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "      ANALISI UTENTI LOCALI DA EXCEL"
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
    }
    catch {
        Write-ErrorMsg "Impossibile caricare le dipendenze richieste (System.Windows.Forms / ImportExcel): $($_.Exception.Message)"
        Pause-Toolkit
        return
    }

    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = "Seleziona il file Excel"
    $Dialog.Filter = "File Excel (*.xlsx)|*.xlsx"

    if ($Dialog.ShowDialog() -ne "OK")
    {
        Write-Host ""
        Write-Host "Operazione annullata." -ForegroundColor Yellow
        Pause-Toolkit
        return
    }

    $ExcelFile = $Dialog.FileName

    Write-Host ""
    Write-Host "File selezionato:"
    Write-Host $ExcelFile -ForegroundColor Green
    Write-Host ""

    $Rows = @()

    $Sheets = Get-ExcelSheetInfo -Path $ExcelFile

    foreach($Sheet in $Sheets)
    {

        if($Sheet.Name -eq "Amministratori locali")
        {
            continue
        }

        Write-Host "Analizzo foglio $($Sheet.Name)..."

        $Data = Import-Excel `
                    -Path $ExcelFile `
                    -WorksheetName $Sheet.Name
                foreach($Row in $Data)
        {

            $Properties = @($Row.PSObject.Properties)

            if($Properties.Count -lt 3)
            {
                continue
            }

            $IP = $Properties[0].Value
            $Computer = $Properties[1].Value

            for($i=2;$i -lt $Properties.Count;$i++)
            {

                $User = "$($Properties[$i].Value)".Trim()

                if([string]::IsNullOrWhiteSpace($User))
                {
                    continue
                }

                $Rows += [PSCustomObject]@{

                    "Indirizzo IP"  = $IP
                    "Nome macchina" = $Computer
                    "Utente"        = $User

                }

            }

        }

    }

    Write-Host ""
    Write-Host "Conteggio utenti..."
    Write-Host ""

    $UniqueUsers =

        $Rows |
        Group-Object Utente |
        Where-Object Count -eq 1 |
        ForEach-Object{

            $Rows |
            Where-Object Utente -eq $_.Name

        }

    Remove-Worksheet `
        -Path $ExcelFile `
        -WorksheetName "Amministratori locali" `
        -ErrorAction SilentlyContinue
        $UniqueUsers |
    Sort-Object Utente |
    Select-Object `
        Utente,
        NomeComputer,
        IP |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName "Amministratori locali" `
        -TableName "AmministratoriLocali" `
        -AutoSize `
        -BoldTopRow `
        -FreezeTopRow `
        -ClearSheet

    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Green
    Write-Host " ANALISI COMPLETATA"
    Write-Host "==============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Creato il foglio:"
    Write-Host "   Amministratori locali" -ForegroundColor Cyan
    Write-Host ""

    Invoke-Item $ExcelFile

    Pause-Toolkit

}