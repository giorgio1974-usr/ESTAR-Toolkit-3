# =====================================================================
# Analyze-LocalUsers.ps1
# ESTAR Admin Toolkit v3
#
# Punto 7 - Analizza report Excel
#
# Funzioni:
#   1. Analizza il foglio "Dati"
#   2. Conta le occorrenze di ogni colonna "Utente - ..."
#   3. Aggiorna l'intestazione con il conteggio:
#        Utente - NomeUtente (N)
#   4. Considera per "Amministratori locali" solo le colonne
#      con conteggio da 1 a 3 compresi
#   5. Crea il foglio "Amministratori locali" con:
#        Indirizzo IP | Nome DNS | User
#
# Compatibile PowerShell 5
# Richiede ImportExcel
# =====================================================================

function Start-AnalyzeLocalUsers
{
    Clear-Host

    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "       ANALISI REPORT EXCEL"
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""

    # -----------------------------------------------------------------
    # Dipendenze
    # -----------------------------------------------------------------

    try {

        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        Import-Module ImportExcel -ErrorAction Stop

    }
    catch {

        Write-ErrorMsg `
            "Impossibile caricare le dipendenze richieste (System.Windows.Forms / ImportExcel): $($_.Exception.Message)"

        Pause-Toolkit

        return

    }

    # -----------------------------------------------------------------
    # Selezione file
    # -----------------------------------------------------------------

    $Dialog = New-Object System.Windows.Forms.OpenFileDialog

    $Dialog.Title  = "Seleziona il file Excel"

    $Dialog.Filter = "File Excel (*.xlsx)|*.xlsx"

    if ($Dialog.ShowDialog() -ne "OK")
    {

        Write-Host ""

        Write-Host `
            "Operazione annullata." `
            -ForegroundColor Yellow

        Pause-Toolkit

        return

    }

    $ExcelFile = $Dialog.FileName

    Write-Host ""

    Write-Host "File selezionato:"

    Write-Host `
        $ExcelFile `
        -ForegroundColor Green

    Write-Host ""

    # -----------------------------------------------------------------
    # Apertura Excel
    # -----------------------------------------------------------------

    try {

        $Excel = Open-ExcelPackage -Path $ExcelFile -ErrorAction Stop

    }
    catch {

        Write-ErrorMsg `
            "Impossibile aprire il file Excel: $($_.Exception.Message)"

        Pause-Toolkit

        return

    }

    try {

        # =============================================================
        # FOGLIO DATI
        # =============================================================

        $DataSheet = $Excel.Workbook.Worksheets["Dati"]

        if (-not $DataSheet)
        {

            Write-ErrorMsg `
                "Il file non contiene il foglio 'Dati'."

            return

        }

        Write-Host "Analizzo foglio Dati..." -ForegroundColor Cyan

        # -------------------------------------------------------------
        # Dimensioni
        # -------------------------------------------------------------

        $LastRow    = $DataSheet.Dimension.End.Row

        $LastColumn = $DataSheet.Dimension.End.Column

        # -------------------------------------------------------------
        # Colonne amministratori da 1 a 3 occorrenze
        # -------------------------------------------------------------

        $AdminColumns = @()

        # =============================================================
        # Analisi colonne
        # =============================================================

        for ($Column = 1; $Column -le $LastColumn; $Column++)
        {

            $Header = $DataSheet.Cells[1, $Column].Text

            # ---------------------------------------------------------
            # Consideriamo solamente le colonne:
            #
            # Utente - NomeUtente
            # ---------------------------------------------------------

            if ($Header -notlike "Utente -*")
            {
                continue
            }

            # ---------------------------------------------------------
            # Nome dell'utente dalla intestazione
            #
            # Rimuove eventuale vecchio conteggio:
            #
            # Utente - LanSweeper (1507)
            #        ↓
            # LanSweeper
            # ---------------------------------------------------------

            $UserName = $Header.Substring(9).Trim()

            $UserName = `
                $UserName -replace '\s*\(\d+\)\s*$', ''

            # ---------------------------------------------------------
            # Conteggio valori presenti nella colonna
            # ---------------------------------------------------------

            $Count = 0

            for ($Row = 2; $Row -le $LastRow; $Row++)
            {

                $Value = $DataSheet.Cells[$Row, $Column].Text

                if (-not [string]::IsNullOrWhiteSpace($Value))
                {

                    $Count++

                }

            }

            # ---------------------------------------------------------
            # Aggiorna intestazione
            #
            # Esempio:
            #
            # Utente - LanSweeper
            #
            # diventa:
            #
            # Utente - LanSweeper (1507)
            # ---------------------------------------------------------

            $NewHeader = `
                "Utente - {0} ({1})" -f $UserName, $Count

            $DataSheet.Cells[1, $Column].Value = $NewHeader

            Write-Host `
                ("  {0} -> {1}" -f $UserName, $Count)

            # ---------------------------------------------------------
            # Se il conteggio è da 1 a 3:
            #
            # la colonna entra nell'elenco Amministratori locali
            # ---------------------------------------------------------

            if ($Count -ge 1 -and $Count -le 3)
            {

                $AdminColumns += [PSCustomObject]@{

                    Column   = $Column

                    UserName = $UserName

                    Count    = $Count

                }

            }

        }

        # =============================================================
        # CREAZIONE FOGLIO AMMINISTRATORI LOCALI
        # =============================================================

        Write-Host ""

        Write-Host `
            "Creo il foglio 'Amministratori locali'..." `
            -ForegroundColor Cyan

        # -------------------------------------------------------------
        # Elimina eventuale vecchio foglio
        # -------------------------------------------------------------

        $OldSheet = $Excel.Workbook.Worksheets[
            "Amministratori locali"
        ]

        if ($OldSheet)
        {

            $Excel.Workbook.Worksheets.Delete(
                $OldSheet
            )

        }

        # -------------------------------------------------------------
        # Crea nuovo foglio
        # -------------------------------------------------------------

        $AdminSheet = `
            $Excel.Workbook.Worksheets.Add(
                "Amministratori locali"
            )

        # -------------------------------------------------------------
        # Intestazioni
        # -------------------------------------------------------------

        $AdminSheet.Cells[1,1].Value = "Indirizzo IP"

        $AdminSheet.Cells[1,2].Value = "Nome DNS"

        $AdminSheet.Cells[1,3].Value = "User"

        # -------------------------------------------------------------
        # Formattazione intestazioni
        # -------------------------------------------------------------

        $HeaderRange = `
            $AdminSheet.Cells[
                1,
                1,
                1,
                3
            ]

        $HeaderRange.Style.Font.Bold = $true

        # =============================================================
        # RACCOLTA UTENTI
        # =============================================================

        $OutputRow = 2

        foreach ($AdminColumn in $AdminColumns)
        {

            $Column   = $AdminColumn.Column

            $UserName = $AdminColumn.UserName

            # ---------------------------------------------------------
            # Cerca ogni riga valorizzata della colonna
            # ---------------------------------------------------------

            for ($Row = 2; $Row -le $LastRow; $Row++)
            {

                $UserValue = `
                    $DataSheet.Cells[$Row, $Column].Text

                if ([string]::IsNullOrWhiteSpace($UserValue))
                {
                    continue
                }

                # -----------------------------------------------------
                # IP
                #
                # Colonna 1
                # -----------------------------------------------------

                $IP = `
                    $DataSheet.Cells[$Row, 1].Text

                # -----------------------------------------------------
                # Nome DNS
                #
                # Colonna 2
                # -----------------------------------------------------

                $DNS = `
                    $DataSheet.Cells[$Row, 2].Text

                # -----------------------------------------------------
                # Scrive risultato
                # -----------------------------------------------------

                $AdminSheet.Cells[
                    $OutputRow,
                    1
                ].Value = $IP

                $AdminSheet.Cells[
                    $OutputRow,
                    2
                ].Value = $DNS

                $AdminSheet.Cells[
                    $OutputRow,
                    3
                ].Value = $UserValue

                $OutputRow++

            }

        }

        # =============================================================
        # FORMATTAZIONE FOGLIO
        # =============================================================

        if ($OutputRow -gt 2)
        {

            $DataRange = `
                $AdminSheet.Cells[
                    1,
                    1,
                    ($OutputRow - 1),
                    3
                ]

            $DataRange.AutoFitColumns()

            $AdminSheet.View.FreezePanes(
                2,
                1
            )

        }

        # -------------------------------------------------------------
        # Salvataggio
        # -------------------------------------------------------------

        Close-ExcelPackage `
            -ExcelPackage $Excel

        $Excel = $null

        # =============================================================
        # RISULTATO
        # =============================================================

        Write-Host ""

        Write-Host `
            "==============================================" `
            -ForegroundColor Green

        Write-Host `
            " ANALISI COMPLETATA" `
            -ForegroundColor Green

        Write-Host `
            "==============================================" `
            -ForegroundColor Green

        Write-Host ""

        Write-Host `
            "Conteggi aggiornati nel foglio Dati." `
            -ForegroundColor Cyan

        Write-Host ""

        Write-Host `
            "Colonne con 1-3 occorrenze: $($AdminColumns.Count)" `
            -ForegroundColor Cyan

        Write-Host ""

        Write-Host `
            "Creato il foglio:" 

        Write-Host `
            "   Amministratori locali" `
            -ForegroundColor Cyan

        Write-Host ""

        Invoke-Item $ExcelFile

        Pause-Toolkit

    }
    catch {

        # -------------------------------------------------------------
        # Gestione errore
        # -------------------------------------------------------------

        if ($Excel)
        {

            try {

                Close-ExcelPackage `
                    -ExcelPackage $Excel

            }
            catch {}

        }

        Write-ErrorMsg `
            "Errore durante l'analisi del report: $($_.Exception.Message)"

        Pause-Toolkit

    }

}