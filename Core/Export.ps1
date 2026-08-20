# ============================================================
# ESTAR Admin Toolkit
# File: Core\Export.ps1
# ============================================================

Set-StrictMode -Version Latest

function Test-ExcelEngine {

    [CmdletBinding()]
    param()

    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw "Modulo ImportExcel non installato. Installarlo con: Install-Module ImportExcel -Scope CurrentUser"
        }

        Import-Module ImportExcel -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $DefaultExcelTemplate)) {
            throw "Template Excel non trovato: $DefaultExcelTemplate"
        }

        if (-not (Test-Path -LiteralPath $ReportsFolder)) {
            New-Item -ItemType Directory -Path $ReportsFolder -Force | Out-Null
        }

        Write-Success "Excel Engine pronto."
        return $true
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        return $false
    }
}

function Get-ExcelTableName {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Worksheet
    )

    $safeName = $Worksheet -replace '[^A-Za-z0-9_]', '_'

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = 'Data'
    }

    return "Tbl_{0}" -f $safeName
}

function Initialize-StatisticsWorksheet {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Worksheet,

        [Parameter(Mandatory)]
        [string]$Module,

        [Parameter(Mandatory)]
        [int]$TotalRecords,

        [Nullable[int]]$OnlineCount,

        [Parameter(Mandatory)]
        [hashtable]$Metadata
    )

    $Worksheet.Cells['A1'].Value = 'ESTAR Admin Toolkit - Riepilogo'
    $Worksheet.Cells['A2'].Value = 'Modulo'
    $Worksheet.Cells['A3'].Value = 'Versione toolkit'
    $Worksheet.Cells['A4'].Value = 'Data report'
    $Worksheet.Cells['A5'].Value = 'Utente'
    $Worksheet.Cells['A7'].Value = 'Host analizzati'
    $Worksheet.Cells['A8'].Value = 'Online'
    $Worksheet.Cells['A9'].Value = 'Offline'
    $Worksheet.Cells['A10'].Value = 'Durata scansione'
    $Worksheet.Cells['A11'].Value = 'Subnet'

    $Worksheet.Cells['B2'].Value = $Module
    $Worksheet.Cells['B3'].Value = $ToolkitVersion
    $Worksheet.Cells['B4'].Value = Get-Date
    $Worksheet.Cells['B5'].Value = $env:USERNAME
    $Worksheet.Cells['B7'].Value = $TotalRecords
    if ($null -ne $OnlineCount) {
        $Worksheet.Cells['B8'].Value = $OnlineCount
        $Worksheet.Cells['B9'].Value = $TotalRecords - $OnlineCount
    }
    else {
        $Worksheet.Cells['B8:B9'].Clear()
    }
    $Worksheet.Cells['B10'].Value = $Metadata['ScanTime']
    $Worksheet.Cells['B11'].Value = $Metadata['Subnet']

    $Worksheet.Cells['A1:B1'].Merge = $true
    $Worksheet.Cells['A1:B1'].Style.Font.Bold = $true
    $Worksheet.Cells['A1:B1'].Style.Font.Size = 14
    $Worksheet.Cells['A1:B1'].Style.Fill.PatternType = 'Solid'
    $Worksheet.Cells['A1:B1'].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SteelBlue)
    $Worksheet.Cells['A1:B1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $Worksheet.Cells['A2:A11'].Style.Font.Bold = $true
    $Worksheet.Cells['B4'].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm:ss'
    $Worksheet.Cells['A1:B11'].AutoFitColumns()
}

function Export-ToolkitExcel {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Module,

        [Parameter(Mandatory)]
        [string]$Worksheet,

        [Parameter(Mandatory)]
        [object[]]$Data,

        [string]$ReportName = $Module,

        [hashtable]$Metadata = @{}
    )

    $excelPackage = $null

    try {
        if (-not $Data -or $Data.Count -eq 0) {
            throw 'Nessun dato da esportare.'
        }

        if (-not (Test-ExcelEngine)) {
            throw 'Excel Engine non disponibile.'
        }

        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outputFile = Join-Path $ReportsFolder ('{0}_{1}.xlsx' -f $ReportName, $timestamp)

        $tableName = Get-ExcelTableName -Worksheet $Worksheet

        Copy-Item -LiteralPath $DefaultExcelTemplate -Destination $outputFile -Force

        # Export-Excel -PassThru restituisce il package EPPlus effettivamente
        # usato dal modulo. È più affidabile fra le diverse versioni di
        # ImportExcel rispetto a riaprire il file con Open-ExcelPackage.
        $excelPackage = @(
            $Data | Export-Excel `
                -Path $outputFile `
                -PassThru `
                -WorksheetName $Worksheet `
                -StartRow 1 `
                -StartColumn 1 `
                -AutoFilter `
                -AutoSize `
                -TableName $tableName `
                -TableStyle Medium2 `
                -ClearSheet
        ) | Where-Object { $null -ne $_.PSObject.Properties['Workbook'] } | Select-Object -Last 1

        if ($null -eq $excelPackage -or $null -eq $excelPackage.Workbook) {
            throw "ImportExcel non ha restituito un workbook valido per: $outputFile"
        }

        $dataSheet = $excelPackage.Workbook.Worksheets[$Worksheet]

        if ($null -eq $dataSheet -or $null -eq $dataSheet.Dimension) {
            throw "Il foglio '$Worksheet' non contiene dati esportati."
        }

        $lastRow = $dataSheet.Dimension.End.Row
        $lastColumn = $dataSheet.Dimension.End.Column






        
        if ($lastRow -gt 1) {
            $dataSheet.View.FreezePanes(2, 1)
        }

        $dataSheet.Cells.AutoFitColumns()

        $statisticsSheet = $excelPackage.Workbook.Worksheets['Statistiche']

        if ($null -eq $statisticsSheet) {
            $statisticsSheet = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName 'Statistiche'
        }

        $onlineCount = $null

        if ($Data[0].PSObject.Properties.Name -contains 'Online') {
            $onlineCount = @($Data | Where-Object { $_.Online -eq 'SI' }).Count
        }

        Initialize-StatisticsWorksheet `
            -Worksheet $statisticsSheet `
            -Module $Module `
            -TotalRecords $Data.Count `
            -OnlineCount $onlineCount `
            -Metadata $Metadata

        $onlineColumn = 0

        for ($column = 1; $column -le $lastColumn; $column++) {
            if ($dataSheet.Cells[1, $column].Text -eq 'Online') {
                $onlineColumn = $column
                break
            }
        }

        if ($onlineColumn -gt 0 -and $lastRow -gt 1) {
            $range = $dataSheet.Cells[2, $onlineColumn, $lastRow, $onlineColumn].Address

            Add-ConditionalFormatting `
                -Worksheet $dataSheet `
                -Address $range `
                -RuleType ContainsText `
                -ConditionValue 'SI' `
                -BackgroundColor Green

            Add-ConditionalFormatting `
                -Worksheet $dataSheet `
                -Address $range `
                -RuleType ContainsText `
                -ConditionValue 'NO' `
                -BackgroundColor LightPink
        }

        Close-ExcelPackage -ExcelPackage $excelPackage
        $excelPackage = $null

        Write-Success "Report creato: $outputFile"
        return $outputFile
    }
    catch {
        Write-ErrorMsg $_.Exception.Message
        throw
    }
    finally {
        if ($null -ne $excelPackage) {
            $excelPackage.Dispose()
        }
    }
}

function Open-ToolkitReport {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Report non trovato: $Path"
    }

    Invoke-Item -LiteralPath $Path
}
