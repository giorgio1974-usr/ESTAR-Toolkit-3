# ============================================================
# ESTAR Admin Toolkit
# File: Config.ps1
# Version: 2.0.0
# ============================================================


#============================================================
# Inizializzo variabili di configurazione
#============================================================

$Global:OpenExcelAfterExport = $false
#============================================================

#============================================================
# Network
#============================================================

$DefaultNetwork = "10.64"

$DefaultVlanStart = 20
$DefaultVlanEnd   = 30

$DefaultHostStart = 1
$DefaultHostEnd   = 254
#============================================================

# -----------------------------
# Toolkit
# -----------------------------
$ToolkitVersion = "2.0.0"
$ToolkitBuild   = "001"

if (-not $ToolkitRoot) {
    $ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# -----------------------------
# Folder structure
# -----------------------------
$CoreFolder     = Join-Path $ToolkitRoot "Core"
$ModulesFolder  = Join-Path $ToolkitRoot "Modules"
$ReportsFolder  = Join-Path $ToolkitRoot "Reports"
$LogsFolder     = Join-Path $ToolkitRoot "Logs"
$TempFolder     = Join-Path $ToolkitRoot "Temp"
$TestsFolder    = Join-Path $ToolkitRoot "Tests"

# -----------------------------
# Performance
# -----------------------------
$Global:MaxThreads   = 30
$Global:PingTimeout  = 200
$Global:DnsTimeout   = 500
$Global:WMITimeout   = 5
$Global:WinRMTimeout = 5

# -----------------------------
# Export
# -----------------------------
$ExportCsv  = $true
$ExportXlsx = $true

# -----------------------------
# Logging
# -----------------------------
$EnableLog = $true
$LogLevel  = "INFO"

# -----------------------------
# Date format
# -----------------------------
$DateFormat = "yyyy-MM-dd HH:mm:ss"

# -----------------------------
# Create required folders
# -----------------------------
foreach($Folder in @(
    $CoreFolder,
    $ModulesFolder,
    $ReportsFolder,
    $LogsFolder,
    $TempFolder,
    $TestsFolder
)){
    if(-not (Test-Path $Folder)){
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    }
}

$LogFile = Join-Path $LogsFolder ("Toolkit_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

#============================================================
# Excel
#============================================================

$TemplatesFolder = Join-Path $ToolkitRoot "Templates"

$DefaultExcelTemplate = Join-Path `
    $TemplatesFolder `
    "Standard.xlsx"

#============================================================
# Runspace
#============================================================

$ShowProgressBar = $true
