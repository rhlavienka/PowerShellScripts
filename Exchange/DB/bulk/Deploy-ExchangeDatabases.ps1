<#
.SYNOPSIS
    Creation of mailbox databases, their copies, and setting ActivationPreference
    in the Exchange DAG based on a CSV file.

.DESCRIPTION
    The script reads a CSV (1 row = 1 database copy) and, in three phases:
      PHASE 1 - creates the primary databases (rows with ActivationPreference = 1)
                via New-MailboxDatabase on the target server with the given EDB and Log path.
      PHASE 2 - adds passive copies (ActivationPreference 2 and 3) via
                Add-MailboxDatabaseCopy with the corresponding ActivationPreference.
      PHASE 3 - verifies the resulting layout and prints a per-server count check.

    Prerequisites:
      - The script runs in Exchange Management Shell (EMS) with sufficient permissions.
      - Disk volumes (letters H..S) are created, formatted (ReFS 64K, integrity
        streams OFF) and available on each server.
      - Circular logging stays disabled (default) because of Veeam backups.

.PARAMETER CsvPath
    Path to the CSV file with the database layout.

.PARAMETER WhatIfMode
    If $true, the script only prints what it would do, without making real changes.

.NOTES
    Version: 1.0
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    Expected CSV columns:
      DatabaseName, HomeSite, VolumeLetter, EdbFilePath, LogFolderPath,
      Server, ActivationPreference, CopyRole
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,

    [Parameter(Mandatory=$false)]
    [bool]$WhatIfMode = $true
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts][$Level] $Message"
}

# ------------------------------------------------------------------
# Loading and validating the CSV
# ------------------------------------------------------------------
if (-not (Test-Path $CsvPath)) { throw "CSV file not found: $CsvPath" }
$data = Import-Csv -Path $CsvPath
if (-not $data) { throw "CSV is empty." }

$required = 'DatabaseName','VolumeLetter','EdbFilePath','LogFolderPath','Server','ActivationPreference'
foreach ($c in $required) {
    if (-not ($data[0].PSObject.Properties.Name -contains $c)) {
        throw "Required column missing from CSV: $c"
    }
}

Write-Log "Loaded $($data.Count) rows (copies) from $CsvPath"
if ($WhatIfMode) { Write-Log "MODE: WhatIf (no real changes). To deploy, run with -WhatIfMode `$false" 'WARN' }

# ------------------------------------------------------------------
# PHASE 1 - creating primary databases (ActivationPreference = 1)
# ------------------------------------------------------------------
Write-Log "=== PHASE 1: Creating primary databases (AP1) ==="
$primaries = $data | Where-Object { [int]$_.ActivationPreference -eq 1 } |
             Sort-Object { [int]($_.DatabaseName -replace '\D','') }

foreach ($row in $primaries) {
    $db = $row.DatabaseName
    if (Get-MailboxDatabase -Identity $db -ErrorAction SilentlyContinue) {
        Write-Log "Database '$db' already exists - skipping creation." 'WARN'
        continue
    }
    Write-Log "New-MailboxDatabase '$db' on server $($row.Server) | EDB: $($row.EdbFilePath)"
    if (-not $WhatIfMode) {
        New-MailboxDatabase -Name $db `
                            -Server $row.Server `
                            -EdbFilePath $row.EdbFilePath `
                            -LogFolderPath $row.LogFolderPath | Out-Null

        # Recommended properties (adjust to environment needs)
        Set-MailboxDatabase -Identity $db -CircularLoggingEnabled $false

        # Mount the new database
        Mount-Database -Identity $db
    }
}

if (-not $WhatIfMode) {
    Write-Log "Waiting 60s to settle after database creation..."
    Start-Sleep -Seconds 60
}

# ------------------------------------------------------------------
# PHASE 2 - adding passive copies (ActivationPreference 2 and 3)
# ------------------------------------------------------------------
Write-Log "=== PHASE 2: Adding passive copies (AP2, AP3) ==="
$copies = $data | Where-Object { [int]$_.ActivationPreference -gt 1 } |
          Sort-Object { [int]($_.DatabaseName -replace '\D','') }, ActivationPreference

foreach ($row in $copies) {
    $db  = $row.DatabaseName
    $srv = $row.Server
    $ap  = [int]$row.ActivationPreference

    $existing = Get-MailboxDatabaseCopyStatus -Identity "$db\$srv" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "Copy '$db' on '$srv' already exists - skipping." 'WARN'
        continue
    }
    Write-Log "Add-MailboxDatabaseCopy '$db' -> $srv (ActivationPreference $ap)"
    if (-not $WhatIfMode) {
        Add-MailboxDatabaseCopy -Identity $db `
                                -MailboxServer $srv `
                                -ActivationPreference $ap | Out-Null
    }
}

# ------------------------------------------------------------------
# PHASE 3 - verifying the layout
# ------------------------------------------------------------------
Write-Log "=== PHASE 3: Checking the resulting layout ==="
if (-not $WhatIfMode) {
    Write-Log "Copy count per server (expected 72 each):"
    Get-MailboxDatabaseCopyStatus -Server * 2>$null |
        Group-Object { ($_.Identity -split '\\')[1] } |
        Select-Object Name, Count | Format-Table -AutoSize | Out-String | Write-Host

    Write-Log "Layout by ActivationPreference:"
    Get-MailboxDatabase |
        Select-Object Name, @{n='AP';e={($_.ActivationPreference | ForEach-Object {"$($_.Key.Name):$($_.Value)"}) -join ' '}} |
        Sort-Object Name | Format-Table -AutoSize | Out-String | Write-Host
} else {
    # Check directly from the CSV (WhatIf mode)
    $data | Group-Object Server | ForEach-Object {
        $ap1 = ($_.Group | Where-Object {[int]$_.ActivationPreference -eq 1}).Count
        $ap2 = ($_.Group | Where-Object {[int]$_.ActivationPreference -eq 2}).Count
        $ap3 = ($_.Group | Where-Object {[int]$_.ActivationPreference -eq 3}).Count
        [PSCustomObject]@{
            Server=$_.Name; Copies=$_.Count; AP1=$ap1; AP2=$ap2; AP3=$ap3
        }
    } | Sort-Object Server | Format-Table -AutoSize | Out-String | Write-Host
}

Write-Log "=== DONE ==="
if ($WhatIfMode) { Write-Log "This was only a WhatIf run. For real deployment: -WhatIfMode `$false" 'WARN' }
