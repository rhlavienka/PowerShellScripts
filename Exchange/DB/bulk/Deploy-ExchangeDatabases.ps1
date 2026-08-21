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

    Between PHASE 1 and PHASE 2 the script polls every primary database until it
    reports Mounted, or aborts if MountTimeoutSeconds is exceeded (see PARAMETER
    section below). If it aborts:
      1. Wait for the affected server(s) to settle, then check database health
         manually, e.g.:
           Get-MailboxDatabaseCopyStatus -Identity "<DatabaseName>\<Server>"
           Get-EventLog -LogName Application -Source MSExchangeIS -Newest 20
         Look for the cause (slow disk, EDB/Log path not reachable, server not
         ready, etc.) and fix it.
      2. Once the primary database(s) mount cleanly, simply re-run the script
         with the same CsvPath. PHASE 1 skips databases that already exist
         (Get-MailboxDatabase check), so it safely resumes at the mount check
         and continues into PHASE 2/3.

    Prerequisites:
      - The script runs in Exchange Management Shell (EMS) with sufficient permissions.
      - Disk volumes (letters H..S) are created, formatted (ReFS 64K, integrity
        streams OFF) and available on each server.
      - Circular logging stays disabled (default) because of Veeam backups.

.PARAMETER CsvPath
    Path to the CSV file with the database layout.

.PARAMETER WhatIfMode
    If $true, the script only prints what it would do, without making real changes.

.PARAMETER MountTimeoutSeconds
    Maximum time (seconds) to wait, after PHASE 1, for every primary database to
    report a Mounted status before PHASE 2 starts adding passive copies.

.PARAMETER MountCheckIntervalSeconds
    Polling interval (seconds) used while waiting for primary databases to mount.

.EXAMPLE
    .\Deploy-ExchangeDatabases.ps1 -CsvPath "C:\Temp\databases.csv"
    Runs in WhatIf mode (default) and only prints the planned actions for all three phases.

.EXAMPLE
    .\Deploy-ExchangeDatabases.ps1 -CsvPath "C:\Temp\databases.csv" -WhatIfMode $false
    Performs the actual deployment: creates the primary databases, adds the passive
    copies, and prints the resulting per-server layout check.

.NOTES
    Version: 1.1 (2026-08-21)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Changelog:
      1.1 (2026-08-21) - Added a real Mounted-status check for primary databases
                          before PHASE 2 (MountTimeoutSeconds / MountCheckIntervalSeconds),
                          replacing the fixed 60s Start-Sleep.
      1.0 (2026-08-20) - Initial version.

    Expected CSV columns:
      DatabaseName, HomeSite, VolumeLetter, EdbFilePath, LogFolderPath,
      Server, ActivationPreference, CopyRole
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,

    [Parameter(Mandatory=$false)]
    [bool]$WhatIfMode = $true,

    [Parameter(Mandatory=$false)]
    [int]$MountTimeoutSeconds = 300,

    [Parameter(Mandatory=$false)]
    [int]$MountCheckIntervalSeconds = 10
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
    Write-Log "Waiting for primary databases to report Mounted (timeout ${MountTimeoutSeconds}s)..."

    $pending = $primaries | ForEach-Object { "$($_.DatabaseName)\$($_.Server)" }
    $elapsed = 0

    while ($pending.Count -gt 0 -and $elapsed -lt $MountTimeoutSeconds) {
        $pending = $pending | Where-Object {
            $status = Get-MailboxDatabaseCopyStatus -Identity $_ -ErrorAction SilentlyContinue
            $status.Status -ne 'Mounted'
        }
        if ($pending.Count -gt 0) {
            Start-Sleep -Seconds $MountCheckIntervalSeconds
            $elapsed += $MountCheckIntervalSeconds
        }
    }

    if ($pending.Count -gt 0) {
        throw "Primary database(s) not Mounted after ${MountTimeoutSeconds}s, aborting before PHASE 2: $($pending -join ', ')"
    }

    Write-Log "All primary databases are Mounted."
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
