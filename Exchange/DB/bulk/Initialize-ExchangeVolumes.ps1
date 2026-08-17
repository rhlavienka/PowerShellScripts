<#
.SYNOPSIS
    Initialization of disk volumes for Exchange - assigning drive letters H..S,
    formatting to ReFS (64K allocation unit, integrity streams OFF)
    and preparing the directory structure.

.DESCRIPTION
    The script runs LOCALLY on each mailbox server (EXSVR-S1-1..4, EXSVR-S2-1..4).
    For each data disk intended for databases:
      1. Brings the disk online and clears the read-only flag.
      2. Initializes the disk (GPT), if not already initialized.
      3. Creates a partition and assigns a letter in the H..S range.
      4. Formats it as ReFS with a 64K allocation unit.
      5. Disables integrity streams (Exchange requirement for DB volumes).

    Disk selection:
      - By default processes all RAW (uninitialized) disks of a suitable size.
      - Alternatively, specific disk numbers can be provided via -DiskNumbers.

.PARAMETER StartLetter
    First letter of the volume range. Default 'H'.

.PARAMETER VolumeCount
    Number of volumes per server. Default 12 (H..S).

.PARAMETER DiskNumbers
    Optional: explicit list of disk numbers to process (in order).
    If not specified, all RAW disks sorted by number are used.

.PARAMETER MinSizeGB
    Minimum disk size (GB) to be considered a data disk. Default 1000.

.PARAMETER LabelPrefix
    Prefix for the volume label. Default 'EXVOL'. Result e.g. EXVOL_H.

.PARAMETER WhatIfMode
    If $true (default), only prints the plan without making real changes.

.NOTES
    Version: 1.0
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    WARNING: Formatting is IRREVERSIBLE. Run only on the correct server and disks.
    Recommendation: run first in the default WhatIf mode and review the plan.

    DATA PROTECTION: The script formats ONLY disks that are RAW (completely unformatted).
    Any already formatted disk (has a file system) is entirely SKIPPED, regardless
    of whether it contains data. Existing volumes are therefore never overwritten.
#>

[CmdletBinding()]
param(
    [char]  $StartLetter = 'H',
    [int]   $VolumeCount = 12,
    [int[]] $DiskNumbers,
    [int]   $MinSizeGB = 1000,
    [string]$LabelPrefix = 'EXVOL',
    [bool]  $WhatIfMode = $true
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts][$Level] $Message"
}

# ------------------------------------------------------------------
# Preparing the list of target letters (H..S)
# ------------------------------------------------------------------
$letters = 0..($VolumeCount-1) | ForEach-Object { [char]([int][char]$StartLetter + $_) }
Write-Log "Target volume letters: $($letters -join ', ')"

# Check for collisions with already used letters
$used = (Get-Volume | Where-Object DriveLetter).DriveLetter
$conflict = $letters | Where-Object { $used -contains $_ }
if ($conflict) {
    Write-Log "WARNING: the following letters are already in use: $($conflict -join ', ')" 'WARN'
    Write-Log "Check the state before continuing." 'WARN'
}

if ($WhatIfMode) { Write-Log "MODE: WhatIf (no changes). For real formatting: -WhatIfMode `$false" 'WARN' }

# ------------------------------------------------------------------
# Selecting disks to process
# ------------------------------------------------------------------
if ($DiskNumbers) {
    $disks = $DiskNumbers | ForEach-Object { Get-Disk -Number $_ }
    Write-Log "Using explicitly specified disks: $($DiskNumbers -join ', ')"
} else {
    $disks = Get-Disk |
             Where-Object { $_.PartitionStyle -eq 'RAW' -and ($_.Size/1GB) -ge $MinSizeGB } |
             Sort-Object Number
    Write-Log "Automatically selected RAW disks (>= $MinSizeGB GB): $(( $disks.Number ) -join ', ')"
}

if (-not $disks) { throw "No suitable disks found to process." }
if ($disks.Count -lt $letters.Count) {
    Write-Log "Number of disks found ($($disks.Count)) is less than the requested volumes ($($letters.Count))." 'WARN'
}

# ------------------------------------------------------------------
# Processing disks
# ------------------------------------------------------------------
$pairCount = [Math]::Min($disks.Count, $letters.Count)
for ($i = 0; $i -lt $pairCount; $i++) {
    $disk   = $disks[$i]
    $letter = $letters[$i]
    $label  = "{0}_{1}" -f $LabelPrefix, $letter
    $sizeGB = [Math]::Round($disk.Size/1GB, 0)

    Write-Log "----------------------------------------------------------------"
    Write-Log "Disk #$($disk.Number) ($sizeGB GB)  ->  $letter :  (label $label)"

    # --- Format ONLY a RAW (unformatted) disk. Skip an already formatted one entirely. ---
    $current = Get-Disk -Number $disk.Number
    if ($current.PartitionStyle -ne 'RAW') {
        Write-Log "  SKIPPING disk #$($disk.Number) - already formatted ($($current.PartitionStyle))." 'WARN'
        continue
    }

    if ($WhatIfMode) {
        Write-Log "  [WhatIf] Online + Init GPT + Partition ${letter}: + Format ReFS 64K + IntegrityStreams OFF"
        continue
    }

    # 1) Bring online + clear read-only
    if ($disk.IsOffline)  { Set-Disk -Number $disk.Number -IsOffline $false }
    if ($disk.IsReadOnly) { Set-Disk -Number $disk.Number -IsReadOnly $false }

    # 2) GPT initialization (disk is RAW)
    Initialize-Disk -Number $disk.Number -PartitionStyle GPT

    # 3) Create partition and assign letter
    $part = New-Partition -DiskNumber $disk.Number -UseMaximumSize -DriveLetter $letter

    # 4) Format ReFS, 64K allocation unit
    Format-Volume -Partition $part -FileSystem ReFS -AllocationUnitSize 65536 `
                  -NewFileSystemLabel $label -Confirm:$false | Out-Null

    # 5) Disable integrity streams for the volume root (Exchange requirement)
    Set-FileIntegrity "$($letter):\" -Enable $false

    Write-Log "  Done: ${letter}: formatted (ReFS 64K, integrity OFF)."
}

# ------------------------------------------------------------------
# Verification
# ------------------------------------------------------------------
Write-Log "=== Result check ==="
if (-not $WhatIfMode) {
    Get-Volume | Where-Object { $letters -contains $_.DriveLetter } |
        Select-Object DriveLetter, FileSystemLabel, FileSystem,
                      @{n='SizeGB';e={[Math]::Round($_.Size/1GB,0)}},
                      @{n='FreeGB';e={[Math]::Round($_.SizeRemaining/1GB,0)}} |
        Sort-Object DriveLetter | Format-Table -AutoSize | Out-String | Write-Host
} else {
    Write-Log "WhatIf run - no real volumes were created." 'WARN'
}

Write-Log "=== DONE ==="
Write-Log "Note: database directories (e.g. H:\\DBSE101) are created by the"
Write-Log "Deploy-ExchangeDatabases.ps1 script during New-MailboxDatabase / Add-MailboxDatabaseCopy."
