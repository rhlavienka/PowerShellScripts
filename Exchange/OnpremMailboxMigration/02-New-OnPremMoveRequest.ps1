<#
.SYNOPSIS
    Creates the local mailbox move requests for a migration wave from the Ready
    CSV produced by 01-Get-OnPremMigrationScope.ps1. GUI-free replacement for the
    "new local move" wizard in the Exchange admin center.

.DESCRIPTION
    Step 2 of the OnpremMailboxMigration toolset. For every row in the Ready CSV
    it runs New-MoveRequest, tagging them all with -BatchName <WaveName> so 03/04/
    05/06 can address the whole wave with one filter.

    Local moves (both databases in the same org) are carried out by the Mailbox
    Replication Service. There is no migration endpoint and no MRSProxy involved.

    Target database:
      - -TargetDatabase <name>  : every mailbox goes to that one database
      - a "TargetDatabase" column in the CSV : per-row override (wins)
      - neither                 : -TargetDatabase is not passed and Exchange
                                  selects AND load-balances the target itself
                                  (automatic mailbox distribution across databases
                                  with provisioning enabled). This is the normal
                                  case. The script never round-robins or checks
                                  capacity - that is Exchange's job.

    Finalization:
      - -CompleteAfter <datetime>       keep syncing, finalize no earlier than this
      - -SuspendWhenReadyToComplete     sync to ~95%, then AutoSuspended until 06
                                        (or Resume-MoveRequest) finalizes it
      - neither                         the move completes as soon as it can

    Pre-flight (what the wizard does silently):
      - on-premises Exchange Management Shell
      - CSV exists, has an EmailAddress column, no duplicates
      - every address resolves to a local mailbox
      - no move request already exists for any of them
      - every target database named (parameter or CSV) exists and is mounted;
        excluded-from-provisioning is a WARN
      - if no target database anywhere, the automatic-distribution candidate
        databases are listed and confirmation is required

    Output (wave folder):
      <WaveName>_MoveRequests_<ts>.csv   one row per address: Action
                                        (Created / Skipped / Failed), TargetMode,
                                        Message
      <WaveName>_MoveRequests_<ts>.log   console transcript

    Track the wave with 03-Get-OnPremMigrationStatus.ps1.

.PARAMETER CsvPath
    The <WaveName>_Ready_*.csv from 01 (or a _partNN file). Column
    "EmailAddress" required; optional "TargetDatabase" / "TargetArchiveDatabase"
    columns are honoured per row.

.PARAMETER WaveName
    Passed to New-MoveRequest -BatchName. How the rest of the toolset finds this
    wave. Max 64 characters. Defaults to the CSV file's wave prefix
    (text before "_Ready_").

.PARAMETER TargetDatabase
    Single database for every mailbox that has no per-row TargetDatabase. Omit
    for automatic mailbox distribution (the normal case).

.PARAMETER CompleteAfter
    Do not finalize any mailbox before this local date/time. Mutually exclusive
    with -SuspendWhenReadyToComplete.

.PARAMETER SuspendWhenReadyToComplete
    Sync then hold at ~95% until 06-Invoke-OnPremMoveRequestControl.ps1 (or
    Resume-MoveRequest) finalizes it. Mutually exclusive with -CompleteAfter.

.PARAMETER StartAfter
    MRS does not start the move before this local date/time.

.PARAMETER BadItemLimit
    Corrupt items to skip before the move fails. Default 0.

.PARAMETER LargeItemLimit
    Oversized items to skip before the move fails. Default 0.

.PARAMETER Priority
    MRS scheduling weight: Normal (default), High, ...

.PARAMETER PrimaryOnly
    Move only the primary mailbox (leave the archive where it is).

.PARAMETER ArchiveOnly
    Move only the archive. Pairs with a "TargetArchiveDatabase" column or
    -TargetArchiveDatabase.

.PARAMETER TargetArchiveDatabase
    Archive target database for rows without a "TargetArchiveDatabase" column.

.PARAMETER Delimiter
    Delimiter of the input CSV. Default ",".

.PARAMETER OutputFolder
    Parent folder; the wave folder <OutputFolder>\<WaveName>\ is used / created.
    Default: the script folder, or OutputRoot from OnPremMigration.Settings.psd1.

.EXAMPLE
    .\02-New-OnPremMoveRequest.ps1 -CsvPath .\Praha-Sales-W1\Praha-Sales-W1_Ready_2026-09-09_09-00-00.csv `
        -CompleteAfter '2026-09-13 22:00' -WhatIf

.EXAMPLE
    .\02-New-OnPremMoveRequest.ps1 -CsvPath .\wave2_Ready.csv -WaveName "Wave2" `
        -SuspendWhenReadyToComplete -BadItemLimit 10

.NOTES
    Version: 1.0 (2026-09-10)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer),
              the "Move Mailboxes" RBAC role. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [string]$WaveName,

    [string]$TargetDatabase,

    [datetime]$CompleteAfter,

    [switch]$SuspendWhenReadyToComplete,

    [datetime]$StartAfter,

    [int]$BadItemLimit = 0,

    [int]$LargeItemLimit = 0,

    [string]$Priority,

    [switch]$PrimaryOnly,

    [switch]$ArchiveOnly,

    [string]$TargetArchiveDatabase,

    [string]$Delimiter = ',',

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

# [datetime] params default to DateTime.MinValue (which is truthy) - test binding, not value.
$hasCompleteAfter = $PSBoundParameters.ContainsKey('CompleteAfter')
$hasStartAfter    = $PSBoundParameters.ContainsKey('StartAfter')

if ($hasCompleteAfter -and $SuspendWhenReadyToComplete)
{
    throw "-CompleteAfter and -SuspendWhenReadyToComplete cannot be used together."
}
if ($PrimaryOnly -and $ArchiveOnly)
{
    throw "-PrimaryOnly and -ArchiveOnly cannot be used together."
}

# ---------------------------------------------------------------------------
# Settings + wave folder
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV '$CsvPath' not found." }
$csvFullPath = (Resolve-Path -LiteralPath $CsvPath).Path
$csvName     = [System.IO.Path]::GetFileNameWithoutExtension($csvFullPath)

if (-not $WaveName)
{
    $WaveName = if ($csvName -match '^(.*?)_Ready_') { $matches[1] } else { $csvName }
}
if ($WaveName.Length -gt 64) { throw "WaveName '$WaveName' is longer than 64 characters (New-MoveRequest -BatchName limit)." }

$settingsPath = Join-Path $PSScriptRoot 'OnPremMigration.Settings.psd1'
$settings = @{}
if (Test-Path -LiteralPath $settingsPath)
{
    try { $settings = Import-PowerShellDataFile -LiteralPath $settingsPath } catch { $settings = @{} }
}
if (-not $OutputFolder)
{
    $OutputFolder = if ($settings.OutputRoot) { [string]$settings.OutputRoot } else { $PSScriptRoot }
}
if ($settings.DefaultBadItemLimit   -and -not $PSBoundParameters.ContainsKey('BadItemLimit'))   { $BadItemLimit   = [int]$settings.DefaultBadItemLimit }
if ($settings.DefaultLargeItemLimit -and -not $PSBoundParameters.ContainsKey('LargeItemLimit')) { $LargeItemLimit = [int]$settings.DefaultLargeItemLimit }

$timestamp  = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$waveFolder = Join-Path $OutputFolder $WaveName
if (-not (Test-Path -LiteralPath $waveFolder)) { New-Item -ItemType Directory -Path $waveFolder -Force | Out-Null }
$logFile    = Join-Path $waveFolder ("{0}_MoveRequests_{1}.log" -f $WaveName, $timestamp)
$resultCsv  = Join-Path $waveFolder ("{0}_MoveRequests_{1}.csv" -f $WaveName, $timestamp)

function Write-Log
{
    param([string]$Message, [ValidateSet('Info', 'Ok', 'Warn', 'Fail', 'Head')] [string]$Level = 'Info')
    $tag, $color = switch ($Level)
    {
        'Ok'   { '[ OK ] ', 'Green' }
        'Warn' { '[WARN] ', 'Yellow' }
        'Fail' { '[FAIL] ', 'Red' }
        'Head' { '', 'Cyan' }
        default { '', 'White' }
    }
    $line = "$tag$Message"
    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $logFile -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $line)
}

Write-Log "On-prem move-request creation - wave '$WaveName' - $timestamp" -Level Head

# ---------------------------------------------------------------------------
# Environment
# ---------------------------------------------------------------------------
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    throw "Get-ExchangeServer not available. Run this from the on-premises Exchange Management Shell."
}
if (-not (Get-Command New-MoveRequest -ErrorAction SilentlyContinue))
{
    throw "New-MoveRequest not available - your RBAC role is missing 'Move Mailboxes'."
}
Write-Log "On-premises Exchange session confirmed." -Level Ok

# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
$rows = Import-Csv -LiteralPath $csvFullPath -Delimiter $Delimiter
if (-not $rows) { throw "CSV '$CsvPath' is empty." }
if (-not ($rows | Get-Member -Name 'EmailAddress' -MemberType NoteProperty))
{
    throw "CSV '$CsvPath' must contain an 'EmailAddress' column."
}
$hasDbColumn      = [bool]($rows | Get-Member -Name 'TargetDatabase' -MemberType NoteProperty)
$hasArchiveColumn = [bool]($rows | Get-Member -Name 'TargetArchiveDatabase' -MemberType NoteProperty)

$addresses = $rows.EmailAddress | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }
if (-not $addresses) { throw "CSV '$CsvPath' has an 'EmailAddress' column but no values." }

$dupes = $addresses | Group-Object -Property { $_.ToLower() } | Where-Object Count -gt 1
if ($dupes)
{
    $dupes | ForEach-Object { Write-Log "Duplicate address: $($_.Name)" -Level Fail }
    throw "Remove duplicate rows from the CSV and run again."
}
Write-Log "CSV OK - $($addresses.Count) unique address(es)." -Level Ok

# ---------------------------------------------------------------------------
# Resolve mailboxes, check for existing move requests
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Pre-flight" -Level Head

$plan = New-Object System.Collections.Generic.List[object]
$blockers = 0
$targetDbNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($row in $rows)
{
    $addr = ([string]$row.EmailAddress).Trim()
    if (-not $addr) { continue }

    $rowDb        = if ($hasDbColumn) { ([string]$row.TargetDatabase).Trim() } else { '' }
    $rowArchiveDb = if ($hasArchiveColumn) { ([string]$row.TargetArchiveDatabase).Trim() } else { '' }
    $effectiveDb  = if ($rowDb) { $rowDb } elseif ($TargetDatabase) { $TargetDatabase } else { '' }
    $effectiveArchiveDb = if ($rowArchiveDb) { $rowArchiveDb } elseif ($TargetArchiveDatabase) { $TargetArchiveDatabase } else { '' }
    if ($effectiveDb) { [void]$targetDbNames.Add($effectiveDb) }
    if ($effectiveArchiveDb) { [void]$targetDbNames.Add($effectiveArchiveDb) }

    $mbx = Get-Mailbox -Identity $addr -ErrorAction SilentlyContinue
    if (-not $mbx)
    {
        Write-Log "  $addr : no on-premises mailbox" -Level Fail
        $plan.Add([PSCustomObject]@{ Address = $addr; Guid = $null; CurrentDb = $null; TargetDb = $effectiveDb; ArchiveDb = $effectiveArchiveDb; Blocked = 'no mailbox' })
        $blockers++
        continue
    }

    $existing = Get-MoveRequest -Identity $mbx.Guid.ToString() -ErrorAction SilentlyContinue
    if ($existing)
    {
        Write-Log "  $addr : move request already exists ($([string]$existing.Status))" -Level Fail
        $plan.Add([PSCustomObject]@{ Address = $addr; Guid = $mbx.Guid.ToString(); CurrentDb = [string]$mbx.Database; TargetDb = $effectiveDb; ArchiveDb = $effectiveArchiveDb; Blocked = "existing move request ($([string]$existing.Status))" })
        $blockers++
        continue
    }

    if ($effectiveDb -and [string]$mbx.Database -eq $effectiveDb)
    {
        Write-Log "  $addr : already on target database '$effectiveDb'" -Level Warn
        $plan.Add([PSCustomObject]@{ Address = $addr; Guid = $mbx.Guid.ToString(); CurrentDb = [string]$mbx.Database; TargetDb = $effectiveDb; ArchiveDb = $effectiveArchiveDb; Blocked = "already on target database" })
        $blockers++
        continue
    }

    $plan.Add([PSCustomObject]@{ Address = $addr; Guid = $mbx.Guid.ToString(); CurrentDb = [string]$mbx.Database; TargetDb = $effectiveDb; ArchiveDb = $effectiveArchiveDb; Blocked = $null })
}

$runnable = @($plan | Where-Object { -not $_.Blocked })
if (-not $runnable) { throw "Nothing to do - every row is blocked. See the log." }

# ---------------------------------------------------------------------------
# Target databases
# ---------------------------------------------------------------------------
if ($targetDbNames.Count -gt 0)
{
    foreach ($db in $targetDbNames)
    {
        $mdb = Get-MailboxDatabase -Identity $db -Status -ErrorAction SilentlyContinue
        if (-not $mdb)
        {
            throw "Target database '$db' not found."
        }
        if (-not $mdb.Mounted)
        {
            Write-Log "Target database '$db' is not mounted." -Level Fail
            throw "Mount '$db' or fix the CSV / -TargetDatabase, then run again."
        }
        if ($mdb.IsExcludedFromProvisioning -or $mdb.IsSuspendedFromProvisioning)
        {
            Write-Log "Target database '$db' is excluded/suspended from provisioning (moving into it explicitly still works)." -Level Warn
        }
        else
        {
            Write-Log "Target database '$db' - mounted, provisioning enabled." -Level Ok
        }
    }
}
else
{
    Write-Log "" -Level Head
    Write-Log "No target database specified - Exchange automatic mailbox distribution will place and balance every mailbox." -Level Warn
    $candidates = Get-MailboxDatabase -Status |
        Where-Object { -not $_.IsExcludedFromProvisioning -and -not $_.IsExcludedFromProvisioningByOperator -and -not $_.IsExcludedFromProvisioningBySpaceMonitoring -and -not $_.IsSuspendedFromProvisioning -and $_.Mounted }
    if (-not $candidates)
    {
        throw "Automatic distribution has no candidate: no mounted database currently has provisioning enabled. Specify -TargetDatabase."
    }
    Write-Log "Candidate databases ($(@($candidates).Count)):" -Level Info
    $candidates | Sort-Object Name | ForEach-Object { Write-Log ("  {0}" -f $_.Name) -Level Info }
}

# ---------------------------------------------------------------------------
# Summary + confirm
# ---------------------------------------------------------------------------
$targetMode = if ($hasDbColumn) { "per-row from CSV" }
              elseif ($TargetDatabase) { "fixed: $TargetDatabase" }
              else { "automatic mailbox distribution" }

$finalizeMode = if ($hasCompleteAfter) { "CompleteAfter $($CompleteAfter.ToString('yyyy-MM-dd HH:mm')) (UTC $($CompleteAfter.ToUniversalTime().ToString('yyyy-MM-dd HH:mm')))" }
                elseif ($SuspendWhenReadyToComplete) { "SuspendWhenReadyToComplete (manual finalize via 06)" }
                else { "complete as soon as possible" }

$summary = @"

About to create local move requests:
  Wave (BatchName)  : $WaveName
  Mailboxes         : $($runnable.Count)  (blocked / skipped: $blockers)
  Target database   : $targetMode
  Finalization      : $finalizeMode
  StartAfter        : $(if ($hasStartAfter) { $StartAfter.ToString('yyyy-MM-dd HH:mm') } else { '(now)' })
  BadItemLimit      : $BadItemLimit
  LargeItemLimit    : $LargeItemLimit
  Priority          : $(if ($Priority) { $Priority } else { '(default)' })
  Scope             : $(if ($PrimaryOnly) { 'PrimaryOnly' } elseif ($ArchiveOnly) { 'ArchiveOnly' } else { 'primary + archive (if present)' })
"@
Write-Log $summary -Level Head

# ---------------------------------------------------------------------------
# Create
# ---------------------------------------------------------------------------
$n = 0
$outcome = foreach ($item in $plan)
{
    $n++
    if ($item.Blocked)
    {
        [PSCustomObject]@{ Index = $n; Address = $item.Address; Action = 'Skipped'; TargetMode = $targetMode; Message = $item.Blocked }
        continue
    }

    $p = @{
        Identity       = $item.Guid
        BatchName      = $WaveName
        BadItemLimit   = $BadItemLimit
        LargeItemLimit = $LargeItemLimit
        Confirm        = $false
    }
    if ($item.TargetDb)        { $p.TargetDatabase        = $item.TargetDb }
    if ($item.ArchiveDb)       { $p.ArchiveTargetDatabase = $item.ArchiveDb }
    if ($hasCompleteAfter)     { $p.CompleteAfter               = $CompleteAfter }
    if ($SuspendWhenReadyToComplete) { $p.SuspendWhenReadyToComplete = $true }
    if ($hasStartAfter)        { $p.StartAfter            = $StartAfter }
    if ($Priority)             { $p.Priority              = $Priority }
    if ($PrimaryOnly)          { $p.PrimaryOnly           = $true }
    if ($ArchiveOnly)          { $p.ArchiveOnly           = $true }

    if (-not $PSCmdlet.ShouldProcess($item.Address, "New-MoveRequest -BatchName '$WaveName'"))
    {
        [PSCustomObject]@{ Index = $n; Address = $item.Address; Action = 'WhatIf'; TargetMode = $targetMode; Message = 'not created' }
        continue
    }

    try
    {
        New-MoveRequest @p -ErrorAction Stop | Out-Null
        Write-Log ("[{0,4}] {1,-45} created" -f $n, $item.Address) -Level Ok
        [PSCustomObject]@{ Index = $n; Address = $item.Address; Action = 'Created'; TargetMode = $targetMode; Message = '' }
    }
    catch
    {
        Write-Log ("[{0,4}] {1,-45} FAILED - {2}" -f $n, $item.Address, $_.Exception.Message) -Level Fail
        [PSCustomObject]@{ Index = $n; Address = $item.Address; Action = 'Failed'; TargetMode = $targetMode; Message = $_.Exception.Message }
    }
}

$outcome | Export-Csv -LiteralPath $resultCsv -NoTypeInformation -Encoding UTF8

# ---------------------------------------------------------------------------
# Tally
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Result" -Level Head
$outcome | Group-Object Action | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0,-9} {1}" -f $_.Name, $_.Count) -Level Info
}
Write-Log "" -Level Info
Write-Log "Result CSV : $resultCsv" -Level Info
Write-Log "Log        : $logFile" -Level Info

if (@($outcome | Where-Object Action -eq 'Created').Count)
{
    Write-Log "Track the wave: .\03-Get-OnPremMigrationStatus.ps1 -WaveName '$WaveName'" -Level Info
}
