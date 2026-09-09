<#
.SYNOPSIS
    Progress dashboard for an on-premises mailbox migration wave. Classifies every
    move request, keeps a run-over-run history so stalled mailboxes stand out, and
    shows where automatic mailbox distribution actually placed each mailbox.

.DESCRIPTION
    Step 3 of the OnpremMailboxMigration toolset. Collects per-mailbox state from
    Get-MoveRequest + Get-MoveRequestStatistics and classifies it from Status /
    StatusDetail:

      Queued        waiting for an MRS slot (throttling) - not a problem
      Provisioning  InProgress, still building the initial folder hierarchy
      Syncing       InProgress, copying messages
      Synced        AutoSuspended - initial sync done, waiting to be finalized
      Completing    CompletionInProgress
      Completed     Completed / CompletedWithWarning
      Suspended     manually suspended
      Failed        Failed / CompletionFailed
      Stalled       still Syncing and StalledSinceTimestamp is set, OR
                    PercentComplete + ItemsTransferred unchanged since the
                    previous run of this script and -StallHours elapsed
      CorruptItems  Failed only because bad/large items exceeded the limit

    Scope (pick one):
      -WaveName <name>   Get-MoveRequest -BatchName <name>
      -CsvPath <file>    the addresses in the CSV
      (neither)          every move request in the organization

    Output (wave folder when -WaveName is used, else -OutputFolder):
      <label>_Status_<ts>.csv        per-mailbox snapshot
      <label>_Status_<ts>.log        console transcript
      MigrationStatus-History.csv    appended every run - keep it, it is what
                                     makes Stalled detection work

    -GridView also opens the snapshot in Out-GridView (sortable / filterable)
    when that cmdlet exists on the host.

.PARAMETER WaveName
    Restrict to one wave (New-MoveRequest -BatchName value from 02).

.PARAMETER CsvPath
    Restrict to the addresses in this CSV.

.PARAMETER IdentityColumn
    Identity column in -CsvPath. Default "EmailAddress".

.PARAMETER StallHours
    A Syncing mailbox counts as Stalled when its progress is unchanged since the
    previous run and its LastUpdateTimestamp is older than this many hours.
    Default 6.

.PARAMETER IncludeCompleted
    Also list mailboxes that are already Completed (hidden by default).

.PARAMETER GridView
    Send the snapshot to Out-GridView in addition to the console and CSV.

.PARAMETER OutputFolder
    Where snapshot / log / history live. Default: the wave folder
    (<script or OutputRoot>\<WaveName>) when -WaveName is given, otherwise the
    script folder.

.EXAMPLE
    .\03-Get-OnPremMigrationStatus.ps1 -WaveName "Praha-Sales-W1" -GridView

.EXAMPLE
    while ($true) { .\03-Get-OnPremMigrationStatus.ps1 -WaveName "Wave2"; Start-Sleep 1800 }

.NOTES
    Version: 1.0 (2026-09-09)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer).
              View-Only Recipients is enough. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding()]
param(
    [string]$WaveName,

    [string]$CsvPath,

    [string]$IdentityColumn = 'EmailAddress',

    [int]$StallHours = 6,

    [switch]$IncludeCompleted,

    [switch]$GridView,

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Output folder
# ---------------------------------------------------------------------------
$settingsPath = Join-Path $PSScriptRoot 'OnPremMigration.Settings.psd1'
$settings = @{}
if (Test-Path -LiteralPath $settingsPath)
{
    try { $settings = Import-PowerShellDataFile -LiteralPath $settingsPath } catch { $settings = @{} }
}
$outputRoot = if ($settings.OutputRoot) { [string]$settings.OutputRoot } else { $PSScriptRoot }

if (-not $OutputFolder)
{
    $OutputFolder = if ($WaveName) { Join-Path $outputRoot $WaveName } else { $PSScriptRoot }
}
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$runTime   = Get-Date
$label     = if ($WaveName) { $WaveName -replace '[^\w\-]', '_' }
             elseif ($CsvPath) { [System.IO.Path]::GetFileNameWithoutExtension($CsvPath) }
             else { 'AllMoveRequests' }

$logFile     = Join-Path $OutputFolder ("{0}_Status_{1}.log" -f $label, $timestamp)
$snapshotCsv = Join-Path $OutputFolder ("{0}_Status_{1}.csv" -f $label, $timestamp)
$historyCsv  = Join-Path $OutputFolder 'MigrationStatus-History.csv'

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

Write-Log "On-prem migration status - $label - $timestamp" -Level Head

# ---------------------------------------------------------------------------
# Environment
# ---------------------------------------------------------------------------
if (-not (Get-Command Get-MoveRequest -ErrorAction SilentlyContinue))
{
    throw "Get-MoveRequest not available. Run this from the on-premises Exchange Management Shell."
}

# ---------------------------------------------------------------------------
# Build the move-request list
# ---------------------------------------------------------------------------
if ($WaveName)
{
    $requests = Get-MoveRequest -BatchName $WaveName -ResultSize Unlimited -ErrorAction SilentlyContinue
    if (-not $requests) { Write-Log "No move requests with BatchName '$WaveName'." -Level Warn; return }
}
elseif ($CsvPath)
{
    if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV '$CsvPath' not found." }
    $rows = Import-Csv -LiteralPath $CsvPath
    if (-not ($rows | Get-Member -Name $IdentityColumn -MemberType NoteProperty))
    {
        throw "CSV '$CsvPath' has no '$IdentityColumn' column."
    }
    $wanted = $rows.$IdentityColumn | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }
    $requests = foreach ($w in $wanted)
    {
        $mr = Get-MoveRequest -Identity $w -ErrorAction SilentlyContinue
        if ($mr) { $mr } else { Write-Log "No move request for '$w'." -Level Warn }
    }
}
else
{
    $requests = Get-MoveRequest -ResultSize Unlimited
}

$requests = @($requests | Where-Object { $_ })
if (-not $requests) { Write-Log "Nothing in scope." -Level Warn; return }
Write-Log "$($requests.Count) move request(s) in scope." -Level Info

# ---------------------------------------------------------------------------
# Previous run (for stall detection)
# ---------------------------------------------------------------------------
$previous = @{}
if (Test-Path -LiteralPath $historyCsv)
{
    $hist = Import-Csv -LiteralPath $historyCsv
    foreach ($g in ($hist | Group-Object Identity))
    {
        $previous[$g.Name.ToLower()] = $g.Group | Sort-Object { [datetime]$_.Timestamp } | Select-Object -Last 1
    }
}

# ---------------------------------------------------------------------------
# Per-mailbox snapshot
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Per-mailbox status" -Level Head

$n = 0
$snapshot = foreach ($mr in $requests)
{
    $n++
    $id = [string]$mr.Identity
    if (-not $id) { $id = [string]$mr.DisplayName }

    $st = Get-MoveRequestStatistics -Identity $mr.Guid.ToString() -ErrorAction SilentlyContinue

    $status       = [string]$mr.Status
    $detail       = if ($st) { [string]$st.StatusDetail } else { [string]$mr.StatusDetail }
    $percent      = if ($st -and $null -ne $st.PercentComplete) { [int]$st.PercentComplete } else { $null }
    $itemsDone    = if ($st) { [int64]$st.ItemsTransferred } else { $null }
    $bytesDone    = if ($st -and $st.BytesTransferred) { [int64]($st.BytesTransferred.ToString() -replace '.*\(| bytes\)|,', '') } else { $null }
    $sourceDb     = if ($st) { [string]$st.SourceDatabase } else { [string]$mr.SourceDatabase }
    $targetDb     = if ($st) { [string]$st.TargetDatabase } else { [string]$mr.TargetDatabase }
    $lastUpdate   = if ($st) { $st.LastUpdateTimestamp } else { $null }
    $stalledSince = if ($st) { $st.StalledSinceTimestamp } else { $null }
    $failureType  = if ($st) { [string]$st.FailureType } else { [string]$mr.FailureType }
    $message      = if ($st -and $st.Message) { [string]$st.Message } else { [string]$mr.Message }
    $badItems     = if ($st) { [int]$st.BadItemsEncountered } else { $null }
    $largeItems   = if ($st) { [int]$st.LargeItemsEncountered } else { $null }

    # classify
    $class = switch -Regex ($status)
    {
        'CompletedWithWarning|^Completed' { 'Completed'; break }
        'CompletionInProgress'            { 'Completing'; break }
        'CompletionFailed|^Failed'        { 'Failed'; break }
        'AutoSuspended'                   { 'Synced'; break }
        '^Suspended'                      { 'Suspended'; break }
        '^Queued'                         { 'Queued'; break }
        '^InProgress'
        {
            if ($detail -match 'CreatingFolderHierarchy|CreatingInitialSyncCheckpoint|LoadingMessages|InitialSeeding|Queued')
            { 'Provisioning' } else { 'Syncing' }
            break
        }
        default { 'Provisioning' }
    }

    if ($class -eq 'Failed' -and ($failureType -match 'TooMany(Bad|Large)Items' -or $message -match 'bad items|corrupt items|large items|allowed limit'))
    {
        $class = 'CorruptItems'
    }

    # stall detection
    if ($class -eq 'Syncing')
    {
        $stall = $false
        if ($stalledSince) { $stall = $true }
        else
        {
            $prev = $previous[$id.ToLower()]
            if ($prev)
            {
                $same = ([string]$prev.PercentComplete -eq [string]$percent) -and ([string]$prev.ItemsTransferred -eq [string]$itemsDone)
                $old  = (-not $lastUpdate) -or ($lastUpdate -lt $runTime.AddHours(-$StallHours))
                if ($same -and $old) { $stall = $true }
            }
        }
        if ($stall) { $class = 'Stalled' }
    }

    $level = switch ($class)
    {
        { $_ -in 'Completed', 'Synced' }                   { 'Ok' }
        { $_ -in 'Failed', 'Stalled', 'CorruptItems' }     { 'Fail' }
        'Suspended'                                         { 'Warn' }
        default                                             { 'Info' }
    }
    Write-Log ("[{0,4}] {1,-42} {2,-13} {3,3}%  {4} -> {5}  {6}" -f `
        $n, $id, $class, $percent, $sourceDb, $targetDb, (($message -split "`r?`n")[0])) -Level $level

    [PSCustomObject]@{
        Timestamp           = $runTime.ToString('s')
        Index               = $n
        Identity            = $id
        Wave                = [string]$mr.BatchName
        Class               = $class
        Status              = $status
        StatusDetail        = $detail
        PercentComplete     = $percent
        ItemsTransferred    = $itemsDone
        BytesTransferred    = $bytesDone
        SourceDatabase      = $sourceDb
        TargetDatabase      = $targetDb
        StalledSince        = if ($stalledSince) { $stalledSince.ToString('s') } else { $null }
        LastUpdateTimestamp = if ($lastUpdate) { $lastUpdate.ToString('s') } else { $null }
        BadItemsEncountered = $badItems
        LargeItemsEncountered = $largeItems
        FailureType         = $failureType
        Message             = ($message -split "`r?`n")[0]
    }
}

$snapshot = @($snapshot)

# ---------------------------------------------------------------------------
# Output + history + grid
# ---------------------------------------------------------------------------
$view = if ($IncludeCompleted) { $snapshot } else { $snapshot | Where-Object Class -ne 'Completed' }
$view | Sort-Object Class, Identity |
    Format-Table Index, Identity, Class, PercentComplete, ItemsTransferred, SourceDatabase, TargetDatabase, LastUpdateTimestamp -AutoSize |
    Out-Host

$snapshot | Export-Csv -LiteralPath $snapshotCsv -NoTypeInformation -Encoding UTF8
$snapshot |
    Select-Object Timestamp, Identity, Wave, Class, Status, PercentComplete, ItemsTransferred, BytesTransferred, TargetDatabase, LastUpdateTimestamp, Message |
    Export-Csv -LiteralPath $historyCsv -NoTypeInformation -Encoding UTF8 -Append

if ($GridView)
{
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue)
    {
        $snapshot | Out-GridView -Title ("On-prem migration status - {0} - {1}" -f $label, $timestamp)
    }
    else
    {
        Write-Log "Out-GridView is not available on this host - skipping -GridView." -Level Warn
    }
}

# ---------------------------------------------------------------------------
# Summaries
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Class summary" -Level Head
$snapshot | Group-Object Class | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0,-13} {1}" -f $_.Name, $_.Count) -Level Info
}

Write-Log "" -Level Head
Write-Log "Target-database landing (where mailboxes are going)" -Level Head
$snapshot | Where-Object TargetDatabase | Group-Object TargetDatabase | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0,-30} {1}" -f $_.Name, $_.Count) -Level Info
}

$problems = $snapshot | Where-Object Class -in 'Failed', 'Stalled', 'CorruptItems'
if ($problems)
{
    Write-Log "" -Level Head
    Write-Log "Failure groups" -Level Head
    $problems |
        Group-Object { if ($_.FailureType) { $_.FailureType } else { ($_.Message -split "`r?`n")[0].Trim() } } |
        Sort-Object Count -Descending |
        ForEach-Object {
            Write-Log ("  x{0}  {1}" -f $_.Count, $_.Name) -Level Fail
            $_.Group | ForEach-Object { Write-Log ("        - {0}" -f $_.Identity) -Level Info }
        }
    Write-Log "" -Level Info
    Write-Log "Deep-dive: .\04-Get-OnPremMigrationFailureReport.ps1 -WaveName '$label'" -Level Warn
}

# rough throughput
$active = $snapshot | Where-Object { $_.Class -in 'Syncing', 'Provisioning', 'Completing' -and $_.BytesTransferred }
if ($active)
{
    $gb = [math]::Round((($active | Measure-Object BytesTransferred -Sum).Sum) / 1GB, 1)
    Write-Log "" -Level Head
    Write-Log ("In-flight: {0} mailbox(es), {1} GB transferred so far." -f @($active).Count, $gb) -Level Info
}

Write-Log "" -Level Head
Write-Log "Snapshot : $snapshotCsv" -Level Info
Write-Log "History  : $historyCsv" -Level Info
Write-Log "Log      : $logFile" -Level Info
