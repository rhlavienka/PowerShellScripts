<#
.SYNOPSIS
    Builds the next migration batch from the mailboxes that still sit on databases
    excluded from provisioning, runs the source-side readiness checks, and writes
    the Ready CSV that 02-New-OnPremMoveRequest.ps1 consumes.

.DESCRIPTION
    Alternative step 1 of the OnpremMailboxMigration toolset (sibling of
    01-Get-OnPremMigrationScope.ps1). Instead of selecting by OU / City /
    Department it drains databases: every mailbox whose primary database has
    IsExcludedFromProvisioning = $true is a candidate. Exchange's automatic
    mailbox distribution never picks an excluded database as a target, so the
    moved mailboxes land on the databases that are still open for provisioning.

    One run creates ONE batch of at most -BatchSize mailboxes. Run it again for
    the next batch - mailboxes that already belong to another batch are skipped:

      - a move request exists for the mailbox (any status, any BatchName) -
        Get-MoveRequest -ResultSize Unlimited is read once at start
      - the mailbox is listed in a Ready CSV (<name>_Ready_*.csv) under
        -OutputFolder that has not been submitted yet, i.e. no
        <name>_MoveRequests_*.csv in the same folder reports it as "Created".
        Once 02 has created the move request, the live move request is what
        counts; if that request is later removed, the mailbox becomes eligible
        again. To release a prepared but abandoned batch, delete its folder.

    Candidates are walked in -SortBy order - by default spread round-robin
    over the source servers and databases so one batch does not hammer a
    single database or server (see -SortBy Spread); each one gets the same readiness
    checks as 01 (movable type, no pending Restore / Import / Export request,
    ExchangeGuid stamped, size / item-count WARN, archive, holds). FAIL rows are
    reported and skipped, and the batch is topped up with the next candidate
    until -BatchSize PASS/WARN mailboxes are collected or candidates run out.

    Output (all in the batch folder <OutputFolder>\<BatchName>\):

      <BatchName>_BatchReport_<ts>.csv   every mailbox checked in this run
                                         (batch members + FAIL rows)
      <BatchName>_Ready_<ts>.csv         EmailAddress column, batch members only;
                                         optional TargetDatabase /
                                         TargetArchiveDatabase columns
      <BatchName>_BatchReport_<ts>.log   console transcript

.PARAMETER BatchSize
    Maximum number of mailboxes in the new batch (PASS + WARN rows).

.PARAMETER BatchName
    Name of the batch. Becomes the batch sub-folder name and the file-name
    prefix, and 02 passes it to New-MoveRequest -BatchName. Max 64 characters.
    Default: the current date and time, "yyyyMMdd_HHmmss". The script refuses a
    name that is already used by a move request or by an existing Ready CSV.

.PARAMETER Database
    Optional. Restrict the candidates to these databases (each must be excluded
    from provisioning). Default: all databases with IsExcludedFromProvisioning.

.PARAMETER SortBy
    Order in which candidates fill the batch:
      Spread          - spread the read load over the source side (default):
                        round-robin over the source servers (the server that
                        hosts the active copy of each database), within each
                        server round-robin over its databases, mailboxes in
                        random order inside each database. A batch of N from
                        S servers therefore takes ~N/S mailboxes per server,
                        split evenly across that server's databases; a
                        server or database that runs out simply drops out of
                        the rotation.
      Database        - database name, then display name
      SizeAscending   - smallest mailboxes first
      SizeDescending  - largest mailboxes first

.PARAMETER RandomSeed
    Optional seed for -SortBy Spread. The same seed over the same candidates
    gives the same order (reproducible dry runs). Default: a new random order
    on every run.

.PARAMETER TargetDatabase
    Optional. When given, a "TargetDatabase" column with this single value is
    written to every Ready-CSV row. It must not be excluded from provisioning.
    When omitted Exchange picks and load-balances the target itself.

.PARAMETER IncludeArchive
    Also emit a "TargetArchiveDatabase" column in the Ready CSV (value =
    -TargetArchiveDatabase, or blank for automatic distribution).

.PARAMETER TargetArchiveDatabase
    Value for the "TargetArchiveDatabase" column when -IncludeArchive is set.

.PARAMETER MaxMailboxSizeGB
    Primary mailboxes larger than this raise a WARN. Default 50.

.PARAMETER MaxItemCount
    Primary mailboxes with more items than this raise a WARN. Default 200000.

.PARAMETER OutputFolder
    Parent folder; the batch folder <OutputFolder>\<BatchName>\ is created under
    it, and existing Ready CSVs are searched under it. Default: the script
    folder, or OutputRoot from OnPremMigration.Settings.psd1 if that file exists
    next to the script.

.EXAMPLE
    .\01b-New-OnPremMigrationBatch.ps1 -BatchSize 50

    Next 50 mailboxes from all excluded databases, spread across the source
    servers and databases, batch named e.g. 20260928_141500.

.EXAMPLE
    .\01b-New-OnPremMigrationBatch.ps1 -BatchSize 100 -BatchName "DB01-Drain-03" `
        -Database "DB01" -SortBy SizeAscending

.NOTES
    Version: 1.1 (2026-09-28)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer).
              View-Only Recipients is enough. No Exchange Online / Graph modules.

    Changelog:
    1.1 (2026-09-28) - -SortBy Spread (new default): round-robin over source
                       servers and databases, random order inside a database;
                       -RandomSeed; SourceServer column and per-server summary.
    1.0 (2026-09-28) - Initial version.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateRange(1, 100000)]
    [int]$BatchSize,

    [string]$BatchName = (Get-Date -Format 'yyyyMMdd_HHmmss'),

    [string[]]$Database,

    [ValidateSet('Spread', 'Database', 'SizeAscending', 'SizeDescending')]
    [string]$SortBy = 'Spread',

    [int]$RandomSeed,

    [string]$TargetDatabase,

    [switch]$IncludeArchive,

    [string]$TargetArchiveDatabase,

    [int]$MaxMailboxSizeGB = 50,

    [int]$MaxItemCount = 200000,

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Batch name validation
# ---------------------------------------------------------------------------
$BatchName = $BatchName.Trim()
if (-not $BatchName) { throw "BatchName is empty." }
if ($BatchName.Length -gt 64) { throw "BatchName '$BatchName' is longer than 64 characters (New-MoveRequest -BatchName limit)." }
if ($BatchName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0)
{
    throw "BatchName '$BatchName' contains characters that are not valid in a folder name."
}
if ($BatchName -match '_Ready_') { throw "BatchName must not contain '_Ready_' (02 derives the batch name from the CSV file name)." }

# ---------------------------------------------------------------------------
# Settings file (optional) + output folder
# ---------------------------------------------------------------------------
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
if ($settings.MaxMailboxSizeGB -and -not $PSBoundParameters.ContainsKey('MaxMailboxSizeGB')) { $MaxMailboxSizeGB = [int]$settings.MaxMailboxSizeGB }

$batchFolder = Join-Path $OutputFolder $BatchName
if ((Test-Path -LiteralPath $batchFolder) -and (Get-ChildItem -LiteralPath $batchFolder -Filter '*_Ready_*.csv' -File -ErrorAction SilentlyContinue))
{
    throw "Batch folder '$batchFolder' already contains a Ready CSV. Choose another -BatchName."
}

# Nothing is written before the environment and the name checks pass.
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    throw "Get-ExchangeServer not available. Run this from the on-premises Exchange Management Shell."
}

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
if (-not (Test-Path -LiteralPath $batchFolder)) { New-Item -ItemType Directory -Path $batchFolder -Force | Out-Null }

$logFile    = Join-Path $batchFolder ("{0}_BatchReport_{1}.log" -f $BatchName, $timestamp)
$reportFile = Join-Path $batchFolder ("{0}_BatchReport_{1}.csv" -f $BatchName, $timestamp)
$readyFile  = Join-Path $batchFolder ("{0}_Ready_{1}.csv"       -f $BatchName, $timestamp)

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

function Get-SizeGB
{
    param($TotalItemSize)
    if (-not $TotalItemSize) { return $null }
    $text = $TotalItemSize.ToString()
    if ($text -match '\(([\d,]+) bytes\)')
    {
        return [math]::Round(([double]($matches[1] -replace ',', '')) / 1GB, 2)
    }
    return $null
}

function Get-SmtpAddressList
{
    # All SMTP proxy addresses of a mailbox, lower-case, without the prefix.
    param($Mailbox)
    $list = @($Mailbox.EmailAddresses | ForEach-Object { [string]$_ } |
        Where-Object { $_ -match '^smtp:' } | ForEach-Object { ($_ -replace '^smtp:', '').ToLowerInvariant() })
    $primary = ([string]$Mailbox.PrimarySmtpAddress).ToLowerInvariant()
    if ($primary -and $list -notcontains $primary) { $list += $primary }
    $list
}

Write-Log "On-prem migration batch - '$BatchName' (size $BatchSize) - $timestamp" -Level Head
Write-Log "On-premises Exchange session confirmed." -Level Ok

$movableTypes = 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'

# ---------------------------------------------------------------------------
# Source databases (IsExcludedFromProvisioning)
# ---------------------------------------------------------------------------
$allDbs      = @(Get-MailboxDatabase)
$excludedDbs = @($allDbs | Where-Object { $_.IsExcludedFromProvisioning -eq $true })

if ($Database)
{
    foreach ($d in $Database)
    {
        $match = $allDbs | Where-Object { $_.Name -eq $d }
        if (-not $match) { throw "Database '$d' not found." }
        if (-not $match.IsExcludedFromProvisioning) { throw "Database '$d' is not excluded from provisioning - it is not a drain source." }
    }
    $excludedDbs = @($excludedDbs | Where-Object { $Database -contains $_.Name })
}

if (-not $excludedDbs)
{
    Write-Log "No mailbox database has IsExcludedFromProvisioning = `$true. Nothing to batch." -Level Warn
    return
}
Write-Log ("Source databases (excluded from provisioning): {0}" -f (($excludedDbs.Name | Sort-Object) -join ', ')) -Level Info

if ($TargetDatabase)
{
    $tdb = $allDbs | Where-Object { $_.Name -eq $TargetDatabase }
    if (-not $tdb) { throw "Target database '$TargetDatabase' not found." }
    if ($tdb.IsExcludedFromProvisioning) { throw "Target database '$TargetDatabase' is itself excluded from provisioning." }
}

# ---------------------------------------------------------------------------
# Existing batch membership
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Existing batch membership" -Level Head

# live move requests, keyed by ExchangeGuid
$moveByGuid = @{}
$liveBatchNames = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
foreach ($mr in @(Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue))
{
    $bn = [string]$mr.BatchName
    if ($bn) { [void]$liveBatchNames.Add($bn) }
    $moveByGuid[[string]$mr.ExchangeGuid] = [PSCustomObject]@{ BatchName = $bn; Status = [string]$mr.Status }
}
Write-Log "$($moveByGuid.Count) existing move request(s) in $($liveBatchNames.Count) batch(es)." -Level Info

if ($liveBatchNames.Contains($BatchName))
{
    throw "BatchName '$BatchName' is already used by existing move requests. Choose another -BatchName."
}

# prepared (not yet submitted) Ready CSVs under OutputFolder
$reservedByAddress = @{}
$readyCsvs = @(Get-ChildItem -LiteralPath $OutputFolder -Filter '*_Ready_*.csv' -File -Recurse -ErrorAction SilentlyContinue)
foreach ($csv in $readyCsvs)
{
    $prepName = if ($csv.BaseName -match '^(.*?)_Ready_') { $matches[1] } else { $csv.BaseName }

    # addresses 02 already created a move request for (tracked by Get-MoveRequest from then on)
    $submitted = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    foreach ($mrCsv in @(Get-ChildItem -LiteralPath $csv.DirectoryName -Filter '*_MoveRequests_*.csv' -File -ErrorAction SilentlyContinue))
    {
        try
        {
            Import-Csv -LiteralPath $mrCsv.FullName | Where-Object { $_.Action -eq 'Created' -and $_.Address } |
                ForEach-Object { [void]$submitted.Add(([string]$_.Address).Trim()) }
        }
        catch { Write-Log "Cannot read '$($mrCsv.FullName)': $($_.Exception.Message)" -Level Warn }
    }

    try { $rows = @(Import-Csv -LiteralPath $csv.FullName) }
    catch { Write-Log "Cannot read '$($csv.FullName)': $($_.Exception.Message)" -Level Warn; continue }

    foreach ($row in $rows)
    {
        $addr = ([string]$row.EmailAddress).Trim()
        if (-not $addr -or $submitted.Contains($addr)) { continue }
        $key = $addr.ToLowerInvariant()
        if (-not $reservedByAddress.ContainsKey($key)) { $reservedByAddress[$key] = $prepName }
    }
}
Write-Log "$($reservedByAddress.Count) mailbox(es) reserved by $($readyCsvs.Count) prepared Ready CSV(s) not yet submitted." -Level Info

# ---------------------------------------------------------------------------
# Candidates
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Candidates" -Level Head

$candidates = New-Object System.Collections.Generic.List[object]
foreach ($db in $excludedDbs)
{
    $dbName   = [string]$db.Name
    $dbServer = [string]$db.Server    # server hosting the active copy = where the move reads from
    $mbxs   = @(Get-Mailbox -Database $dbName -ResultSize Unlimited -ErrorAction SilentlyContinue)

    # one statistics call per database instead of one per mailbox
    $statsByGuid = @{}
    foreach ($s in @(Get-MailboxStatistics -Database $dbName -ErrorAction SilentlyContinue))
    {
        if (-not $s.DisconnectDate) { $statsByGuid[[string]$s.MailboxGuid] = $s }
    }

    Write-Log ("  {0,-30} {1,-20} {2} mailbox(es)" -f $dbName, $dbServer, $mbxs.Count) -Level Info
    foreach ($m in $mbxs)
    {
        $st = $statsByGuid[[string]$m.ExchangeGuid]
        $candidates.Add([PSCustomObject]@{
                Mailbox  = $m
                Stats    = $st
                Database = $dbName
                Server   = $dbServer
                SizeGB   = if ($st) { Get-SizeGB $st.TotalItemSize } else { $null }
            })
    }
}

$rng = if ($PSBoundParameters.ContainsKey('RandomSeed')) { [System.Random]::new($RandomSeed) } else { [System.Random]::new() }

function Get-Shuffled
{
    # Fisher-Yates shuffle; returns the array as one object so callers keep it intact
    param([object[]]$Items)
    $a = @($Items)
    for ($i = $a.Count - 1; $i -gt 0; $i--)
    {
        $j = $rng.Next($i + 1)
        $a[$i], $a[$j] = $a[$j], $a[$i]
    }
    , $a
}

function Get-SpreadOrder
{
    # server ring -> database queues -> shuffled mailboxes; take one mailbox per
    # server per round, rotating through that server's databases
    param([object[]]$Items)
    $rings = [System.Collections.Generic.List[object]]::new()
    foreach ($sg in (Get-Shuffled @($Items | Group-Object Server)))
    {
        $queues = [System.Collections.Generic.List[object]]::new()
        foreach ($dg in (Get-Shuffled @($sg.Group | Group-Object Database)))
        {
            $queues.Add([System.Collections.Generic.Queue[object]]::new([object[]](Get-Shuffled $dg.Group)))
        }
        $rings.Add([PSCustomObject]@{ Queues = $queues; Next = 0 })
    }

    $out = [System.Collections.Generic.List[object]]::new()
    while ($rings.Count)
    {
        $s = 0
        while ($s -lt $rings.Count)
        {
            $ring = $rings[$s]
            $q    = $ring.Queues[$ring.Next]
            $out.Add($q.Dequeue())
            if ($q.Count -eq 0) { $ring.Queues.RemoveAt($ring.Next) } else { $ring.Next++ }
            if ($ring.Queues.Count -eq 0) { $rings.RemoveAt($s); continue }
            $ring.Next = $ring.Next % $ring.Queues.Count
            $s++
        }
    }
    $out
}

$ordered = switch ($SortBy)
{
    'Spread'         { Get-SpreadOrder $candidates }
    'SizeAscending'  { $candidates | Sort-Object { [double]$_.SizeGB }, { [string]$_.Mailbox.DisplayName } }
    'SizeDescending' { $candidates | Sort-Object @{ Expression = { [double]$_.SizeGB }; Descending = $true }, @{ Expression = { [string]$_.Mailbox.DisplayName } } }
    default          { $candidates | Sort-Object Database, { [string]$_.Mailbox.DisplayName } }
}
$ordered = @($ordered)
Write-Log ("{0} candidate mailbox(es) on excluded databases, order: {1}{2}." -f $ordered.Count, $SortBy,
    $(if ($SortBy -eq 'Spread' -and $PSBoundParameters.ContainsKey('RandomSeed')) { " (seed $RandomSeed)" } else { '' })) -Level Info

if (-not $ordered) { Write-Log "No mailboxes left on the source databases." -Level Ok; return }

# ---------------------------------------------------------------------------
# Fill the batch
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Readiness checks" -Level Head

$inOtherBatch = @{}     # batch name -> count, for the summary
$results  = New-Object System.Collections.Generic.List[object]
$members  = 0
$n        = 0

foreach ($c in $ordered)
{
    if ($members -ge $BatchSize) { break }

    $mbx       = $c.Mailbox
    $primary   = [string]$mbx.PrimarySmtpAddress
    $rtd       = [string]$mbx.RecipientTypeDetails
    $currentDb = $c.Database

    # already part of another batch?
    $other = $null; $otherDetail = $null
    $mr = $moveByGuid[[string]$mbx.ExchangeGuid]
    if ($mr)
    {
        $other       = if ($mr.BatchName) { $mr.BatchName } else { '(move request without batch name)' }
        $otherDetail = "move request $($mr.Status)"
    }
    else
    {
        foreach ($a in (Get-SmtpAddressList $mbx))
        {
            if ($reservedByAddress.ContainsKey($a)) { $other = $reservedByAddress[$a]; $otherDetail = 'prepared, not submitted'; break }
        }
    }
    if ($other)
    {
        $inOtherBatch[$other] = 1 + [int]$inOtherBatch[$other]
        Write-Verbose "$primary skipped - already in batch '$other' ($otherDetail)"
        continue
    }

    $n++
    $fail = [System.Collections.Generic.List[string]]::new()
    $warn = [System.Collections.Generic.List[string]]::new()

    if ($movableTypes -notcontains $rtd)
    {
        $fail.Add("RecipientTypeDetails '$rtd' is not a movable mailbox type")
    }

    # pending non-move requests
    $existing = $null
    foreach ($pair in @(
            @{ Cmd = 'Get-MailboxRestoreRequest'; Label = 'RestoreRequest' },
            @{ Cmd = 'Get-MailboxImportRequest';  Label = 'ImportRequest' },
            @{ Cmd = 'Get-MailboxExportRequest';  Label = 'ExportRequest' }))
    {
        if (Get-Command $pair.Cmd -ErrorAction SilentlyContinue)
        {
            $pending = & $pair.Cmd -Mailbox $mbx.Guid.ToString() -ErrorAction SilentlyContinue |
                Where-Object { [string]$_.Status -notin 'Completed', 'Failed' }
            if ($pending)
            {
                $warn.Add("pending $($pair.Label) ($([string]@($pending)[0].Status)) - it will block completion")
                if (-not $existing) { $existing = $pair.Label }
            }
        }
    }

    # ExchangeGuid
    $eg = [guid]::Empty
    [void][guid]::TryParse([string]$mbx.ExchangeGuid, [ref]$eg)
    if ($eg -eq [guid]::Empty) { $fail.Add("ExchangeGuid is empty (mailbox not provisioned)") }

    # statistics
    $sizeGB = $c.SizeGB; $items = $null
    if ($c.Stats)
    {
        $items = [int]$c.Stats.ItemCount
        if ($sizeGB -and $sizeGB -gt $MaxMailboxSizeGB) { $warn.Add("size ${sizeGB} GB > ${MaxMailboxSizeGB} GB - slow move") }
        if ($items  -and $items  -gt $MaxItemCount)     { $warn.Add("$items items > $MaxItemCount - slow move") }
    }
    else
    {
        $warn.Add("Get-MailboxStatistics returned nothing (never logged on?)")
    }

    # archive
    $archiveState = 'None'; $archiveDb = $null
    if ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [guid]::Empty)
    {
        $archiveDb = [string]$mbx.ArchiveDatabase
        $aStats = Get-MailboxStatistics -Identity $mbx.Guid.ToString() -Archive -ErrorAction SilentlyContinue
        $aSize  = if ($aStats) { Get-SizeGB $aStats.TotalItemSize } else { $null }
        $archiveState = if ($aSize) { "Present ($aSize GB)" } else { 'Present' }
        if (-not $IncludeArchive)
        {
            $warn.Add("mailbox has an archive on '$archiveDb' - it stays put unless you run 01b with -IncludeArchive")
        }
    }

    # holds
    $holds = [System.Collections.Generic.List[string]]::new()
    if ($mbx.LitigationHoldEnabled)                            { $holds.Add('Litigation') }
    if ($mbx.ComplianceTagHoldApplied)                         { $holds.Add('ComplianceTag') }
    if ($mbx.DelayHoldApplied -or $mbx.DelayReleaseHoldApplied) { $holds.Add('DelayHold') }
    if (@($mbx.InPlaceHolds).Count -gt 0)                      { $holds.Add('InPlaceHold') }
    if ($holds.Count) { $warn.Add("on hold: $($holds -join ',') (move works, recoverable-items space counts)") }

    $status = if ($fail.Count) { 'FAIL' } elseif ($warn.Count) { 'WARN' } else { 'PASS' }
    if ($status -ne 'FAIL') { $members++ }
    switch ($status)
    {
        'FAIL' { Write-Log ("[{0,4}] {1,-45} FAIL - {2}" -f $n, $primary, ($fail -join ' | ')) -Level Fail }
        'WARN' { Write-Log ("[{0,4}] {1,-45} WARN - {2}" -f $n, $primary, ($warn -join ' | ')) -Level Warn }
        'PASS' { Write-Log ("[{0,4}] {1,-45} ready ({2}, {3} GB)" -f $n, $primary, $currentDb, $sizeGB) -Level Ok }
    }

    $results.Add([PSCustomObject]@{
            Index                = $n
            EmailAddress         = $primary
            DisplayName          = [string]$mbx.DisplayName
            RecipientTypeDetails = $rtd
            PrimarySmtpAddress   = $primary
            CurrentDatabase      = $currentDb
            SourceServer         = $c.Server
            MailboxSizeGB        = $sizeGB
            ItemCount            = $items
            ArchiveState         = $archiveState
            ArchiveDatabase      = $archiveDb
            Holds                = ($holds -join ',')
            ExistingRequest      = $existing
            InBatch              = ($status -ne 'FAIL')
            Status               = $status
            Issues               = (@($fail) + @($warn)) -join ' | '
        })
}

# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------
if ($results.Count) { $results | Export-Csv -LiteralPath $reportFile -NoTypeInformation -Encoding UTF8 }

$ready = @($results | Where-Object InBatch)
if ($ready.Count)
{
    $ready | ForEach-Object {
        $row = [ordered]@{ EmailAddress = $_.EmailAddress }
        if ($PSBoundParameters.ContainsKey('TargetDatabase')) { $row.TargetDatabase = $TargetDatabase }
        if ($IncludeArchive) { $row.TargetArchiveDatabase = $TargetArchiveDatabase }
        [PSCustomObject]$row
    } | Export-Csv -LiteralPath $readyFile -NoTypeInformation -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$pass  = @($results | Where-Object Status -eq 'PASS').Count
$warnC = @($results | Where-Object Status -eq 'WARN').Count
$failC = @($results | Where-Object Status -eq 'FAIL').Count
$skipC = ($inOtherBatch.Values | Measure-Object -Sum).Sum
$left  = $ordered.Count - [int]$skipC - $results.Count

Write-Log "" -Level Head
Write-Log "Summary: batch '$BatchName' has $($ready.Count)/$BatchSize mailbox(es) ($pass PASS, $warnC WARN); $failC FAIL skipped." -Level Head
if ($ready.Count -lt $BatchSize) { Write-Log "Candidates ran out before the batch was full." -Level Warn }

if ($inOtherBatch.Count)
{
    Write-Log "" -Level Head
    Write-Log "Skipped - already part of another batch ($skipC)" -Level Head
    $inOtherBatch.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Log ("  {0,-40} {1}" -f $_.Name, $_.Value) -Level Info
    }
}
Write-Log "Candidates not examined in this run (left for later batches): $([math]::Max(0, $left))" -Level Info

if ($ready.Count)
{
    Write-Log "" -Level Head
    Write-Log "Batch members by source server / database" -Level Head
    $ready | Group-Object SourceServer | Sort-Object Name | ForEach-Object {
        $gb = ($_.Group | Measure-Object -Property MailboxSizeGB -Sum).Sum
        Write-Log ("  {0,-30} {1,5} mailbox(es) {2,10:N1} GB" -f $_.Name, $_.Count, [double]$gb) -Level Info
        $_.Group | Group-Object CurrentDatabase | Sort-Object Name | ForEach-Object {
            Write-Log ("    {0,-28} {1,5}" -f $_.Name, $_.Count) -Level Info
        }
    }
}

Write-Log "" -Level Head
if ($results.Count) { Write-Log "Batch report: $reportFile" -Level Info }
if ($ready.Count)   { Write-Log "Ready CSV   : $readyFile  (feed this to 02-New-OnPremMoveRequest.ps1)" -Level Info }
else                { Write-Log "No batch members - no Ready CSV written." -Level Warn }
Write-Log "Log         : $logFile" -Level Info

if ($failC)     { Write-Log "FAIL rows are not in the batch; fix them and they are picked up by a later batch." -Level Fail }
elseif ($warnC) { Write-Log "Review the WARN rows, then run 02." -Level Warn }
elseif ($ready.Count) { Write-Log "All checks passed." -Level Ok }
