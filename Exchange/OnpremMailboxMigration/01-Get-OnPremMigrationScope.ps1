<#
.SYNOPSIS
    Selects the mailboxes for an on-premises database-to-database migration wave
    from directory attributes (OU / City / Department), runs the source-side
    readiness checks, and writes the CSV that 02-New-OnPremMoveRequest.ps1 consumes.

.DESCRIPTION
    This is step 1 of the OnpremMailboxMigration toolset - moves BETWEEN mailbox
    databases inside ONE on-premises Exchange organization (consolidation, storage
    refresh, new DAG, retiring a database). Nothing here touches Exchange Online.

    Selection is server-side wherever possible: an OPATH -Filter on City /
    Department plus -OrganizationalUnit is handed to Get-Recipient, so the query
    stays fast in a large directory. -Identity / -CsvPath bypass the attribute
    filter and check an explicit list instead.

    For every selected recipient the script checks the things that make a local
    move request fail, stall, or land somewhere unexpected:

      - the recipient is a movable mailbox type (UserMailbox / SharedMailbox /
        RoomMailbox / EquipmentMailbox) - anything else FAILs
      - no move request already exists for the mailbox (Get-MoveRequest)
      - no pending Restore / MailboxImport / MailboxExport request
      - ExchangeGuid is stamped (an all-zero GUID means MRS cannot match it)
      - the current database is not in -ExcludeDatabase / not already a target
      - Get-MailboxStatistics: mailbox is not disconnected; size and item count
        are within -MaxMailboxSizeGB / -MaxItemCount (WARN only - big mailboxes
        just take longer)
      - archive presence and its database (informational; drives -IncludeArchive)
      - holds (Litigation / ComplianceTag / DelayHold / InPlaceHold) -
        informational, recoverable-items space counts against the move

    Output (all in the wave folder <OutputFolder>\<WaveName>\):

      <WaveName>_ScopeReport_<ts>.csv   full per-mailbox result, every row
      <WaveName>_Ready_<ts>.csv         EmailAddress column, PASS/WARN rows only;
                                        optional TargetDatabase / TargetArchiveDatabase
                                        columns; split into _partNN when a wave
                                        limit is exceeded
      <WaveName>_ScopeReport_<ts>.log   console transcript

    The Ready CSV is the ONLY file 02 needs. Fix the FAIL rows, review the WARN
    rows, then run 02.

.PARAMETER OrganizationalUnit
    Restrict the selection to this OU subtree. Canonical
    ("contoso.com/Users/Sales") or distinguished name. Passed straight to
    Get-Recipient -OrganizationalUnit.

.PARAMETER City
    One or more values for the AD "l" (City) attribute. Multiple values are
    OR'd together. Combined with -Department using AND.

.PARAMETER Department
    One or more values for the Department attribute. Multiple values are OR'd
    together. Combined with -City using AND.

.PARAMETER RecipientTypeDetails
    Optional, no default. When given, narrows the Get-Recipient query to these
    types up front. When omitted, no type filter is applied and the readiness
    pass simply FAILs any non-movable recipient it encounters.

.PARAMETER ExcludeDatabase
    Mailboxes whose current database is in this list are dropped from the wave
    (already where you want them).

.PARAMETER Identity
    Explicit mailbox identities to check instead of an attribute query. Mutually
    exclusive with -CsvPath and the attribute filters.

.PARAMETER CsvPath
    CSV with a column named by -IdentityColumn (default "EmailAddress"); those
    identities are checked instead of an attribute query.

.PARAMETER IdentityColumn
    Identity column name in -CsvPath. Default "EmailAddress".

.PARAMETER TargetDatabase
    Optional. When given, a "TargetDatabase" column with this single value is
    written to every Ready-CSV row. When omitted the column is left out and
    Exchange picks and load-balances the target database itself (automatic
    mailbox distribution). The script never spreads mailboxes across databases.

.PARAMETER IncludeArchive
    Also emit a "TargetArchiveDatabase" column in the Ready CSV (value =
    -TargetArchiveDatabase, or blank for automatic distribution). Use it when the
    archive should move too.

.PARAMETER TargetArchiveDatabase
    Value for the "TargetArchiveDatabase" column when -IncludeArchive is set.

.PARAMETER MaxMailboxSizeGB
    Primary mailboxes larger than this raise a WARN. Default 50.

.PARAMETER MaxItemCount
    Primary mailboxes with more items than this raise a WARN. Default 200000.

.PARAMETER MaxWaveMailboxes
    When the PASS/WARN count exceeds this, the Ready CSV is split into
    <name>_part01, _part02, ... Default 500. 0 disables the split.

.PARAMETER MaxWaveSizeGB
    Same split, by summed primary-mailbox size. Default 2000. 0 disables it.

.PARAMETER WaveName
    Label for this wave. Becomes the wave sub-folder name and the file-name
    prefix, and 02 passes it to New-MoveRequest -BatchName. Default
    "Wave_<yyyyMMdd_HHmmss>".

.PARAMETER OutputFolder
    Parent folder; the wave folder <OutputFolder>\<WaveName>\ is created under it.
    Default: the script folder, or OutputRoot from OnPremMigration.Settings.psd1
    if that file exists next to the script.

.EXAMPLE
    .\01-Get-OnPremMigrationScope.ps1 -OrganizationalUnit "contoso.com/Users/Praha" `
        -City "Praha" -Department "Sales","Marketing" -WaveName "Praha-Sales-W1"

.EXAMPLE
    .\01-Get-OnPremMigrationScope.ps1 -CsvPath .\handpicked.csv -WaveName "Adhoc-2026-09"

.NOTES
    Version: 1.0 (2026-09-09)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer).
              View-Only Recipients is enough. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding(DefaultParameterSetName = 'Filter')]
param(
    [Parameter(ParameterSetName = 'Filter')]
    [string]$OrganizationalUnit,

    [Parameter(ParameterSetName = 'Filter')]
    [string[]]$City,

    [Parameter(ParameterSetName = 'Filter')]
    [string[]]$Department,

    [Parameter(ParameterSetName = 'Filter')]
    [string[]]$RecipientTypeDetails,

    [Parameter(Mandatory, ParameterSetName = 'Identity')]
    [string[]]$Identity,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Csv')]
    [string]$IdentityColumn = 'EmailAddress',

    [string[]]$ExcludeDatabase,

    [string]$TargetDatabase,

    [switch]$IncludeArchive,

    [string]$TargetArchiveDatabase,

    [int]$MaxMailboxSizeGB = 50,

    [int]$MaxItemCount = 200000,

    [int]$MaxWaveMailboxes = 500,

    [int]$MaxWaveSizeGB = 2000,

    [string]$WaveName = ("Wave_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss')),

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

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
if ($settings.MaxWaveMailboxes -and -not $PSBoundParameters.ContainsKey('MaxWaveMailboxes')) { $MaxWaveMailboxes = [int]$settings.MaxWaveMailboxes }

$timestamp  = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$waveFolder = Join-Path $OutputFolder $WaveName
if (-not (Test-Path -LiteralPath $waveFolder)) { New-Item -ItemType Directory -Path $waveFolder -Force | Out-Null }

$logFile    = Join-Path $waveFolder ("{0}_ScopeReport_{1}.log" -f $WaveName, $timestamp)
$reportFile = Join-Path $waveFolder ("{0}_ScopeReport_{1}.csv" -f $WaveName, $timestamp)
$readyBase  = Join-Path $waveFolder ("{0}_Ready_{1}"           -f $WaveName, $timestamp)

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

Write-Log "On-prem migration scope - wave '$WaveName' - $timestamp" -Level Head

# ---------------------------------------------------------------------------
# Environment
# ---------------------------------------------------------------------------
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
    throw "Get-ExchangeServer not available. Run this from the on-premises Exchange Management Shell."
}
Write-Log "On-premises Exchange session confirmed." -Level Ok

$movableTypes = 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'

# ---------------------------------------------------------------------------
# Build the recipient list
# ---------------------------------------------------------------------------
function Join-Or
{
    param([string[]]$Values, [string]$Property)
    if (-not $Values) { return $null }
    '(' + (($Values | ForEach-Object { "$Property -eq '$($_ -replace "'", "''")'" }) -join ' -or ') + ')'
}

$recipients = @()

switch ($PSCmdlet.ParameterSetName)
{
    'Identity'
    {
        Write-Log "Selection: explicit -Identity list ($($Identity.Count))." -Level Info
        $recipients = foreach ($i in $Identity)
        {
            $r = Get-Recipient -Identity $i -ErrorAction SilentlyContinue
            if ($r) { $r } else { Write-Log "Not found: '$i'." -Level Fail }
        }
    }
    'Csv'
    {
        if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV '$CsvPath' not found." }
        $rows = Import-Csv -LiteralPath $CsvPath
        if (-not ($rows | Get-Member -Name $IdentityColumn -MemberType NoteProperty))
        {
            throw "CSV '$CsvPath' has no '$IdentityColumn' column. Use -IdentityColumn."
        }
        $wanted = $rows.$IdentityColumn | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ } | Select-Object -Unique
        Write-Log "Selection: $($wanted.Count) identities from '$([System.IO.Path]::GetFileName($CsvPath))'." -Level Info
        $recipients = foreach ($w in $wanted)
        {
            $r = Get-Recipient -Identity $w -ErrorAction SilentlyContinue
            if ($r) { $r } else { Write-Log "Not found: '$w'." -Level Fail }
        }
    }
    'Filter'
    {
        $clauses = @(Join-Or $City 'City'; Join-Or $Department 'Department') | Where-Object { $_ }
        $filter  = if ($clauses) { $clauses -join ' -and ' } else { $null }

        if (-not $filter -and -not $OrganizationalUnit -and -not $RecipientTypeDetails)
        {
            throw "Refusing to select the whole organization. Supply at least one of -OrganizationalUnit / -City / -Department (or use -Identity / -CsvPath)."
        }

        $gr = @{ ResultSize = 'Unlimited' }
        if ($filter)               { $gr.Filter               = $filter }
        if ($OrganizationalUnit)   { $gr.OrganizationalUnit   = $OrganizationalUnit }
        if ($RecipientTypeDetails) { $gr.RecipientTypeDetails = $RecipientTypeDetails }

        Write-Log ("Selection: Get-Recipient{0}{1}{2}" -f `
            $(if ($OrganizationalUnit) { " -OrganizationalUnit '$OrganizationalUnit'" } else { '' }),
            $(if ($filter) { " -Filter `"$filter`"" } else { '' }),
            $(if ($RecipientTypeDetails) { " -RecipientTypeDetails $($RecipientTypeDetails -join ',')" } else { '' })) -Level Info

        $recipients = Get-Recipient @gr
    }
}

$recipients = @($recipients | Where-Object { $_ } | Sort-Object -Property Guid -Unique)
if (-not $recipients) { Write-Log "Nothing selected." -Level Warn; return }
Write-Log "$($recipients.Count) recipient(s) selected. Running readiness checks ..." -Level Info

# ---------------------------------------------------------------------------
# Per-mailbox readiness
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Readiness checks" -Level Head

$n = 0
$results = foreach ($rcp in $recipients)
{
    $n++
    $fail = [System.Collections.Generic.List[string]]::new()
    $warn = [System.Collections.Generic.List[string]]::new()

    $id   = [string]$rcp.PrimarySmtpAddress
    if (-not $id) { $id = [string]$rcp.Name }
    $rtd  = [string]$rcp.RecipientTypeDetails

    if ($movableTypes -notcontains $rtd)
    {
        $fail.Add("RecipientTypeDetails '$rtd' is not a movable mailbox type")
        [PSCustomObject]@{
            Index = $n; EmailAddress = $id; DisplayName = [string]$rcp.DisplayName
            RecipientTypeDetails = $rtd; PrimarySmtpAddress = $id; CurrentDatabase = $null
            MailboxSizeGB = $null; ItemCount = $null; ArchiveState = $null; ArchiveDatabase = $null
            Holds = $null; ExistingRequest = $null
            Status = 'FAIL'; Issues = ($fail -join ' | ')
        }
        Write-Log ("[{0,4}] {1,-45} FAIL - {2}" -f $n, $id, ($fail -join ' | ')) -Level Fail
        continue
    }

    $mbx = Get-Mailbox -Identity $rcp.Guid.ToString() -ErrorAction SilentlyContinue
    if (-not $mbx)
    {
        $fail.Add("Get-Mailbox returned nothing for this recipient")
        [PSCustomObject]@{
            Index = $n; EmailAddress = $id; DisplayName = [string]$rcp.DisplayName
            RecipientTypeDetails = $rtd; PrimarySmtpAddress = $id; CurrentDatabase = $null
            MailboxSizeGB = $null; ItemCount = $null; ArchiveState = $null; ArchiveDatabase = $null
            Holds = $null; ExistingRequest = $null
            Status = 'FAIL'; Issues = ($fail -join ' | ')
        }
        Write-Log ("[{0,4}] {1,-45} FAIL - {2}" -f $n, $id, ($fail -join ' | ')) -Level Fail
        continue
    }

    $primary     = [string]$mbx.PrimarySmtpAddress
    $currentDb   = [string]$mbx.Database

    # existing requests
    $existing = $null
    $mr = Get-MoveRequest -Identity $mbx.Guid.ToString() -ErrorAction SilentlyContinue
    if ($mr)
    {
        $existing = "MoveRequest ($([string]$mr.Status))"
        $fail.Add("a move request already exists ($([string]$mr.Status)) - remove it first (Remove-MoveRequest) or exclude this mailbox")
    }
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

    # current DB vs exclusions / target
    if ($ExcludeDatabase -and $ExcludeDatabase -contains $currentDb)
    {
        $fail.Add("current database '$currentDb' is in -ExcludeDatabase")
    }
    if ($TargetDatabase -and $currentDb -eq $TargetDatabase)
    {
        $fail.Add("mailbox is already on the target database '$TargetDatabase'")
    }

    # statistics
    $sizeGB = $null; $items = $null
    $stats = Get-MailboxStatistics -Identity $mbx.Guid.ToString() -ErrorAction SilentlyContinue
    if ($stats)
    {
        $sizeGB = Get-SizeGB $stats.TotalItemSize
        $items  = [int]$stats.ItemCount
        if ($stats.DisconnectDate) { $fail.Add("mailbox is disconnected (DisconnectDate $($stats.DisconnectDate))") }
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
            $warn.Add("mailbox has an archive on '$archiveDb' - it stays put unless you run 01 with -IncludeArchive")
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
    switch ($status)
    {
        'FAIL' { Write-Log ("[{0,4}] {1,-45} FAIL - {2}" -f $n, $primary, ($fail -join ' | ')) -Level Fail }
        'WARN' { Write-Log ("[{0,4}] {1,-45} WARN - {2}" -f $n, $primary, ($warn -join ' | ')) -Level Warn }
        'PASS' { Write-Log ("[{0,4}] {1,-45} ready ({2}, {3} GB)" -f $n, $primary, $currentDb, $sizeGB) -Level Ok }
    }

    [PSCustomObject]@{
        Index                = $n
        EmailAddress         = $primary
        DisplayName          = [string]$mbx.DisplayName
        RecipientTypeDetails = $rtd
        PrimarySmtpAddress   = $primary
        CurrentDatabase      = $currentDb
        MailboxSizeGB        = $sizeGB
        ItemCount            = $items
        ArchiveState         = $archiveState
        ArchiveDatabase      = $archiveDb
        Holds                = ($holds -join ',')
        ExistingRequest      = $existing
        Status               = $status
        Issues               = (@($fail) + @($warn)) -join ' | '
    }
}

$results = @($results)

# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------
$results | Export-Csv -LiteralPath $reportFile -NoTypeInformation -Encoding UTF8

$ready = @($results | Where-Object Status -ne 'FAIL' | Sort-Object { [double]($_.MailboxSizeGB) } -Descending)
$readyRows = foreach ($r in $ready)
{
    $row = [ordered]@{ EmailAddress = $r.EmailAddress }
    if ($PSBoundParameters.ContainsKey('TargetDatabase')) { $row.TargetDatabase = $TargetDatabase }
    if ($IncludeArchive) { $row.TargetArchiveDatabase = $TargetArchiveDatabase }
    [PSCustomObject]$row
}
$readyRows = @($readyRows)

# wave split
$parts = New-Object System.Collections.Generic.List[object]
if ($readyRows.Count -eq 0)
{
    Write-Log "No PASS/WARN mailboxes - no Ready CSV written." -Level Warn
}
else
{
    $splitByCount = $MaxWaveMailboxes -gt 0 -and $readyRows.Count -gt $MaxWaveMailboxes
    $totalGB      = ($ready | Measure-Object -Property MailboxSizeGB -Sum).Sum
    $splitBySize  = $MaxWaveSizeGB -gt 0 -and $totalGB -gt $MaxWaveSizeGB

    if (-not $splitByCount -and -not $splitBySize)
    {
        $file = "$readyBase.csv"
        $readyRows | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8
        $parts.Add($file)
    }
    else
    {
        Write-Log ("Wave exceeds a limit (mailboxes {0}/{1}, size {2:N0}/{3} GB) - splitting the Ready CSV." -f `
            $readyRows.Count, $MaxWaveMailboxes, [double]$totalGB, $MaxWaveSizeGB) -Level Warn

        $bucket = New-Object System.Collections.Generic.List[object]
        $bucketGB = 0.0; $partNo = 1
        for ($i = 0; $i -lt $readyRows.Count; $i++)
        {
            $sz = [double]($ready[$i].MailboxSizeGB)
            $countHit = $MaxWaveMailboxes -gt 0 -and $bucket.Count -ge $MaxWaveMailboxes
            $sizeHit  = $MaxWaveSizeGB -gt 0 -and $bucket.Count -gt 0 -and ($bucketGB + $sz) -gt $MaxWaveSizeGB
            if ($countHit -or $sizeHit)
            {
                $file = "{0}_part{1:D2}.csv" -f $readyBase, $partNo
                $bucket | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8
                $parts.Add($file); $partNo++
                $bucket = New-Object System.Collections.Generic.List[object]; $bucketGB = 0.0
            }
            $bucket.Add($readyRows[$i]); $bucketGB += $sz
        }
        if ($bucket.Count)
        {
            $file = "{0}_part{1:D2}.csv" -f $readyBase, $partNo
            $bucket | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8
            $parts.Add($file)
        }
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$pass = @($results | Where-Object Status -eq 'PASS').Count
$warnC = @($results | Where-Object Status -eq 'WARN').Count
$failC = @($results | Where-Object Status -eq 'FAIL').Count

Write-Log "" -Level Head
Write-Log "Summary: $pass PASS, $warnC WARN, $failC FAIL (of $($results.Count))." -Level Head

Write-Log "" -Level Head
Write-Log "Current database distribution (selected mailboxes)" -Level Head
$results | Where-Object CurrentDatabase | Group-Object CurrentDatabase | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0,-30} {1}" -f $_.Name, $_.Count) -Level Info
}

Write-Log "" -Level Head
Write-Log "Full report : $reportFile" -Level Info
foreach ($p in $parts) { Write-Log "Ready CSV   : $p  (feed this to 02-New-OnPremMoveRequest.ps1)" -Level Info }
Write-Log "Log         : $logFile" -Level Info

if ($failC) { Write-Log "Fix the FAIL rows before running 02." -Level Fail }
elseif ($warnC) { Write-Log "Review the WARN rows, then run 02." -Level Warn }
else { Write-Log "All checks passed." -Level Ok }
