<#
.SYNOPSIS
    Reconciles a finished (or finishing) on-premises migration wave against the
    mailboxes it was supposed to move, and produces a management-friendly rollup.

.DESCRIPTION
    Step 5 of the OnpremMailboxMigration toolset. Takes the original Ready CSV
    from 01 (or a wave name) and, for every mailbox that was in scope, reports:

      - is the move Completed, still running, Failed, or missing a request?
      - the mailbox's database now vs. before (and vs. the requested target, when
        one was specified)
      - source vs. target item-count delta, bad / large items encountered
      - per-mailbox duration and bytes moved
      - a wave rollup: mailbox count, total GB, wall-clock window, mailboxes/hour

    Use it at the end of a wave to prove every requested mailbox landed, and to
    catch anything left behind.

    Output (wave folder when -WaveName is used, else -OutputFolder):
      <label>_Completion_<ts>.csv    per-mailbox reconciliation
      <label>_Completion_<ts>.txt    short text summary (paste into a change record)
      <label>_Completion_<ts>.log    console transcript

.PARAMETER WaveName
    The wave to report on (New-MoveRequest -BatchName). Move requests are read
    with Get-MoveRequest -BatchName.

.PARAMETER CsvPath
    The original Ready CSV (or scope CSV) - defines the set of mailboxes that
    SHOULD have moved. Rows with no matching move request are flagged as
    "NoRequest". Can be combined with -WaveName.

.PARAMETER IdentityColumn
    Identity column in -CsvPath. Default "EmailAddress".

.PARAMETER ExpectDatabase
    Optional. The database (or one of several) the wave targeted. Mailboxes that
    completed onto a different database are flagged "UnexpectedDatabase". Omit
    when automatic mailbox distribution was used.

.PARAMETER IncludeGridView
    Also open the per-mailbox reconciliation in Out-GridView.

.PARAMETER OutputFolder
    Default: the wave folder when -WaveName is given, otherwise the script folder.

.EXAMPLE
    .\05-Get-OnPremMigrationCompletionReport.ps1 -WaveName "Praha-Sales-W1" `
        -CsvPath .\Praha-Sales-W1\Praha-Sales-W1_Ready_2026-09-09_09-00-00.csv

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

    [string[]]$ExpectDatabase,

    [switch]$IncludeGridView,

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

if (-not $WaveName -and -not $CsvPath)
{
    throw "Supply -WaveName, -CsvPath, or both."
}

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
$label     = if ($WaveName) { $WaveName -replace '[^\w\-]', '_' }
             elseif ($CsvPath) { [System.IO.Path]::GetFileNameWithoutExtension($CsvPath) }
             else { 'Wave' }

$logFile     = Join-Path $OutputFolder ("{0}_Completion_{1}.log" -f $label, $timestamp)
$reconCsv    = Join-Path $OutputFolder ("{0}_Completion_{1}.csv" -f $label, $timestamp)
$summaryTxt  = Join-Path $OutputFolder ("{0}_Completion_{1}.txt" -f $label, $timestamp)

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

function Get-Bytes
{
    param($Value)
    if (-not $Value) { return $null }
    $text = $Value.ToString()
    if ($text -match '\(([\d,]+) bytes\)') { return [int64]($matches[1] -replace ',', '') }
    return $null
}

Write-Log "On-prem migration completion report - $label - $timestamp" -Level Head

if (-not (Get-Command Get-MoveRequest -ErrorAction SilentlyContinue))
{
    throw "Get-MoveRequest not available. Run this from the on-premises Exchange Management Shell."
}

# ---------------------------------------------------------------------------
# Expected set (from CSV) and actual move requests (from wave)
# ---------------------------------------------------------------------------
$expected = @()
if ($CsvPath)
{
    if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV '$CsvPath' not found." }
    $rows = Import-Csv -LiteralPath $CsvPath
    if (-not ($rows | Get-Member -Name $IdentityColumn -MemberType NoteProperty)) { throw "CSV '$CsvPath' has no '$IdentityColumn' column." }
    $expected = $rows.$IdentityColumn | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }
    Write-Log "Expected set: $($expected.Count) mailbox(es) from the CSV." -Level Info
}

$requests = @()
if ($WaveName)
{
    $requests = @(Get-MoveRequest -BatchName $WaveName -ResultSize Unlimited -ErrorAction SilentlyContinue)
    Write-Log "Wave '$WaveName': $($requests.Count) move request(s)." -Level Info
}

# union of identities to report on
$identities = [System.Collections.Generic.List[string]]::new()
foreach ($e in $expected) { if (-not $identities.Contains($e)) { $identities.Add($e) } }
foreach ($r in $requests)
{
    $rid = [string]$r.Identity
    if ($rid -and -not ($identities | Where-Object { $_ -ieq $rid })) { $identities.Add($rid) }
}
if (-not $identities) { Write-Log "Nothing to reconcile." -Level Warn; return }

# quick lookup of move requests by a few keys
$reqByKey = @{}
foreach ($r in $requests)
{
    foreach ($k in @([string]$r.Identity, [string]$r.DisplayName, [string]$r.Alias))
    {
        if ($k) { $reqByKey[$k.ToLower()] = $r }
    }
}

# ---------------------------------------------------------------------------
# Per-mailbox reconciliation
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Reconciliation" -Level Head

$n = 0
$recon = foreach ($id in $identities)
{
    $n++
    $wasExpected = [bool]($expected | Where-Object { $_ -ieq $id })

    $mr = $reqByKey[$id.ToLower()]
    if (-not $mr) { $mr = Get-MoveRequest -Identity $id -ErrorAction SilentlyContinue }

    $mbx = Get-Mailbox -Identity $id -ErrorAction SilentlyContinue
    $currentDb = [string]$mbx.Database

    if (-not $mr)
    {
        Write-Log ("[{0,4}] {1,-45} NoRequest (current DB {2})" -f $n, $id, $currentDb) -Level Fail
        [PSCustomObject]@{
            Index = $n; Identity = $id; Expected = $wasExpected; Outcome = 'NoRequest'
            Status = $null; PercentComplete = $null; SourceDatabase = $null; TargetDatabase = $null
            CurrentDatabase = $currentDb; DatabaseCheck = 'n/a'
            SourceItemCount = $null; TargetItemCount = $null; ItemDelta = $null
            BadItems = $null; LargeItems = $null; MovedGB = $null; DurationHours = $null
            CompletedTimestamp = $null; Message = 'no move request found'
        }
        continue
    }

    $st = Get-MoveRequestStatistics -Identity $mr.Guid.ToString() -ErrorAction SilentlyContinue
    $status  = [string]$mr.Status
    $percent = if ($st) { [int]$st.PercentComplete } else { $null }
    $srcDb   = if ($st) { [string]$st.SourceDatabase } else { [string]$mr.SourceDatabase }
    $tgtDb   = if ($st) { [string]$st.TargetDatabase } else { [string]$mr.TargetDatabase }
    $srcItems = if ($st -and $null -ne $st.TotalMailboxItemCount) { [int64]$st.TotalMailboxItemCount } else { $null }
    $tgtItems = if ($st -and $null -ne $st.ItemsTransferred) { [int64]$st.ItemsTransferred } else { $null }
    $badEnc   = if ($st) { [int]$st.BadItemsEncountered } else { $null }
    $largeEnc = if ($st) { [int]$st.LargeItemsEncountered } else { $null }
    $movedGB  = if ($st) { $b = Get-Bytes $st.BytesTransferred; if ($b) { [math]::Round($b / 1GB, 2) } else { $null } } else { $null }
    $durH     = $null
    if ($st)
    {
        foreach ($prop in 'TotalInProgressDuration', 'OverallDuration')
        {
            $raw = $st.$prop
            if ($raw)
            {
                try { $durH = [math]::Round(([timespan]::Parse($raw.ToString())).TotalHours, 1); break } catch { }
            }
        }
    }
    $completed = if ($st) { $st.CompletionTimestamp } else { $null }

    $outcome = switch -Regex ($status)
    {
        'CompletedWithWarning|^Completed' { 'Completed'; break }
        'Fail'                            { 'Failed'; break }
        'AutoSuspended'                   { 'WaitingFinalize'; break }
        '^Suspended'                      { 'Suspended'; break }
        default                           { 'InProgress' }
    }

    $dbCheck = if (-not $ExpectDatabase) { 'auto' }
              elseif ($outcome -ne 'Completed') { 'pending' }
              elseif ($ExpectDatabase -contains $currentDb) { 'match' }
              else { 'UnexpectedDatabase' }

    $delta = if ($null -ne $srcItems -and $null -ne $tgtItems) { $tgtItems - $srcItems } else { $null }

    $lvl = switch ($outcome)
    {
        'Completed'  { if ($dbCheck -eq 'UnexpectedDatabase') { 'Warn' } else { 'Ok' } }
        'Failed'     { 'Fail' }
        default      { 'Info' }
    }
    Write-Log ("[{0,4}] {1,-45} {2,-16} {3,3}%  {4} -> {5}  {6}" -f $n, $id, $outcome, $percent, $srcDb, $tgtDb, $dbCheck) -Level $lvl

    [PSCustomObject]@{
        Index              = $n
        Identity           = $id
        Expected           = $wasExpected
        Outcome            = $outcome
        Status             = $status
        PercentComplete    = $percent
        SourceDatabase     = $srcDb
        TargetDatabase     = $tgtDb
        CurrentDatabase    = $currentDb
        DatabaseCheck      = $dbCheck
        SourceItemCount    = $srcItems
        TargetItemCount    = $tgtItems
        ItemDelta          = $delta
        BadItems           = $badEnc
        LargeItems         = $largeEnc
        MovedGB            = $movedGB
        DurationHours      = $durH
        CompletedTimestamp = if ($completed) { $completed.ToString('s') } else { $null }
        Message            = ($([string]$mr.Message) -split "`r?`n")[0]
    }
}

$recon = @($recon)
$recon | Export-Csv -LiteralPath $reconCsv -NoTypeInformation -Encoding UTF8

if ($IncludeGridView -and (Get-Command Out-GridView -ErrorAction SilentlyContinue))
{
    $recon | Out-GridView -Title ("Completion - {0} - {1}" -f $label, $timestamp)
}

# ---------------------------------------------------------------------------
# Rollup
# ---------------------------------------------------------------------------
$byOutcome = $recon | Group-Object Outcome | Sort-Object Name
$completedRows = $recon | Where-Object Outcome -eq 'Completed'
$totalGB   = [math]::Round((($completedRows | Measure-Object MovedGB -Sum).Sum), 1)
$times     = $completedRows | Where-Object CompletedTimestamp | ForEach-Object { [datetime]$_.CompletedTimestamp }
$window    = if ($times) { "{0} .. {1}" -f ($times | Measure-Object -Minimum).Minimum.ToString('s'), ($times | Measure-Object -Maximum).Maximum.ToString('s') } else { 'n/a' }
$rate      = if ($times -and @($times).Count -gt 1)
             {
                 $span = (($times | Measure-Object -Maximum).Maximum - ($times | Measure-Object -Minimum).Minimum).TotalHours
                 if ($span -gt 0) { [math]::Round(@($completedRows).Count / $span, 1) } else { $null }
             } else { $null }

$lines = @()
$lines += "On-prem migration completion - $label - $timestamp"
$lines += ""
$lines += "Mailboxes reported : $($recon.Count)"
foreach ($g in $byOutcome) { $lines += ("  {0,-16} {1}" -f $g.Name, $g.Count) }
$lines += ""
$lines += "Completed data     : $totalGB GB"
$lines += "Completion window  : $window"
if ($rate) { $lines += "Throughput         : ~$rate mailboxes/hour (completed)" }
$notExpected = $recon | Where-Object { -not $_.Expected }
if ($notExpected) { $lines += ""; $lines += "In the wave but NOT in the expected CSV: $($notExpected.Count)" ; $notExpected | ForEach-Object { $lines += "  - $($_.Identity)" } }
$leftBehind = $recon | Where-Object { $_.Expected -and $_.Outcome -ne 'Completed' }
if ($leftBehind) { $lines += ""; $lines += "Expected but NOT completed: $($leftBehind.Count)"; $leftBehind | ForEach-Object { $lines += "  - $($_.Identity)  [$($_.Outcome)]" } }
$badDb = $recon | Where-Object DatabaseCheck -eq 'UnexpectedDatabase'
if ($badDb) { $lines += ""; $lines += "Completed onto an unexpected database: $($badDb.Count)"; $badDb | ForEach-Object { $lines += "  - $($_.Identity) -> $($_.CurrentDatabase)" } }

$lines | Set-Content -LiteralPath $summaryTxt -Encoding UTF8

Write-Log "" -Level Head
$lines | ForEach-Object { Write-Log $_ -Level Info }

Write-Log "" -Level Head
Write-Log "Reconciliation CSV : $reconCsv" -Level Info
Write-Log "Text summary       : $summaryTxt" -Level Info
Write-Log "Log                : $logFile" -Level Info

if ($leftBehind) { Write-Log "Wave is NOT fully complete - $($leftBehind.Count) expected mailbox(es) outstanding." -Level Fail }
elseif ($recon | Where-Object Outcome -eq 'Failed') { Write-Log "Some move requests failed - see 04-Get-OnPremMigrationFailureReport.ps1." -Level Fail }
else { Write-Log "Every expected mailbox is Completed." -Level Ok }
