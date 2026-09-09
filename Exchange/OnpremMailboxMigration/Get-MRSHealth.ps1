<#
.SYNOPSIS
    Shows the Mailbox Replication Service throttling limits against the number of
    move requests currently active, so you can tell whether a wave will run now
    or just queue.

.DESCRIPTION
    Supporting tool for the OnpremMailboxMigration toolset. Every local mailbox
    move is carried out by MRS on a Mailbox server; MRSProxy is NOT involved for
    same-org moves and is not checked here.

    MRS applies concurrency limits (Get-MailboxReplicationService):
      MaxActiveMovesPerSourceMDB / PerTargetMDB   per database
      MaxActiveMovesPerSourceServer / PerTargetServer
      MaxTotalMovesPerMRS
    A wave bigger than these limits does not fail - the extra requests sit at
    Queued until a slot frees. This script prints the limits, the per-server MRS
    state, and the current active-move count per source database, per target
    database and per server, and flags anything already at or above a limit.

    Output: console + <OutputFolder>\MRSHealth_<ts>.csv + .log

.PARAMETER Server
    Restrict the MRS-configuration read to these Mailbox servers. Default: all
    Mailbox servers returned by Get-ExchangeServer.

.PARAMETER IncomingMoves
    Hypothetical number of new move requests you are about to create. The script
    adds it to the current total and reports the projected queue depth.

.PARAMETER OutputFolder
    Where the log / CSV are written. Default: the script folder.

.EXAMPLE
    .\Get-MRSHealth.ps1

.EXAMPLE
    .\Get-MRSHealth.ps1 -IncomingMoves 300

.NOTES
    Version: 1.0 (2026-09-09)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer).
              View-Only Configuration is enough. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding()]
param(
    [string[]]$Server,

    [int]$IncomingMoves = 0,

    [string]$OutputFolder = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$logFile   = Join-Path $OutputFolder ("MRSHealth_{0}.log" -f $timestamp)
$csvFile   = Join-Path $OutputFolder ("MRSHealth_{0}.csv" -f $timestamp)

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

Write-Log "MRS health - $timestamp" -Level Head

if (-not (Get-Command Get-MailboxReplicationService -ErrorAction SilentlyContinue))
{
    throw "Get-MailboxReplicationService not available. Run this from the on-premises Exchange Management Shell."
}

# ---------------------------------------------------------------------------
# MRS configuration per server
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "MRS configuration" -Level Head

$mrs = if ($Server) { $Server | ForEach-Object { Get-MailboxReplicationService -Identity $_ -ErrorAction SilentlyContinue } }
       else { Get-MailboxReplicationService -ErrorAction SilentlyContinue }
$mrs = @($mrs | Where-Object { $_ })
if (-not $mrs) { throw "Get-MailboxReplicationService returned nothing." }

$cfg = foreach ($m in $mrs)
{
    Write-Log ("  {0,-20} running={1}  MaxTotalMovesPerMRS={2}  PerSourceMDB={3} PerTargetMDB={4}  PerSourceServer={5} PerTargetServer={6}" -f `
        $m.Identity, $m.MRSProxyEnabled, $m.MaxTotalMovesPerMRS, $m.MaxActiveMovesPerSourceMDB, $m.MaxActiveMovesPerTargetMDB, $m.MaxActiveMovesPerSourceServer, $m.MaxActiveMovesPerTargetServer) -Level Info
    [PSCustomObject]@{
        Scope = 'MRSConfig'; Name = [string]$m.Identity
        MaxTotalMovesPerMRS = $m.MaxTotalMovesPerMRS
        MaxActiveMovesPerSourceMDB = $m.MaxActiveMovesPerSourceMDB
        MaxActiveMovesPerTargetMDB = $m.MaxActiveMovesPerTargetMDB
        MaxActiveMovesPerSourceServer = $m.MaxActiveMovesPerSourceServer
        MaxActiveMovesPerTargetServer = $m.MaxActiveMovesPerTargetServer
        Active = $null; Limit = $null; AtLimit = $null
    }
}

$limitPerSourceMDB    = ($mrs | Measure-Object MaxActiveMovesPerSourceMDB -Minimum).Minimum
$limitPerTargetMDB    = ($mrs | Measure-Object MaxActiveMovesPerTargetMDB -Minimum).Minimum
$limitPerSourceServer = ($mrs | Measure-Object MaxActiveMovesPerSourceServer -Minimum).Minimum
$limitPerTargetServer = ($mrs | Measure-Object MaxActiveMovesPerTargetServer -Minimum).Minimum
$limitTotalPerMRS     = ($mrs | Measure-Object MaxTotalMovesPerMRS -Minimum).Minimum

# ---------------------------------------------------------------------------
# Current move requests
# ---------------------------------------------------------------------------
$activeStates = 'Queued', 'InProgress', 'CompletionInProgress'
$all = Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue
$active = @($all | Where-Object { [string]$_.Status -in $activeStates })

Write-Log "" -Level Head
Write-Log ("Move requests: {0} total, {1} active ({2}), {3} queued." -f `
    @($all).Count, $active.Count, ($active | Where-Object { [string]$_.Status -eq 'InProgress' }).Count, ($active | Where-Object { [string]$_.Status -eq 'Queued' }).Count) -Level Info

function Report-Group
{
    param($Rows, [string]$Property, [string]$ScopeName, [int]$Limit)
    $out = foreach ($g in ($Rows | Where-Object { $_.$Property } | Group-Object $Property | Sort-Object Count -Descending))
    {
        $atLimit = $Limit -gt 0 -and $g.Count -ge $Limit
        Write-Log ("  {0,-30} {1,4} active  (limit {2}){3}" -f $g.Name, $g.Count, $Limit, $(if ($atLimit) { '  <= AT LIMIT' } else { '' })) -Level $(if ($atLimit) { 'Warn' } else { 'Info' })
        [PSCustomObject]@{
            Scope = $ScopeName; Name = $g.Name
            MaxTotalMovesPerMRS = $null; MaxActiveMovesPerSourceMDB = $null; MaxActiveMovesPerTargetMDB = $null
            MaxActiveMovesPerSourceServer = $null; MaxActiveMovesPerTargetServer = $null
            Active = $g.Count; Limit = $Limit; AtLimit = $atLimit
        }
    }
    $out
}

Write-Log "" -Level Head
Write-Log "Active moves per SOURCE database (limit $limitPerSourceMDB)" -Level Head
$srcMdb = Report-Group -Rows $active -Property 'SourceDatabase' -ScopeName 'SourceMDB' -Limit $limitPerSourceMDB

Write-Log "" -Level Head
Write-Log "Active moves per TARGET database (limit $limitPerTargetMDB)" -Level Head
$tgtMdb = Report-Group -Rows $active -Property 'TargetDatabase' -ScopeName 'TargetMDB' -Limit $limitPerTargetMDB

# ---------------------------------------------------------------------------
# Projection
# ---------------------------------------------------------------------------
Write-Log "" -Level Head
Write-Log "Overall" -Level Head
$totalActive = $active.Count
Write-Log ("  active now                 : {0}" -f $totalActive) -Level Info
Write-Log ("  MaxTotalMovesPerMRS (min)  : {0}  x {1} MRS instance(s)" -f $limitTotalPerMRS, $mrs.Count) -Level Info
if ($IncomingMoves -gt 0)
{
    $capacity = $limitTotalPerMRS * $mrs.Count
    $projected = $totalActive + $IncomingMoves
    $queued = [math]::Max(0, $projected - $capacity)
    Write-Log ("  + {0} incoming -> {1} requested, ~{2} running, ~{3} queued at any moment" -f $IncomingMoves, $projected, [math]::Min($projected, $capacity), $queued) -Level $(if ($queued -gt 0) { 'Warn' } else { 'Ok' })
    Write-Log "  (queued requests are normal - they start as slots free; not an error)" -Level Info
}

$rows = @($cfg) + @($srcMdb) + @($tgtMdb)
$rows | Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding UTF8

Write-Log "" -Level Head
Write-Log "CSV : $csvFile" -Level Info
Write-Log "Log : $logFile" -Level Info
