<#
.SYNOPSIS
    Drives the move requests of an on-premises migration wave - suspend, resume,
    finalize, adjust limits, or remove - without the Exchange admin center.

.DESCRIPTION
    Step 6 of the OnpremMailboxMigration toolset. A thin, auditable wrapper around
    the *-MoveRequest cmdlets so wave control (especially finalization at a
    planned window) is a one-liner with a summary and a confirmation.

    Actions:
      Suspend       Suspend-MoveRequest on everything still running
      Resume        Resume-MoveRequest on suspended / auto-suspended requests
      Complete      finalize now: Set-MoveRequest -CompleteAfter <now> on syncing
                    requests, Resume-MoveRequest on AutoSuspended ones
      SetLimit      Set-MoveRequest -BadItemLimit / -LargeItemLimit / -Priority
      Remove        Remove-MoveRequest (refuses non-Completed unless -Force)

    Scope: the whole wave (-WaveName) or named mailboxes (-Identity), optionally
    narrowed to one class with -OnlyStatus.

.PARAMETER Action
    Suspend | Resume | Complete | SetLimit | Remove.

.PARAMETER WaveName
    The wave (New-MoveRequest -BatchName). Required unless -Identity is used.

.PARAMETER Identity
    One or more mailbox identities instead of / in addition to the whole wave.

.PARAMETER OnlyStatus
    Only act on move requests whose Status matches one of these values
    (e.g. AutoSuspended, InProgress, Failed, Queued).

.PARAMETER BadItemLimit
    New BadItemLimit for -Action SetLimit.

.PARAMETER LargeItemLimit
    New LargeItemLimit for -Action SetLimit.

.PARAMETER Priority
    New Priority for -Action SetLimit.

.PARAMETER SuspendComment
    Comment stored on the request for -Action Suspend.

.PARAMETER Force
    Allow -Action Remove to delete move requests that are not Completed.

.PARAMETER OutputFolder
    Where the log is written. Default: the wave folder when -WaveName is given,
    otherwise the script folder.

.EXAMPLE
    # finalize the whole wave at the change window
    .\06-Invoke-OnPremMoveRequestControl.ps1 -Action Complete -WaveName "Praha-Sales-W1"

.EXAMPLE
    # pause a wave that is hammering a server
    .\06-Invoke-OnPremMoveRequestControl.ps1 -Action Suspend -WaveName "Wave2" -SuspendComment "storage maintenance"

.EXAMPLE
    .\06-Invoke-OnPremMoveRequestControl.ps1 -Action SetLimit -Identity jan.novak@contoso.com -BadItemLimit 25

.NOTES
    Version: 1.0 (2026-09-09)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer),
              the "Move Mailboxes" RBAC role. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Suspend', 'Resume', 'Complete', 'SetLimit', 'Remove')]
    [string]$Action,

    [string]$WaveName,

    [string[]]$Identity,

    [string[]]$OnlyStatus,

    [int]$BadItemLimit,

    [int]$LargeItemLimit,

    [string]$Priority,

    [string]$SuspendComment,

    [switch]$Force,

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

if (-not $WaveName -and -not $Identity) { throw "Supply -WaveName and/or -Identity." }
if ($Action -eq 'SetLimit' -and -not ($PSBoundParameters.ContainsKey('BadItemLimit') -or $PSBoundParameters.ContainsKey('LargeItemLimit') -or $Priority))
{
    throw "-Action SetLimit needs at least one of -BadItemLimit / -LargeItemLimit / -Priority."
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
$logFile   = Join-Path $OutputFolder ("Control_{0}_{1}.log" -f $Action, $timestamp)

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

Write-Log "On-prem move-request control - $Action - $timestamp" -Level Head

if (-not (Get-Command Set-MoveRequest -ErrorAction SilentlyContinue))
{
    throw "Set-MoveRequest not available - your RBAC role is missing 'Move Mailboxes'."
}

# ---------------------------------------------------------------------------
# Collect the target move requests
# ---------------------------------------------------------------------------
$requests = [System.Collections.Generic.List[object]]::new()
if ($WaveName)
{
    Get-MoveRequest -BatchName $WaveName -ResultSize Unlimited -ErrorAction SilentlyContinue | ForEach-Object { $requests.Add($_) }
}
foreach ($i in @($Identity))
{
    $mr = Get-MoveRequest -Identity $i -ErrorAction SilentlyContinue
    if ($mr) { if (-not ($requests | Where-Object { $_.Guid -eq $mr.Guid })) { $requests.Add($mr) } }
    else { Write-Log "No move request for '$i'." -Level Warn }
}

$requests = @($requests)
if ($OnlyStatus) { $requests = @($requests | Where-Object { [string]$_.Status -in $OnlyStatus }) }
if (-not $requests) { Write-Log "No move requests in scope." -Level Warn; return }

Write-Log "$($requests.Count) move request(s) in scope:" -Level Info
$requests | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Log ("  {0,-22} {1}" -f $_.Name, $_.Count) -Level Info }

# ---------------------------------------------------------------------------
# Act
# ---------------------------------------------------------------------------
$now = Get-Date
$done = 0; $skip = 0; $failn = 0

foreach ($mr in $requests)
{
    $id     = [string]$mr.Identity
    $status = [string]$mr.Status
    $guid   = $mr.Guid.ToString()

    $plan = switch ($Action)
    {
        'Suspend'
        {
            if ($status -in 'InProgress', 'Queued') { 'Suspend-MoveRequest' } else { $null }
        }
        'Resume'
        {
            if ($status -in 'Suspended', 'AutoSuspended', 'Failed', 'CompletionFailed') { 'Resume-MoveRequest' } else { $null }
        }
        'Complete'
        {
            if ($status -eq 'AutoSuspended') { 'Resume-MoveRequest' }
            elseif ($status -in 'InProgress', 'Queued', 'Suspended') { 'Set-MoveRequest -CompleteAfter now' }
            else { $null }
        }
        'SetLimit'      { 'Set-MoveRequest -limits' }
        'Remove'
        {
            if ($status -match 'Completed' -or $Force) { 'Remove-MoveRequest' } else { $null }
        }
    }

    if (-not $plan)
    {
        Write-Log ("  skip {0,-45} ({1})" -f $id, $status) -Level Info
        $skip++
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($id, "$Action ($plan; current status $status)"))
    {
        $skip++
        continue
    }

    try
    {
        switch ($plan)
        {
            'Suspend-MoveRequest'
            {
                $p = @{ Identity = $guid; Confirm = $false }
                if ($SuspendComment) { $p.SuspendComment = $SuspendComment }
                Suspend-MoveRequest @p
            }
            'Resume-MoveRequest' { Resume-MoveRequest -Identity $guid -Confirm:$false }
            'Set-MoveRequest -CompleteAfter now'
            {
                Set-MoveRequest -Identity $guid -CompleteAfter $now -SuspendWhenReadyToComplete:$false -Confirm:$false
                if ($status -eq 'Suspended') { Resume-MoveRequest -Identity $guid -Confirm:$false }
            }
            'Set-MoveRequest -limits'
            {
                $p = @{ Identity = $guid; Confirm = $false }
                if ($PSBoundParameters.ContainsKey('BadItemLimit'))   { $p.BadItemLimit = $BadItemLimit; if ($BadItemLimit -ge 51) { $p.AcceptLargeDataLoss = $true } }
                if ($PSBoundParameters.ContainsKey('LargeItemLimit')) { $p.LargeItemLimit = $LargeItemLimit; if ($LargeItemLimit -ge 51) { $p.AcceptLargeDataLoss = $true } }
                if ($Priority) { $p.Priority = $Priority }
                Set-MoveRequest @p
            }
            'Remove-MoveRequest' { Remove-MoveRequest -Identity $guid -Confirm:$false }
        }
        Write-Log ("  ok   {0,-45} {1}" -f $id, $plan) -Level Ok
        $done++
    }
    catch
    {
        Write-Log ("  FAIL {0,-45} {1}" -f $id, $_.Exception.Message) -Level Fail
        $failn++
    }
}

Write-Log "" -Level Head
Write-Log "Done: $done acted, $skip skipped, $failn failed." -Level $(if ($failn) { 'Fail' } else { 'Ok' })
Write-Log "Log: $logFile" -Level Info
if ($Action -in 'Complete', 'Resume', 'Suspend')
{
    Write-Log "Check progress: .\03-Get-OnPremMigrationStatus.ps1 $(if ($WaveName) { "-WaveName '$WaveName'" })" -Level Info
}
