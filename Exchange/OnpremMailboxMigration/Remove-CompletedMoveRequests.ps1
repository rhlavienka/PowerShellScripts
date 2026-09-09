<#
.SYNOPSIS
    Housekeeping - removes completed on-premises move requests so they stop
    blocking a later move of the same mailbox.

.DESCRIPTION
    Supporting tool for the OnpremMailboxMigration toolset. A move request stays
    in Get-MoveRequest after it finishes; New-MoveRequest then refuses that
    mailbox until the old request is cleared. Run this once a wave is verified
    complete (05-Get-OnPremMigrationCompletionReport.ps1 clean).

    By default it only touches requests whose Status is Completed or
    CompletedWithWarning and whose completion is older than -OlderThanDays. Add
    -IncludeFailed to also clear Failed / CompletionFailed requests (use with
    care - you lose the diagnostic statistics).

    Output: <OutputFolder>\RemoveCompletedMoveRequests_<ts>.csv + .log

.PARAMETER WaveName
    Restrict to one wave (New-MoveRequest -BatchName). Omit to consider every
    completed move request in the organization.

.PARAMETER OlderThanDays
    Only remove requests that completed at least this many days ago. Default 7.
    0 = no age filter.

.PARAMETER IncludeFailed
    Also remove Failed / CompletionFailed requests. Their MoveRequestStatistics
    (and the failure report) are lost once removed.

.PARAMETER OutputFolder
    Where the log / CSV are written. Default: the wave folder when -WaveName is
    given, otherwise the script folder.

.EXAMPLE
    .\Remove-CompletedMoveRequests.ps1 -WaveName "Praha-Sales-W1" -WhatIf

.EXAMPLE
    .\Remove-CompletedMoveRequests.ps1 -OlderThanDays 14

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
    [string]$WaveName,

    [int]$OlderThanDays = 7,

    [switch]$IncludeFailed,

    [string]$OutputFolder
)

$ErrorActionPreference = 'Stop'

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
$logFile   = Join-Path $OutputFolder ("RemoveCompletedMoveRequests_{0}.log" -f $timestamp)
$csvFile   = Join-Path $OutputFolder ("RemoveCompletedMoveRequests_{0}.csv" -f $timestamp)

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

Write-Log "Remove completed move requests - $timestamp" -Level Head

if (-not (Get-Command Get-MoveRequest -ErrorAction SilentlyContinue))
{
    throw "Get-MoveRequest not available. Run this from the on-premises Exchange Management Shell."
}

$wantStatus = @('Completed', 'CompletedWithWarning')
if ($IncludeFailed) { $wantStatus += 'Failed', 'CompletionFailed' }

$gm = @{ ResultSize = 'Unlimited' }
if ($WaveName) { $gm.BatchName = $WaveName }
$all = Get-MoveRequest @gm | Where-Object { [string]$_.Status -in $wantStatus }

if (-not $all) { Write-Log "No matching move requests." -Level Ok; return }

$cutoff = if ($OlderThanDays -gt 0) { (Get-Date).AddDays(-$OlderThanDays) } else { $null }

$n = 0
$rows = foreach ($mr in $all)
{
    $n++
    $id = [string]$mr.Identity
    $st = Get-MoveRequestStatistics -Identity $mr.Guid.ToString() -ErrorAction SilentlyContinue
    $completedAt = if ($st) { $st.CompletionTimestamp } else { $null }
    if (-not $completedAt -and $st) { $completedAt = $st.LastUpdateTimestamp }

    $tooNew = $cutoff -and $completedAt -and ($completedAt -gt $cutoff)
    $action = if ($tooNew) { 'Kept (too recent)' } else { 'Remove' }

    if ($action -eq 'Remove')
    {
        if ($PSCmdlet.ShouldProcess($id, "Remove-MoveRequest (status $([string]$mr.Status))"))
        {
            try
            {
                Remove-MoveRequest -Identity $mr.Guid.ToString() -Confirm:$false -ErrorAction Stop
                Write-Log ("[{0,4}] {1,-45} removed ({2})" -f $n, $id, [string]$mr.Status) -Level Ok
                $action = 'Removed'
            }
            catch
            {
                Write-Log ("[{0,4}] {1,-45} FAILED - {2}" -f $n, $id, $_.Exception.Message) -Level Fail
                $action = "Failed: $($_.Exception.Message)"
            }
        }
        else
        {
            $action = 'WhatIf'
        }
    }
    else
    {
        Write-Log ("[{0,4}] {1,-45} kept - completed {2}" -f $n, $id, $completedAt) -Level Info
    }

    [PSCustomObject]@{
        Identity    = $id
        Wave        = [string]$mr.BatchName
        Status      = [string]$mr.Status
        CompletedAt = if ($completedAt) { $completedAt.ToString('s') } else { $null }
        Action      = $action
    }
}

$rows | Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding UTF8

Write-Log "" -Level Head
$rows | Group-Object Action | Sort-Object Name | ForEach-Object { Write-Log ("  {0,-20} {1}" -f $_.Name, $_.Count) -Level Info }
Write-Log "" -Level Info
Write-Log "CSV : $csvFile" -Level Info
Write-Log "Log : $logFile" -Level Info
