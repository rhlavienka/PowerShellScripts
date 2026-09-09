<#
.SYNOPSIS
    Deep diagnostic dump for failing or stalled on-premises move requests. Use it
    when 03-Get-OnPremMigrationStatus.ps1 shows Failed / Stalled / CorruptItems.

.DESCRIPTION
    Step 4 of the OnpremMailboxMigration toolset. For each mailbox in scope it
    collects everything the Mailbox Replication Service exposes about the move and
    writes it to a per-mailbox folder:

      <mbx>\MoveRequestStatistics.txt   Get-MoveRequestStatistics -IncludeReport
                                        -DiagnosticInfo "verbose;showtimeslots;showtimeline"
      <mbx>\MoveRequestStatistics.xml   the same object, Export-Clixml
      <mbx>\Failures.csv                Report.Failures (timestamp, type, message)
      <mbx>\BadItems.csv / LargeItems.csv
      <mbx>\Entries.txt                 Report.Entries (the move log)
      <mbx>\MailboxStatistics.txt       source (and target, if finalized) counts

    Top level:
      FailureSummary.csv    one row per mailbox: Class, PercentComplete,
                            FailureType, FailureSide, BadItemsEncountered,
                            LargeItemsEncountered, LastError, LikelyCause,
                            SuggestedAction
      <label>_FailureReport_<ts>.zip    the whole folder, zipped for a case
      <label>_FailureReport_<ts>.log    console transcript

    Scope (pick one): -Identity (one or more), -WaveName, or -CsvPath.

.PARAMETER Identity
    One or more mailbox identities to diagnose.

.PARAMETER WaveName
    Diagnose the problem mailboxes of this wave (New-MoveRequest -BatchName).

.PARAMETER CsvPath
    Diagnose the addresses in this CSV (column -IdentityColumn).

.PARAMETER IdentityColumn
    Identity column in -CsvPath. Default "EmailAddress".

.PARAMETER IncludeAll
    With -WaveName / -CsvPath, dump every mailbox, not just the failed / stalled
    ones.

.PARAMETER OutputFolder
    Parent for the report folder. Default: the wave folder when -WaveName is
    given, otherwise the script folder.

.EXAMPLE
    .\04-Get-OnPremMigrationFailureReport.ps1 -WaveName "Praha-Sales-W1"

.EXAMPLE
    .\04-Get-OnPremMigrationFailureReport.ps1 -Identity jan.novak@contoso.com

.NOTES
    Version: 1.0 (2026-09-09)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Requires: on-premises Exchange Management Shell (Exchange 2013 or newer).
              View-Only Recipients is enough. No Exchange Online / Graph modules.

    Changelog:
    1.0 (2026-09-09) - Initial version.
#>

[CmdletBinding(DefaultParameterSetName = 'Identity')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Identity')]
    [string[]]$Identity,

    [Parameter(Mandatory, ParameterSetName = 'Wave')]
    [string]$WaveName,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Csv')]
    [string]$IdentityColumn = 'EmailAddress',

    [Parameter(ParameterSetName = 'Wave')]
    [Parameter(ParameterSetName = 'Csv')]
    [switch]$IncludeAll,

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
$label     = switch ($PSCmdlet.ParameterSetName)
{
    'Wave' { $WaveName -replace '[^\w\-]', '_' }
    'Csv'  { [System.IO.Path]::GetFileNameWithoutExtension($CsvPath) }
    default { 'Mailboxes' }
}

$reportRoot = Join-Path $OutputFolder ("{0}_FailureReport_{1}" -f $label, $timestamp)
$zipPath    = "$reportRoot.zip"
$logFile    = "$reportRoot.log"
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null

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

# ---------------------------------------------------------------------------
# failure type -> plain-language cause + action (on-prem local moves)
# ---------------------------------------------------------------------------
function Resolve-FailureAdvice
{
    param([string]$FailureType, [string]$Message)
    $t = "$FailureType $Message"
    switch -Regex ($t)
    {
        'TooManyBadItems|corrupt items'                 { return 'Corrupt items above BadItemLimit', 'Confirm the items are expendable, then Set-MoveRequest -BadItemLimit <n> and Resume-MoveRequest. Consider New-MailboxRepairRequest first.' }
        'TooManyLargeItems|LargeItem|allowed limit'     { return 'Items over the message-size limit', 'Set-MoveRequest -LargeItemLimit <n> to skip them, or reduce MaxSendSize/MaxReceiveSize differences between source and target.' }
        'StoragePermanent|QuotaExceeded|out of space|insufficient'    { return 'Target database / volume out of space or over quota', 'Free space on the target database volume or pick another -TargetDatabase; check the mailbox quota on the target.' }
        'MapiExceptionNotFound|RelinquishedMailbox|does not exist'    { return 'Source or target object changed / removed', 'Remove-MoveRequest and re-create it from 02; verify the mailbox still exists and is not disconnected.' }
        'MailboxReplicationTransient|Transient|timeout|connection was closed|CommunicationError' { return 'Transient MRS / network / store error', 'Usually self-heals; Resume-MoveRequest and watch. If it repeats, check MRS load (Get-MRSHealth) and the source/target server health.' }
        'MapiExceptionCorrupt|Corrupt|Isam|-1018|-1019'              { return 'Database / mailbox corruption on the source', 'Run New-MailboxRepairRequest against the source mailbox, then Resume-MoveRequest.' }
        'DataMoveReplicationConstraint|replication.*health|CopyQueueLength' { return 'Target DAG copies not healthy enough to accept the move', 'Wait for the target database copies to catch up, or Set-MoveRequest -DataMoveReplicationConstraint None (only if you accept the risk).' }
        'AccessDenied|not have permission|MigrationPermanentException.*credential'  { return 'RBAC / permission problem', 'Confirm the running account holds the "Move Mailboxes" role; check for a deny ACL on the mailbox object.' }
        'QuarantinedMailbox|poison'                     { return 'Source mailbox is quarantined', 'Clear the poison-mailbox quarantine on the source Mailbox server, then Resume-MoveRequest.' }
        'ResourceReservation|resource.*unhealthy|server busy|throttl' { return 'Server resource health / throttling', 'No action needed - MRS retries. If chronic, lower concurrency or raise WorkloadManagement limits.' }
        default                                          { return 'See MoveRequestStatistics / Failures', 'Read Failures.csv and Entries.txt in this mailbox folder; open a support case with the XML if unclear.' }
    }
}

function Get-Bytes
{
    param($Value)
    if (-not $Value) { return $null }
    $text = $Value.ToString()
    if ($text -match '\(([\d,]+) bytes\)') { return [int64]($matches[1] -replace ',', '') }
    return $null
}

# ---------------------------------------------------------------------------
Write-Log "On-prem migration failure report - $label - $timestamp" -Level Head

if (-not (Get-Command Get-MoveRequestStatistics -ErrorAction SilentlyContinue))
{
    throw "Get-MoveRequestStatistics not available. Run this from the on-premises Exchange Management Shell."
}

# ---------------------------------------------------------------------------
# build scope
# ---------------------------------------------------------------------------
switch ($PSCmdlet.ParameterSetName)
{
    'Identity'
    {
        $targets = foreach ($i in $Identity)
        {
            $mr = Get-MoveRequest -Identity $i -ErrorAction SilentlyContinue
            if ($mr) { $mr } else { Write-Log "No move request for '$i'." -Level Warn }
        }
    }
    'Wave'
    {
        $targets = Get-MoveRequest -BatchName $WaveName -ResultSize Unlimited -ErrorAction SilentlyContinue
        if (-not $targets) { throw "No move requests with BatchName '$WaveName'." }
    }
    'Csv'
    {
        if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV '$CsvPath' not found." }
        $rows = Import-Csv -LiteralPath $CsvPath
        if (-not ($rows | Get-Member -Name $IdentityColumn -MemberType NoteProperty)) { throw "CSV '$CsvPath' has no '$IdentityColumn' column." }
        $targets = foreach ($w in ($rows.$IdentityColumn | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }))
        {
            $mr = Get-MoveRequest -Identity $w -ErrorAction SilentlyContinue
            if ($mr) { $mr } else { Write-Log "No move request for '$w'." -Level Warn }
        }
    }
}

$targets = @($targets | Where-Object { $_ })
if (-not $targets) { Write-Log "Nothing in scope." -Level Warn; return }

if (-not $IncludeAll -and $PSCmdlet.ParameterSetName -ne 'Identity')
{
    $targets = $targets | Where-Object {
        [string]$_.Status -match 'Fail' -or
        ([string]$_.Status -eq 'InProgress' -and $_.Message)
    }
    if (-not $targets) { Write-Log "No failed / stalled move requests in scope. Use -IncludeAll to dump everyone." -Level Ok; return }
}
Write-Log "$(@($targets).Count) mailbox(es) to diagnose." -Level Info

# ---------------------------------------------------------------------------
# per-mailbox dump
# ---------------------------------------------------------------------------
$n = 0
$summary = foreach ($mr in $targets)
{
    $n++
    $id = [string]$mr.Identity
    if (-not $id) { $id = [string]$mr.DisplayName }
    Write-Log "" -Level Head
    Write-Log "[$n/$(@($targets).Count)] $id" -Level Head

    $dir = Join-Path $reportRoot ($id -replace '[^\w\.\-@]', '_')
    New-Item -ItemType Directory -Path $dir -Force | Out-Null

    $failureType = $null; $failureSide = $null; $percent = $null
    $badEnc = $null; $largeEnc = $null; $lastError = [string]$mr.Message

    try
    {
        $st = Get-MoveRequestStatistics -Identity $mr.Guid.ToString() -IncludeReport -DiagnosticInfo 'verbose;showtimeslots;showtimeline' -ErrorAction Stop
        $st | Export-Clixml -LiteralPath (Join-Path $dir 'MoveRequestStatistics.xml')
        $st | Format-List * | Out-File -LiteralPath (Join-Path $dir 'MoveRequestStatistics.txt')

        $failureType = [string]$st.FailureType
        $failureSide = [string]$st.FailureSide
        $percent     = $st.PercentComplete
        $badEnc      = $st.BadItemsEncountered
        $largeEnc    = $st.LargeItemsEncountered
        if (-not $lastError -and $st.Message) { $lastError = [string]$st.Message }

        if ($st.Report.Failures)
        {
            $st.Report.Failures | Select-Object Timestamp, FailureType, Message |
                Export-Csv -LiteralPath (Join-Path $dir 'Failures.csv') -NoTypeInformation -Encoding UTF8
        }
        if ($st.Report.BadItems)   { $st.Report.BadItems   | Export-Csv -LiteralPath (Join-Path $dir 'BadItems.csv')   -NoTypeInformation -Encoding UTF8 }
        if ($st.Report.LargeItems) { $st.Report.LargeItems | Export-Csv -LiteralPath (Join-Path $dir 'LargeItems.csv') -NoTypeInformation -Encoding UTF8 }
        if ($st.Report.Entries)
        {
            $st.Report.Entries |
                ForEach-Object { "{0}  [{1}]  {2}" -f $_.CreationTime, $_.Type, $_.Message } |
                Out-File -LiteralPath (Join-Path $dir 'Entries.txt')
        }
        Write-Log ("  {0} / {1} / {2}%  FailureType '{3}' ({4})" -f $st.Status, $st.StatusDetail, $percent, $failureType, $failureSide) -Level Info
    }
    catch
    {
        Write-Log "  Get-MoveRequestStatistics failed: $($_.Exception.Message)" -Level Fail
    }

    # mailbox statistics (source, and target if finalized)
    try
    {
        $src = Get-MailboxStatistics -Identity $id -ErrorAction SilentlyContinue
        $out = "SOURCE`r`n" + ($src | Format-List DisplayName, Database, ItemCount, TotalItemSize, DeletedItemCount, TotalDeletedItemSize | Out-String)
        $out | Out-File -LiteralPath (Join-Path $dir 'MailboxStatistics.txt')
    }
    catch { }

    $cause, $action = Resolve-FailureAdvice -FailureType $failureType -Message $lastError
    Write-Log "  Likely cause : $cause" -Level Fail
    Write-Log "  Suggested    : $action" -Level Warn

    [PSCustomObject]@{
        Identity              = $id
        Wave                  = [string]$mr.BatchName
        Status                = [string]$mr.Status
        PercentComplete       = $percent
        FailureType           = $failureType
        FailureSide           = $failureSide
        BadItemsEncountered   = $badEnc
        LargeItemsEncountered = $largeEnc
        LastError             = ($lastError -split "`r?`n")[0]
        LikelyCause           = $cause
        SuggestedAction       = $action
    }
}

# ---------------------------------------------------------------------------
# summary + zip
# ---------------------------------------------------------------------------
$summaryCsv = Join-Path $reportRoot 'FailureSummary.csv'
$summary | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding UTF8

Write-Log "" -Level Head
Write-Log "Failure summary" -Level Head
$summary | Format-Table Identity, Status, PercentComplete, FailureType, LikelyCause -AutoSize | Out-Host

Write-Log "" -Level Head
Write-Log "Grouped by likely cause" -Level Head
$summary | Group-Object LikelyCause | Sort-Object Count -Descending | ForEach-Object {
    Write-Log ("  x{0}  {1}" -f $_.Count, $_.Name) -Level Fail
}

try
{
    Compress-Archive -Path (Join-Path $reportRoot '*') -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop
    Write-Log "" -Level Head
    Write-Log "Report folder : $reportRoot" -Level Info
    Write-Log "Zipped        : $zipPath" -Level Ok
}
catch
{
    Write-Log "Compress-Archive failed: $($_.Exception.Message) (folder is still there)." -Level Warn
}
Write-Log "Log           : $logFile" -Level Info
