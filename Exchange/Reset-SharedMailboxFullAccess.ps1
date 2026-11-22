<#
.SYNOPSIS
    Backs up, removes and re-adds explicit FullAccess permissions on shared
    mailboxes of a given e-mail domain.

.DESCRIPTION
    For every shared mailbox whose primary SMTP address ends with @<Domain>
    (or a single mailbox given by -Identity) the script:

      1. reads all explicit FullAccess entries (Get-MailboxPermission),
      2. writes them to a backup CSV BEFORE anything is changed,
      3. removes each FullAccess entry (Remove-MailboxPermission),
      4. waits -DelaySeconds,
      5. adds each successfully removed entry back (Add-MailboxPermission)
         with the original InheritanceType and the -AutoMapping value.

    Typical use: re-trigger Outlook auto-mapping or repair broken FullAccess
    entries after a migration.

    Entries that are NOT touched (logged and kept in the backup):
      - inherited entries (IsInherited = True),
      - Deny entries,
      - NT AUTHORITY\SELF and other NT AUTHORITY\* accounts,
      - orphaned entries shown only as a SID (S-1-5-...) - they could not be
        added back after removal.

    If removing an entry fails, it is not re-added (it is still in place).
    If re-adding fails, the entry is reported as Failed in the result CSV and
    can be restored later with -RestoreFromCsv <backup file>.

    Works in Exchange Online (Connect-ExchangeOnline) and in Exchange
    on-premises (Exchange Management Shell) - only Get-Mailbox,
    Get-/Add-/Remove-MailboxPermission are used.

    Supports -WhatIf: the backup CSV, log and result CSV are still written,
    but no permission is removed or added.

.PARAMETER Domain
    E-mail domain of the shared mailboxes to process, e.g. contoso.com
    (a leading '@' is accepted). Default: value of $DefaultDomain in the script.

.PARAMETER Identity
    Process only this one shared mailbox (any identity Get-Mailbox accepts).
    Use it to test the procedure on a single mailbox. -Domain is ignored.

.PARAMETER MatchAnyAddress
    Match the domain against all SMTP addresses (proxy addresses), not only
    the primary SMTP address.

.PARAMETER RestoreFromCsv
    Restore mode: reads a backup CSV created by a previous run and adds every
    eligible FullAccess entry that is currently missing. Nothing is removed.

.PARAMETER AutoMapping
    Value passed to Add-MailboxPermission -AutoMapping. Default: $true.

.PARAMETER DelaySeconds
    Pause between removing and re-adding the permissions of one mailbox.
    Default: 2.

.PARAMETER LogFolder
    Folder for the log, backup CSV and result CSV. Created if missing.
    Default: a "Logs" subfolder next to this script.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Identity shared-test@contoso.com -WhatIf

    Dry run on one mailbox: backs up its FullAccess entries and shows what
    would be removed and re-added.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Identity shared-test@contoso.com

    Real run on one mailbox only.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Domain contoso.com -WhatIf
    .\Reset-SharedMailboxFullAccess.ps1 -Domain contoso.com

    Dry run, then real run for all shared mailboxes with a @contoso.com
    primary address.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -RestoreFromCsv .\Logs\Reset-SharedMailboxFullAccess_20261122_101500_backup.csv

    Adds back every FullAccess entry from the backup that is currently missing.

.NOTES
    Version: 1.0 (2026-11-22)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Output files (in -LogFolder, UTF-8, ';' delimited):
      Reset-SharedMailboxFullAccess_<timestamp>.log         - text log (INFO/OK/WARN/ERROR)
      Reset-SharedMailboxFullAccess_<timestamp>_backup.csv  - FullAccess entries before any change
      Reset-SharedMailboxFullAccess_<timestamp>_result.csv  - one line per permission entry

    Result status values (Removed / Added columns):
      Done, WhatIf, Failed, NotNeeded, Skipped.

    Changelog:
      1.0 (2026-11-22) - Initial version.
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Domain')]
param(
    [Parameter(ParameterSetName = 'Domain')]
    [string]$Domain,

    [Parameter(ParameterSetName = 'Domain')]
    [switch]$MatchAnyAddress,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [string]$Identity,

    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$RestoreFromCsv,

    [bool]$AutoMapping = $true,

    [ValidateRange(0, 300)]
    [int]$DelaySeconds = 2,

    [string]$LogFolder = (Join-Path -Path $PSScriptRoot -ChildPath 'Logs')
)

# Default domain used when -Domain is not given on the command line.
$DefaultDomain = 'contoso.com'

#region Logging
# -WhatIf:$false everywhere below: logs and CSVs must be written even in a dry run.
if (-not (Test-Path -LiteralPath $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force -WhatIf:$false | Out-Null
}
$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$scriptName = 'Reset-SharedMailboxFullAccess'
$LogFile    = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$timestamp.log"
$BackupFile = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$($timestamp)_backup.csv"
$ResultFile = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$($timestamp)_result.csv"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 -WhatIf:$false
}

function Get-ErrorText {
    param($ErrorRecord)
    ($ErrorRecord.Exception.Message -replace '\s+', ' ').Trim()
}
#endregion

$isWhatIf = [bool]$WhatIfPreference
Write-Log "Start. Mode: $($PSCmdlet.ParameterSetName) | AutoMapping: $AutoMapping | DelaySeconds: $DelaySeconds | WhatIf: $isWhatIf"
Write-Log "Log file: $LogFile"

#region Prerequisites
foreach ($cmd in 'Get-Mailbox', 'Get-MailboxPermission', 'Add-MailboxPermission', 'Remove-MailboxPermission') {
    if (-not (Get-Command -Name $cmd -ErrorAction SilentlyContinue)) {
        Write-Log "Cmdlet $cmd is not available. Connect to Exchange Online (Connect-ExchangeOnline) or run in the Exchange Management Shell." -Level ERROR
        exit 1
    }
}
#endregion

#region Helpers
# Returns $null when the entry may be reset, otherwise the reason why it is skipped.
function Get-SkipReason {
    param($Entry)
    $user = "$($Entry.User)"
    if ($Entry.IsInherited)                   { return 'Inherited' }
    if ($Entry.Deny)                          { return 'Deny entry' }
    if ($user -like 'NT AUTHORITY\*')         { return 'System account' }
    if ($user -match '^S-1-\d+(-\d+)+$')      { return 'Orphaned SID' }
    return $null
}

function Test-FullAccess {
    param([string]$Mailbox, [string]$User)
    $perm = Get-MailboxPermission -Identity $Mailbox -User $User -ErrorAction SilentlyContinue |
        Where-Object { -not $_.IsInherited -and -not $_.Deny -and ($_.AccessRights -join ',') -match 'FullAccess' }
    return [bool]$perm
}

function New-ResultRow {
    param($Mailbox, $Entry)
    [ordered]@{
        Mailbox         = "$($Mailbox.PrimarySmtpAddress)"
        DisplayName     = "$($Mailbox.DisplayName)"
        User            = "$($Entry.User)"
        InheritanceType = $(if ("$($Entry.InheritanceType)") { "$($Entry.InheritanceType)" } else { 'All' })
        Removed         = 'Skipped'
        Added           = 'Skipped'
        Note            = ''
    }
}
#endregion

$results = New-Object System.Collections.Generic.List[object]

if ($PSCmdlet.ParameterSetName -eq 'Restore') {
    #region Restore mode
    try {
        $backupRows = @(Import-Csv -LiteralPath $RestoreFromCsv -Delimiter ';' -Encoding UTF8 -ErrorAction Stop)
    }
    catch {
        Write-Log "Cannot read backup CSV '$RestoreFromCsv': $(Get-ErrorText $_)" -Level ERROR
        exit 1
    }
    Write-Log "Restore from '$RestoreFromCsv': $($backupRows.Count) row(s)."

    foreach ($row in $backupRows) {
        $result = [ordered]@{
            Mailbox         = $row.Mailbox
            DisplayName     = $row.DisplayName
            User            = $row.User
            InheritanceType = $row.InheritanceType
            Removed         = 'NotNeeded'
            Added           = 'Skipped'
            Note            = ''
        }
        if ($row.SkipReason) {
            $result.Note = "Not restored: $($row.SkipReason)"
            $results.Add([pscustomobject]$result)
            continue
        }
        try {
            if (Test-FullAccess -Mailbox $row.Mailbox -User $row.User) {
                $result.Added = 'NotNeeded'
                $result.Note  = 'Already present'
                Write-Log "$($row.Mailbox): FullAccess for '$($row.User)' already present."
            }
            elseif ($PSCmdlet.ShouldProcess($row.Mailbox, "Add FullAccess for $($row.User)")) {
                $inh = if ($row.InheritanceType) { $row.InheritanceType } else { 'All' }
                Add-MailboxPermission -Identity $row.Mailbox -User $row.User -AccessRights FullAccess `
                    -InheritanceType $inh -AutoMapping $AutoMapping -Confirm:$false -ErrorAction Stop | Out-Null
                $result.Added = 'Done'
                Write-Log "$($row.Mailbox): FullAccess for '$($row.User)' restored." -Level OK
            }
            else {
                $result.Added = 'WhatIf'
            }
        }
        catch {
            $result.Added = 'Failed'
            $result.Note  = Get-ErrorText $_
            Write-Log "$($row.Mailbox): restoring FullAccess for '$($row.User)' FAILED: $($result.Note)" -Level ERROR
        }
        $results.Add([pscustomobject]$result)
    }
    #endregion
}
else {
    #region Collect mailboxes
    try {
        if ($PSCmdlet.ParameterSetName -eq 'Single') {
            $mailboxes = @(Get-Mailbox -Identity $Identity -ErrorAction Stop)
            if ($mailboxes[0].RecipientTypeDetails -ne 'SharedMailbox') {
                Write-Log "Mailbox '$Identity' is $($mailboxes[0].RecipientTypeDetails), not a SharedMailbox. Aborting." -Level ERROR
                exit 1
            }
            Write-Log "Single mailbox mode: $($mailboxes[0].PrimarySmtpAddress)"
        }
        else {
            if (-not $Domain) { $Domain = $DefaultDomain }
            $Domain = $Domain.Trim().TrimStart('@')
            if (-not $Domain) { throw 'Domain is empty.' }
            $suffix = "@$Domain"

            Write-Log "Loading shared mailboxes (domain '$Domain', match any address: $([bool]$MatchAnyAddress))..."
            $allShared = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
            $mailboxes = @($allShared | Where-Object {
                if ($MatchAnyAddress) {
                    @($_.EmailAddresses | Where-Object { "$_" -like 'smtp:*' -and "$_".EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
                }
                else {
                    "$($_.PrimarySmtpAddress)".EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase)
                }
            })
            Write-Log "Shared mailboxes total: $($allShared.Count), matching '$suffix': $($mailboxes.Count)"
        }
    }
    catch {
        Write-Log "Cannot load mailboxes: $(Get-ErrorText $_)" -Level ERROR
        exit 1
    }

    if ($mailboxes.Count -eq 0) {
        Write-Log 'No mailbox to process.' -Level WARN
        exit 0
    }
    #endregion

    #region Backup (before any change)
    $backup  = New-Object System.Collections.Generic.List[object]
    $permMap = @{}
    foreach ($mbx in $mailboxes) {
        $id = "$($mbx.PrimarySmtpAddress)"
        try {
            $entries = @(Get-MailboxPermission -Identity $id -ErrorAction Stop |
                Where-Object { ($_.AccessRights -join ',') -match 'FullAccess' })
        }
        catch {
            Write-Log "$($id): cannot read permissions: $(Get-ErrorText $_)" -Level ERROR
            $permMap[$id] = $null
            continue
        }
        $permMap[$id] = $entries
        foreach ($e in $entries) {
            $backup.Add([pscustomobject][ordered]@{
                Mailbox         = $id
                DisplayName     = "$($mbx.DisplayName)"
                User            = "$($e.User)"
                AccessRights    = ($e.AccessRights -join ',')
                InheritanceType = "$($e.InheritanceType)"
                IsInherited     = [bool]$e.IsInherited
                Deny            = [bool]$e.Deny
                SkipReason      = (Get-SkipReason $e)
            })
        }
    }

    try {
        $backup | Export-Csv -LiteralPath $BackupFile -Delimiter ';' -NoTypeInformation -Encoding UTF8 -WhatIf:$false -ErrorAction Stop
        Write-Log "Backup CSV: $BackupFile ($($backup.Count) FullAccess entr(y/ies), $(@($backup | Where-Object { -not $_.SkipReason }).Count) eligible for reset)" -Level OK
    }
    catch {
        Write-Log "Cannot write backup CSV '$BackupFile': $(Get-ErrorText $_). Aborting - no change was made." -Level ERROR
        exit 1
    }
    #endregion

    #region Remove and re-add
    $i = 0
    foreach ($mbx in $mailboxes) {
        $i++
        $id = "$($mbx.PrimarySmtpAddress)"
        Write-Progress -Activity 'Reset FullAccess' -Status $id -PercentComplete ($i / $mailboxes.Count * 100)

        if ($null -eq $permMap[$id]) {
            $results.Add([pscustomobject][ordered]@{
                Mailbox = $id; DisplayName = "$($mbx.DisplayName)"; User = ''; InheritanceType = ''
                Removed = 'Failed'; Added = 'Skipped'; Note = 'Cannot read permissions'
            })
            continue
        }

        $rows = @()
        foreach ($e in $permMap[$id]) {
            $row  = New-ResultRow -Mailbox $mbx -Entry $e
            $skip = Get-SkipReason $e
            if ($skip) {
                $row.Note = $skip
                if ($skip -eq 'Orphaned SID') {
                    Write-Log "$($id): '$($row.User)' is an orphaned SID - left untouched." -Level WARN
                }
            }
            $rows += , $row
        }

        $eligible = @($rows | Where-Object { -not $_.Note })
        Write-Log "[$i/$($mailboxes.Count)] $($id): $($eligible.Count) FullAccess entr(y/ies) to reset$(if ($eligible.Count) { ': ' + (($eligible | ForEach-Object { $_.User }) -join ', ') })"

        # Step 1: remove
        foreach ($row in $eligible) {
            if (-not $PSCmdlet.ShouldProcess($id, "Remove FullAccess for $($row.User)")) {
                $row.Removed = 'WhatIf'
                continue
            }
            try {
                Remove-MailboxPermission -Identity $id -User $row.User -AccessRights FullAccess `
                    -InheritanceType $row.InheritanceType -Confirm:$false -ErrorAction Stop
                $row.Removed = 'Done'
                Write-Log "$($id): FullAccess for '$($row.User)' removed." -Level OK
            }
            catch {
                $row.Removed = 'Failed'
                $row.Note    = "Remove: $(Get-ErrorText $_)"
                Write-Log "$($id): removing FullAccess for '$($row.User)' FAILED: $(Get-ErrorText $_) - it will not be re-added." -Level ERROR
            }
        }

        # Step 2: re-add only what was removed (or would be removed in WhatIf)
        $toAdd = @($eligible | Where-Object { $_.Removed -in 'Done', 'WhatIf' })
        if ($toAdd.Count -and -not $isWhatIf -and $DelaySeconds -gt 0) {
            Start-Sleep -Seconds $DelaySeconds
        }
        foreach ($row in $toAdd) {
            if (-not $PSCmdlet.ShouldProcess($id, "Add FullAccess for $($row.User) (AutoMapping $AutoMapping)")) {
                $row.Added = 'WhatIf'
                continue
            }
            try {
                Add-MailboxPermission -Identity $id -User $row.User -AccessRights FullAccess `
                    -InheritanceType $row.InheritanceType -AutoMapping $AutoMapping -Confirm:$false -ErrorAction Stop | Out-Null
                $row.Added = 'Done'
                Write-Log "$($id): FullAccess for '$($row.User)' added back." -Level OK
            }
            catch {
                $row.Added = 'Failed'
                $row.Note  = "Add: $(Get-ErrorText $_)"
                Write-Log "$($id): re-adding FullAccess for '$($row.User)' FAILED: $(Get-ErrorText $_). Restore with -RestoreFromCsv '$BackupFile'." -Level ERROR
            }
        }

        foreach ($row in $rows) { $results.Add([pscustomobject]$row) }
    }
    Write-Progress -Activity 'Reset FullAccess' -Completed
    #endregion
}

#region Summary
try {
    $results | Export-Csv -LiteralPath $ResultFile -Delimiter ';' -NoTypeInformation -Encoding UTF8 -WhatIf:$false -ErrorAction Stop
    Write-Log "Result CSV: $ResultFile"
}
catch {
    Write-Log "Cannot write result CSV '$ResultFile': $(Get-ErrorText $_)" -Level ERROR
}

$failed = @($results | Where-Object { $_.Removed -eq 'Failed' -or $_.Added -eq 'Failed' })
Write-Log ('Summary: entries={0}, removed={1}, added={2}, whatif={3}, skipped={4}, failed={5}' -f `
    $results.Count,
    @($results | Where-Object Removed -eq 'Done').Count,
    @($results | Where-Object Added -eq 'Done').Count,
    @($results | Where-Object { $_.Removed -eq 'WhatIf' -or $_.Added -eq 'WhatIf' }).Count,
    @($results | Where-Object { $_.Removed -eq 'Skipped' }).Count,
    $failed.Count) -Level $(if ($failed.Count) { 'WARN' } else { 'OK' })
if ($failed.Count -and $PSCmdlet.ParameterSetName -ne 'Restore') {
    Write-Log "Some entries failed. Missing permissions can be restored with: .\$scriptName.ps1 -RestoreFromCsv '$BackupFile'" -Level WARN
}
Write-Log 'Done.'
#endregion
