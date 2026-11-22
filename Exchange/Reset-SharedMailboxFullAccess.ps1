<#
.SYNOPSIS
    Backs up, removes and re-adds explicit FullAccess permissions on shared
    mailboxes of a given e-mail domain. Built for large environments.

.DESCRIPTION
    For every shared mailbox whose primary SMTP address ends with @<Domain>
    (or a single mailbox given by -Identity) the script:

      1. reads the explicit (non-inherited) FullAccess entries,
      2. appends them to the backup CSV BEFORE the mailbox is changed,
      3. removes each eligible FullAccess entry (Remove-MailboxPermission),
      4. waits -DelaySeconds (default 0),
      5. adds each successfully removed entry back (Add-MailboxPermission)
         with the original InheritanceType and the -AutoMapping value.

    Typical use: re-trigger Outlook auto-mapping or repair broken FullAccess
    entries after a migration.

    Entries that are backed up but NOT touched (SkipReason column):
      - Deny entries,
      - NT AUTHORITY\SELF and other NT AUTHORITY\* accounts,
      - orphaned entries shown only as a SID (S-1-5-...) - they could not be
        added back after removal.
    Inherited entries and NT AUTHORITY\SELF (the mailbox's own access, present
    on every mailbox) are ignored completely (not backed up, not touched).

    If removing an entry fails, it is not re-added (it is still in place).
    If re-adding fails, the entry is reported as Failed in the result CSV and
    can be restored later with -RestoreFromCsv <backup file>.

    Trustee details: Get-MailboxPermission returns only the trustee name
    (User). PrimarySmtpAddress, Guid and RecipientType(Details) of the trustee
    are looked up with Get-EXORecipient / Get-Recipient, once per unique
    trustee (cached). Trustees that are not mail-enabled (e.g. a plain AD
    security group) stay empty in these columns. -SkipTrusteeLookup turns the
    lookup off completely.

    Performance (designed for ~100 000 mailboxes):
      - Exchange Online is detected automatically and the REST-based
        Get-EXOMailbox / Get-EXOMailboxPermission / Get-EXORecipient cmdlets
        are used; on-premises the classic cmdlets are used.
      - Only shared mailboxes are loaded (-RecipientTypeDetails SharedMailbox),
        reduced to the few needed properties. On-premises the domain is also
        filtered server-side (-Filter "EmailAddresses -like '*@<Domain>'").
        In Exchange Online a leading wildcard in -Filter is not officially
        supported, so there the domain is matched client-side (Minimum
        property set keeps the transfer small).
      - Exactly one permission read per mailbox.
      - Backup, result and log are streamed to disk line by line (nothing is
        accumulated in memory) - an interrupted run keeps everything written
        so far.
      - The console shows one line per mailbox, progress with ETA every
        -ProgressInterval mailboxes, warnings and errors; per-entry details
        go to the log file only.
    The remaining cost is one Remove + one Add call per FullAccess entry,
    which cannot be batched.

    Works in Exchange Online (Connect-ExchangeOnline) and in Exchange
    on-premises (Exchange Management Shell / remote PowerShell).

    Supports -WhatIf: the backup CSV, log and result CSV are still written,
    but no permission is removed or added.

.PARAMETER Domain
    E-mail domain of the shared mailboxes to process, e.g. contoso.com
    (a leading '@' is accepted). Default: value of $DefaultDomain in the script.

.PARAMETER Identity
    Process only this one shared mailbox (any identity Get-Mailbox accepts).
    Use it to test the procedure on a single mailbox. -Identity takes
    precedence over -Domain: if both are given, only this mailbox is processed
    and a warning is logged when it does not match the domain.

.PARAMETER MatchAnyAddress
    Match the domain against all SMTP addresses (proxy addresses), not only
    the primary SMTP address.

.PARAMETER BackupOnly
    Only write the backup CSV; no permission is removed or added.

.PARAMETER SkipTrusteeLookup
    Do not look up trustee PrimarySmtpAddress / Guid / RecipientType
    (fastest; the Trustee* columns stay empty).

.PARAMETER RestoreFromCsv
    Restore mode: reads a backup CSV created by a previous run and adds every
    eligible FullAccess entry that is currently missing. Nothing is removed.

.PARAMETER AutoMapping
    Value passed to Add-MailboxPermission -AutoMapping. Default: $true.

.PARAMETER DelaySeconds
    Pause between removing and re-adding the permissions of one mailbox.
    Default: 0. Every second here costs ~1 second per mailbox.

.PARAMETER ProgressInterval
    Show progress (console + log) every N mailboxes. Default: 100.

.PARAMETER LogFolder
    Folder for the log, backup CSV and result CSV. Created if missing.
    Default: a "Logs" subfolder next to this script.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Identity shared-test@contoso.com -WhatIf

    Dry run on one mailbox: backs up its FullAccess entries and shows what
    would be removed and re-added.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Domain contoso.com -BackupOnly

    Only backs up FullAccess entries of all @contoso.com shared mailboxes.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -Domain contoso.com -WhatIf
    .\Reset-SharedMailboxFullAccess.ps1 -Domain contoso.com

    Dry run, then real run for all shared mailboxes with a @contoso.com
    primary address.

.EXAMPLE
    .\Reset-SharedMailboxFullAccess.ps1 -RestoreFromCsv .\Logs\Reset-SharedMailboxFullAccess_20261122_101500_backup.csv

    Adds back every FullAccess entry from the backup that is currently missing.

.NOTES
    Version: 2.4 (2026-11-22)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Output files (in -LogFolder, UTF-8, ';' delimited):
      Reset-SharedMailboxFullAccess_<timestamp>.log         - text log (INFO/OK/WARN/ERROR)
      Reset-SharedMailboxFullAccess_<timestamp>_backup.csv  - FullAccess entries before any change
      Reset-SharedMailboxFullAccess_<timestamp>_result.csv  - one line per permission entry

    Backup columns:
      Mailbox, MailboxGuid, MailboxRecipientTypeDetails, DisplayName, User,
      TrusteePrimarySmtpAddress, TrusteeGuid, TrusteeRecipientType,
      TrusteeRecipientTypeDetails, AccessRights, InheritanceType, Deny, SkipReason
    Result columns: backup columns + Removed, Added, Note.

    Result status values (Removed / Added columns):
      Done, WhatIf, Failed, NotNeeded, Skipped.

    Changelog:
      2.4 (2026-11-22) - Exchange Online again matches the domain client-side (leading wildcard in -Filter is not officially supported); on-premises keeps the server-side filter.
      2.3 (2026-11-22) - Server-side domain filter also in Exchange Online via EmailAddresses -like '*@domain'; client-side check keeps the primary-address rule.
      2.2 (2026-11-22) - Fix empty TrusteeGuid in EXO (Get-EXORecipient -Properties Guid); NT AUTHORITY\SELF no longer backed up; one console line per mailbox (incl. mailboxes without entries).
      2.1 (2026-11-22) - Fix: domain filter returned no mailbox in Exchange Online (leading wildcard not supported in -Filter); EXO now matches the domain client-side, on-premises falls back to an unfiltered load when the filter returns nothing.
      2.0 (2026-11-22) - Performance rewrite for ~100k mailboxes: EXO REST cmdlets, server-side domain filter, streamed CSV/log output, per-mailbox backup; added mailbox Guid/RecipientTypeDetails and trustee PrimarySmtpAddress/Guid/RecipientType columns, -BackupOnly, -SkipTrusteeLookup, -ProgressInterval; inherited entries no longer backed up; DelaySeconds default 2 -> 0.
      1.1 (2026-11-22) - -Domain and -MatchAnyAddress can be combined with -Identity (domain mismatch is logged as a warning).
      1.0 (2026-11-22) - Initial version.
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Domain')]
param(
    [Parameter(ParameterSetName = 'Domain')]
    [Parameter(ParameterSetName = 'Single')]
    [ValidatePattern('^@?[A-Za-z0-9][A-Za-z0-9.-]*$')]
    [string]$Domain,

    [Parameter(ParameterSetName = 'Domain')]
    [Parameter(ParameterSetName = 'Single')]
    [switch]$MatchAnyAddress,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [string]$Identity,

    [Parameter(ParameterSetName = 'Domain')]
    [Parameter(ParameterSetName = 'Single')]
    [switch]$BackupOnly,

    [Parameter(ParameterSetName = 'Domain')]
    [Parameter(ParameterSetName = 'Single')]
    [switch]$SkipTrusteeLookup,

    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$RestoreFromCsv,

    [bool]$AutoMapping = $true,

    [ValidateRange(0, 300)]
    [int]$DelaySeconds = 0,

    [ValidateRange(1, 1000000)]
    [int]$ProgressInterval = 100,

    [string]$LogFolder = (Join-Path -Path $PSScriptRoot -ChildPath 'Logs')
)

# Default domain used when -Domain is not given on the command line.
$DefaultDomain = 'contoso.com'

#region Output files (streamed)
if (-not (Test-Path -LiteralPath $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force -WhatIf:$false | Out-Null
}
$LogFolder  = (Resolve-Path -LiteralPath $LogFolder).ProviderPath
$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$scriptName = 'Reset-SharedMailboxFullAccess'
$LogFile    = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$timestamp.log"
$BackupFile = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$($timestamp)_backup.csv"
$ResultFile = Join-Path -Path $LogFolder -ChildPath "$($scriptName)_$($timestamp)_result.csv"

$BackupColumns = @(
    'Mailbox', 'MailboxGuid', 'MailboxRecipientTypeDetails', 'DisplayName', 'User',
    'TrusteePrimarySmtpAddress', 'TrusteeGuid', 'TrusteeRecipientType', 'TrusteeRecipientTypeDetails',
    'AccessRights', 'InheritanceType', 'Deny', 'SkipReason'
)
$ResultColumns = $BackupColumns + @('Removed', 'Added', 'Note')

# StreamWriter instead of Add-Content / Export-Csv: no file re-open per line,
# no in-memory accumulation, not affected by -WhatIf.
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
function New-OutputWriter {
    param([string]$Path, [string[]]$Header)
    $w = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
    $w.AutoFlush = $true
    if ($Header) { $w.WriteLine((ConvertTo-CsvLine -Values $Header)) }
    return $w
}

function ConvertTo-CsvLine {
    param([object[]]$Values)
    $sb = New-Object System.Text.StringBuilder
    for ($n = 0; $n -lt $Values.Count; $n++) {
        if ($n) { [void]$sb.Append(';') }
        [void]$sb.Append('"').Append(("$($Values[$n])").Replace('"', '""')).Append('"')
    }
    $sb.ToString()
}

function Write-CsvRow {
    param($Writer, $Row, [string[]]$Columns)
    # String conversion keeps empty values as "" (a bare $null/empty would be dropped and shift columns).
    $vals = foreach ($c in $Columns) { "$($Row.$c)" }
    $Writer.WriteLine((ConvertTo-CsvLine -Values $vals))
}

$LogWriter    = $null
$BackupWriter = $null
$ResultWriter = $null

# -FileOnly: detail line for the log file only (WARN/ERROR always reach the console too).
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [switch]$FileOnly
    )
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($script:LogWriter) { $script:LogWriter.WriteLine($line) }
    if (-not $FileOnly -or $Level -eq 'WARN' -or $Level -eq 'ERROR') {
        $color = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
        Write-Host $line -ForegroundColor $color
    }
}

function Get-ErrorText {
    param($ErrorRecord)
    ($ErrorRecord.Exception.Message -replace '\s+', ' ').Trim()
}
#endregion

#region Helpers
# Returns $null when the entry may be reset, otherwise the reason why it is skipped.
function Get-SkipReason {
    param($Entry)
    $user = "$($Entry.User)"
    if ($Entry.Deny)                      { return 'Deny entry' }
    if ($user -like 'NT AUTHORITY\*')     { return 'System account' }
    if ($user -match '^S-1-\d+(-\d+)+$')  { return 'Orphaned SID' }
    return $null
}

function Test-DomainMatch {
    param($Mailbox, [string]$Suffix)
    if ($MatchAnyAddress) {
        foreach ($a in $Mailbox.EmailAddresses) {
            if ($a -like 'smtp:*' -and $a.EndsWith($Suffix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
        return $false
    }
    return $Mailbox.PrimarySmtpAddress.EndsWith($Suffix, [StringComparison]::OrdinalIgnoreCase)
}

# Reduces a mailbox object to the few properties the script needs (low memory).
function Select-MailboxInfo {
    process {
        [pscustomobject]@{
            Guid                 = "$($_.Guid)"
            PrimarySmtpAddress   = "$($_.PrimarySmtpAddress)"
            DisplayName          = "$($_.DisplayName)"
            RecipientTypeDetails = "$($_.RecipientTypeDetails)"
            EmailAddresses       = [string[]]@($_.EmailAddresses | ForEach-Object { "$_" })
        }
    }
}

function Get-MailboxInfo {
    param([hashtable]$Params)
    if ($UseEXO) { Get-EXOMailbox @Params -PropertySets Minimum -ErrorAction Stop | Select-MailboxInfo }
    else         { Get-Mailbox    @Params -ErrorAction Stop | Select-MailboxInfo }
}

# Explicit (non-inherited) FullAccess entries of one mailbox - one call per mailbox.
function Get-FullAccessEntry {
    param([string]$MailboxId)
    $perms = if ($UseEXO) { Get-EXOMailboxPermission -Identity $MailboxId -ErrorAction Stop }
             else         { Get-MailboxPermission    -Identity $MailboxId -ErrorAction Stop }
    foreach ($p in $perms) {
        # NT AUTHORITY\SELF is the mailbox's own access (present on every mailbox) - not backed up.
        if (-not $p.IsInherited -and "$($p.User)" -ne 'NT AUTHORITY\SELF' -and ($p.AccessRights -join ',') -match 'FullAccess') { $p }
    }
}

$TrusteeCache = @{}
function Resolve-Trustee {
    param([string]$User)
    if ($SkipTrusteeLookup -or -not $User) { return $null }
    if ($TrusteeCache.ContainsKey($User)) { return $TrusteeCache[$User] }
    $info = $null
    try {
        # Guid is not part of the EXO Minimum property set - request it explicitly.
        $r = @(if ($UseEXO) { Get-EXORecipient -Identity $User -Properties Guid -ErrorAction Stop }
               else         { Get-Recipient    -Identity $User -ErrorAction Stop })
        if ($r.Count -eq 1) {
            $info = [pscustomobject]@{
                PrimarySmtpAddress   = "$($r[0].PrimarySmtpAddress)"
                Guid                 = "$($r[0].Guid)"
                RecipientType        = "$($r[0].RecipientType)"
                RecipientTypeDetails = "$($r[0].RecipientTypeDetails)"
            }
        }
        else {
            Write-Log "Trustee '$User' is ambiguous ($($r.Count) recipients) - details left empty." -Level WARN -FileOnly
        }
    }
    catch {
        Write-Log "Trustee '$User' not resolved as a recipient: $(Get-ErrorText $_)" -FileOnly
    }
    $TrusteeCache[$User] = $info
    return $info
}

function New-EntryRow {
    param($Mailbox, $Entry)
    $skip    = Get-SkipReason $Entry
    $user    = "$($Entry.User)"
    $trustee = if (-not $skip) { Resolve-Trustee -User $user } else { $null }
    $inh     = "$($Entry.InheritanceType)"
    [ordered]@{
        Mailbox                     = $Mailbox.PrimarySmtpAddress
        MailboxGuid                 = $Mailbox.Guid
        MailboxRecipientTypeDetails = $Mailbox.RecipientTypeDetails
        DisplayName                 = $Mailbox.DisplayName
        User                        = $user
        TrusteePrimarySmtpAddress   = $(if ($trustee) { $trustee.PrimarySmtpAddress })
        TrusteeGuid                 = $(if ($trustee) { $trustee.Guid })
        TrusteeRecipientType        = $(if ($trustee) { $trustee.RecipientType })
        TrusteeRecipientTypeDetails = $(if ($trustee) { $trustee.RecipientTypeDetails })
        AccessRights                = ($Entry.AccessRights -join ',')
        InheritanceType             = $(if ($inh) { $inh } else { 'All' })
        Deny                        = [bool]$Entry.Deny
        SkipReason                  = $skip
        Removed                     = 'Skipped'
        Added                       = 'Skipped'
        Note                        = ''
    }
}

function Write-ProgressInfo {
    param([int]$Current, [int]$Total, [System.Diagnostics.Stopwatch]$Watch, [string]$Activity)
    $pct = [math]::Round($Current / [math]::Max($Total, 1) * 100, 1)
    $eta = if ($Current -gt 0) { [timespan]::FromSeconds($Watch.Elapsed.TotalSeconds / $Current * ($Total - $Current)) } else { [timespan]::Zero }
    Write-Progress -Activity $Activity -Status "$Current / $Total ($pct %) - ETA $($eta.ToString('hh\:mm\:ss'))" -PercentComplete $pct
    Write-Log ("Progress: {0}/{1} ({2} %), elapsed {3}, ETA {4}" -f $Current, $Total, $pct, $Watch.Elapsed.ToString('hh\:mm\:ss'), $eta.ToString('hh\:mm\:ss'))
}
#endregion

$stats = [ordered]@{
    Mailboxes = 0; MailboxesWithEntries = 0; ReadFailed = 0; Entries = 0; Eligible = 0
    Removed = 0; Added = 0; NotNeeded = 0; WhatIf = 0; Failed = 0
}

try {
    $LogWriter = New-OutputWriter -Path $LogFile
    $isWhatIf  = [bool]$WhatIfPreference
    Write-Log "Start. Mode: $($PSCmdlet.ParameterSetName) | BackupOnly: $([bool]$BackupOnly) | AutoMapping: $AutoMapping | DelaySeconds: $DelaySeconds | WhatIf: $isWhatIf"
    Write-Log "Log file: $LogFile"

    #region Prerequisites / environment detection
    $UseEXO = $false
    if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
        try {
            $conn = Get-ConnectionInformation -ErrorAction Stop |
                Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
            if ($conn) {
                $UseEXO = $true
                Write-Log "Exchange Online detected (tenant $($conn.TenantID), $($conn.UserPrincipalName)) - using REST cmdlets."
            }
        }
        catch { }
    }
    $required = if ($UseEXO) { 'Get-EXOMailbox', 'Get-EXOMailboxPermission', 'Get-EXORecipient', 'Add-MailboxPermission', 'Remove-MailboxPermission' }
                else         { 'Get-Mailbox', 'Get-MailboxPermission', 'Get-Recipient', 'Add-MailboxPermission', 'Remove-MailboxPermission' }
    foreach ($cmd in $required) {
        if (-not (Get-Command -Name $cmd -ErrorAction SilentlyContinue)) {
            Write-Log "Cmdlet $cmd is not available. Connect to Exchange Online (Connect-ExchangeOnline) or run in the Exchange Management Shell." -Level ERROR
            exit 1
        }
    }
    if (-not $UseEXO) { Write-Log 'Exchange on-premises cmdlets detected.' }
    #endregion

    if ($PSCmdlet.ParameterSetName -eq 'Restore') {
        #region Restore mode
        try {
            $backupRows = @(Import-Csv -LiteralPath $RestoreFromCsv -Delimiter ';' -Encoding UTF8 -ErrorAction Stop)
        }
        catch {
            Write-Log "Cannot read backup CSV '$RestoreFromCsv': $(Get-ErrorText $_)" -Level ERROR
            exit 1
        }
        $ResultWriter = New-OutputWriter -Path $ResultFile -Header $ResultColumns
        $groups = @($backupRows | Group-Object -Property { if ($_.MailboxGuid) { $_.MailboxGuid } else { $_.Mailbox } })
        Write-Log "Restore from '$RestoreFromCsv': $($backupRows.Count) row(s), $($groups.Count) mailbox(es)."

        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $i = 0
        foreach ($g in $groups) {
            $i++
            if ($i -eq 1 -or $i % $ProgressInterval -eq 0) { Write-ProgressInfo -Current $i -Total $groups.Count -Watch $watch -Activity 'Restore FullAccess' }
            $stats.Mailboxes++

            $present  = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $readErr  = $null
            try { foreach ($p in (Get-FullAccessEntry -MailboxId $g.Name)) { if (-not $p.Deny) { [void]$present.Add("$($p.User)") } } }
            catch { $readErr = Get-ErrorText $_; $stats.ReadFailed++; Write-Log "$($g.Name): cannot read permissions: $readErr" -Level ERROR }

            foreach ($b in $g.Group) {
                $row = [ordered]@{}
                foreach ($c in $BackupColumns) { $row[$c] = $b.$c }
                $row.Removed = 'NotNeeded'; $row.Added = 'Skipped'; $row.Note = ''
                $stats.Entries++
                $mbxId = if ($b.MailboxGuid) { $b.MailboxGuid } else { $b.Mailbox }

                if ($b.SkipReason) {
                    $row.Note = "Not restored: $($b.SkipReason)"
                }
                elseif ($readErr) {
                    $row.Added = 'Failed'; $row.Note = "Read: $readErr"; $stats.Failed++
                }
                elseif ($present.Contains("$($b.User)")) {
                    $row.Added = 'NotNeeded'; $stats.NotNeeded++
                    Write-Log "$($b.Mailbox): FullAccess for '$($b.User)' already present." -FileOnly
                }
                elseif ($PSCmdlet.ShouldProcess($b.Mailbox, "Add FullAccess for $($b.User)")) {
                    try {
                        $inh = if ($b.InheritanceType) { $b.InheritanceType } else { 'All' }
                        Add-MailboxPermission -Identity $mbxId -User $b.User -AccessRights FullAccess `
                            -InheritanceType $inh -AutoMapping $AutoMapping -Confirm:$false -ErrorAction Stop | Out-Null
                        $row.Added = 'Done'; $stats.Added++
                        Write-Log "$($b.Mailbox): FullAccess for '$($b.User)' restored." -Level OK -FileOnly
                    }
                    catch {
                        $row.Added = 'Failed'; $row.Note = Get-ErrorText $_; $stats.Failed++
                        Write-Log "$($b.Mailbox): restoring FullAccess for '$($b.User)' FAILED: $($row.Note)" -Level ERROR
                    }
                }
                else {
                    $row.Added = 'WhatIf'; $stats.WhatIf++
                }
                Write-CsvRow -Writer $ResultWriter -Row $row -Columns $ResultColumns
            }
        }
        #endregion
    }
    else {
        #region Collect mailboxes
        try {
            if ($PSCmdlet.ParameterSetName -eq 'Single') {
                $mailboxes = @(Get-MailboxInfo -Params @{ Identity = $Identity })
                if ($mailboxes[0].RecipientTypeDetails -ne 'SharedMailbox') {
                    Write-Log "Mailbox '$Identity' is $($mailboxes[0].RecipientTypeDetails), not a SharedMailbox. Aborting." -Level ERROR
                    exit 1
                }
                Write-Log "Single mailbox mode: $($mailboxes[0].PrimarySmtpAddress)"
                if ($Domain) {
                    $Domain = $Domain.TrimStart('@')
                    if (-not (Test-DomainMatch -Mailbox $mailboxes[0] -Suffix "@$Domain")) {
                        Write-Log "Mailbox '$($mailboxes[0].PrimarySmtpAddress)' does not match domain '$Domain' (match any address: $([bool]$MatchAnyAddress)). -Identity takes precedence - processing it anyway." -Level WARN
                    }
                }
            }
            else {
                if (-not $Domain) { $Domain = $DefaultDomain }
                $Domain = $Domain.TrimStart('@')
                $suffix = "@$Domain"
                $params = @{ RecipientTypeDetails = 'SharedMailbox'; ResultSize = 'Unlimited' }
                $watch  = [System.Diagnostics.Stopwatch]::StartNew()

                if ($UseEXO) {
                    # A leading wildcard (-Filter "... -like '*@domain'") is not officially
                    # supported in Exchange Online, so its result is not relied on. Load all
                    # shared mailboxes (Minimum property set - cheap compared to the
                    # Remove/Add calls) and match the domain client-side.
                    Write-Log 'Loading all shared mailboxes (Exchange Online; domain matched client-side)...'
                    $found = @(Get-MailboxInfo -Params $params)
                }
                else {
                    # On-premises supports a leading wildcard; EmailAddresses covers the
                    # primary address too, the client-side check below narrows the result.
                    $filter = "EmailAddresses -like '*$suffix'"
                    Write-Log "Loading shared mailboxes with server-side filter: $filter"
                    try {
                        $found = @(Get-MailboxInfo -Params ($params + @{ Filter = $filter }))
                        if ($found.Count -eq 0) {
                            Write-Log 'Server-side filter returned no mailbox - re-checking by loading all shared mailboxes.' -Level WARN
                            $found = @(Get-MailboxInfo -Params $params)
                        }
                    }
                    catch {
                        Write-Log "Server-side filter failed ($(Get-ErrorText $_)) - falling back to loading all shared mailboxes." -Level WARN
                        $found = @(Get-MailboxInfo -Params $params)
                    }
                }
                # Client-side re-check guards against loose server-side wildcard matching.
                $mailboxes = @($found | Where-Object { Test-DomainMatch -Mailbox $_ -Suffix $suffix })
                Write-Log ("Shared mailboxes loaded: {0}, matching '{1}': {2} (in {3:n1} s)" -f $found.Count, $suffix, $mailboxes.Count, $watch.Elapsed.TotalSeconds)
                $found = $null
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

        #region Backup + remove + re-add (one mailbox at a time)
        $BackupWriter = New-OutputWriter -Path $BackupFile -Header $BackupColumns
        Write-Log "Backup CSV: $BackupFile"
        if (-not $BackupOnly) {
            $ResultWriter = New-OutputWriter -Path $ResultFile -Header $ResultColumns
            Write-Log "Result CSV: $ResultFile"
        }

        $total = $mailboxes.Count
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        for ($i = 0; $i -lt $total; $i++) {
            $mbx = $mailboxes[$i]
            $mailboxes[$i] = $null   # release memory as we go
            if ($i -eq 0 -or ($i + 1) % $ProgressInterval -eq 0) {
                Write-ProgressInfo -Current ($i + 1) -Total $total -Watch $watch -Activity 'Reset FullAccess'
            }
            $id = $mbx.Guid
            $stats.Mailboxes++

            try {
                $entries = @(Get-FullAccessEntry -MailboxId $id)
            }
            catch {
                $stats.ReadFailed++
                Write-Log "$($mbx.PrimarySmtpAddress): cannot read permissions: $(Get-ErrorText $_)" -Level ERROR
                if ($ResultWriter) {
                    $row = New-EntryRow -Mailbox $mbx -Entry ([pscustomobject]@{ User = ''; AccessRights = @(); InheritanceType = ''; Deny = $false })
                    $row.Removed = 'Failed'; $row.Note = "Read: $(Get-ErrorText $_)"
                    Write-CsvRow -Writer $ResultWriter -Row $row -Columns $ResultColumns
                }
                continue
            }
            if ($entries.Count -eq 0) {
                Write-Log "[$($i + 1)/$total] $($mbx.PrimarySmtpAddress): no explicit FullAccess entry."
                continue
            }
            $stats.MailboxesWithEntries++

            $rows = @(foreach ($e in $entries) { New-EntryRow -Mailbox $mbx -Entry $e })
            $stats.Entries += $rows.Count

            # Backup of this mailbox is on disk before it is changed.
            foreach ($row in $rows) {
                Write-CsvRow -Writer $BackupWriter -Row $row -Columns $BackupColumns
                if ($row.SkipReason -eq 'Orphaned SID') {
                    Write-Log "$($mbx.PrimarySmtpAddress): '$($row.User)' is an orphaned SID - left untouched." -Level WARN -FileOnly
                }
            }
            $eligible = @($rows | Where-Object { -not $_.SkipReason })
            $users    = if ($eligible.Count) { ': ' + (($eligible | ForEach-Object { $_.User }) -join ', ') } else { '' }
            $action   = if ($BackupOnly) { 'backed up' } else { 'to reset' }
            Write-Log "[$($i + 1)/$total] $($mbx.PrimarySmtpAddress): $($rows.Count) entr(y/ies) $action, $($eligible.Count) eligible$users"
            if ($BackupOnly) { continue }
            $stats.Eligible += $eligible.Count

            # Step 1: remove
            foreach ($row in $eligible) {
                if (-not $PSCmdlet.ShouldProcess($mbx.PrimarySmtpAddress, "Remove FullAccess for $($row.User)")) {
                    $row.Removed = 'WhatIf'
                    continue
                }
                try {
                    Remove-MailboxPermission -Identity $id -User $row.User -AccessRights FullAccess `
                        -InheritanceType $row.InheritanceType -Confirm:$false -ErrorAction Stop
                    $row.Removed = 'Done'; $stats.Removed++
                    Write-Log "$($mbx.PrimarySmtpAddress): FullAccess for '$($row.User)' removed." -Level OK -FileOnly
                }
                catch {
                    $row.Removed = 'Failed'; $row.Note = "Remove: $(Get-ErrorText $_)"; $stats.Failed++
                    Write-Log "$($mbx.PrimarySmtpAddress): removing FullAccess for '$($row.User)' FAILED: $(Get-ErrorText $_) - it will not be re-added." -Level ERROR
                }
            }

            # Step 2: re-add only what was removed (or would be removed in WhatIf)
            $toAdd = @($eligible | Where-Object { $_.Removed -eq 'Done' -or $_.Removed -eq 'WhatIf' })
            if ($DelaySeconds -gt 0 -and -not $isWhatIf -and @($toAdd | Where-Object { $_.Removed -eq 'Done' }).Count) {
                Start-Sleep -Seconds $DelaySeconds
            }
            foreach ($row in $toAdd) {
                if (-not $PSCmdlet.ShouldProcess($mbx.PrimarySmtpAddress, "Add FullAccess for $($row.User) (AutoMapping $AutoMapping)")) {
                    $row.Added = 'WhatIf'; $stats.WhatIf++
                    continue
                }
                try {
                    Add-MailboxPermission -Identity $id -User $row.User -AccessRights FullAccess `
                        -InheritanceType $row.InheritanceType -AutoMapping $AutoMapping -Confirm:$false -ErrorAction Stop | Out-Null
                    $row.Added = 'Done'; $stats.Added++
                    Write-Log "$($mbx.PrimarySmtpAddress): FullAccess for '$($row.User)' added back." -Level OK -FileOnly
                }
                catch {
                    $row.Added = 'Failed'; $row.Note = "Add: $(Get-ErrorText $_)"; $stats.Failed++
                    Write-Log "$($mbx.PrimarySmtpAddress): re-adding FullAccess for '$($row.User)' FAILED: $(Get-ErrorText $_)" -Level ERROR
                }
            }

            foreach ($row in $rows) { Write-CsvRow -Writer $ResultWriter -Row $row -Columns $ResultColumns }
        }
        Write-ProgressInfo -Current $total -Total $total -Watch $watch -Activity 'Reset FullAccess'
        Write-Progress -Activity 'Reset FullAccess' -Completed
        #endregion
    }

    #region Summary
    Write-Log ('Summary: mailboxes={0}, with FullAccess={1}, read failed={2}, entries={3}, eligible={4}, removed={5}, added={6}, already present={7}, whatif={8}, failed={9}, trustees resolved={10}' -f `
        $stats.Mailboxes, $stats.MailboxesWithEntries, $stats.ReadFailed, $stats.Entries, $stats.Eligible,
        $stats.Removed, $stats.Added, $stats.NotNeeded, $stats.WhatIf, $stats.Failed, $TrusteeCache.Count) `
        -Level $(if ($stats.Failed -or $stats.ReadFailed) { 'WARN' } else { 'OK' })
    if ($stats.Failed -and $PSCmdlet.ParameterSetName -ne 'Restore') {
        Write-Log "Some entries failed. Missing permissions can be restored with: .\$scriptName.ps1 -RestoreFromCsv '$BackupFile'" -Level WARN
    }
    Write-Log 'Done.'
    #endregion
}
finally {
    foreach ($w in $ResultWriter, $BackupWriter, $LogWriter) {
        if ($w) { $w.Dispose() }
    }
}
