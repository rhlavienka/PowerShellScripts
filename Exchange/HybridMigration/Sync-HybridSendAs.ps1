<#
.SYNOPSIS
    Adds missing SendAs permissions in Exchange Online for migrated users who
    have SendAs on mailboxes that are still on-prem (typically shared mailboxes).

.DESCRIPTION
    1. Reads a CSV with an EmailAddress column (primary SMTP address of the
       migrated users).
    2. On-prem (session with prefix OnPrem): finds the user's RemoteMailbox
       object (SamAccountName, UPN) and scans all on-prem mailboxes on which
       the user has an explicit (not inherited, not Deny) Send-As right.
    3. EXO (session with prefix Cloud): verifies that the user has a mailbox in
       EXO and that the on-prem mailbox is visible in EXO (MailUser); if SendAs
       is missing there, adds it with Add-RecipientPermission.
       Trustee = the user's UPN (UPN may differ from PrimarySmtpAddress).
    4. Writes the result to a CSV report. MailboxType is the on-prem type,
       CloudRecipientType the actual type of the object in EXO (MailUser
       while the mailbox is still on-prem, e.g. SharedMailbox once migrated).

    Cmdlets are imported with prefixes so on-prem and EXO do not collide:
      On-prem: Get-OnPremRemoteMailbox, Get-OnPremMailbox, Get-OnPremADPermission
      EXO:     Get-CloudMailbox, Get-CloudRecipient, Get-CloudRecipientPermission,
               Add-CloudRecipientPermission

    Run in plain Windows PowerShell 5.1 (not in the Exchange Management Shell)
    with the ExchangeOnlineManagement module (v3+).

.PARAMETER CsvPath
    Path to the input CSV (';' delimiter) with an EmailAddress column.

.PARAMETER ExchangeServer
    FQDN of the on-prem Exchange server. Default: the local server.

.PARAMETER ExoAdminUpn
    Admin UPN for Connect-ExchangeOnline (optional).

.PARAMETER SharedOnly
    Scan only on-prem SharedMailbox objects (faster). Without this switch all
    on-prem mailboxes are scanned.

.PARAMETER ReportPath
    Path of the output report (UTF-8, ';' delimiter).

.EXAMPLE
    .\Sync-HybridSendAs.ps1 -CsvPath .\users.csv -WhatIf
    Only shows what would be added.

.EXAMPLE
    .\Sync-HybridSendAs.ps1 -CsvPath .\users.csv -SharedOnly -ExoAdminUpn admin@contoso.com

.NOTES
    - Requires an on-prem account with permission to read mailboxes and
      AD permissions (Kerberos remoting to the Exchange server).
    - Requires an EXO role that allows Add-RecipientPermission
      (e.g. Recipient Management).

    Version: 1.1 (2026-11-19)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [string]$ExchangeServer = ([System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName),

    [string]$ExoAdminUpn,

    [switch]$SharedOnly,

    [string]$ReportPath = (Join-Path $PSScriptRoot ("SendAsReport_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)))
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host ("[{0:HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

# --- Input -------------------------------------------------------------------
if (-not (Test-Path $CsvPath)) { throw "CSV '$CsvPath' does not exist." }
$inputUsers = Import-Csv -Path $CsvPath -Delimiter ';' |
    Where-Object { $_.EmailAddress -and $_.EmailAddress.Trim() } |
    ForEach-Object { $_.EmailAddress.Trim() } |
    Sort-Object -Unique
if (-not $inputUsers) { throw "CSV contains no values in the EmailAddress column." }
Write-Log "Users loaded from CSV: $(@($inputUsers).Count)"

$report = New-Object System.Collections.Generic.List[object]
$onPremSession = $null

try {
    # --- Connect on-prem ------------------------------------------------------
    Write-Log "Connecting to on-prem Exchange: $ExchangeServer"
    $onPremSession = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos
    Import-PSSession -Session $onPremSession -Prefix OnPrem -DisableNameChecking -AllowClobber `
        -CommandName Get-RemoteMailbox, Get-Mailbox, Get-ADPermission | Out-Null

    # --- Connect EXO ----------------------------------------------------------
    Write-Log "Connecting to Exchange Online"
    Import-Module ExchangeOnlineManagement
    $exoParams = @{
        Prefix      = 'Cloud'
        ShowBanner  = $false
        CommandName = 'Get-Mailbox', 'Get-Recipient', 'Get-RecipientPermission', 'Add-RecipientPermission'
    }
    if ($ExoAdminUpn) { $exoParams.UserPrincipalName = $ExoAdminUpn }
    Connect-ExchangeOnline @exoParams

    # --- 1. Resolve users on-prem --------------------------------------------
    # key = SamAccountName (lowercase), value = user info
    $trustees = @{}
    foreach ($email in $inputUsers) {
        $rm = Get-OnPremRemoteMailbox -Identity $email -ErrorAction SilentlyContinue
        if (-not $rm) {
            Write-Log "$email - not a RemoteMailbox on-prem (not migrated or does not exist), skipping." 'WARN'
            $report.Add([pscustomobject]@{
                User = $email; UserUPN = ''; Mailbox = ''; MailboxType = ''; CloudRecipientType = ''
                Status = 'UserNotRemoteMailbox'; Detail = 'Get-RemoteMailbox found no object'
            })
            continue
        }

        $exoMbx = Get-CloudMailbox -Identity $rm.UserPrincipalName -ErrorAction SilentlyContinue
        if (-not $exoMbx) {
            Write-Log "$email ($($rm.UserPrincipalName)) - has no mailbox in EXO, skipping." 'WARN'
            $report.Add([pscustomobject]@{
                User = $email; UserUPN = $rm.UserPrincipalName; Mailbox = ''; MailboxType = ''; CloudRecipientType = ''
                Status = 'UserNotInEXO'; Detail = 'Get-Mailbox in EXO found no object'
            })
            continue
        }

        $trustees[$rm.SamAccountName.ToLower()] = [pscustomobject]@{
            Email = $email
            Sam   = $rm.SamAccountName
            UPN   = $exoMbx.UserPrincipalName
        }
        Write-Log "$email -> UPN $($exoMbx.UserPrincipalName), sAM $($rm.SamAccountName)"
    }

    if ($trustees.Count -eq 0) {
        Write-Log "No valid user to process." 'WARN'
        return
    }

    # --- 2. Find SendAs on on-prem mailboxes ---------------------------------
    $mbxParams = @{ ResultSize = 'Unlimited' }
    if ($SharedOnly) { $mbxParams.RecipientTypeDetails = 'SharedMailbox' }
    $onPremMailboxes = @(Get-OnPremMailbox @mbxParams | Select-Object DisplayName, PrimarySmtpAddress, DistinguishedName, RecipientTypeDetails)
    Write-Log "Scanning SendAs on $($onPremMailboxes.Count) on-prem mailboxes (this may take a while)..."

    $pairs = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($mbx in $onPremMailboxes) {
        $i++
        Write-Progress -Activity 'Get-ADPermission' -Status "$i / $($onPremMailboxes.Count): $($mbx.PrimarySmtpAddress)" `
            -PercentComplete (($i / [math]::Max($onPremMailboxes.Count, 1)) * 100)

        try {
            $aces = Get-OnPremADPermission -Identity $mbx.DistinguishedName
        }
        catch {
            Write-Log "Get-ADPermission failed for $($mbx.PrimarySmtpAddress): $($_.Exception.Message)" 'ERROR'
            continue
        }

        foreach ($ace in $aces) {
            if ($ace.IsInherited -or $ace.Deny) { continue }
            if ((@($ace.ExtendedRights) -join ',') -notmatch 'Send-As') { continue }

            # After deserialization User is the string "DOMAIN\sam"
            $sam = ([string]$ace.User).Split('\')[-1].ToLower()
            if ($trustees.ContainsKey($sam)) {
                $pairs.Add([pscustomobject]@{
                    Mailbox     = [string]$mbx.PrimarySmtpAddress
                    MailboxType = [string]$mbx.RecipientTypeDetails
                    Trustee     = $trustees[$sam]
                    AceUser     = [string]$ace.User
                })
            }
        }
    }
    Write-Progress -Activity 'Get-ADPermission' -Completed
    Write-Log "SendAs pairs found (on-prem mailbox -> migrated user): $($pairs.Count)"

    # Users without any SendAs are written to the report as well
    foreach ($t in $trustees.Values) {
        if (-not ($pairs | Where-Object { $_.Trustee.Sam -eq $t.Sam })) {
            $report.Add([pscustomobject]@{
                User = $t.Email; UserUPN = $t.UPN; Mailbox = ''; MailboxType = ''; CloudRecipientType = ''
                Status = 'NoOnPremSendAs'; Detail = 'User has no SendAs on any on-prem mailbox'
            })
        }
    }

    # --- 3. Add SendAs in EXO ------------------------------------------------
    foreach ($p in $pairs) {
        $row = [pscustomobject]@{
            User = $p.Trustee.Email; UserUPN = $p.Trustee.UPN
            Mailbox = $p.Mailbox; MailboxType = $p.MailboxType; CloudRecipientType = ''
            Status = ''; Detail = "On-prem ACE: $($p.AceUser)"
        }

        $cloudObj = Get-CloudRecipient -Identity $p.Mailbox -ErrorAction SilentlyContinue
        if (-not $cloudObj) {
            Write-Log "$($p.Mailbox) - not visible in EXO (Entra Connect sync?)." 'WARN'
            $row.Status = 'MailboxNotInEXO'
            $report.Add($row); continue
        }
        # Actual object type in EXO: MailUser while the mailbox is still on-prem,
        # SharedMailbox / UserMailbox / ... once it has been migrated
        $row.CloudRecipientType = [string]$cloudObj.RecipientTypeDetails

        $existing = Get-CloudRecipientPermission -Identity $p.Mailbox -AccessRights SendAs -ErrorAction SilentlyContinue |
            Where-Object { $_.Trustee -eq $p.Trustee.UPN -and $_.AccessControlType -eq 'Allow' }
        if ($existing) {
            Write-Log "$($p.Mailbox) <- $($p.Trustee.UPN): SendAs already exists in EXO." 'OK'
            $row.Status = 'AlreadyExists'
            $report.Add($row); continue
        }

        if ($PSCmdlet.ShouldProcess("$($p.Mailbox) (EXO)", "Add-RecipientPermission SendAs for $($p.Trustee.UPN)")) {
            try {
                # EXO cmdlets report errors as non-terminating, hence -ErrorAction Stop
                Add-CloudRecipientPermission -Identity $p.Mailbox -Trustee $p.Trustee.UPN -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null

                $verify = Get-CloudRecipientPermission -Identity $p.Mailbox -AccessRights SendAs -ErrorAction SilentlyContinue |
                    Where-Object { $_.Trustee -eq $p.Trustee.UPN -and $_.AccessControlType -eq 'Allow' }
                if ($verify) {
                    Write-Log "$($p.Mailbox) <- $($p.Trustee.UPN): SendAs added in EXO." 'OK'
                    $row.Status = 'Added'
                }
                else {
                    Write-Log "$($p.Mailbox) <- $($p.Trustee.UPN): Add returned no error, but SendAs is not visible in EXO." 'WARN'
                    $row.Status = 'NotVerified'
                }
            }
            catch {
                Write-Log "$($p.Mailbox) <- $($p.Trustee.UPN): error - $($_.Exception.Message)" 'ERROR'
                $row.Status = 'Error'
                $row.Detail = $_.Exception.Message
            }
        }
        else {
            $row.Status = 'WhatIf'
        }
        $report.Add($row)
    }
}
finally {
    if ($report.Count -gt 0) {
        # -WhatIf:$false - the report and disconnect must run even with -WhatIf
        $report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';' -WhatIf:$false
        Write-Log "Report: $ReportPath"
        $report | Group-Object Status | Select-Object Name, Count | Format-Table -AutoSize | Out-String | Write-Host
    }
    if ($onPremSession) { Remove-PSSession $onPremSession -ErrorAction SilentlyContinue -WhatIf:$false }
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue -WhatIf:$false
}
