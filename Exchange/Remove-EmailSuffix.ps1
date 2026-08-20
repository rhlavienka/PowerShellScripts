<#
.SYNOPSIS
    Fast (parallel) removal of email addresses with a given domain suffix from
    Exchange Online recipients.

.DESCRIPTION
    Removes all email addresses ending in the given domain suffix from EXO recipients
    (mailboxes, mail users, mail contacts, distribution/security groups, Microsoft 365
    groups). Recipients are fetched with a single bulk query, split into batches, and
    processed in parallel via ForEach-Object -Parallel.

    Parallelization works WITHOUT a registered (app-only) application - it relies
    on Connect-ExchangeOnline, which has multithreading enabled natively and
    is designed exactly for the ForEach-Object -Parallel scenario. Each runspace
    is isolated and does not share the main thread's connection, so each thread
    first verifies/restores the connection (without a repeated MFA prompt - the
    already existing auth token is recycled).

    Pattern source: https://techcommunity.microsoft.com/blog/exchange/more-efficient-bulk-operations-with-powershell-parallelism/4409693

    You MUST FIRST connect interactively in the main window:
        Connect-ExchangeOnline

.PARAMETER SuffixToRemove
    Domain suffix (without "@") whose addresses should be removed from recipients.
    Default: "old-domain.com"

.PARAMETER ThrottleLimit
    Number of parallel runspaces (batches) used to process recipients. Default: 8

.PARAMETER AdminUPN
    UPN of the administrator account that is already connected interactively in the
    main window; used to reconnect inside a parallel runspace if its own connection
    check fails.

.PARAMETER WhatIf
    Dry-run switch. When set, the script only reports which addresses would be
    removed, without making any changes.

.PARAMETER LogPath
    Path to the CSV file where per-recipient results (OK / FAIL / DryRun) are logged.
    Default: ".\Remove-EmailSuffix-<timestamp>.csv"

.EXAMPLE
    Connect-ExchangeOnline -UserPrincipalName admin@contoso.com
    .\Remove-EmailSuffix.ps1 -AdminUPN admin@contoso.com -WhatIf
    Connects interactively, then previews which "@old-domain.com" addresses would be
    removed, without making any changes.

.EXAMPLE
    .\Remove-EmailSuffix.ps1 -AdminUPN admin@contoso.com -SuffixToRemove "old-domain.com" -ThrottleLimit 8
    Removes all "@old-domain.com" addresses from EXO recipients using 8 parallel
    runspaces and logs the results to a timestamped CSV file.

.NOTES
    Version: 1.0 (2026-08-20)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>

#Requires -Version 7.0

param(
    [string]$SuffixToRemove = "old-domain.com",
    [int]$ThrottleLimit = 8,
    [Parameter(Mandatory=$true)]
    [string]$AdminUPN,                 # UPN of the admin you are already connected with interactively
    [switch]$WhatIf,                   # DRY-RUN switch
    [string]$LogPath = ".\Remove-EmailSuffix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

# ---------- 0) Verify we are already connected in the main thread ----------
try {
    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
} catch {
    Write-Host "Not connected to Exchange Online. Run this first:" -ForegroundColor Red
    Write-Host "  Connect-ExchangeOnline -UserPrincipalName $AdminUPN" -ForegroundColor Yellow
    return
}

if ($WhatIf) {
    Write-Host "*** DRY-RUN MODE - no changes will be written ***" -ForegroundColor Magenta
}

# ---------- 1) ONE bulk query instead of 10 separate ones ----------
$recipientTypesToCheck = @(
    "GroupMailbox","GuestMailUser","MailContact","MailUniversalDistributionGroup",
    "MailUniversalSecurityGroup","MailUser","RoomMailbox","SharedMailbox",
    "SchedulingMailbox","UserMailbox"
)

Write-Host "Fetching recipients..." -ForegroundColor Cyan
$affected = Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $recipientTypesToCheck |
    Where-Object { $_.EmailAddresses -like "*@$SuffixToRemove" } |
    Select-Object ExternalDirectoryObjectId, RecipientTypeDetails, PrimarySmtpAddress, EmailAddresses

Write-Host "Found $($affected.Count) affected recipients." -ForegroundColor Yellow
if (-not $affected) { Write-Host "Nothing to do." -ForegroundColor Green; return }

# thread-safe collection for logging from parallel threads
$results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

# ---------- 2) Split into batches according to ThrottleLimit ----------
# each runspace processes its batch sequentially, but batches run in parallel
$chunks = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $ThrottleLimit; $i++) { $chunks.Add([System.Collections.Generic.List[object]]::new()) }
for ($i = 0; $i -lt $affected.Count; $i++) { $chunks[$i % $ThrottleLimit].Add($affected[$i]) }

# ---------- 3) Parallel processing ----------
$chunks | Where-Object { $_.Count -gt 0 } | ForEach-Object -Parallel {

    $suffix     = $using:SuffixToRemove
    $whatIfMode = $using:WhatIf
    $resultsBag = $using:results
    $adminUpn   = $using:AdminUPN
    $batch      = $_

    # Each runspace is isolated -> verify/restore the connection.
    # Multithreading is enabled natively, so it recycles the already
    # existing auth token and does NOT ask again for MFA/password.
    try {
        $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
    } catch {
        Connect-ExchangeOnline -UserPrincipalName $adminUpn -ShowBanner:$false -ErrorAction Stop
    }

    foreach ($recipient in $batch) {
        try {
            $full = Get-Recipient -Identity $recipient.ExternalDirectoryObjectId -ErrorAction Stop
            $removed = $full.EmailAddresses | Where-Object { $_ -like "*@$suffix" }
            $newAddresses = $full.EmailAddresses | Where-Object { $_ -notlike "*@$suffix" }

            if (-not $removed) { continue }  # may have already been changed in the meantime

            if ($whatIfMode) {
                $resultsBag.Add([pscustomobject]@{
                    Timestamp = Get-Date; Identity = $recipient.PrimarySmtpAddress
                    Type = $recipient.RecipientTypeDetails; Action = "WOULD REMOVE"
                    RemovedAddresses = ($removed -join "; "); Status = "DryRun"; Error = ""
                })
                Write-Host "[DRY-RUN] $($recipient.PrimarySmtpAddress) -> would remove: $($removed -join ', ')" -ForegroundColor DarkYellow
                continue
            }

            switch ($recipient.RecipientTypeDetails) {
                { $_ -in @("MailUser","GuestMailUser") }                                     { Set-MailUser -Identity $recipient.ExternalDirectoryObjectId -EmailAddresses $newAddresses -Confirm:$false }
                "MailContact"                                                                 { Set-MailContact -Identity $recipient.ExternalDirectoryObjectId -EmailAddresses $newAddresses -Confirm:$false }
                { $_ -in @("MailUniversalDistributionGroup","MailUniversalSecurityGroup") }    { Set-DistributionGroup -Identity $recipient.ExternalDirectoryObjectId -EmailAddresses $newAddresses -Confirm:$false }
                { $_ -in @("UserMailbox","SharedMailbox","RoomMailbox","SchedulingMailbox") }  { Set-Mailbox -Identity $recipient.ExternalDirectoryObjectId -EmailAddresses $newAddresses -Confirm:$false }
                "GroupMailbox"                                                                { Set-UnifiedGroup -Identity $recipient.ExternalDirectoryObjectId -EmailAddresses $newAddresses -Confirm:$false }
            }

            $resultsBag.Add([pscustomobject]@{
                Timestamp = Get-Date; Identity = $recipient.PrimarySmtpAddress
                Type = $recipient.RecipientTypeDetails; Action = "REMOVED"
                RemovedAddresses = ($removed -join "; "); Status = "OK"; Error = ""
            })
            Write-Host "[OK] $($recipient.PrimarySmtpAddress)" -ForegroundColor Green
        }
        catch {
            $resultsBag.Add([pscustomobject]@{
                Timestamp = Get-Date; Identity = $recipient.PrimarySmtpAddress
                Type = $recipient.RecipientTypeDetails; Action = "ERROR"
                RemovedAddresses = ""; Status = "FAIL"; Error = $_.Exception.Message
            })
            Write-Host "[FAIL] $($recipient.PrimarySmtpAddress): $_" -ForegroundColor Red
        }
    }

} -ThrottleLimit $ThrottleLimit

# ---------- 4) Export log ----------
$results | Sort-Object Timestamp | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
Write-Host "`nLog saved to: $LogPath" -ForegroundColor Cyan
Write-Host "Done. Processed: $($results.Count), errors: $(($results | Where-Object Status -eq 'FAIL').Count)" -ForegroundColor Cyan