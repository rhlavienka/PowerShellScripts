<#
.SYNOPSIS
    Analyzes Exchange SMTP protocol logs to report relay usage by client IP and receive connector.

.DESCRIPTION
    Parses the most recent Exchange FrontEnd SMTP Receive protocol logs and counts
    sent messages (MAIL FROM), grouped by client source IP address and Receive Connector.
    Results are exported to CSV, opened in Out-GridView, and printed to the console.

    The log folder is not hardcoded - it is resolved at runtime from the local
    Frontend Transport service configuration (Get-FrontendTransportService
    ReceiveProtocolLogPath), so the script follows whatever path is actually
    configured on the server instead of assuming the Exchange default.

.PARAMETER RelayOnly
    If set, only Receive Connectors identified as relay connectors are included in
    the output - i.e. connectors with AnonymousUsers permission group where
    "NT AUTHORITY\ANONYMOUS LOGON" has been granted the "Ms-Exch-SMTP-Accept-Any-Recipient"
    extended right (the same pattern used by New-OpenRelayConnector.ps1 in this repo).
    Without this switch, all connectors are included, same as before.

.PARAMETER LastLogs
    Number of most recent protocol log files to process. Default: 20.

.PARAMETER CsvFile
    Path to the output CSV file. Default: "C:\Temp\SMTP_Relay_Usage_<timestamp>.csv"
    (or "C:\Temp\SMTP_RelayOnly_Usage_<timestamp>.csv" when -RelayOnly is used).

.EXAMPLE
    .\get-SMTPlog_IPlist.ps1
    Processes the 20 most recent SMTP Receive protocol logs on the local Exchange
    server, exports a CSV to C:\Temp, and opens the results in Out-GridView.

.EXAMPLE
    .\get-SMTPlog_IPlist.ps1 -RelayOnly
    Same as above, but only includes traffic through connectors configured for
    anonymous relay (Ms-Exch-SMTP-Accept-Any-Recipient granted to Anonymous Logon).

.NOTES
    Version: 1.1 (2026-08-24)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Changelog:
      1.1 (2026-08-24) - Added -RelayOnly switch to filter results down to relay
                          connectors (identified via the Ms-Exch-SMTP-Accept-Any-Recipient
                          AD permission on Anonymous Logon). Log folder is now resolved
                          dynamically from Get-FrontendTransportService instead of a
                          hardcoded path.
      1.0 (2026-08-20) - Initial version.

    Run in Exchange Management Shell on the Exchange server whose logs you want to analyze.
#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Only include relay connectors (Anonymous + Ms-Exch-SMTP-Accept-Any-Recipient)")]
    [switch]$RelayOnly,

    [Parameter(Mandatory=$false, HelpMessage="Number of most recent protocol log files to process")]
    [int]$LastLogs = 20,

    [Parameter(Mandatory=$false, HelpMessage="Path to the output CSV file")]
    [string]$CsvFile
)

# Resolve the SMTP Receive protocol log folder from the local Frontend Transport service configuration
try {
    $LogPath = (Get-FrontendTransportService $env:COMPUTERNAME -ErrorAction Stop).ReceiveProtocolLogPath.PathName
}
catch {
    Write-Host "X Failed to read ReceiveProtocolLogPath from Get-FrontendTransportService: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Make sure this script runs in Exchange Management Shell." -ForegroundColor Yellow
    exit 1
}

if (-not $LogPath -or -not (Test-Path $LogPath)) {
    Write-Host "X SMTP Receive protocol log folder not found: $LogPath" -ForegroundColor Red
    exit 1
}

Write-Host "Log folder     : $LogPath" -ForegroundColor Cyan

# Output file
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not $CsvFile) {
    $Suffix  = if ($RelayOnly) { "RelayOnly" } else { "Relay_Usage" }
    $CsvFile = "C:\Temp\SMTP_${Suffix}_$TimeStamp.csv"
}

# Identify relay connectors (only needed when -RelayOnly is used)
$RelayConnectors = @()
if ($RelayOnly) {
    Write-Host ""
    Write-Host "Identifying relay connectors (Anonymous + Ms-Exch-SMTP-Accept-Any-Recipient)..." -ForegroundColor Cyan

    $RelayConnectors = Get-ReceiveConnector -Server $env:COMPUTERNAME |
        Where-Object {
            $_.PermissionGroups -like '*AnonymousUsers*' -and
            (Get-ADPermission $_.Identity |
                Where-Object {
                    $_.User -like '*ANONYMOUS LOGON*' -and
                    $_.ExtendedRights -match 'Ms-Exch-SMTP-Accept-Any-Recipient'
                })
        } |
        Select-Object -ExpandProperty Identity

    if (-not $RelayConnectors) {
        Write-Warning "No relay connectors found on $env:COMPUTERNAME - RelayOnly output will be empty."
    }
    else {
        Write-Host "Relay connectors found:" -ForegroundColor Green
        $RelayConnectors | ForEach-Object { Write-Host "  $_" }
    }
}

Write-Host ""
Write-Host "Loading SMTP Protocol Logs..." -ForegroundColor Cyan

$AllRecords = foreach ($File in (Get-ChildItem "$LogPath\*.log" |
                                 Sort-Object LastWriteTime |
                                 Select-Object -Last $LastLogs))
{
    Write-Host "Processing: $($File.Name)" -ForegroundColor Yellow

    $Lines = Get-Content $File.FullName

    $FieldsLine = $Lines |
        Where-Object { $_ -like '#Fields:*' } |
        Select-Object -First 1

    if (-not $FieldsLine)
    {
        Write-Warning "Skipping $($File.Name) - no #Fields header found"
        continue
    }

    $Headers = ($FieldsLine -replace '^#Fields:\s*','') -split ','

    $DataLines = $Lines |
        Where-Object {
            $_ -and
            $_ -notmatch '^#'
        }

    $Records = $DataLines |
        ConvertFrom-Csv -Header $Headers

    foreach ($Record in $Records)
    {
        [PSCustomObject]@{
            Connector      = $Record.'connector-id'
            RemoteEndpoint = $Record.'remote-endpoint'
            Event          = $Record.event
            Data           = $Record.data
        }
    }
}

Write-Host ""
Write-Host "Analyzing MAIL FROM records..." -ForegroundColor Cyan

$Results = $AllRecords |
    Where-Object {

        # Actual message send
        $_.Data -match 'MAIL FROM:'
    } |
    Where-Object {

        # When -RelayOnly is set, keep only traffic through relay connectors
        (-not $RelayOnly) -or ($_.Connector -in $RelayConnectors)
    } |
    ForEach-Object {

        $IP = $null

        if ($_.RemoteEndpoint)
        {
            $IP = ($_.RemoteEndpoint -split ':')[0]
        }

        [PSCustomObject]@{
            ClientIP  = $IP
            Connector = $_.Connector
        }
    } |
    Where-Object {
        $_.ClientIP
    } |
    Group-Object ClientIP,Connector |
    Sort-Object Count -Descending |
    Select-Object @{
            Name = "Messages"
            Expression = { $_.Count }
        },
        @{
            Name = "ClientIP"
            Expression = { $_.Group[0].ClientIP }
        },
        @{
            Name = "Connector"
            Expression = { $_.Group[0].Connector }
        }

# Export CSV
$Results |
    Export-Csv `
        -Path $CsvFile `
        -NoTypeInformation `
        -Encoding UTF8

# GridView
$GridTitle = if ($RelayOnly) { "SMTP Relay Usage (relay connectors only)" } else { "SMTP Relay Usage" }
$Results |
    Out-GridView -Title $GridTitle

# Console
$Results | Format-Table -AutoSize

Write-Host ""
Write-Host "Results found : $($Results.Count)" -ForegroundColor Green
Write-Host "CSV exported  : $CsvFile" -ForegroundColor Green
