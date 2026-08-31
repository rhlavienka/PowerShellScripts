<#
.SYNOPSIS
    Analyzes Exchange IMAP4 protocol logs to report client logon usage by IP address.

.DESCRIPTION
    Parses the most recent Exchange IMAP4 protocol logs from the Client
    Access / front-end log files ("IMAP4*.LOG") and counts client LOGIN /
    AUTHENTICATE commands, grouped by account, client source IP address,
    and the server endpoint (IP:Port) the client connected to. Mailbox
    server back-end logs ("IMAP4BE*.LOG") are excluded, since they just
    mirror the same client session on the server-to-server hop and would
    otherwise double-count it. Results are exported to CSV, opened in
    Out-GridView, and printed to the console.

    If -LogPath is not specified, the script auto-detects the IMAP4 log
    folder(s) by probing $env:ExchangeInstallPath and the default Exchange
    installation paths, then recursively locating folders that contain
    IMAP4*.LOG files.

    IMAP4 protocol logging is disabled by default on Exchange Server and must
    be enabled before this script has any data to analyze. Run on the
    Exchange server (or against every Mailbox server if load balanced) as
    Administrator:

        Set-ImapSettings -ProtocolLogEnabled $true `
            -LogFileLocation "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4"
        Restart-Service MSExchangeIMAP4
        Restart-Service MSExchangeIMAP4BE

    The log folder normally requires elevated (Administrator) rights to
    read. If your account uses eligible/PIM-style administrative rights
    rather than a permanent Administrators membership, Windows may deny
    access until that elevation has been activated - typically by opening
    the folder once in File Explorer (or launching Explorer "as
    administrator") and confirming the elevation/UAC prompt. Do this before
    running the script if you hit an "Access is denied" error.

.PARAMETER LogPath
    Explicit path to the folder containing IMAP4 protocol logs. When
    omitted, the script attempts to auto-detect it.

.PARAMETER LastLogs
    Number of most recent logs (across all discovered log folders) to
    process. Defaults to 0, which processes every log file found (no
    limit). Pass a positive number to cap it to only the N most recent
    logs.

.PARAMETER CsvFile
    Output CSV path. Defaults to a timestamped file in C:\Temp.

.EXAMPLE
    .\get-IMAPlog_IPlist.ps1
    Auto-detects the IMAP4 log folder(s) on the local Exchange server,
    processes every IMAP4 log file found, exports a CSV to C:\Temp, and
    opens the results in Out-GridView.

.EXAMPLE
    .\get-IMAPlog_IPlist.ps1 -LogPath "D:\ExchangeLogs\Imap4" -LastLogs 50
    Only processes the 50 most recent logs instead of all of them.

.NOTES
    Version: 1.1 (2026-08-25)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    Run locally on an Exchange server with elevated (Administrator /
    activated eligible role) access to the IMAP4 protocol log folder.

    Changelog:
      1.1 (2026-08-25) - Fixed field mapping to match the real IMAP4 log schema (sIp/cIp/command/parameters, not ServerIP/ClientIP/event/data); added automatic log-folder discovery; added clear handling for access-denied/eligible-rights scenarios; added diagnostic counters; added the account (mailbox/user) identity as an output column; excluded backend (IMAP4BE*.LOG) connections since they duplicate the front-end session; -LastLogs now defaults to 0, processing every log file found instead of being capped at 20.
      1.0 (2026-08-21) - Initial version
#>

param(
    [string]$LogPath,
    [int]$LastLogs = 0,
    [string]$CsvFile
)

# ---------------------------------------------------------------------------
# Elevation advisory
# ---------------------------------------------------------------------------
$CurrentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Warning "This session is not running elevated. The IMAP4 log folder typically requires Administrator rights to read."
    Write-Warning "If your account uses eligible/PIM-style admin rights, open the log folder once in File Explorer (or run Explorer as Administrator) to activate elevation, then re-run this script elevated."
}

# ---------------------------------------------------------------------------
# Log folder discovery
# ---------------------------------------------------------------------------
function Find-Imap4LogFolders
{
    $CandidateRoots = @(
        $env:ExchangeInstallPath
        "C:\Program Files\Microsoft\Exchange Server\V15"
        "C:\Program Files\Microsoft\Exchange Server\V14"
        "D:\Program Files\Microsoft\Exchange Server\V15"
    ) | Where-Object { $_ } | Select-Object -Unique

    $SearchErrors = @()

    $Found = foreach ($Root in $CandidateRoots)
    {
        $LoggingRoot = Join-Path $Root "Logging"
        if (-not (Test-Path $LoggingRoot)) { continue }

        Get-ChildItem -Path $LoggingRoot -Recurse -Filter "IMAP4*.LOG" -File `
                -ErrorAction SilentlyContinue -ErrorVariable +SearchErrors |
            Select-Object -ExpandProperty DirectoryName -Unique
    }

    if ($SearchErrors -and -not $Found)
    {
        Write-Warning "Some folders were inaccessible while searching for IMAP4 logs (Access denied)."
        Write-Warning "If your account uses eligible/PIM-style admin rights, open the expected log folder once in File Explorer to activate elevation, then re-run this script."
    }

    $Found | Select-Object -Unique
}

if ($LogPath)
{
    $LogFolders = @($LogPath)
}
else
{
    Write-Host "No -LogPath specified, attempting to auto-detect the IMAP4 log folder..." -ForegroundColor Cyan
    $LogFolders = @(Find-Imap4LogFolders)

    if (-not $LogFolders)
    {
        Write-Error "Could not auto-detect an IMAP4 protocol log folder. Specify one explicitly with -LogPath."
        return
    }

    Write-Host "Found log folder(s): $($LogFolders -join ', ')" -ForegroundColor Green
}

if (-not $CsvFile)
{
    $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $CsvFile = "C:\Temp\IMAP_Logon_Usage_$TimeStamp.csv"
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Split-IPEndpoint
{
    param([string]$Endpoint)

    # Handles both IPv4 ("1.2.3.4:993") and bracketed IPv6 ("[fe80::1%13]:993") forms
    if ($Endpoint -match '^(?<IP>\[.+\]|[^:]+):(?<Port>\d+)$')
    {
        [PSCustomObject]@{ IP = $Matches.IP; Port = $Matches.Port }
    }
    else
    {
        [PSCustomObject]@{ IP = $Endpoint; Port = $null }
    }
}

# ---------------------------------------------------------------------------
# Load logs
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Loading IMAP4 Protocol Logs..." -ForegroundColor Cyan

try
{
    # Backend (IMAP4BE*.LOG) entries are the Mailbox-server side of the same
    # client session the front-end (IMAP4*.LOG) already logged - excluded
    # here so backend connections don't show up in the output.
    $LogFiles = $LogFolders |
        ForEach-Object { Get-ChildItem -Path "$_\*.LOG" -Exclude "IMAP4BE*.LOG" -ErrorAction Stop } |
        Sort-Object LastWriteTime

    # -LastLogs 0 (or negative) means "no limit, process every log file found"
    if ($LastLogs -gt 0)
    {
        $LogFiles = $LogFiles | Select-Object -Last $LastLogs
    }
}
catch [System.UnauthorizedAccessException]
{
    Write-Error "Access is denied to the log folder. If your account uses eligible/PIM-style admin rights, open the folder manually in File Explorer first to activate elevation, then re-run this script."
    return
}

$AllRecords = foreach ($File in $LogFiles)
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

    # Only real data rows start with an ISO-8601 timestamp; this also skips
    # the metadata ('#...') lines and any stray duplicate header line.
    $DataLines = $Lines |
        Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}' }

    $Records = $DataLines |
        ConvertFrom-Csv -Header $Headers

    foreach ($Record in $Records)
    {
        $ClientEndpoint = Split-IPEndpoint $Record.cIp
        $ServerEndpoint = Split-IPEndpoint $Record.sIp

        [PSCustomObject]@{
            ClientIP   = $ClientEndpoint.IP
            ServerIP   = $ServerEndpoint.IP
            ServerPort = $ServerEndpoint.Port
            Command    = $Record.command
            User       = $Record.user
        }
    }
}

Write-Host ""
Write-Host "Analyzing LOGIN / AUTHENTICATE records..." -ForegroundColor Cyan
Write-Host "Log files processed  : $($LogFiles.Count)" -ForegroundColor DarkGray
Write-Host "Total entries parsed : $($AllRecords.Count)" -ForegroundColor DarkGray

$LoginRecords = $AllRecords |
    Where-Object {

        # Client logon commands (username/password or SASL)
        $_.Command -match '(?i)^(login|authenticate)$'
    } |
    Where-Object {
        $_.ClientIP
    }

Write-Host "LOGIN/AUTHENTICATE   : $($LoginRecords.Count)" -ForegroundColor DarkGray

$Results = $LoginRecords |
    Group-Object ClientIP,ServerIP,ServerPort,User |
    Sort-Object Count -Descending |
    Select-Object @{
            Name = "Logons"
            Expression = { $_.Count }
        },
        @{
            Name = "Account"
            Expression = { $_.Group[0].User }
        },
        @{
            Name = "ClientIP"
            Expression = { $_.Group[0].ClientIP }
        },
        @{
            Name = "ServerEndpoint"
            Expression = { "$($_.Group[0].ServerIP):$($_.Group[0].ServerPort)" }
        }

# Export CSV
$Results |
    Export-Csv `
        -Path $CsvFile `
        -NoTypeInformation `
        -Encoding UTF8

# GridView
$Results |
    Out-GridView -Title "IMAP4 Logon Usage"

# Console
$Results | Format-Table -AutoSize

Write-Host ""
Write-Host "Results found : $($Results.Count)" -ForegroundColor Green
Write-Host "CSV exported  : $CsvFile" -ForegroundColor Green
