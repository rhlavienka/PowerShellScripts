<#
.SYNOPSIS
    Identifies currently connected clients on an Exchange server (by local port, default 443),
    their authentication method (Kerberos/NTLM) from the Security log, and the user identity
    from the IIS log. Optionally also resolves OAuth bearer-token clients from the Exchange
    HttpProxy logs (-IncludeOAuth).

.DESCRIPTION
    Run directly ON the Exchange server (front-end), in a PowerShell console with
    Administrator rights (without elevation the Security log is not readable - the
    result is "No events found" even when the log contains matching records).

    Windows-authenticated clients (Kerberos/NTLM/Basic) are resolved from the
    Security log + IIS log. Clients that authenticate with an OAuth bearer token
    (Outlook with Modern/Hybrid Modern Auth, Outlook mobile, REST apps) leave no
    matching Security-log logon and no cs-username in the IIS log; pass
    -IncludeOAuth to additionally parse the Exchange HttpProxy logs, where their
    identity (AuthenticationType=OAuth, AuthenticatedUser) is recorded.

.PARAMETER Port
    Local port on which to look for active connections. Default 443.

.PARAMETER MinutesBack
    How far back in time to search the Security and IIS logs. Default 60 minutes.

.PARAMETER SiteName
    Name of the IIS website whose logs are read. Default "Default Web Site".

.PARAMETER IncludeOAuth
    Also parse the Exchange HttpProxy logs
    (%ExchangeInstallPath%Logging\HttpProxy\*) to resolve clients that
    authenticate with an OAuth bearer token, which the Security and IIS logs
    cannot identify. Populates the User_OAuthLog and OAuth_Protocols columns
    (empty without this switch).
    Requires the script to run on an Exchange server (install path is read from
    the ExchangeInstallPath env var, with a registry fallback).

.PARAMETER ExportCsv
    If set, the result is also saved to a CSV file.

.PARAMETER CsvPath
    Path to the CSV file used when -ExportCsv is set.
    Default: ".\ConnectedClients_<timestamp>.csv"

.EXAMPLE
    .\Get-ConnectedClients.ps1
    Lists clients currently connected on port 443, using the last 60 minutes of logs.

.EXAMPLE
    .\Get-ConnectedClients.ps1 -MinutesBack 120 -ExportCsv
    Looks back 120 minutes and also exports the result to a timestamped CSV file.

.EXAMPLE
    .\Get-ConnectedClients.ps1 -IncludeOAuth
    Also parses the Exchange HttpProxy logs so OAuth bearer-token clients
    (Modern/Hybrid Modern Auth) are resolved instead of showing up with blank
    user columns.

.NOTES
    Version: 1.3 (2026-09-02)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Note: without -IncludeOAuth this script only resolves the user identity for
    Windows-authenticated (Kerberos/NTLM/Basic) connections. Clients that
    authenticate with an OAuth bearer token (Outlook with Modern/Hybrid Modern
    Auth, Outlook mobile, REST apps) produce no matching Security-log logon and
    no cs-username in the IIS log, so for those the User_SecurityLog /
    User_IISLog / AuthPackage columns stay blank. Pass -IncludeOAuth to also
    read the Exchange HttpProxy logs (%ExchangeInstallPath%Logging\HttpProxy\*,
    columns AuthenticationType=OAuth and AuthenticatedUser).

    Changelog:
      1.3 (2026-09-02) - Added -IncludeOAuth: parses the Exchange HttpProxy logs
        to resolve OAuth bearer-token clients (AuthenticationType OAuth/Bearer),
        populating the User_OAuthLog and OAuth_Protocols columns. Install path comes
        from ExchangeInstallPath with an HKLM\...\ExchangeServer\v15\Setup
        fallback; HttpProxy timestamps are UTC and handled the same way as IIS.
      1.2 (2026-09-02) - IIS W3C timestamps are always UTC; compare them against
        a UTC threshold and convert to local time for display, so clients are no
        longer silently dropped in non-UTC time zones (e.g. -MinutesBack 60 on a
        UTC+2 server used to filter out every IIS entry). Documented the OAuth
        bearer-token limitation.
      1.1 (2026-09-02) - Fixed IIS log parsing: the "#Fields:" header is now read
        from the top of the file instead of from the last 5000 lines (on a busy
        server the header sits far outside that tail window, which produced
        "Could not read the IIS log format (#Fields missing)" even for valid W3C
        logs). The log is now read streaming via [System.IO.File]::ReadLines,
        "#Fields:" is re-read if it changes mid-file (service restart), the two
        newest log files are scanned so a UTC/local offset or midnight rollover
        does not drop entries, and a non-W3C log format is now reported clearly.
      1.0 (2026-08-20) - Initial version.
#>

[CmdletBinding()]
param(
    [int]$Port = 443,
    [int]$MinutesBack = 60,
    [string]$SiteName = "Default Web Site",
    [switch]$IncludeOAuth,
    [switch]$ExportCsv,
    [string]$CsvPath = ".\ConnectedClients_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# --- Elevation check (without it, the Security log returns "No events found") ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    throw "This script must run as Administrator (right-click > Run as Administrator), otherwise the Security log is not readable."
}

$logonTypeMap = @{
    "2"  = "Interactive"
    "3"  = "Network"
    "4"  = "Batch"
    "5"  = "Service"
    "7"  = "Unlock"
    "8"  = "NetworkCleartext"
    "9"  = "NewCredentials"
    "10" = "RemoteInteractive"
    "11" = "CachedInteractive"
}

function Get-ActiveClientConnections {
    param([int]$Port)
    Get-NetTCPConnection -LocalPort $Port -State Established -ErrorAction SilentlyContinue |
        Where-Object { $_.RemoteAddress -notmatch '^(127\.|::1)' } |
        Select-Object -ExpandProperty RemoteAddress -Unique
}

function Resolve-ClientHostname {
    param([string]$IpAddress)
    try { ([System.Net.Dns]::GetHostEntry($IpAddress)).HostName }
    catch { $null }
}

function Get-SecurityLogonEvents {
    param([int]$MinutesBack)
    $startTime = (Get-Date).AddMinutes(-$MinutesBack)
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4624,4625; StartTime=$startTime} -ErrorAction Stop
    }
    catch {
        Write-Warning "Nothing in the Security log for the last $MinutesBack minutes (check auditpol /get /subcategory:'Logon')."
        return @()
    }

    $events | ForEach-Object {
        $xml  = [xml]$_.ToXml()
        $data = $xml.Event.EventData.Data
        [PSCustomObject]@{
            Time        = $_.TimeCreated
            EventId     = $_.Id
            User        = ($data | Where-Object Name -eq 'TargetUserName').'#text'
            Domain      = ($data | Where-Object Name -eq 'TargetDomainName').'#text'
            LogonType   = ($data | Where-Object Name -eq 'LogonType').'#text'
            AuthPackage = ($data | Where-Object Name -eq 'AuthenticationPackageName').'#text'
            SourceIP    = ($data | Where-Object Name -eq 'IpAddress').'#text'
        }
    } | Where-Object { $_.SourceIP -and $_.SourceIP -ne '-' }
}

function Get-IisLogEntries {
    param([string]$SiteName, [int]$MinutesBack, [string[]]$IpFilter)

    Import-Module WebAdministration -ErrorAction SilentlyContinue
    $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if (-not $site) { Write-Warning "IIS site '$SiteName' was not found - skipping IIS logs."; return @() }

    if ($site.logFile.logFormat -and $site.logFile.logFormat -ne 'W3C') {
        Write-Warning "IIS site '$SiteName' uses '$($site.logFile.logFormat)' log format, not W3C - this script can only parse W3C logs."
        return @()
    }

    $logDir  = [System.Environment]::ExpandEnvironmentVariables($site.logFile.directory)
    $logPath = Join-Path $logDir "W3SVC$($site.id)"

    # W3C log timestamps are ALWAYS in UTC (regardless of the "local time
    # rollover" setting), whereas Get-Date / the Security log are local time.
    # Compare against a UTC threshold, then convert each entry back to local
    # for display so it lines up with the Security log column.
    $thresholdUtc = (Get-Date).ToUniversalTime().AddMinutes(-$MinutesBack)

    # Scanning the two newest files makes sure the UTC/local offset or a
    # midnight rollover doesn't drop the entries we need.
    $logFiles = Get-ChildItem $logPath -Filter "u_ex*.log" -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 2
    if (-not $logFiles) { Write-Warning "No IIS log file found in $logPath."; return @() }

    foreach ($file in $logFiles) {

        # "#Fields:" sits at the top of the file (and is re-emitted whenever the
        # column set changes, e.g. after a service restart) - read it from the
        # header rather than from a fixed-size tail, which on a busy server
        # never reaches back far enough.
        $fields = $null
        foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {

            if ($line.StartsWith('#')) {
                if ($line.StartsWith('#Fields:')) {
                    $fields = ($line -replace '#Fields:\s*', '') -split '\s+'
                }
                continue
            }
            if (-not $fields -or $line.Trim() -eq '') { continue }

            $vals = $line -split '\s+'
            if ($vals.Count -lt $fields.Count) { continue }

            $obj = [ordered]@{}
            for ($i = 0; $i -lt $fields.Count; $i++) { $obj[$fields[$i]] = $vals[$i] }
            $entry = [PSCustomObject]$obj

            if ($IpFilter -notcontains $entry.'c-ip') { continue }

            [datetime]$entryTime = 0
            if (-not [DateTime]::TryParse("$($entry.date) $($entry.time)", [ref]$entryTime)) { continue }
            $entryUtc = [DateTime]::SpecifyKind($entryTime, [DateTimeKind]::Utc)

            if ($entryUtc -ge $thresholdUtc) {
                [PSCustomObject]@{
                    Time     = $entryUtc.ToLocalTime()
                    ClientIP = $entry.'c-ip'
                    User     = $entry.'cs-username'
                    Uri      = $entry.'cs-uri-stem'
                    Status   = $entry.'sc-status'
                }
            }
        }
    }
}

function Get-HttpProxyOAuthEntries {
    param([int]$MinutesBack, [string[]]$IpFilter)

    $exPath = $env:ExchangeInstallPath
    if (-not $exPath) {
        foreach ($ver in 'v15', 'v16') {
            $p = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\ExchangeServer\$ver\Setup" -ErrorAction SilentlyContinue).MsiInstallPath
            if ($p) { $exPath = $p; break }
        }
    }
    if (-not $exPath) {
        Write-Warning "Exchange install path not found (ExchangeInstallPath env var / registry) - skipping HttpProxy logs. Run this on an Exchange server."
        return @()
    }

    $proxyRoot = Join-Path $exPath 'Logging\HttpProxy'
    if (-not (Test-Path $proxyRoot)) {
        Write-Warning "HttpProxy log folder not found: $proxyRoot"
        return @()
    }

    # HttpProxy DateTime values are UTC, same as the IIS logs.
    $thresholdUtc = (Get-Date).ToUniversalTime().AddMinutes(-$MinutesBack)
    # Files roll by size, so a recent entry can sit in a file whose LastWriteTime
    # is already an hour old - widen the file window accordingly.
    $fileCutoff = (Get-Date).AddMinutes(-$MinutesBack).AddHours(-1)

    $logFiles = Get-ChildItem $proxyRoot -Recurse -Filter '*.LOG' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $fileCutoff }
    if (-not $logFiles) { Write-Warning "No recent HttpProxy log files under $proxyRoot."; return @() }

    foreach ($file in $logFiles) {

        $fields    = $null
        $dataLines = [System.Collections.Generic.List[string]]::new()

        foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {
            if ($line.StartsWith('#')) {
                if ($line.StartsWith('#Fields:')) {
                    $fields = $line.Substring(8).Trim() -split '\s*,\s*'
                }
                continue
            }
            # Cheap prefilter before the (relatively expensive) CSV parse - the
            # literal "OAuth"/"Bearer" almost only appears as the auth type.
            if ($fields -and ($line -like '*OAuth*' -or $line -like '*Bearer*')) {
                $dataLines.Add($line)
            }
        }
        if (-not $fields -or $dataLines.Count -eq 0) { continue }

        $dataLines | ConvertFrom-Csv -Header $fields | ForEach-Object {
            if ($_.AuthenticationType -notmatch 'OAuth|Bearer') { return }
            if ($IpFilter -notcontains $_.ClientIpAddress)      { return }

            [datetime]$t = 0
            if (-not [DateTime]::TryParse($_.DateTime, [ref]$t)) { return }
            $tUtc = $t.ToUniversalTime()
            if ($tUtc -lt $thresholdUtc) { return }

            [PSCustomObject]@{
                Time     = $tUtc.ToLocalTime()
                ClientIP = $_.ClientIpAddress
                User     = $_.AuthenticatedUser
                Protocol = $_.Protocol
                Auth     = $_.AuthenticationType
            }
        }
    }
}

# --- Main ---
Write-Host "Looking for active connections on port $Port..." -ForegroundColor Cyan
$activeIps = Get-ActiveClientConnections -Port $Port

if (-not $activeIps) {
    Write-Host "No active connections on port $Port." -ForegroundColor Yellow
    return
}

Write-Host "Found $($activeIps.Count) unique IPs. Reading Security log (last $MinutesBack min)..." -ForegroundColor Cyan
$secEvents = Get-SecurityLogonEvents -MinutesBack $MinutesBack

Write-Host "Reading IIS logs for user identity..." -ForegroundColor Cyan
$iisEntries = Get-IisLogEntries -SiteName $SiteName -MinutesBack $MinutesBack -IpFilter $activeIps

$oauthEntries = @()
if ($IncludeOAuth) {
    Write-Host "Reading Exchange HttpProxy logs for OAuth (bearer) clients..." -ForegroundColor Cyan
    $oauthEntries = Get-HttpProxyOAuthEntries -MinutesBack $MinutesBack -IpFilter $activeIps
}

$results = foreach ($ip in $activeIps) {
    $hostname = Resolve-ClientHostname -IpAddress $ip
    $lastSec  = $secEvents | Where-Object { $_.SourceIP -eq $ip } | Sort-Object Time -Descending | Select-Object -First 1
    $iisUsers = ($iisEntries | Where-Object { $_.ClientIP -eq $ip -and $_.User -and $_.User -ne '-' } |
                 Select-Object -ExpandProperty User -Unique) -join ", "
    $iisUris  = ($iisEntries | Where-Object { $_.ClientIP -eq $ip } |
                 Select-Object -ExpandProperty Uri -Unique | Select-Object -First 3) -join ", "

    $oauthForIp = $oauthEntries | Where-Object { $_.ClientIP -eq $ip }
    $oauthUsers = ($oauthForIp | Where-Object { $_.User -and $_.User -ne '-' } |
                   Select-Object -ExpandProperty User -Unique) -join ", "
    $oauthProto = ($oauthForIp | Where-Object { $_.Protocol } |
                   Select-Object -ExpandProperty Protocol -Unique) -join ", "

    $logonTypeLabel = if ($lastSec.LogonType -and $logonTypeMap.ContainsKey($lastSec.LogonType)) {
        "$($lastSec.LogonType) ($($logonTypeMap[$lastSec.LogonType]))"
    } else { $lastSec.LogonType }

    [PSCustomObject]@{
        ClientIP         = $ip
        Hostname         = $hostname
        User_SecurityLog = $lastSec.User
        User_IISLog      = $iisUsers
        User_OAuthLog    = $oauthUsers
        AuthPackage      = if ($lastSec.AuthPackage) { $lastSec.AuthPackage }
                           elseif ($oauthForIp)      { 'OAuth/Bearer' }
                           else                      { $null }
        LogonType        = $logonTypeLabel
        LastLogonTime    = $lastSec.Time
        OAuth_Protocols  = $oauthProto
        SampleEndpoints  = $iisUris
    }
}

$results | Sort-Object ClientIP | Format-Table -AutoSize -Wrap

if ($ExportCsv) {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to: $CsvPath" -ForegroundColor Green
}