<#
.SYNOPSIS
    Identifies currently connected clients on an Exchange server (by local port, default 443),
    their authentication method (Kerberos/NTLM) from the Security log, and the user identity
    from the IIS log.

.DESCRIPTION
    Run directly ON the Exchange server (front-end), in a PowerShell console with
    Administrator rights (without elevation the Security log is not readable - the
    result is "No events found" even when the log contains matching records).

.PARAMETER Port
    Local port on which to look for active connections. Default 443.

.PARAMETER MinutesBack
    How far back in time to search the Security and IIS logs. Default 60 minutes.

.PARAMETER SiteName
    Name of the IIS website whose logs are read. Default "Default Web Site".

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

.NOTES
    Version: 1.0 (2026-08-20)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>

[CmdletBinding()]
param(
    [int]$Port = 443,
    [int]$MinutesBack = 60,
    [string]$SiteName = "Default Web Site",
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

    $logDir  = [System.Environment]::ExpandEnvironmentVariables($site.logFile.directory)
    $logPath = Join-Path $logDir "W3SVC$($site.id)"
    $latest  = Get-ChildItem $logPath -Filter "u_ex*.log" -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) { Write-Warning "No IIS log file found in $logPath."; return @() }

    $lines = Get-Content $latest.FullName -Tail 5000
    $fieldLine = $lines | Where-Object { $_ -like '#Fields:*' } | Select-Object -Last 1
    if (-not $fieldLine) { Write-Warning "Could not read the IIS log format (#Fields missing)."; return @() }
    $fields = ($fieldLine -replace '#Fields:\s*', '') -split '\s+'

    $threshold = (Get-Date).AddMinutes(-$MinutesBack)

    # Note: IIS logs are typically in UTC by default (unless set to Local Time),
    # while the Security log uses the server's local time - account for this offset when comparing times.
    $lines | Where-Object { $_ -notlike '#*' -and $_.Trim() -ne '' } | ForEach-Object {
        $vals = $_ -split '\s+'
        if ($vals.Count -lt $fields.Count) { return }
        $obj = [ordered]@{}
        for ($i = 0; $i -lt $fields.Count; $i++) { $obj[$fields[$i]] = $vals[$i] }
        $entry = [PSCustomObject]$obj

        [datetime]$entryTime = 0
        if ([DateTime]::TryParse("$($entry.date) $($entry.time)", [ref]$entryTime)) {
            if ($entryTime -ge $threshold -and $IpFilter -contains $entry.'c-ip') {
                [PSCustomObject]@{
                    Time     = $entryTime
                    ClientIP = $entry.'c-ip'
                    User     = $entry.'cs-username'
                    Uri      = $entry.'cs-uri-stem'
                    Status   = $entry.'sc-status'
                }
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

$results = foreach ($ip in $activeIps) {
    $hostname = Resolve-ClientHostname -IpAddress $ip
    $lastSec  = $secEvents | Where-Object { $_.SourceIP -eq $ip } | Sort-Object Time -Descending | Select-Object -First 1
    $iisUsers = ($iisEntries | Where-Object { $_.ClientIP -eq $ip -and $_.User -and $_.User -ne '-' } |
                 Select-Object -ExpandProperty User -Unique) -join ", "
    $iisUris  = ($iisEntries | Where-Object { $_.ClientIP -eq $ip } |
                 Select-Object -ExpandProperty Uri -Unique | Select-Object -First 3) -join ", "

    $logonTypeLabel = if ($lastSec.LogonType -and $logonTypeMap.ContainsKey($lastSec.LogonType)) {
        "$($lastSec.LogonType) ($($logonTypeMap[$lastSec.LogonType]))"
    } else { $lastSec.LogonType }

    [PSCustomObject]@{
        ClientIP         = $ip
        Hostname         = $hostname
        User_SecurityLog = $lastSec.User
        User_IISLog      = $iisUsers
        AuthPackage      = $lastSec.AuthPackage
        LogonType        = $logonTypeLabel
        LastLogonTime    = $lastSec.Time
        SampleEndpoints  = $iisUris
    }
}

$results | Sort-Object ClientIP | Format-Table -AutoSize -Wrap

if ($ExportCsv) {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExported to: $CsvPath" -ForegroundColor Green
}