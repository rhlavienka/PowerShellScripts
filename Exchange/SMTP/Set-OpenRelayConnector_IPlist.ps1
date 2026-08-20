<#
.SYNOPSIS
    Imports the allowed IP addresses from a file into an Exchange Receive Connector.

.DESCRIPTION
    This script reads a list of IP addresses or ranges from a file and sets them
    as the allowed RemoteIPRanges for the specified Receive Connector.
    WARNING: This command overwrites all existing IP addresses on the connector!

.PARAMETER ConnectorIdentity
    Identifier of the receive connector in the form "Server\ConnectorName"
    Example: "EXCH02\Application Relay"

.PARAMETER ImportFile
    Path to the file with IP addresses (one per line)
    Default value: "C:\Temp\RelayIPs.txt"

.EXAMPLE
    .\Set-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH02\Application Relay"
    Imports IP addresses from the default file into the connector

.EXAMPLE
    .\Set-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH02\Application Relay" -ImportFile "C:\Backup\IPs.txt"
    Imports IP addresses from a custom file

.NOTES
    Requires Exchange Management Shell and appropriate permissions
    This command overwrites existing IP addresses - make a backup first!
    Version: 1.0 (2026-07-30)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Identifier of the receive connector (Server\Name)")]
    [string]$ConnectorIdentity,

    [Parameter(Mandatory=$false, HelpMessage="Path to the file with IP addresses")]
    [string]$ImportFile = "C:\Temp\RelayIPs.txt"
)

# Check whether the file exists
if (-not (Test-Path $ImportFile)) {
    Write-Host "X File not found: $ImportFile" -ForegroundColor Red
    exit 1
}

$backupFile = $null

try {
    # Load IP addresses from the file (remove blank lines)
    Write-Host "Loading IP addresses from file: $ImportFile..." -ForegroundColor Cyan
    $IPs = Get-Content $ImportFile |
           Where-Object { $_ -and $_.Trim() -ne "" }

    if ($IPs.Count -eq 0) {
        Write-Host "X File contains no valid IP addresses!" -ForegroundColor Red
        exit 1
    }

    Write-Host "Loaded $($IPs.Count) IP ranges" -ForegroundColor Yellow
    Write-Host "First 5 addresses:" -ForegroundColor Cyan
    $IPs | Select-Object -First 5 | ForEach-Object { Write-Host "  $_" }
    if ($IPs.Count -gt 5) {
        Write-Host "  ..." -ForegroundColor Gray
    }

    # Back up the current state
    Write-Host "`nBacking up current IP addresses..." -ForegroundColor Yellow
    $connector = Get-ReceiveConnector $ConnectorIdentity -ErrorAction Stop
    $backupFile = "C:\Temp\RelayIPs_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $connector.RemoteIPRanges | ForEach-Object { $_.ToString() } | Set-Content $backupFile
    Write-Host "Backup saved: $backupFile" -ForegroundColor Green

    # Set the new IP ranges
    Write-Host "`nSetting new IP ranges on $ConnectorIdentity..." -ForegroundColor Cyan
    Set-ReceiveConnector `
        -Identity $ConnectorIdentity `
        -RemoteIPRanges $IPs `
        -ErrorAction Stop

    Write-Host "OK Imported $($IPs.Count) IP ranges into $ConnectorIdentity" -ForegroundColor Green

    # Verification
    $updatedConnector = Get-ReceiveConnector $ConnectorIdentity
    Write-Host "Current number of IP ranges on the connector: $($updatedConnector.RemoteIPRanges.Count)" -ForegroundColor Cyan
}
catch {
    Write-Host "X Error during import: $($_.Exception.Message)" -ForegroundColor Red
    if ($backupFile) {
        Write-Host "Backup is available at: $backupFile" -ForegroundColor Yellow
    }
    exit 1
}
