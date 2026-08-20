<#
.SYNOPSIS
    Exports the allowed IP addresses from an existing Exchange Receive Connector to a file.

.DESCRIPTION
    This script exports the list of allowed IP addresses (RemoteIPRanges)
    from an existing Receive Connector on an Exchange server to a text file.
    Each IP address or range is saved on a separate line.

.PARAMETER ConnectorIdentity
    Identifier of the receive connector in the form "Server\ConnectorName"
    Example: "EXCH01\Application Relay"

.PARAMETER ExportFile
    Path to the output file. Default value: "C:\Temp\RelayIPs.txt"

.EXAMPLE
    .\Get-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH01\Default Frontend EXCH01"
    Exports IP addresses from the connector to the default file C:\Temp\RelayIPs.txt

.EXAMPLE
    .\Get-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH01\Application Relay" -ExportFile "C:\Backup\IPs.txt"
    Exports IP addresses to a custom file

.NOTES
    Requires Exchange Management Shell and appropriate permissions
    Version: 1.0 (2026-07-30)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Identifier of the receive connector (Server\Name)")]
    [string]$ConnectorIdentity,

    [Parameter(Mandatory=$false, HelpMessage="Path to the output file")]
    [string]$ExportFile = "C:\Temp\RelayIPs.txt"
)

# Check whether the target directory exists
$directory = Split-Path -Path $ExportFile -Parent
if ($directory -and -not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    Write-Host "Created directory: $directory" -ForegroundColor Yellow
}

try {
    # Retrieve the connector
    Write-Host "Loading connector: $ConnectorIdentity..." -ForegroundColor Cyan
    $connector = Get-ReceiveConnector $ConnectorIdentity -ErrorAction Stop

    # Export the IP ranges to the file
    $connector.RemoteIPRanges |
        ForEach-Object { $_.ToString() } |
        Set-Content $ExportFile

    Write-Host "OK Exported $($connector.RemoteIPRanges.Count) IP ranges to $ExportFile" -ForegroundColor Green
    Write-Host "File contains:" -ForegroundColor Cyan
    Get-Content $ExportFile | Select-Object -First 5 | ForEach-Object { Write-Host "  $_" }
    if ($connector.RemoteIPRanges.Count -gt 5) {
        Write-Host "  ..." -ForegroundColor Gray
    }
}
catch {
    Write-Host "X Error during export: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
