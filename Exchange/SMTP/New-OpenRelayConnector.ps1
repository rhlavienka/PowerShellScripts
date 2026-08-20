<#
.SYNOPSIS
    Creates a new Open Relay Receive Connector on an Exchange server.

.DESCRIPTION
    This script creates a new Receive Connector configured as an open relay
    for applications and devices that need to send email through the Exchange server.
    The connector is created with anonymous authentication and relay permissions.

    SECURITY WARNING:
    - This connector allows relay for anonymous users!
    - After creation you MUST restrict RemoteIPRanges to trusted IP addresses only!
    - The default value of 255.255.255.255 is only a placeholder and MUST be changed!

.PARAMETER Server
    Name of the Exchange server on which the connector should be created
    Default value: "EXCH01"

.PARAMETER ConnectorName
    Name of the new receive connector
    Default value: "Application Relay"

.EXAMPLE
    .\New-OpenRelayConnector.ps1
    Creates a connector with default settings (EXCH01\Application Relay)

.EXAMPLE
    .\New-OpenRelayConnector.ps1 -Server "EXCH02" -ConnectorName "Custom Relay"
    Creates a connector on server EXCH02 with a custom name

.EXAMPLE
    # Complete workflow for creating and populating a connector:
    # 1. Export IP addresses from the old server
    .\Get-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH01\Application Relay" -ExportFile "C:\Temp\IPs.txt"

    # 2. Create a new connector on the new server
    .\New-OpenRelayConnector.ps1 -Server "EXCH02" -ConnectorName "Application Relay"

    # 3. Import the IP addresses into the new connector
    .\Set-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity "EXCH02\Application Relay" -ImportFile "C:\Temp\IPs.txt"

.NOTES
    Requires Exchange Management Shell and Organization Management permissions
    The connector is created with the FrontendTransport role on port 25
    After creation, it is recommended to use Set-OpenRelayConnector_IPlist.ps1 to configure allowed IP addresses
    Version: 1.0 (2026-07-30)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
#>

param(
    [Parameter(Mandatory=$false, HelpMessage="Name of the Exchange server")]
    [string]$Server = "EXCH01",

    [Parameter(Mandatory=$false, HelpMessage="Name of the new connector")]
    [string]$ConnectorName = "Application Relay"
)

if ([string]::IsNullOrWhiteSpace($Server) -or [string]::IsNullOrWhiteSpace($ConnectorName)) {
    Write-Host "X Server and ConnectorName cannot be empty." -ForegroundColor Red
    exit 1
}

$connectorIdentity = "$Server\$ConnectorName"

# Check whether the connector already exists
$existing = Get-ReceiveConnector $connectorIdentity -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "X Connector '$connectorIdentity' already exists!" -ForegroundColor Red
    Write-Host "Use a different name or remove the existing connector with:" -ForegroundColor Yellow
    Write-Host "  Remove-ReceiveConnector '$connectorIdentity' -Confirm:`$false" -ForegroundColor Cyan
    exit 1
}

Write-Host "Creating Open Relay Connector..." -ForegroundColor Cyan
Write-Host "  Server: $Server" -ForegroundColor White
Write-Host "  Name: $ConnectorName" -ForegroundColor White
Write-Host "  Identity: $connectorIdentity" -ForegroundColor White
Write-Host ""

try {
    # Create the receive connector
    Write-Host "Step 1/2: Creating Receive Connector..." -ForegroundColor Yellow
    New-ReceiveConnector `
        -Name $ConnectorName `
        -Server $Server `
        -TransportRole FrontendTransport `
        -Bindings 0.0.0.0:25 `
        -RemoteIPRanges 255.255.255.255 `
        -PermissionGroups AnonymousUsers `
        -ErrorAction Stop | Out-Null

    Write-Host "OK Receive Connector created" -ForegroundColor Green

    # Grant relay permissions
    try {
        Write-Host "Step 2/2: Setting relay permissions..." -ForegroundColor Yellow
        Get-ReceiveConnector $connectorIdentity |
            Add-ADPermission `
                -User "NT AUTHORITY\ANONYMOUS LOGON" `
                -ExtendedRights "Ms-Exch-SMTP-Accept-Any-Recipient" `
                -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "X Failed to set relay permissions: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Rolling back: removing partially created connector '$connectorIdentity'..." -ForegroundColor Yellow
        Remove-ReceiveConnector $connectorIdentity -Confirm:$false -ErrorAction SilentlyContinue
        if (Get-ReceiveConnector $connectorIdentity -ErrorAction SilentlyContinue) {
            Write-Host "X Rollback failed: connector '$connectorIdentity' still exists without relay permissions. Remove it manually." -ForegroundColor Red
        }
        else {
            Write-Host "Rollback complete. No connector was left behind." -ForegroundColor Yellow
        }
        exit 1
    }

    Write-Host "OK Relay permissions set" -ForegroundColor Green
    Write-Host ""
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host "OK Relay connector successfully created: $connectorIdentity" -ForegroundColor Green
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "! IMPORTANT SECURITY STEPS:" -ForegroundColor Red
    Write-Host "1. The connector has RemoteIPRanges set to 255.255.255.255 (temporary)" -ForegroundColor Yellow
    Write-Host "2. You MUST restrict the allowed IP addresses!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Recommended next steps:" -ForegroundColor Cyan
    Write-Host "  # If you are copying IP addresses from another server:" -ForegroundColor Gray
    Write-Host "  .\Get-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity 'OLD_SERVER\CONNECTOR'" -ForegroundColor White
    Write-Host "  .\Set-OpenRelayConnector_IPlist.ps1 -ConnectorIdentity '$connectorIdentity'" -ForegroundColor White
    Write-Host ""
    Write-Host "  # Or configure manually:" -ForegroundColor Gray
    Write-Host "  Set-ReceiveConnector '$connectorIdentity' -RemoteIPRanges @('192.168.1.10','192.168.1.11')" -ForegroundColor White
    Write-Host ""

    # Show current configuration
    Write-Host "Current configuration:" -ForegroundColor Cyan
    $connector = Get-ReceiveConnector $connectorIdentity
    Write-Host "  TransportRole: $($connector.TransportRole)" -ForegroundColor White
    Write-Host "  Bindings: $($connector.Bindings -join ', ')" -ForegroundColor White
    Write-Host "  RemoteIPRanges: $($connector.RemoteIPRanges -join ', ')" -ForegroundColor White
    Write-Host "  PermissionGroups: $($connector.PermissionGroups)" -ForegroundColor White
}
catch {
    Write-Host ""
    Write-Host "X Error creating connector: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible causes:" -ForegroundColor Yellow
    Write-Host "  - Exchange Management Shell is not loaded" -ForegroundColor Gray
    Write-Host "  - Insufficient permissions (Organization Management required)" -ForegroundColor Gray
    Write-Host "  - Server '$Server' does not exist or is unreachable" -ForegroundColor Gray
    Write-Host "  - A connector with this name already exists" -ForegroundColor Gray
    exit 1
}
