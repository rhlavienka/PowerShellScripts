<#
.SYNOPSIS
    Exports all on-premises Exchange transport (mail flow) rules for offline review.

.DESCRIPTION
    Collects every transport rule from the local on-premises Exchange organization
    and writes it out in three formats:

      1. JSON  - one object per rule with all relevant properties, including the
                 auto-generated plain-English "Description" Exchange builds from
                 the rule's conditions/exceptions/actions. This is the primary
                 file intended for offline analysis (e.g. deciding which rules
                 also need to exist in Exchange Online for a centralized/hybrid
                 mail transport setup).
      2. CSV   - a flat summary (Name, Priority, State, Mode, Comments,
                 Description) for a quick spreadsheet/Out-GridView look.
      3. XML   - the native Exchange transport rule collection produced by
                 Export-TransportRuleCollection. This is the official
                 backup/restore format (re-importable on-prem with
                 Import-TransportRuleCollection) and serves as the authoritative
                 source if anything needs to be cross-checked against the raw
                 rule definition.

    This script only reads and exports rules - it does not create, modify, or
    remove anything. It must be run from the on-premises Exchange Management
    Shell (or a remote session to an on-premises Exchange server), not from
    Exchange Online PowerShell.

.PARAMETER OutputFolder
    Folder the JSON, CSV, and XML export files are written to. Created if it
    does not already exist. Defaults to "C:\Temp".

.EXAMPLE
    .\Export-TransportRules.ps1
    Exports all on-premises transport rules to timestamped JSON/CSV/XML files
    in C:\Temp.

.EXAMPLE
    .\Export-TransportRules.ps1 -OutputFolder "D:\Exports\TransportRules"
    Exports to a custom folder.

.NOTES
    Requires Exchange Management Shell (on-premises) with an account that has
    View-Only/Organization Management rights to read transport rules.

    Version: 1.0 (2026-08-25)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Changelog:
      1.0 (2026-08-25) - Initial version
#>

param(
    [Parameter(Mandatory = $false, HelpMessage = "Folder the export files are written to")]
    [string]$OutputFolder = "C:\Temp"
)

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------
if (-not (Get-Command Get-TransportRule -ErrorAction SilentlyContinue))
{
    Write-Error "Get-TransportRule is not available. Run this script from the on-premises Exchange Management Shell (or a remote session to an on-premises Exchange server), not from Exchange Online PowerShell."
    return
}

if (-not (Test-Path $OutputFolder))
{
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Yellow
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$JsonFile  = Join-Path $OutputFolder "TransportRules_$TimeStamp.json"
$CsvFile   = Join-Path $OutputFolder "TransportRules_$TimeStamp.csv"
$XmlFile   = Join-Path $OutputFolder "TransportRules_$TimeStamp.xml"

# ---------------------------------------------------------------------------
# Collect rules
# ---------------------------------------------------------------------------
Write-Host "Retrieving transport rules from the on-premises organization..." -ForegroundColor Cyan

try
{
    $Rules = Get-TransportRule -ErrorAction Stop | Sort-Object Priority
}
catch
{
    Write-Error "Failed to retrieve transport rules: $($_.Exception.Message)"
    return
}

if (-not $Rules)
{
    Write-Warning "No transport rules were found in this organization."
    return
}

Write-Host "Found $($Rules.Count) transport rule(s)." -ForegroundColor Green

# ---------------------------------------------------------------------------
# JSON export (full detail, primary analysis input)
# ---------------------------------------------------------------------------
$RuleDetails = $Rules | Select-Object Name, Priority, State, Mode, Comments, Description, `
    Conditions, Exceptions, Actions, `
    RuleErrorAction, SenderAddressLocation, `
    WhenChanged, WhenCreated

$RuleDetails |
    ConvertTo-Json -Depth 10 |
    Set-Content -Path $JsonFile -Encoding UTF8

Write-Host "JSON export : $JsonFile" -ForegroundColor Green

# ---------------------------------------------------------------------------
# CSV export (flat summary)
# ---------------------------------------------------------------------------
$Rules |
    Select-Object Name, Priority, State, Mode, Comments, Description |
    Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

Write-Host "CSV export  : $CsvFile" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Native rule collection backup (authoritative, re-importable on-prem)
# ---------------------------------------------------------------------------
try
{
    $Collection = Export-TransportRuleCollection -ErrorAction Stop
    [System.IO.File]::WriteAllBytes($XmlFile, $Collection.FileData)
    Write-Host "XML backup  : $XmlFile" -ForegroundColor Green
}
catch
{
    Write-Warning "Export-TransportRuleCollection failed, XML backup was skipped: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Summary by state:" -ForegroundColor Cyan
$Rules | Group-Object State | Select-Object Name, Count | Format-Table -AutoSize

Write-Host "Summary by mode:" -ForegroundColor Cyan
$Rules | Group-Object Mode | Select-Object Name, Count | Format-Table -AutoSize

$Rules | Select-Object Priority, Name, State, Mode | Format-Table -AutoSize
