<#
.SYNOPSIS
    Creates custom outbound sync rules in Microsoft Entra Connect (Sync Rule Editor)
    that rewrite the domain in the userPrincipalName, mail and proxyAddresses
    attributes from the old domain to the new one. All expressions have an IIF safeguard.

.PARAMETER OldDomain
    Original domain without the "@", e.g. "old-domain.sk"

.PARAMETER NewDomain
    New domain without the "@", e.g. "new-domain.sk"
    MUST already be added and verified in Entra ID!

.PARAMETER BasePrecedence
    Preferred starting Precedence value for the 3 new rules. The script scans
    existing Outbound rules and, if this value (or the next ones) is already
    taken, automatically picks the next free Precedence numbers instead of
    failing. Default 50.

.EXAMPLE
    .\New-DomainRewriteRules.ps1 -OldDomain "old-domain.sk" -NewDomain "new-domain.sk"
    Creates the 3 outbound sync rules (userPrincipalName, mail, proxyAddresses)
    starting at the default Precedence 50, rewriting old-domain.sk to new-domain.sk.

.EXAMPLE
    .\New-DomainRewriteRules.ps1 -OldDomain "old-domain.sk" -NewDomain "new-domain.sk" -BasePrecedence 100
    Same as above, but starts searching for free Precedence values from 100 instead of 50.

.NOTES
    Version: 1.0 (2026-08-20)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)
    - Run on the Entra Connect server as Administrator.
    - Before running, back up the configuration (Get-ADSyncRule | Export-Clixml backup.xml).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$OldDomain,

    [Parameter(Mandatory = $true)]
    [string]$NewDomain,

    [int]$BasePrecedence = 50
)

Import-Module ADSync -ErrorAction Stop

# --- Find the AAD (Entra) connector ---
$AADConnector = Get-ADSyncConnector | Where-Object { $_.ConnectorTypeName -eq 'AAD' }
if (-not $AADConnector) {
    throw "Could not find the AAD/Entra connector. Check: Get-ADSyncConnector | ft Name,ConnectorTypeName"
}

# --- Find 3 free Precedence values, starting at -BasePrecedence, that do not collide ---
# --- with any existing Outbound rule (New-ADSyncRule fails on a duplicate Precedence) ---
$usedPrecedences = (Get-ADSyncRule | Where-Object { $_.Direction -eq 'Outbound' }).Precedence
$precedences = [System.Collections.Generic.List[int]]::new()
$candidate = $BasePrecedence
while ($precedences.Count -lt 3) {
    if ($usedPrecedences -notcontains $candidate) { $precedences.Add($candidate) }
    $candidate++
}
if (($precedences[0] -ne $BasePrecedence) -or ($precedences[2] -ne $BasePrecedence + 2)) {
    Write-Warning "Precedence $BasePrecedence-$($BasePrecedence+2) is partially taken by existing Outbound rules. Using free values instead: $($precedences -join ', ')"
}
Write-Host "Using Precedence values: $($precedences -join ', ')" -ForegroundColor Cyan

# --- Check whether a similar rule already exists (name collision) ---
$existing = Get-ADSyncRule | Where-Object { $_.Name -like "*$OldDomain*$NewDomain*" }
if ($existing) {
    Write-Warning "Found existing rule(s) with a similar name:"
    $existing | ForEach-Object { Write-Warning " - $($_.Name) (Precedence $($_.Precedence))" }
    $confirm = Read-Host "Continue anyway? (Y/N)"
    if ($confirm -ne 'Y') { throw "Cancelled by user." }
}

function New-DomainRewriteRule {
    param(
        [string]$Name,
        [string]$Description,
        [int]$Precedence,
        [string]$TargetAttribute,
        [string]$ScopeAttribute,
        [string]$Expression
    )

    Write-Host "Creating rule: $Name (precedence $Precedence)" -ForegroundColor Cyan

    New-ADSyncRule `
        -Name $Name `
        -Identifier ([guid]::NewGuid()) `
        -Description $Description `
        -Direction 'Outbound' `
        -Precedence $Precedence `
        -PrecedenceAfter  '00000000-0000-0000-0000-000000000000' `
        -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
        -SourceObjectType 'person' `
        -TargetObjectType 'user' `
        -Connector $AADConnector.Identifier.Guid `
        -LinkType 'Join' `
        -SoftDeleteExpiryInterval 0 `
        -ImmutableTag '' `
        -OutVariable syncRule | Out-Null

    Add-ADSyncAttributeFlowMapping `
        -SynchronizationRule $syncRule[0] `
        -Destination $TargetAttribute `
        -FlowType 'Expression' `
        -ValueMergeType 'Update' `
        -Expression $Expression `
        -OutVariable syncRule | Out-Null

    $scopeCondition = New-Object `
        -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
        -ArgumentList $ScopeAttribute, "@$OldDomain", 'ENDSWITH'

    Add-ADSyncScopeConditionGroup `
        -SynchronizationRule $syncRule[0] `
        -ScopeConditions @($scopeCondition) `
        -OutVariable syncRule | Out-Null

    Add-ADSyncRule -SynchronizationRule $syncRule[0]

    Write-Host "  -> OK: '$Name' added." -ForegroundColor Green
}

# --- 1) userPrincipalName (with IIF safeguard) ---
$exprUpn = "IIF(InStr([userPrincipalName],`"@$OldDomain`")>0, " +
           "Mid([userPrincipalName],1,InStr([userPrincipalName],`"@`")-1) & `"@$NewDomain`", " +
           "[userPrincipalName])"

New-DomainRewriteRule `
    -Name "Custom Out to AAD - UPN rewrite $OldDomain to $NewDomain" `
    -Description "Replaces UPN suffix $OldDomain with $NewDomain (with IIF safeguard)" `
    -Precedence $precedences[0] `
    -TargetAttribute 'userPrincipalName' `
    -ScopeAttribute 'userPrincipalName' `
    -Expression $exprUpn

# --- 2) mail (with IIF safeguard) ---
$exprMail = "IIF(InStr([mail],`"@$OldDomain`")>0, " +
            "Mid([mail],1,InStr([mail],`"@`")-1) & `"@$NewDomain`", " +
            "[mail])"

New-DomainRewriteRule `
    -Name "Custom Out to AAD - Mail rewrite $OldDomain to $NewDomain" `
    -Description "Replaces the domain in the mail attribute from $OldDomain to $NewDomain (with IIF safeguard)" `
    -Precedence $precedences[1] `
    -TargetAttribute 'mail' `
    -ScopeAttribute 'mail' `
    -Expression $exprMail

# --- 3) proxyAddresses (multivalued, already includes an IIF safeguard) ---
$exprProxy = "IIF(InStr([proxyAddresses],`"@$OldDomain`")>0, " +
             "Word([proxyAddresses],1,`"@`") & `"@$NewDomain`", " +
             "[proxyAddresses])"

New-DomainRewriteRule `
    -Name "Custom Out to AAD - ProxyAddresses rewrite $OldDomain to $NewDomain" `
    -Description "Replaces domain $OldDomain with $NewDomain in all proxyAddresses values (with IIF safeguard)" `
    -Precedence $precedences[2] `
    -TargetAttribute 'proxyAddresses' `
    -ScopeAttribute 'proxyAddresses' `
    -Expression $exprProxy

Write-Host ""
Write-Host "Done. Next steps:" -ForegroundColor Yellow
Write-Host "1. Verify the rules in the Synchronization Rules Editor (GUI) or via Get-ADSyncRule."
Write-Host "2. Run a Preview on a test object."
Write-Host "3. Run a Full sync: Start-ADSyncSyncCycle -PolicyType Initial"