<#
.SYNOPSIS
    Creates custom INBOUND sync rules in Microsoft Entra Connect (Sync Rule Editor)
    on the AD DS connector, scoped to objects flagged for pre-migration
    (source AD attribute, default extensionAttribute10, equals "PreMig").

.DESCRIPTION
    Builds up to 5 independent inbound rules, all gated by the same PreMig scope
    condition. Each rule is created by its own function below and can be run alone
    via the -Rule parameter:

      Upn              Rewrites userPrincipalName's domain to <SLD-label>.<NewDomainSuffix>,
                        where <SLD-label> is extracted per-object from whatever
                        domain the UPN currently has (so firma1.sk, firma2.sk,
                        firma3.sk, ... all normalize to their own *.rhlab.in).
      Mail             Same per-object domain rewrite, applied to the mail attribute.
      ProxyAddresses   Rebuilds proxyAddresses: keeps X500: entries and smtp: aliases
                        ending in the tenant's *.mail.onmicrosoft.com (MOERA) domain,
                        rewrites the primary SMTP: address the same way as Upn/Mail,
                        drops everything else.
      CustomAttribute5 Direct-flows the original (pre-rewrite) userPrincipalName into
                        customAttribute5, so the real production UPN stays available
                        for reference after the Upn rule above overwrites it.
      HideFromGal      Forces msExchHideFromAddressLists to True via a Constant flow
                        (not imported from AD) for all objects in scope.

    Domain math happens INSIDE the sync rule expression, per object, not in this
    script: the domain part after "@" is split on ".", and its first label (the
    "second-level" part, assuming a 2-label old domain like "firma1.sk") is
    combined with -NewDomainSuffix, e.g. "firma1.sk" -> "firma1.rhlab.in" and
    "firma2.sk" -> "firma2.rhlab.in" - no need to know the old domain(s) in advance.
    An object whose domain already ends in .NewDomainSuffix is left unchanged
    (idempotency guard, safe to re-run).

.PARAMETER NewDomainSuffix
    Suffix appended to the per-object extracted second-level label to build the
    new domain. Default "rhlab.in" (so "firma1.sk" -> "firma1.rhlab.in",
    "firma2.sk" -> "firma2.rhlab.in"). The resulting domains MUST already be
    added and verified in Entra ID!

.PARAMETER MoeraDomain
    The tenant's default *.onmicrosoft.com mail domain (without "@"), e.g.
    "contoso.mail.onmicrosoft.com". smtp: aliases ending in this domain are the
    only secondary aliases preserved by the ProxyAddresses rule.

.PARAMETER PreMigSourceAttribute
    AD DS (connector space) attribute that carries the PreMig marker. Default
    "extensionAttribute10" (the AD attribute that by default flows into the
    metaverse's customAttribute10 via the out-of-box inbound rule).

.PARAMETER PreMigValue
    Value that PreMigSourceAttribute must equal for a rule to apply. Default "PreMig".

.PARAMETER Rule
    Which rule(s) to create. One of: All, Upn, Mail, ProxyAddresses,
    CustomAttribute5, HideFromGal. Default "All".

.PARAMETER ADConnectorName
    Name of the AD DS connector to use. Only needed to skip the interactive
    prompt when more than one AD DS connector exists (e.g. a multi-forest lab);
    with a single connector it's auto-detected and this is ignored.

.PARAMETER BasePrecedence
    Preferred starting Precedence value for the new rules. The script scans
    existing Inbound rules on the AD DS connector and, if this value (or the
    next ones) is already taken, automatically picks the next free Precedence
    numbers instead of failing. Default 50.

.EXAMPLE
    .\New-PreMigInboundRules.ps1 -MoeraDomain "contoso.mail.onmicrosoft.com"
    Creates all 5 rules. Each in-scope object's own domain (firma1.sk, firma2.sk, ...)
    is normalized to its own *.rhlab.in at sync time.

.EXAMPLE
    .\New-PreMigInboundRules.ps1 -MoeraDomain "contoso.mail.onmicrosoft.com" -Rule HideFromGal
    Creates only the HideFromGal rule.

.NOTES
    - Run on the Entra Connect server as Administrator.
    - Back up the configuration first: Get-ADSyncRule | Export-Clixml backup.xml
    - Domain extraction assumes a 2-label old domain (e.g. "firma1.sk"). A 3+
      label domain (e.g. "sub.firma1.sk") would take the wrong label ("sub")
      as the second-level part - verify this matches your data before running
      on the full directory.
    - The ProxyAddresses rule rewrites EVERY primary SMTP: address that doesn't
      already end in .NewDomainSuffix - it does not check against a fixed old
      domain, since objects can come from different domains.
    - Test every rule with Preview (Synchronization Service Manager -> Connectors
      -> Search connector space -> pick object -> Preview) before a Full sync.
    - Verify PreMigSourceAttribute actually matches your AD schema/attribute
      mapping - "extensionAttribute10" is the common default but confirm with
      Get-ADSyncRule for the existing "In from AD - User Common"/"...Exchange"
      rule that flows into the metaverse's customAttribute10.
    - Multiple AD DS connectors (multi-forest lab) trigger an interactive
      picker at runtime unless -ADConnectorName is given.

    Version: 2.3 (2026-08-31)
    Author:  Richard Hlavienka (richard.hlavienka@elyvyn.com)

    Changelog:
      2.3 (2026-08-31) - Replaced real domain names in the help examples with
        neutral placeholders (firma3.sk, rhlab.in, contoso.mail.onmicrosoft.com);
        no functional change.
      2.2 (2026-08-25) - Fixed a bug where "-Rule <single rule>" (anything other
        than "All") silently created nothing: the if/else building $rulesToCreate
        got unrolled from a 1-element array back into a plain string, so
        $rulesToCreate[$i] did character-indexing instead of array-indexing and
        the dispatch switch never matched. Wrapped the whole if/else in @(...).
      2.1 (2026-08-25) - Multiple AD DS connectors no longer throw immediately;
        the script now lists them and prompts interactively for a choice
        (still overridable non-interactively via -ADConnectorName).
      2.0 (2026-08-25) - Removed -OldDomain: domain extraction now happens per
        object inside the sync rule expression (Word/InStr/Right), so objects
        with different old domains (firma1.sk, firma2.sk, ...) are each
        normalized to their own *.NewDomainSuffix. Breaking parameter change.
      1.0 (2026-08-25) - Initial version.
#>

param(
    [string]$NewDomainSuffix = 'rhlab.in',

    [Parameter(Mandatory = $true)]
    [string]$MoeraDomain,

    [string]$PreMigSourceAttribute = 'extensionAttribute10',

    [string]$PreMigValue = 'PreMig',

    [ValidateSet('All', 'Upn', 'Mail', 'ProxyAddresses', 'CustomAttribute5', 'HideFromGal')]
    [string]$Rule = 'All',

    [string]$ADConnectorName,

    [int]$BasePrecedence = 50
)

Import-Module ADSync -ErrorAction Stop

# --- Find the AD DS connector ---
$ADConnectorCandidates = Get-ADSyncConnector | Where-Object { $_.ConnectorTypeName -eq 'AD' }
if ($ADConnectorName) {
    $ADConnector = $ADConnectorCandidates | Where-Object { $_.Name -eq $ADConnectorName }
    if (-not $ADConnector) {
        throw "No AD DS connector named '$ADConnectorName' found. Check: Get-ADSyncConnector | ft Name,ConnectorTypeName"
    }
}
elseif (@($ADConnectorCandidates).Count -eq 1) {
    $ADConnector = $ADConnectorCandidates
}
elseif (@($ADConnectorCandidates).Count -gt 1) {
    Write-Host "Multiple AD DS connectors found (multi-forest lab?). Pick one:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $ADConnectorCandidates.Count; $i++) {
        Write-Host "  [$i] $($ADConnectorCandidates[$i].Name)"
    }
    $selection = Read-Host "Enter the number of the connector to use"
    if ($selection -notmatch '^\d+$' -or [int]$selection -ge $ADConnectorCandidates.Count) {
        throw "Invalid selection '$selection'."
    }
    $ADConnector = $ADConnectorCandidates[[int]$selection]
    Write-Host "Using connector: $($ADConnector.Name)" -ForegroundColor Cyan
}
else {
    throw "Could not find an AD DS connector. Check: Get-ADSyncConnector | ft Name,ConnectorTypeName"
}

Write-Host "Normalizing every in-scope object's domain to *.$NewDomainSuffix" -ForegroundColor Cyan

# --- Which rules will actually be created, in a fixed order ---
$allRuleKeys = @('Upn', 'Mail', 'ProxyAddresses', 'CustomAttribute5', 'HideFromGal')
# @(...) around the whole if/else is required: a single-element result from the
# else branch would otherwise get unrolled from a 1-item array back into a plain
# string, which then breaks $rulesToCreate[$i] (string char-indexing instead of
# array-element indexing) and silently no-ops the dispatch switch below.
$rulesToCreate = @( if ($Rule -eq 'All') { $allRuleKeys } else { $Rule } )

# --- Find enough free Precedence values on the AD DS connector's Inbound rules ---
$usedPrecedences = (Get-ADSyncRule | Where-Object { $_.Direction -eq 'Inbound' -and $_.Connector -eq $ADConnector.Identifier.Guid }).Precedence
$precedences = [System.Collections.Generic.List[int]]::new()
$candidate = $BasePrecedence
while ($precedences.Count -lt $rulesToCreate.Count) {
    if ($usedPrecedences -notcontains $candidate) { $precedences.Add($candidate) }
    $candidate++
}
Write-Host "Using Precedence values: $($precedences -join ', ')" -ForegroundColor Cyan

# --- Name collision check ---
$existing = Get-ADSyncRule | Where-Object { $_.Name -like "*PreMig*$NewDomainSuffix*" }
if ($existing) {
    Write-Warning "Found existing rule(s) with a similar name:"
    $existing | ForEach-Object { Write-Warning " - $($_.Name) (Precedence $($_.Precedence))" }
    $confirm = Read-Host "Continue anyway? (Y/N)"
    if ($confirm -ne 'Y') { throw "Cancelled by user." }
}

function New-InboundPreMigRule {
    param(
        [string]$Name,
        [string]$Description,
        [int]$Precedence,
        [string]$TargetAttribute,
        [ValidateSet('Expression', 'Direct', 'Constant')]
        [string]$FlowType,
        [string]$Expression,
        [string]$SourceAttribute,
        [string]$ConstantValue
    )

    Write-Host "Creating rule: $Name (precedence $Precedence)" -ForegroundColor Cyan

    New-ADSyncRule `
        -Name $Name `
        -Identifier ([guid]::NewGuid()) `
        -Description $Description `
        -Direction 'Inbound' `
        -Precedence $Precedence `
        -PrecedenceAfter  '00000000-0000-0000-0000-000000000000' `
        -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
        -SourceObjectType 'user' `
        -TargetObjectType 'person' `
        -Connector $ADConnector.Identifier.Guid `
        -LinkType 'Join' `
        -SoftDeleteExpiryInterval 0 `
        -ImmutableTag '' `
        -OutVariable syncRule | Out-Null

    switch ($FlowType) {
        'Expression' {
            Add-ADSyncAttributeFlowMapping `
                -SynchronizationRule $syncRule[0] `
                -Destination $TargetAttribute `
                -FlowType 'Expression' `
                -ValueMergeType 'Update' `
                -Expression $Expression `
                -OutVariable syncRule | Out-Null
        }
        'Direct' {
            Add-ADSyncAttributeFlowMapping `
                -SynchronizationRule $syncRule[0] `
                -Destination $TargetAttribute `
                -FlowType 'Direct' `
                -ValueMergeType 'Update' `
                -Source $SourceAttribute `
                -OutVariable syncRule | Out-Null
        }
        'Constant' {
            Add-ADSyncAttributeFlowMapping `
                -SynchronizationRule $syncRule[0] `
                -Destination $TargetAttribute `
                -FlowType 'Constant' `
                -ValueMergeType 'Update' `
                -ConstantValue $ConstantValue `
                -OutVariable syncRule | Out-Null
        }
    }

    $scopeCondition = New-Object `
        -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
        -ArgumentList $PreMigSourceAttribute, $PreMigValue, 'EQUAL'

    Add-ADSyncScopeConditionGroup `
        -SynchronizationRule $syncRule[0] `
        -ScopeConditions @($scopeCondition) `
        -OutVariable syncRule | Out-Null

    Add-ADSyncRule -SynchronizationRule $syncRule[0]

    Write-Host "  -> OK: '$Name' added." -ForegroundColor Green
}

# --- Rule: Upn ---
# Extracts the second-level label from whatever domain the UPN currently has
# (e.g. "firma1.sk" -> "firma1", "firma2.sk" -> "firma2") and rebuilds it as
# <label>.$NewDomainSuffix. Left unchanged if already on $NewDomainSuffix.
function New-UpnRewriteRule {
    param([int]$Precedence)

    $expr = "IIF(LCase(Right([userPrincipalName],Len(`".$NewDomainSuffix`")))=LCase(`".$NewDomainSuffix`"), " +
            "[userPrincipalName], " +
            "Left([userPrincipalName],InStr([userPrincipalName],`"@`")-1) & `"@`" & " +
            "Word(Mid([userPrincipalName],InStr([userPrincipalName],`"@`")+1,Len([userPrincipalName])),1,`".`") & " +
            "`".$NewDomainSuffix`")"

    New-InboundPreMigRule `
        -Name "Custom In from AD - PreMig UPN normalize to *.$NewDomainSuffix" `
        -Description "PreMig: rewrites the userPrincipalName domain's second-level label to <label>.$NewDomainSuffix" `
        -Precedence $Precedence `
        -TargetAttribute 'userPrincipalName' `
        -FlowType 'Expression' `
        -Expression $expr
}

# --- Rule: Mail ---
# Same per-object domain extraction/rewrite as Upn, applied to the mail attribute.
function New-MailRewriteRule {
    param([int]$Precedence)

    $expr = "IIF(LCase(Right([mail],Len(`".$NewDomainSuffix`")))=LCase(`".$NewDomainSuffix`"), " +
            "[mail], " +
            "Left([mail],InStr([mail],`"@`")-1) & `"@`" & " +
            "Word(Mid([mail],InStr([mail],`"@`")+1,Len([mail])),1,`".`") & " +
            "`".$NewDomainSuffix`")"

    New-InboundPreMigRule `
        -Name "Custom In from AD - PreMig Mail normalize to *.$NewDomainSuffix" `
        -Description "PreMig: rewrites the mail domain's second-level label to <label>.$NewDomainSuffix" `
        -Precedence $Precedence `
        -TargetAttribute 'mail' `
        -FlowType 'Expression' `
        -Expression $expr
}

# --- Rule: ProxyAddresses ---
# Rebuilds the whole list per object: keep X500: as-is, keep smtp: aliases ending
# in @$MoeraDomain as-is, rewrite the primary SMTP: address's domain the same
# way as Upn/Mail (per-object second-level label + $NewDomainSuffix), drop
# everything else, then de-duplicate.
function New-ProxyAddressesRule {
    param([int]$Precedence)

    $perItem =
        "IIF(Left(`$item,5)=`"X500:`", `$item, " +
            "IIF(Left(`$item,5)=`"SMTP:`", " +
                "IIF(LCase(Right(`$item,Len(`".$NewDomainSuffix`")))=LCase(`".$NewDomainSuffix`"), " +
                    "`$item, " +
                    "Left(`$item,5) & Mid(`$item,6,InStr(`$item,`"@`")-6) & `"@`" & " +
                        "Word(Mid(`$item,InStr(`$item,`"@`")+1,Len(`$item)),1,`".`") & `".$NewDomainSuffix`"), " +
                "IIF(Left(`$item,5)=`"smtp:`", " +
                    "IIF(LCase(Right(`$item,Len(`"@$MoeraDomain`")))=LCase(`"@$MoeraDomain`"), " +
                        "`$item, " +
                        "NULL), " +
                    "NULL)))"

    $expr = "RemoveDuplicates(Where(`$item,Select(`$item,[proxyAddresses],$perItem),False=IsNullOrEmpty(`$item)))"

    New-InboundPreMigRule `
        -Name "Custom In from AD - PreMig ProxyAddresses normalize to *.$NewDomainSuffix" `
        -Description "PreMig: keeps X500 and @$MoeraDomain smtp: aliases, rewrites the primary SMTP: address's domain to <label>.$NewDomainSuffix, drops the rest" `
        -Precedence $Precedence `
        -TargetAttribute 'proxyAddresses' `
        -FlowType 'Expression' `
        -Expression $expr
}

# --- Rule: CustomAttribute5 (direct copy of the original, pre-rewrite UPN) ---
function New-CustomAttribute5Rule {
    param([int]$Precedence)

    New-InboundPreMigRule `
        -Name "Custom In from AD - PreMig CustomAttribute5 = original UPN" `
        -Description "PreMig: copies the current AD userPrincipalName into customAttribute5 before it gets rewritten" `
        -Precedence $Precedence `
        -TargetAttribute 'customAttribute5' `
        -FlowType 'Direct' `
        -SourceAttribute 'userPrincipalName'
}

# --- Rule: HideFromGal (constant True, not imported from AD) ---
function New-HideFromGalRule {
    param([int]$Precedence)

    New-InboundPreMigRule `
        -Name "Custom In from AD - PreMig HideFromGAL = True" `
        -Description "PreMig: forces msExchHideFromAddressLists to True (constant, not imported from AD)" `
        -Precedence $Precedence `
        -TargetAttribute 'msExchHideFromAddressLists' `
        -FlowType 'Constant' `
        -ConstantValue 'True'
}

# --- Dispatch ---
for ($i = 0; $i -lt $rulesToCreate.Count; $i++) {
    switch ($rulesToCreate[$i]) {
        'Upn'              { New-UpnRewriteRule -Precedence $precedences[$i] }
        'Mail'             { New-MailRewriteRule -Precedence $precedences[$i] }
        'ProxyAddresses'   { New-ProxyAddressesRule -Precedence $precedences[$i] }
        'CustomAttribute5' { New-CustomAttribute5Rule -Precedence $precedences[$i] }
        'HideFromGal'      { New-HideFromGalRule -Precedence $precedences[$i] }
    }
}

Write-Host ""
Write-Host "Done. Next steps:" -ForegroundColor Yellow
Write-Host "1. Verify the rules in the Synchronization Rules Editor (GUI) or via Get-ADSyncRule."
Write-Host "2. Preview a test object flagged with $PreMigSourceAttribute = $PreMigValue."
Write-Host "3. Run a Full sync: Start-ADSyncSyncCycle -PolicyType Initial"
