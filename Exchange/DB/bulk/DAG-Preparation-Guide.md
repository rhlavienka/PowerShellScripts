# Preparing a Database Availability Group (DAG) — Two-Site Stretch with DatacenterActivationMode DagOnly

This guide covers the steps needed **before** [Initialize-ExchangeVolumes.ps1](Initialize-ExchangeVolumes.ps1) and [Deploy-ExchangeDatabases.ps1](Deploy-ExchangeDatabases.ps1) can be run against a DAG. It assumes the topology implied by the layout CSVs in this folder ([DBSE_Layout_192DB_12vol_HtoS.csv](DBSE_Layout_192DB_12vol_HtoS.csv), [DBSE_Layout_192DB_12vol_KtoV.csv](DBSE_Layout_192DB_12vol_KtoV.csv)):

- **Site S1**: `EXSVR-S1-1`, `EXSVR-S1-2`, `EXSVR-S1-3`, `EXSVR-S1-4` (4 mailbox servers, primary/active copies)
- **Site S2**: `EXSVR-S2-1`, `EXSVR-S2-2`, `EXSVR-S2-3`, `EXSVR-S2-4` (4 mailbox servers, resilience/AP3 copies)

Because the DAG is stretched across two Active Directory sites, it needs `DatacenterActivationMode DagOnly` so that a full site failure requires a deliberate, manual switchover instead of an automatic one. Replace all placeholders in angle brackets (`<...>`) with real values for your environment.

## 1. Prerequisites

- All 8 members (`EXSVR-S1-1..4`, `EXSVR-S2-1..4`) are joined to the domain, patched, and running the same Exchange version/CU.
- **Failover Clustering** Windows feature is installed on every member (Exchange installs this automatically, but verify):
  ```powershell
  Get-WindowsFeature Failover-Clustering | Select Name, InstallState
  ```
- **`Remote Registry` service is `Automatic`/running on every member being added.** `Add-DatabaseAvailabilityGroupServer` / `AddClusterNode()` reads and writes registry data on the joining node remotely, over Remote Registry — if it's `Disabled` (common on hardened images/security baselines that turn it off as an attack-surface reduction measure), the join fails with `Cluster API failed: "AddClusterNode() (MaxPercentage=12) failed with 0x35. Error: The network path was not found"`, even though the node is otherwise perfectly reachable on the network. Check and fix on every member *before* step 4:
  ```powershell
  Get-Service RemoteRegistry | Select Name, Status, StartType
  # if StartType is not Automatic:
  Set-Service RemoteRegistry -StartupType Automatic
  Start-Service RemoteRegistry
  ```
- Static IP addressing on every member; no DHCP on the MAPI network.
- A single network (MAPI) per server is sufficient — see the note in step 5 on why a dedicated replication network is no longer the default recommendation on modern 10 Gbit+ links.
- Name resolution: all 8 servers can resolve each other and the witness server by short name and FQDN, in both AD sites.
- Firewall rules allow cluster/replication traffic between the two sites (ICMP, RPC, and the ports used by the File Share Witness/SMB — see Microsoft's DAG networking requirements for the exact port list for your Exchange version).
- Round-trip network latency between site S1 and site S2 should meet Microsoft's supported threshold for stretched DAGs (a low, stable RTT — verify against the requirement for your Exchange version before going live).

## 2. Decide Witness Server Placement

The **File Share Witness (FSW)** breaks quorum ties. With `DatacenterActivationMode DagOnly`, the design deliberately gives up *automatic* cross-site failover in exchange for avoiding split-brain — so, unlike a fully automatic multi-site DAG, there is no requirement for a neutral third-location witness here. A witness in each site is the standard pattern for this scenario:

- **Primary witness**, in site S1: `<FSW-SERVER-S1>`. This is the one the DAG actually uses for quorum during normal operations.
- **Alternate witness**, in site S2: `<FSW-SERVER-S2>`. It sits idle and is only brought into play by `Restore-DatabaseAvailabilityGroup -AlternateWitnessServer ...` during a manual datacenter switchover (step 7) — that is the whole point of `DagOnly`: a human decides when S2's witness takes over, instead of Windows Failover Clustering deciding automatically.
- Neither witness may be a DAG member. Being a domain controller is technically supported on Exchange 2016+ but still discouraged. Each just needs the File Server role and enough local disk for the witness share.
- If a third, fully independent location (a third AD site, or Azure) is available, it can still be used instead of/in addition to this pattern for extra resilience — but for a two-site `DagOnly` design it isn't required.

Both the witness directory and the share are created automatically by `New-DatabaseAvailabilityGroup` (for the primary) and `Restore-DatabaseAvailabilityGroup` (for the alternate) — no manual `New-Item`/share setup needed, as long as the Exchange Trusted Subsystem / DAG computer account has local admin rights on the witness server at the time. Grant that first, on both `<FSW-SERVER-S1>` and `<FSW-SERVER-S2>`:

```powershell
Add-LocalGroupMember -Group 'Administrators' -Member '<DomainName>\Exchange Trusted Subsystem'
```

## 3. Create the DAG

Run from Exchange Management Shell on any Exchange server (not necessarily a future DAG member):

```powershell
New-DatabaseAvailabilityGroup -Name '<DAGName>' `
    -WitnessServer '<FSW-SERVER-S1>' `
    -WitnessDirectory 'C:\DAGWitnesses\<DAGName>' `
    -DatabaseAvailabilityGroupIpAddresses 255.255.255.255
```

Notes:

- `255.255.255.255` creates the DAG **without** a cluster administrative access point (an "IP-less" DAG — technically a Failover Cluster with a Distributed Network Name instead of a Network Name + IP resource). This is Microsoft's current recommended default for Exchange 2016/2019/SE, not just a fallback for single-subnet clusters:
  - Exchange itself never talks to the cluster over that IP — DAG members locate each other via Active Directory, so the cluster IP/name is pure Windows Failover Clustering plumbing that Exchange doesn't need.
  - It removes the need to reserve, register in DNS, and keep static a cluster IP in *each* site (S1 and S2) — one less moving part to fail or drift, and one less object that has to "belong" to a specific site.
  - It's fully supported for a multi-subnet/multi-site DAG on Windows Server 2012 R2 and later — Exchange SE requires Windows Server 2019+/2022, so this requirement is already satisfied.
  - The only reason to fall back to real per-site cluster IPs (`-DatabaseAvailabilityGroupIpAddresses ('<Site-S1-Cluster-IP>','<Site-S2-Cluster-IP>')`) is a hard dependency from monitoring/tooling on a resolvable cluster network name+IP, or an unsupported/legacy OS — neither applies here.
- `<DAGName>` becomes both the Active Directory object name and the cluster name — keep it short (≤ 15 characters) and free of special characters.

## 4. Add Members to the DAG

Add all 8 servers, alternating sites is not required but confirm each server's `Get-ExchangeServer | fl Site` matches S1/S2 as expected before adding:

```powershell
$dagMembers = 'EXSVR-S1-1','EXSVR-S1-2','EXSVR-S1-3','EXSVR-S1-4',
              'EXSVR-S2-1','EXSVR-S2-2','EXSVR-S2-3','EXSVR-S2-4'

foreach ($server in $dagMembers) {
    Write-Host "Adding $server to DAG <DAGName>..."
    Add-DatabaseAvailabilityGroupServer -Identity '<DAGName>' -MailboxServer $server
}
```

This step creates the underlying Windows Failover Cluster and joins each server to it. Expect a short service interruption on each server as clustering is configured — run outside business hours.

Verify:

```powershell
Get-DatabaseAvailabilityGroup -Identity '<DAGName>' -Status | Select Name, Servers, WitnessServer, WitnessDirectory
```

## 5. Verify DAG Networks (optional, independent of the IP-less cluster access point)

`255.255.255.255` in step 3 only removes the cluster's *administrative access point* (the network name + IP that Windows Failover Clustering itself would otherwise expose) — it has no effect on DAG networks, which Exchange derives independently from the actual subnets present on each member's NIC. So this step is unrelated to the IP-less decision; it's here only because it's good practice to confirm auto-discovery matched reality, not because anything needs configuring.

After all members are added, Exchange auto-discovers networks. With a single NIC/subnet per server (see step 1), one collapsed network carrying both MAPI and replication traffic is the expected, correct result — no action is needed if it looks right:

```powershell
Get-DatabaseAvailabilityGroupNetwork -Identity '<DAGName>' | Format-List Name, Subnets, ReplicationEnabled, MapiAccessEnabled
```

A dedicated replication network (separate NIC/subnet, `ReplicationEnabled $true` / `MapiAccessEnabled $false` on it and the reverse on the MAPI network) was Microsoft's original Exchange 2010-era guidance, written for 1 Gbit links where MAPI and reseed/log-shipping traffic could genuinely starve each other. On today's 10 Gbit+ links that contention essentially disappears, and Microsoft's own current guidance for Exchange 2016+ favors a single network unless you have a specific reason to isolate traffic (e.g. a WAN-constrained link between S1 and S2, or a compliance requirement to segregate replication traffic). Only add a second, dedicated network if such a reason actually applies here:

```powershell
New-DatabaseAvailabilityGroupNetwork -DatabaseAvailabilityGroup '<DAGName>' `
    -Name 'Replication' `
    -Subnets '<Repl-Subnet-S1>/24','<Repl-Subnet-S2>/24' `
    -ReplicationEnabled $true

Set-DatabaseAvailabilityGroupNetwork -Identity '<DAGName>\MAPI Network' -ReplicationEnabled $false
```

## 6. Enable DatacenterActivationMode DagOnly

With a DAG stretched across two AD sites, automatic cross-site failover is risky (network partition can look identical to a real site loss). `DagOnly` keeps automatic failover **within** a site, but requires a **manual** switchover to fail databases **across** sites.

```powershell
Set-DatabaseAvailabilityGroup -Identity '<DAGName>' -DatacenterActivationMode DagOnly
```

Prerequisites/side effects to know before enabling:
- All DAG members must already be added (step 4) and the DAG must be healthy (`Test-ReplicationHealth` clean on all members).
- Once `DagOnly` is set, if an entire site goes down, **do not** just restart the surviving nodes and expect databases to mount — you must run the manual activation procedure in step 7.
- Re-run `Get-DatabaseAvailabilityGroup -Identity '<DAGName>' -Status | fl DatacenterActivationMode` to confirm it reports `DagOnly`.

## 7. Manual Datacenter Switchover Procedure (reference — not part of initial setup)

Keep this for the runbook; it is **not** run during normal deployment, only during an actual site-S1 (or site-S2) outage.

```powershell
# 1. On a surviving server in the healthy site, stop the DAG on the failed site's
#    servers (only if they are truly down/unreachable — do NOT run this if S1 is
#    merely unreachable due to a transient network blip):
Stop-DatabaseAvailabilityGroup -Identity '<DAGName>' -ActiveDirectorySite '<Failed-Site>' -ConfigurationOnly

# 2. Restore/activate the DAG in the surviving site, pointing at the alternate
#    witness prepared in step 2:
Restore-DatabaseAvailabilityGroup -Identity '<DAGName>' `
    -ActiveDirectorySite '<Surviving-Site>' `
    -AlternateWitnessServer '<Alternate-FSW-Server>' `
    -AlternateWitnessDirectory 'C:\DAGWitnesses\<DAGName>-Alt'

# 3. Once the failed site is back and its servers are rebuilt/verified healthy,
#    start the DAG on those servers again to bring them back into replication:
Start-DatabaseAvailabilityGroup -Identity '<DAGName>' -ActiveDirectorySite '<Failed-Site>'
```

Always validate replication health before and after a switchover:

```powershell
Test-ReplicationHealth -Identity <ServerName>
```

## 8. Proceed with Database Deployment

Once the DAG is created, all 8 members are joined, networks are configured, and `DatacenterActivationMode` is `DagOnly`:

1. Provision the data volumes on every member — see [Initialize-ExchangeVolumes.ps1](Initialize-ExchangeVolumes.ps1) (run locally per server, WhatIf first).
2. Deploy the databases and copies from a layout CSV — see [Deploy-ExchangeDatabases.ps1](Deploy-ExchangeDatabases.ps1), e.g.:
   ```powershell
   .\Deploy-ExchangeDatabases.ps1 -CsvPath '.\DBSE_Layout_192DB_12vol_HtoS.csv' -WhatIfMode $false
   ```
3. After deployment, re-verify overall DAG health:
   ```powershell
   Get-DatabaseAvailabilityGroup -Identity '<DAGName>' -Status
   Get-MailboxDatabaseCopyStatus -Server * | Where-Object { $_.Status -notlike 'Healthy' -and $_.Status -ne 'Mounted' }
   ```
