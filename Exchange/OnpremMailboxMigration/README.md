# On-prem mailbox migration (moves within one Exchange organization)

A GUI-free toolset for moving mailboxes **between mailbox databases inside a
single on-premises Exchange organization** - database consolidation, storage /
hardware refresh, a new DAG, retiring a database, or spreading load off an
overweight database.

This set is **on-premises only**. Everything runs in the on-premises
Exchange Management Shell against `New-MoveRequest` and the Mailbox Replication
Service.

> Section ["Still open"](#still-open--not-yet-built) lists what is deliberately
> not built yet.

**Mechanism: `New-MoveRequest` per mailbox**, tagged with a shared
`-BatchName <wave>` so a whole wave is one filter (`Get-MoveRequest -BatchName`).
`New-MoveRequest` natively provides what this task needs:

| Requirement | `New-MoveRequest` parameter |
|---|---|
| finalization time | `-CompleteAfter <datetime>` (keeps syncing, finalizes no earlier) |
| hold at 95%, finalize manually | `-SuspendWhenReadyToComplete` |
| target database | `-TargetDatabase` (single) - **usually omitted; Exchange auto-selects** a provisioning-enabled database and balances the load itself (automatic mailbox distribution). The scripts never spread mailboxes across databases themselves. |
| corrupt / oversized items | `-BadItemLimit`, `-LargeItemLimit` |
| primary only / archive only | `-PrimaryOnly`, `-ArchiveOnly`, `-ArchiveTargetDatabase` |
| MRS scheduling weight | `-Priority` |

**Scope selection: `Get-Recipient -Filter` (server-side OPATH)** on `City` /
`Department` combined with `-OrganizationalUnit`, so selection stays fast in a
large directory.

---

## Is MRS needed for on-prem moves?

**Yes.** The **Mailbox Replication Service (MRS)** performs *every* mailbox move
in Exchange 2013+, including a local move between two databases in the same org.
`New-MoveRequest` only writes the request into Active Directory / the system
mailbox; MRS on a Mailbox server picks it up and does the actual copy.

What you do **not** need for a local move is **MRSProxy** - that is the
externally published MRS endpoint used only for *remote* moves (cross-forest and
hybrid). For moves inside one org MRS talks to the databases directly.

Practical consequences, and why `Get-MRSHealth.ps1` is in the set:

- MRS throttling still applies: `MaxActiveMovesPerSourceMDB`,
  `MaxActiveMovesPerTargetMDB`, `MaxActiveMovesPerSourceServer`,
  `MaxActiveMovesPerTargetServer`, `MaxTotalMovesPerMRS`
  (`Get-MailboxReplicationService`). A wave larger than these limits does not
  fail - it simply queues, and the extra requests sit at `Queued` until a slot
  frees. That is expected; the status script surfaces it.
- `MSExchangeMailboxReplication` must be running on the Mailbox servers.
- For a **cross-version** move (e.g. 2016 -> 2019 database) MRS on the **target**
  (higher) version runs the move.

---

## Scripts

| # | Script | Purpose |
|---|--------|---------|
| 01 | `01-Get-OnPremMigrationScope.ps1` | Turn OU / City / Department into a validated mailbox list + readiness report |
| 02 | `02-New-OnPremMoveRequest.ps1` | Create the move requests from the CSV, all tagged with one `-BatchName` |
| 03 | `03-Get-OnPremMigrationStatus.ps1` | Progress dashboard, run-over-run history, stall detection, per-database landing |
| 04 | `04-Get-OnPremMigrationFailureReport.ps1` | Per-mailbox deep dump + likely cause / action, zipped for a support case |
| 05 | `05-Get-OnPremMigrationCompletionReport.ps1` | Reconcile the original scope against the final state; management rollup |
| 06 | `06-Invoke-OnPremMoveRequestControl.ps1` | Suspend / Resume / Complete / Remove / Set move requests for a wave or named mailboxes |
| - | `Remove-CompletedMoveRequests.ps1` | Housekeeping - clear finished move requests that would block a re-move |
| - | `Get-MRSHealth.ps1` | MRS throttling limits vs. current active moves per database / server |
| - | `OnPremMigration.Settings.psd1` | (data, git-ignored) site defaults: notification addresses, size caps, output root |

All scripts run in the on-premises Exchange Management Shell (Exchange 2013+).

---

## Typical flow

```
                 ┌──────────────────────────────────┐
                 │ 01-Get-OnPremMigrationScope       │  OU / City / Department  ->  Wave1_Ready_*.csv
                 └─────────────┬────────────────────┘   + Wave1_ScopeReport_*.csv (FAIL/WARN/PASS)
                               │  fix FAIL rows, review WARN rows
                 ┌─────────────▼────────────────────┐
   optional      │ Get-MRSHealth                     │  will MRS run the wave now, or queue it?
                 └─────────────┬────────────────────┘
                 ┌─────────────▼────────────────────┐
                 │ 02-New-OnPremMoveRequest          │  New-MoveRequest per row, -BatchName "Wave1"
                 └─────────────┬────────────────────┘   -CompleteAfter / -SuspendWhenReadyToComplete
                 ┌─────────────▼────────────────────┐
                 │ 03-Get-OnPremMigrationStatus      │  run repeatedly; appends history
                 └─────────────┬────────────────────┘
                               │  Failed / Stalled / CorruptItems?
                 ┌─────────────▼────────────────────┐
                 │ 04-Get-OnPremMigrationFailure...  │  per-user dump + likely cause
                 └─────────────┬────────────────────┘
                               │  at the finalization window
                 ┌─────────────▼────────────────────┐
                 │ 06-Invoke-...Control -Complete     │  (or the request's own -CompleteAfter fires)
                 └─────────────┬────────────────────┘
                 ┌─────────────▼────────────────────┐
                 │ 05-Get-OnPremMigrationCompletion  │  every requested mailbox moved? where did it land?
                 └─────────────┬────────────────────┘
                 ┌─────────────▼────────────────────┐
                 │ Remove-CompletedMoveRequests      │  housekeeping once 05 is clean
                 └──────────────────────────────────┘
```

---

## 01 - Get-OnPremMigrationScope

Builds the "what to migrate" list from directory attributes and runs the
source-side readiness checks.

### Selection parameters (all optional, combined with AND)

| Parameter | Filters on | Notes |
|---|---|---|
| `-OrganizationalUnit <ou>` | `Get-Recipient -OrganizationalUnit` | canonical (`contoso.com/Users/Sales`) or DN; whole subtree |
| `-City <string[]>` | `City` (AD `l`) | OPATH `City -eq '...'`; multiple values OR'd |
| `-Department <string[]>` | `Department` | OPATH `Department -eq '...'`; multiple values OR'd |
| `-RecipientTypeDetails <string[]>` | mailbox type | **optional, no default.** If omitted, no type filter is applied to `Get-Recipient`; the readiness pass still FAILs any non-movable recipient it finds. Pass e.g. `UserMailbox,SharedMailbox` to narrow the selection up front. |
| `-ExcludeDatabase <string[]>` | current `Database` | drop mailboxes already on an acceptable database |
| `-Identity <string[]>` / `-CsvPath` | explicit list | bypass the attribute filter, still run readiness |

At least one of `-OrganizationalUnit` / `-City` / `-Department` / `-Identity` /
`-CsvPath` must be supplied - the script refuses to select the whole org by
accident.

### Guardrail parameters

| Parameter | Default | Effect |
|---|---|---|
| `-MaxMailboxSizeGB` | 50 | WARN only |
| `-MaxItemCount` | 200000 | WARN only |
| `-MaxWaveMailboxes` | 500 | above this the Ready CSV is split into `_part01`, `_part02`, ... (operational guidance; there is no hard `New-MoveRequest` batch limit) |
| `-MaxWaveSizeGB` | 2000 | same, by summed primary-mailbox size. 0 disables it |
| `-TargetDatabase <string>` | (none) | if given, written into a `TargetDatabase` column on every Ready-CSV row; **if omitted the column is left out and Exchange auto-selects and load-balances the target** (the expected default here) |
| `-IncludeArchive` | off | also emit a `TargetArchiveDatabase` column (value = `-TargetArchiveDatabase`, else blank) |
| `-TargetArchiveDatabase <string>` | (none) | value for the `TargetArchiveDatabase` column when `-IncludeArchive` is set |
| `-IdentityColumn` | `EmailAddress` | identity column name in `-CsvPath` |
| `-WaveName` | `Wave_<yyyyMMdd_HHmmss>` | wave sub-folder name + file-name prefix; 02 passes it to `-BatchName` |
| `-OutputFolder` | `$PSScriptRoot` (or `OutputRoot` from the settings file) | the wave sub-folder `<OutputFolder>\<WaveName>` is created under this |

### Readiness checks (per mailbox)

- recipient exists and is a movable mailbox type (`UserMailbox`, `SharedMailbox`,
  `RoomMailbox`, `EquipmentMailbox`) - anything else FAILs
- **no existing move request** for the mailbox (`Get-MoveRequest -Identity`)
- no pending `Restore` / `MailboxImport` / `MailboxExport` request
- `ExchangeGuid` is stamped (non-empty)
- current `Database` is **not** in `-ExcludeDatabase` / not already an acceptable target
- `Get-MailboxStatistics`: not `DisconnectDate`; size / item count vs the caps
- archive present? on which database? (drives `-IncludeArchive`)
- holds (`LitigationHoldEnabled`, `ComplianceTagHoldApplied`, `DelayHold*`,
  `InPlaceHolds`) - informational; recoverable-items space counts against the move

### Output

| File | Content |
|---|---|
| `<wave>_ScopeReport_<ts>.csv` | full per-mailbox result: identity, DisplayName, type, PrimarySmtp, current DB, SizeGB, ItemCount, ArchiveDB, Holds, Status (PASS/WARN/FAIL), Issues |
| `<wave>_Ready_<ts>.csv` | `EmailAddress` (+ `TargetDatabase` / `TargetArchiveDatabase` if requested), PASS/WARN rows only - this feeds 02 |
| `<wave>_ScopeReport_<ts>.log` | console transcript |

---

## 02 - New-OnPremMoveRequest

Reads the `<wave>_Ready_*.csv` and issues one `New-MoveRequest` per row, all
sharing `-BatchName <WaveName>`.

### Parameters

| Parameter | Req | Effect |
|---|---|---|
| `-CsvPath` | yes | Ready CSV from 01. `EmailAddress` column required; optional `TargetDatabase` / `TargetArchiveDatabase` columns are honoured per row |
| `-WaveName` | no | value passed to `New-MoveRequest -BatchName`; how 03/04/05/06 find the wave. <= 64 chars. Defaults to the CSV's `_Ready_` prefix (e.g. `Praha-Sales-W1_Ready_...csv` -> `Praha-Sales-W1`), else the CSV base name |
| `-TargetDatabase <string>` | no | **Usually omitted.** When omitted, `-TargetDatabase` is not passed to `New-MoveRequest` and Exchange uses automatic mailbox distribution - it places and load-balances each mailbox across databases where `IsExcludedFromProvisioning`, `IsExcludedFromProvisioningByOperator`, `IsExcludedFromProvisioningBySpaceMonitoring` and `IsSuspendedFromProvisioning` are all `$false`. When given, every mailbox goes to that one database. A per-row CSV `TargetDatabase` column still wins over the parameter. The scripts never round-robin or capacity-check - that is Exchange's job. |
| `-TargetArchiveDatabase <string>` | no | archive target for rows without a `TargetArchiveDatabase` CSV column |
| `-CompleteAfter <datetime>` | no | **finalization time.** The move syncs continuously but finalizes no earlier than this. Local time in, resolved UTC echoed back. Mutually exclusive with `-SuspendWhenReadyToComplete`. |
| `-SuspendWhenReadyToComplete` | no | sync to ~95%, then `AutoSuspended` until 06 (`Resume`) finalizes it. Mutually exclusive with `-CompleteAfter`. |
| `-StartAfter <datetime>` | no | MRS does not begin the move before then |
| `-BadItemLimit` | no | default 0 |
| `-LargeItemLimit` | no | default 0 |
| `-Priority` | no | `Normal` (default) / `High` / ... - MRS scheduling weight |
| `-PrimaryOnly` / `-ArchiveOnly` | no | move only one of the two stores; `-ArchiveOnly` pairs with `TargetArchiveDatabase` |
| `-Delimiter` | no | for the local CSV validation pass |

`-WhatIf` / `-Confirm` via `[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]`.

> `New-MoveRequest` has **no** completion-notification e-mail (that was a
> migration-batch feature). Completion is tracked by 03 / 05; a notification
> wrapper is [not built yet](#still-open--not-yet-built).

### Pre-flight

- on-prem EMS session (`Get-ExchangeServer` present)
- CSV exists, has `EmailAddress`, no duplicate addresses
- every address resolves to a local mailbox with **no existing move request**
- any `TargetDatabase` value (parameter or CSV column) exists and is mounted;
  excluded-from-provisioning is a WARN, not a FAIL
- if no target database anywhere: note that automatic distribution will be used
  and list the current candidate databases, then require confirmation
- summary block (wave name, mailbox count, target mode, CompleteAfter /
  SuspendWhenReady, limits, priority) before `ShouldProcess`

### Per-row call (shape)

```powershell
$p = @{
    Identity       = $mailboxGuid          # resolved in the pre-flight, not the raw address
    BatchName      = $WaveName
    BadItemLimit   = $BadItemLimit
    LargeItemLimit = $LargeItemLimit
    Confirm        = $false
}
# effective target DB: per-row CSV value, else -TargetDatabase, else omit (auto distribution)
if ($effectiveTargetDb)        { $p.TargetDatabase        = $effectiveTargetDb }
if ($effectiveArchiveDb)       { $p.ArchiveTargetDatabase = $effectiveArchiveDb }
if ($hasCompleteAfter)         { $p.CompleteAfter               = $CompleteAfter }
if ($SuspendWhenReadyToComplete) { $p.SuspendWhenReadyToComplete = $true }
if ($hasStartAfter)            { $p.StartAfter            = $StartAfter }
if ($Priority)                { $p.Priority              = $Priority }
if ($PrimaryOnly)             { $p.PrimaryOnly           = $true }
if ($ArchiveOnly)             { $p.ArchiveOnly           = $true }
New-MoveRequest @p
```

`-CompleteAfter` / `-StartAfter` are tested with `$PSBoundParameters.ContainsKey`,
not for truthiness - an unbound `[datetime]` is `DateTime.MinValue`, which is
truthy. A per-row failure is logged and the loop continues; the run ends with a
created / skipped / failed tally and a `<wave>_MoveRequests_<ts>.csv` of what was
issued.

---

## 03 - Get-OnPremMigrationStatus

Progress dashboard built on `Get-MoveRequest` + `Get-MoveRequestStatistics`.

### Scope (pick one)

- `-WaveName <name>` - `Get-MoveRequest -BatchName <name>`
- `-CsvPath <file>` - the addresses in a CSV (`-IdentityColumn`, default `EmailAddress`)
- neither - every move request in the org (`Get-MoveRequest`)

`-GridView` also sends the per-user snapshot to `Out-GridView` (a sortable,
filterable window) when the cmdlet is available on the host. The CSV / log /
history files are written regardless.

### Classification (from `Status` / `StatusDetail`)

| Class | Meaning |
|---|---|
| Queued | `Queued` - waiting for an MRS slot (throttling) |
| Provisioning | `InProgress`, `CreatingInitialSyncCheckpoint` / folder hierarchy |
| Syncing | `InProgress`, initial sync running |
| Synced | `AutoSuspended` - initial sync done, waiting for finalization |
| Completing | `CompletionInProgress` |
| Completed | `Completed` / `CompletedWithWarning` |
| Suspended | manually `Suspended` |
| Failed | `Failed` / `CompletionFailed` |
| **Stalled** | still Syncing, `StalledSinceTimestamp` set, **or** `PercentComplete` + `ItemsTransferred` unchanged since the previous run and `-StallHours` (default 6) elapsed |
| **CorruptItems** | `Failed`, and the failure is only bad/large items over the limit |

### Output

| File | Content |
|---|---|
| `<wave>_Status_<ts>.csv` | per-mailbox snapshot: Timestamp, Index, Identity, Wave, Class, Status, StatusDetail, PercentComplete, ItemsTransferred, BytesTransferred, SourceDatabase, TargetDatabase, StalledSince, LastUpdateTimestamp, BadItemsEncountered, LargeItemsEncountered, FailureType, Message |
| `<wave>_Status_<ts>.log` | console transcript |
| `MigrationStatus-History.csv` | appended every run - the file that makes stall detection work; **keep it** |

Console: per-user line, then a **Class summary**, a **per-target-database
landing** table (how many mailboxes went where - the check that automatic
distribution behaved), a **failure-group** summary (grouped by `FailureType` /
first line of the message), and transfer-rate / rough-ETA figures.

---

## 04 - Get-OnPremMigrationFailureReport

For each failing (or named) mailbox, dumps:

- `Get-MoveRequestStatistics -IncludeReport -DiagnosticInfo "verbose;showtimeslots;showtimeline"`
- `Report.Failures`, `Report.BadItems`, `Report.LargeItems`, the move-report entries
- `Get-MailboxStatistics` source vs (if finalized) target
- maps the failure to a **likely cause** and a **suggested action**
  (`FailureSummary.csv`), e.g.
  `MapiExceptionNotFound` -> stale/orphaned move request, remove and re-create;
  `TooManyBadItemsPermanentException` -> raise `BadItemLimit` or run a repair;
  `StoragePermanentException` / `QuotaExceededException` -> target database space
  or mailbox quota;
  `MailboxReplicationTransientException` / communication errors -> MRS load or
  transient store issue, resume and watch
- zips the per-run folder for a support case

Scope (pick one): `-Identity <addr[]>`, `-WaveName`, or `-CsvPath`. With
`-WaveName` / `-CsvPath` only the Failed / stalled mailboxes are dumped unless
`-IncludeAll` is given.

---

## 05 - Get-OnPremMigrationCompletionReport

Run after finalization. Needs `-WaveName`, `-CsvPath` (the original Ready CSV),
or both - the CSV defines the set that *should* have moved, the wave gives the
actual move requests. `-ExpectDatabase <db[]>` flags mailboxes that completed
onto an unexpected database (omit for automatic distribution); `-IncludeGridView`
opens the reconciliation in `Out-GridView`. Reconciles against the live state:

- every requested mailbox now `Completed`? list anything still `Queued` /
  `Syncing` / `Synced` / `Failed` / with no move request at all
- each mailbox's current `Database` vs the expected target (or "any new database"
  when automatic distribution was used)
- source vs target `ItemCount` delta, `BadItemsEncountered`,
  `LargeItemsEncountered`
- per-mailbox duration, bytes, average rate; per-wave rollup (count, total GB,
  wall-clock window, mailboxes/hour)
- a management-friendly `<wave>_Completion_<ts>.csv` + a short text summary

---

## 06 - Invoke-OnPremMoveRequestControl

Thin, auditable wrapper so finalization and mid-wave control do not need the EAC.
`-Action` is one of:

| `-Action` | Runs |
|---|---|
| `Suspend` | `Suspend-MoveRequest` on `InProgress` / `Queued` requests (`-SuspendComment` optional) |
| `Resume` | `Resume-MoveRequest` on `Suspended` / `AutoSuspended` / `Failed` / `CompletionFailed` |
| `Complete` | `Resume-MoveRequest` on `AutoSuspended`; `Set-MoveRequest -CompleteAfter <now> -SuspendWhenReadyToComplete:$false` on syncing / queued / suspended |
| `SetLimit` | `Set-MoveRequest -BadItemLimit / -LargeItemLimit / -Priority` (needs at least one; `AcceptLargeDataLoss` added automatically at >= 51) |
| `Remove` | `Remove-MoveRequest` - refuses non-`Completed*` unless `-Force` |

Scope: `-WaveName` (maps to `-BatchName`) and/or `-Identity <string[]>` - at
least one is required. `-OnlyStatus <status[]>` narrows to move requests in a
given state. `SupportsShouldProcess`; a status breakdown is printed and every
action is confirmed before it runs.

---

## Supporting tools

### Remove-CompletedMoveRequests.ps1
`Completed` / `CompletedWithWarning` move requests (optionally `-WaveName <wave>`)
whose completion is older than `-OlderThanDays` (default 7, 0 = no age filter)
-> `Remove-MoveRequest`. `-IncludeFailed` also clears `Failed` / `CompletionFailed`
(you lose their statistics). Completed move requests linger and block a later
move of the same mailbox. `SupportsShouldProcess`, `ConfirmImpact='High'`.

### Get-MRSHealth.ps1
`Get-MailboxReplicationService` per Mailbox server (`-Server` to narrow):
`MaxActiveMovesPerSourceMDB` / `PerTargetMDB` / `PerSourceServer` /
`PerTargetServer` / `MaxTotalMovesPerMRS`, plus the current count of active
(`Queued` / `InProgress` / `CompletionInProgress`) `MoveRequest`s per source
database, per target database and overall - rows already at a limit are flagged.
`-IncomingMoves <n>` projects the queue depth after you add `n` requests. So a
wave that will merely **queue** behind the throttling limits is visible before it
is started. No MRSProxy check - not used for local moves.

### OnPremMigration.Settings.psd1  (git-ignored)
Optional site defaults so the numbered scripts can run with fewer switches:

```powershell
@{
    NotificationEmails    = @('messaging-team@contoso.com')  # used by 03/05 wrappers, not by New-MoveRequest
    MaxMailboxSizeGB      = 50
    MaxWaveMailboxes      = 500
    OutputRoot            = 'D:\Migration\Waves'
    DefaultBadItemLimit   = 0
    DefaultLargeItemLimit = 0
}
```

A committed `OnPremMigration.Settings.template.psd1` documents the shape; the
real file goes into `.gitignore` (same pattern as
`Exchange/CrossTenantAPP/CrossTenant.Settings.psd1`). Explicit parameters always
win over the settings file.

---

## Output layout

Every run of a numbered script writes into a per-wave folder so a wave's
artefacts stay together:

```
<OutputFolder>\
  Praha-Sales-W1\                                     <- the wave folder = <OutputFolder>\<WaveName>
    Praha-Sales-W1_ScopeReport_2026-09-15_09-00-00.csv
    Praha-Sales-W1_Ready_2026-09-15_09-00-00.csv
    Praha-Sales-W1_MoveRequests_2026-09-15_09-30-00.csv
    Praha-Sales-W1_Status_2026-09-15_10-30-00.csv
    MigrationStatus-History.csv
    Praha-Sales-W1_Completion_2026-09-16_02-00-00.csv
    Praha-Sales-W1_FailureReport_2026-09-15_23-10-00\ ...  (+ .zip)
    Control_Complete_2026-09-15_22-00-00.log
```

The wave folder is `<OutputFolder>\<WaveName>` - no date in the folder name, so
02-06 find each other by wave name alone. `-WaveName` defaults to
`Wave_<yyyyMMdd_HHmmss>` on 01 and to the CSV's `_Ready_` prefix on 02.
`-OutputFolder` defaults to `$PSScriptRoot`, or `OutputRoot` from
`OnPremMigration.Settings.psd1` when that file is present.

---

## Requirements

- On-premises **Exchange Management Shell**, Exchange 2013 or newer. No Exchange
  Online / Graph modules.
- RBAC: **Move Mailboxes** management role (in *Recipient Management* and
  *Organization Management*) for 01/02/06; *View-Only Recipients* is enough for
  03/04/05.
- **MRS** (`MSExchangeMailboxReplication`) running on the Mailbox servers; for a
  cross-version move the target database must be on the equal-or-higher version.
- Keep `MigrationStatus-History.csv` between runs of 03.

---

## Still open / not yet built

1. **Cross-version / cross-DAG moves** - if any target databases are on a newer
   Exchange build than the source, `Get-MRSHealth` and the completion report
   could add version-aware checks.
2. **Archive to a *different* database than the primary** - supported via a
   `TargetArchiveDatabase` CSV column / `-TargetArchiveDatabase`, but not yet
   surfaced in 01's readiness output.
3. **Completion notification** - `New-MoveRequest` sends none. A Task Scheduler
   wrapper that runs 03 on a timer and e-mails the summary / escalates on
   Failed/Stalled is not built (the `NotificationEmails` setting is a placeholder
   for it).
4. **Scope-drift re-check** - a `-ReCheck` mode on 01 that diffs the current
   filter result against a live wave just before finalization is not built.
