# SharePoint Online Production-to-Test List Synchronization Plan

**Status:** Implementation-ready package  
**Prepared for:** Kevin Hoyt, Solutions Architect  
**Date:** September 21, 2026  
**Scope:** One-way, scheduled refresh of a few SharePoint Online lists containing tens or hundreds of thousands of items

---

## Implementation package files

- `Invoke-SharePointListSync.ps1` - configurable production implementation.
- `New-ListSyncMappingList.ps1` - provisions the "List Sync Configuration" SharePoint list (see 7.2a) that drives which lists get synced.
- `sharepoint-list-sync.sample.json` - safe sample configuration using the SharePoint-driven list-mapping mode; deletion is disabled by default.
- `sharepoint-list-sync.manual-fields.sample.json` - advanced sample configuration using an explicit, hand-written `lists` array (field renames, subsets, etc.).
- `SharePoint-List-Sync-Implementation-Plan.md` - architecture and deployment plan.
- `SharePoint-List-Sync-Implementation-Plan.docx` - formatted implementation plan (last regenerated before the list-mapping feature below; regenerate if the client needs the narrative in Word form).

The script requires PowerShell 7.2 or later and the current PnP.PowerShell module. Run a dry run first with `-WhatIf`; target-only deletion requires both `deleteTargetOnly: true` (or its per-row equivalent in the list-mapping list) in configuration and the explicit `-AllowDelete` switch.

Connections use interactive Entra ID app sign-in or certificate-based app-only auth, matching the pattern already used in this environment by `Provision-MARFoundationLists.ps1` - see `authentication.mode` in the sample configuration (`Interactive`, `Certificate`, or `ManagedIdentity` if hosted with one).

---

## 1. Executive recommendation

Implement the first release as a **PnP PowerShell 7 reconciliation job hosted in Azure Automation**, authenticated with a **managed identity** and granted least-privilege access to only the source and target SharePoint resources.

Each run should read the complete production snapshot in pages and reconcile the test list by an immutable `SourceID`:

1. Add production items missing from test.
2. Update test items whose normalized `SourceHash` differs.
3. Remove test items whose `SourceID` no longer exists in production.
4. Process reference/parent lists before dependent child lists.
5. Remap SharePoint lookup values through a persisted `SourceID -> TargetID` map.
6. Validate counts, source keys, hashes, and relationships before marking the snapshot `Ready`.

This pattern provides the outcome of a full refresh without deleting and recreating every unchanged item. It minimizes writes, keeps target IDs stable across routine runs, supports restart after failure, and does not depend on a long-lived delta token.

Build the same design as a **.NET worker using PnP Core SDK plus Microsoft Graph/SharePoint REST** if the solution grows into a shared synchronization service, requires extensive automated testing and telemetry, or must support more lists and transformations.

Evaluate **Layer2 Cloud Connector** if product ownership is preferred over custom-code ownership. Consider **ShareGate** or **AvePoint Fly** when those products are already licensed or broader migration capabilities are needed. Retain **Power Automate** for orchestration and notifications rather than as the bulk item-processing engine.

### Recommended decision

| Decision | Recommendation |
|---|---|
| Initial implementation | PnP PowerShell 7 in Azure Automation |
| Copy pattern | Complete-source reconciliation, not clear/reload |
| Matching key | Dedicated indexed `SourceID` column |
| Change detection | Deterministic `SourceHash` of report-relevant fields |
| Authentication | System-assigned managed identity; certificate-based Entra application if managed identity is unavailable |
| Permissions | Selected site/list permissions where feasible; no SharePoint Add-in/ACS authentication |
| Scheduling | On demand plus a schedule for the few planned monthly refreshes |
| Report protection | Separate control record with `Loading`, `Validating`, `Ready`, and `Failed` states |
| Escalation path | Rebuild as .NET service if complexity or reuse materially increases |

---

## 2. Why this pattern fits the workload

### 2.1 Options considered

#### Clear and reload

Delete all target items and recreate them from production.

**Benefits**
- Easiest algorithm to understand.
- Guarantees removal of target-only items.
- Useful for the initial load or a disposable test list.

**Risks**
- Produces the maximum number of write operations and therefore the most throttling exposure.
- Leaves the target incomplete if the job fails midway unless a staging/swap design is used.
- Assigns new target item IDs after each refresh.
- Requires complete parent/child lookup remapping on every run.
- Rebuilds item history rather than preserving stable test items.

#### Complete-source reconciliation — recommended

Read all source rows, compare by `SourceID`, write only differences, and delete target-only rows.

**Benefits**
- Self-healing: every run compares against the complete authoritative snapshot.
- Stable target IDs for unchanged rows.
- Lower write volume than clear/reload.
- Straightforward restart and audit behavior.
- No dependency on a retained delta token.

**Tradeoff**
- Requires a target index and deterministic normalization/hash logic.

#### True incremental synchronization

Read only items changed since the prior run, using `Modified`, a change log, or Microsoft Graph delta where supported.

**Benefits**
- Lowest read volume when very few records change.
- Useful if refreshes become frequent or the full scan no longer fits the operating window.

**Risks**
- More stateful and harder to recover if a checkpoint or delta token is lost.
- Deletions, expired delta state, and schema changes need explicit handling.
- An occasional full reconciliation is still advisable.

### 2.2 Trigger for reconsidering incremental processing

Move to a delta/incremental design only when measured runs show one or more of these conditions:

- the complete source scan does not fit the approved refresh window;
- only a small fraction of records changes and write avoidance is operationally significant;
- attachments make repeated source inspection expensive;
- refresh frequency increases substantially;
- the job becomes a reusable service with durable state management.

Do not adopt incremental processing solely because a list exceeds 5,000 items. The 5,000-item list view threshold governs query behavior; it is not the total list capacity. Properly indexed filters and API pagination are the relevant controls.[1][2]

---

## 3. Option comparison

| Option | Large-list fit | Reliability/recovery | Security/governance | Build and support effort | Best use |
|---|---|---|---|---|---|
| **PnP PowerShell** | High with paging, indexed keys, and batches | Strong when checkpoints, idempotency, validation, and structured logs are added | Supports managed identity and certificate-based Entra authentication | Low-to-medium | Best first custom implementation for a few infrequent list refreshes |
| **.NET + PnP Core/Graph/REST** | Highest; greatest control over memory, concurrency, and API mix | Strongest testing, telemetry, retry, and resumability model | Strong application lifecycle and identity controls | Medium-to-high | Shared service, complex mappings, or larger future scope |
| **Power Automate** | Possible but awkward for hundreds of thousands of item operations | Flow history helps, but large loops, retries, quotas, and partial-run recovery add complexity | Familiar governance, but connection ownership must be managed | Low initially; potentially high operationally | Small workloads, orchestration, approval, and notifications |
| **Layer2 Cloud Connector** | Vendor positions it for large one-way data synchronization; prove with representative volume | Scheduling, key matching, logs, and retry are product features | Requires vendor and Windows service review | License plus product administration | Best purpose-built commercial candidate |
| **ShareGate Migrate** | Strong migration tooling; recurring large-list behavior must be tested | Good migration reporting and rerun capabilities | Established migration product controls | License; persistent execution host may be required | Best when already licensed or copying broader SharePoint structure |
| **AvePoint Fly** | Enterprise migration candidate; validate same-tenant recurring use | Dashboard, job monitoring, and rerun capabilities are vendor-described | Enterprise vendor/governance review required | Quote-based licensing and product administration | Best when broader migration governance is needed |

### Ranking for this use case

1. **PnP PowerShell** — best balance of robustness, simplicity, and cost.
2. **.NET** — best engineered platform, but more than the initial scope requires.
3. **Layer2 Cloud Connector** — strongest commercial sync fit if licensing and infrastructure are acceptable.
4. **ShareGate/AvePoint** — reasonable if already owned or if migration scope is broader than list data.
5. **Power Automate** — useful around the job, not preferred for bulk copying.

---

## 4. Target architecture

```text
Production SharePoint lists
        |
        | paged reads (selected fields; indexed filters)
        v
Azure Automation runbook / .NET scheduled worker
        |
        +--> Normalize values and compute SourceHash
        +--> Load target index: SourceID, TargetID, SourceHash
        +--> Build create/update/delete plan
        +--> Resolve people, taxonomy, and lookup mappings
        |
        | bounded batches + Retry-After handling
        v
Test SharePoint lists
        |
        +--> Validation: count, IDs, hashes, required fields, relationships
        +--> Run log and checkpoints
        +--> Refresh control: Loading -> Validating -> Ready / Failed
```

### Components

1. **Source lists:** Production remains authoritative and read-only to the job.
2. **Target lists:** Test lists contain matching business schema plus technical columns.
3. **Synchronization host:** Azure Automation PowerShell 7 runbook initially; .NET worker is the evolution path.
4. **Identity:** Managed identity or certificate-based Entra application.
5. **Configuration store:** JSON configuration in source control; secrets/certificates in approved secure storage.
6. **Checkpoint store:** Azure Table/Storage, Automation variables, or a dedicated SharePoint control list. Do not store the only recovery state inside a list that the job clears.
7. **Telemetry:** Log Analytics/Application Insights or equivalent central logging.
8. **Refresh control:** One row per target snapshot that prevents reports from accepting partially refreshed data.

### Required technical columns in target lists

| Column | Type | Purpose |
|---|---|---|
| `SourceID` | Number, indexed, unique where possible | Production list item ID or other immutable source key |
| `SourceHash` | Single line of text, indexed only if needed | Hash of normalized report-relevant values |
| `SourceModifiedUtc` | Date/time | Diagnostic comparison with production |
| `SyncRunId` | Single line of text | Identifies the run that last wrote the item |
| `SyncUpdatedUtc` | Date/time | Target synchronization timestamp |

If a durable business key already exists, use it instead of the production list item ID. Do not assume target SharePoint-generated IDs can be forced to match production IDs.

---

## 5. Data mapping and fidelity

### 5.1 Field handling

| SharePoint feature | Implementation rule |
|---|---|
| Text, number, currency, Boolean | Copy normalized values; preserve null versus empty where reports distinguish them |
| Date/time | Convert to UTC for comparison and hashing; write in an API-supported ISO representation |
| Choice/multi-choice | Confirm the target schema contains allowed choices; sort multi-values only if order has no business meaning |
| Person/group | Resolve by stable login/UPN; do not copy source user-information-list IDs |
| Lookup | Copy referenced list first; translate source lookup ID through the source-to-target mapping |
| Multi-lookup | Translate every referenced ID and preserve the intended value set |
| Managed metadata | Prefer term GUIDs where both sites use the same term store; otherwise maintain an explicit mapping |
| Hyperlink | Preserve URL and description components |
| Attachments | Copy after the target item exists; compare file name plus size or content hash if fidelity matters |
| Folders | Recreate parent folders before items or explicitly flatten them; make the policy configurable |
| Calculated/read-only fields | Recreate schema/formula; do not attempt ordinary writes to calculated or system-managed values |
| Version history | Default to current value only; use a migration product if historical versions are a requirement |
| Created/Modified/Author/Editor | Treat as system metadata; preserve only if a supported migration path and a report requirement justify it |

Microsoft Graph is suitable for common list-item operations and paging, but SharePoint-specific features can require PnP Core, CSOM, or REST. In particular, list-item attachment content should be handled with PnP or the SharePoint REST attachment endpoints rather than assuming Graph covers the full list-attachment lifecycle.[3][4]

### 5.2 Parent/child relationships

The internal <File>Interactions.xlsx</File> search result shows list relationships recorded through an item ID stored in a `SourceId`-style column, reinforcing that copied lists may contain foreign-key-like values that must be deliberately preserved or translated.{6066}

Processing rule:

1. Copy the parent/reference list.
2. Record `{ParentSourceID, ParentTargetID}` for every row.
3. For each child row, read the source parent ID.
4. Look up the corresponding target parent ID.
5. Write the target lookup field with the target parent ID.
6. Also retain the source relationship ID in a numeric/text field if the report uses it directly.
7. Fail validation if any required parent mapping is missing.

### 5.3 Canonical hash

Create a deterministic representation of the fields that affect report output:

```text
Title=<trimmed string>|Amount=<invariant decimal>|Date=<UTC ISO value>|
Owner=<normalized UPN>|Choices=<sorted escaped values>|ParentSourceID=<integer>
```

Hash the UTF-8 bytes with SHA-256 and store the hexadecimal value in `SourceHash`. Exclude volatile fields such as target `ID`, synchronization timestamps, and target editor metadata.

---

## 6. Large-list mechanics

- SharePoint Online supports lists substantially larger than 5,000 items; the 5,000-item threshold applies to operations/views and query execution, not the list's total capacity.[1][2]
- Index `SourceID`, any incremental watermark such as `Modified`, and fields used in restrictive server-side filters.
- Request only required fields. Avoid retrieving attachments or expanded lookups in the first pass unless required.
- Follow server-provided continuation links or PnP paging. Never manufacture a continuation token.[5]
- Use bounded batches and bounded concurrency. Batch size is a tuning choice, not a service guarantee.
- On HTTP `429` or `503`, honor `Retry-After`; retry only failed operations where the API returns per-operation results.[6][7]
- Graph JSON batching accepts a limited number of requests per batch and each subrequest has its own status. A successful outer batch response does not prove every write succeeded.[7]
- Run a representative-volume performance test. Record items read/written, batches, retries, elapsed time, throttle responses, and peak memory.

Suggested starting values, to be tuned by testing:

| Setting | Initial value |
|---|---:|
| Read page size | 500-1,000 items |
| Write batch | 50-100 item operations for PnP/REST; respect API-specific limits |
| Concurrent write batches | 1 initially; increase cautiously |
| Retry attempts | 6 |
| Retry policy | Honor `Retry-After`; otherwise exponential delay with jitter |
| Checkpoint frequency | After each successfully verified write batch |

These are implementation starting points, not Microsoft service limits.

---

## 7. PnP PowerShell implementation blueprint

### 7.1 Project structure

```text
SharePointListSync/
  config/
    listsync.dev.json
    listsync.prod.json
  src/
    Invoke-ListSync.ps1
    Modules/
      Authentication.psm1
      ReadSource.psm1
      Normalize.psm1
      Reconcile.psm1
      WriteTarget.psm1
      Validate.psm1
      Telemetry.psm1
  tests/
    Unit/
    Integration/
  runbooks/
    Invoke-ListSyncRunbook.ps1
  docs/
    Operations-Runbook.md
```

### 7.2 Configuration example

```json
{
  "sourceSite": "https://tenant.sharepoint.com/sites/Production",
  "targetSite": "https://tenant.sharepoint.com/sites/Test",
  "lists": [
    {
      "sourceList": "Parents",
      "targetList": "Parents-Test",
      "key": "SourceID",
      "fields": ["Title", "Status", "Modified"],
      "dependsOn": []
    },
    {
      "sourceList": "Children",
      "targetList": "Children-Test",
      "key": "SourceID",
      "fields": ["Title", "Parent", "Modified"],
      "dependsOn": ["Parents"],
      "lookups": [{"field": "Parent", "parentList": "Parents"}]
    }
  ]
}
```

### 7.2a Client-managed list mapping

Which production lists get copied to which test lists does not have to live in the JSON
configuration or the script. `Invoke-SharePointListSync.ps1` can instead read that mapping
from a SharePoint list, `List Sync Configuration`, provisioned by `New-ListSyncMappingList.ps1`.
One item = one source/test list pair:

| Column | Purpose |
|---|---|
| `Title` | Friendly sync name (also used as the run's list label in logs) |
| `SourceListName` | Production list to read from |
| `TargetListName` | Test list to write to |
| `Enabled` | Yes/No; disabled rows are skipped |
| `DeleteTargetOnly` | Yes/No; mirrors the per-list delete flag (still gated by `-AllowDelete`) |
| `ProcessOrder` | Optional manual tie-breaker; real dependency order is detected automatically |
| `Notes` | Free text for the client's own documentation |

The client adds or removes rows in SharePoint to change what the job processes - no script or
JSON edit required. Field-by-field mapping is discovered automatically each run: the script
intersects the source and target list's schemas by internal field name (skipping system/
read-only columns and this job's own technical columns) and copies any field present, by that
name, on both sides. A `Lookup`/`MultiLookup` field is only followed automatically when its
parent list is also a configured pair - if it points elsewhere, the run fails with a clear error
naming the field, rather than silently dropping the relationship, and dependency order between
list pairs is derived from those lookups rather than needing to be configured by hand.

This auto-mapping assumes the target list's relevant columns share internal names with the
source list, which holds whenever the test list was created as a copy of (or from the same
content type/template as) the production list - the normal case for this kind of test
refresh. When a list pair needs field renames, a deliberate subset of fields, or any other
per-field control, use an explicit `lists` entry in the JSON configuration instead (see 7.2 and
`sharepoint-list-sync.manual-fields.sample.json`) - an explicit `lists` array always takes
precedence over `listMapping` when both are present in the same configuration file.

### 7.3 Representative flow

```powershell
param([string]$ConfigurationPath, [string]$RunId = [guid]::NewGuid())

$config = Get-Content $ConfigurationPath -Raw | ConvertFrom-Json
$source = Connect-PnPOnline -Url $config.sourceSite -ManagedIdentity -ReturnConnection
$target = Connect-PnPOnline -Url $config.targetSite -ManagedIdentity -ReturnConnection

Set-RefreshStatus -State 'Loading' -RunId $RunId -Connection $target

try {
    foreach ($list in (Sort-ListsByDependency $config.lists)) {
        $checkpoint = Get-SyncCheckpoint -RunId $RunId -List $list.targetList
        $targetIndex = Get-TargetIndex -List $list -Connection $target
        $seenSourceIds = [System.Collections.Generic.HashSet[int]]::new()

        Get-PnPListItem -List $list.sourceList -Fields $list.fields `
            -PageSize 1000 -Connection $source -ScriptBlock {
                param($items)

                $plan = foreach ($item in $items) {
                    $normalized = ConvertTo-NormalizedRecord -Item $item -Map $list
                    $normalized.SourceID = [int]$item.Id
                    $normalized.SourceHash = Get-RecordHash $normalized
                    [void]$seenSourceIds.Add($normalized.SourceID)
                    Compare-TargetRecord -Source $normalized -TargetIndex $targetIndex
                }

                Invoke-ReconciliationBatches -Plan $plan -Target $target `
                    -RunId $RunId -Checkpoint $checkpoint
            }

        Remove-TargetOnlyRecords -TargetIndex $targetIndex `
            -SeenSourceIds $seenSourceIds -Connection $target -RunId $RunId

        Test-ListSnapshot -List $list -Source $source -Target $target -RunId $RunId
    }

    Set-RefreshStatus -State 'Ready' -RunId $RunId -Connection $target
}
catch {
    Write-SyncFailure -RunId $RunId -Exception $_
    Set-RefreshStatus -State 'Failed' -RunId $RunId -Connection $target
    throw
}
```

The production script must include explicit retry wrappers, per-batch result inspection, durable checkpoints, structured logs, and tests. PnP PowerShell documents page-size support for list reads and supports managed identity and certificate-based authentication.[8][9]

### 7.4 Scheduling

Preferred host: **Azure Automation PowerShell 7 runbook**.

- Enable a system-assigned managed identity.
- Grant only the required SharePoint permissions.
- Import and pin a tested PnP.PowerShell module version.
- Create an on-demand entry point and an approved recurring schedule.
- Stream logs to centralized monitoring.
- Alert on job failure, `Failed` refresh status, validation mismatch, and excessive throttling.

---

## 8. .NET implementation blueprint

### 8.1 Technology selection

Use **.NET 8 or later supported LTS**, with:

- **PnP Core SDK** as the primary SharePoint object model;
- **Microsoft Graph SDK** for Graph-native discovery, list operations, paging, batching, or delta scenarios;
- **SharePoint REST/CSOM** only for SharePoint-specific gaps such as attachment operations or unsupported field behavior.

### 8.2 Solution structure

```text
SharePointListSync.sln
  src/
    Sync.Worker/
    Sync.Application/
    Sync.SharePoint/
    Sync.Domain/
  tests/
    Sync.UnitTests/
    Sync.IntegrationTests/
    Sync.ContractTests/
```

Define interfaces so API choices can be changed without rewriting reconciliation logic:

```csharp
public interface IListSourceReader
{
    IAsyncEnumerable<SourcePage> ReadPagesAsync(
        ListDefinition list, Checkpoint checkpoint, CancellationToken ct);
}

public interface IListTargetWriter
{
    Task<BatchResult> ExecuteAsync(
        IReadOnlyCollection<SyncOperation> operations, CancellationToken ct);
}

public interface ISnapshotValidator
{
    Task<ValidationResult> ValidateAsync(
        ListDefinition list, Guid runId, CancellationToken ct);
}
```

### 8.3 Processing structure

```csharp
foreach (var list in dependencySorter.Sort(configuration.Lists))
{
    var checkpoint = await checkpoints.StartOrResumeAsync(runId, list, ct);
    var targetIndex = await target.ReadIndexAsync(list, ct);
    var seen = new HashSet<int>();

    await foreach (var page in source.ReadPagesAsync(list, checkpoint, ct))
    {
        var operations = reconciler.BuildPlan(page, targetIndex, seen);

        foreach (var batch in operations.Chunk(options.BatchSize))
        {
            var result = await retry.ExecuteAsync(
                token => target.ExecuteAsync(batch, token), ct);

            result.ThrowIfAnyOperationFailed();
            await checkpoints.CommitBatchAsync(runId, list, page, result, ct);
        }
    }

    await reconciler.DeleteTargetOnlyAsync(list, targetIndex, seen, ct);
    await validator.RequireSuccessAsync(list, runId, ct);
}
```

### 8.4 .NET-specific controls

- Unit-test normalization, hashes, lookup translation, and reconciliation planning.
- Contract-test every field type against disposable SharePoint lists.
- Use dependency injection for readers, writers, telemetry, and checkpoints.
- Use `HttpClientFactory` and a retry policy that honors `Retry-After`.
- Bound channel/worker concurrency to avoid memory growth and throttling.
- Emit correlation IDs, list IDs, run IDs, page tokens, batch IDs, operation counts, and API request IDs.
- Package as an Azure Function, Container Apps Job, or scheduled worker according to platform standards.

---

## 9. Validation, failure recovery, and report readiness

### 9.1 Validation levels

**Level 1 — mandatory every run**
- Source and target item counts match after applying any documented exclusions.
- Every source `SourceID` exists exactly once in target.
- No unexpected target-only `SourceID` remains.
- No batch operation remains failed.
- Required parent mappings all resolve.

**Level 2 — mandatory for production acceptance**
- Aggregate hash or sorted key/hash comparison matches.
- Required fields are non-null where expected.
- Lookup/person/taxonomy resolution failure count is zero.
- Attachment counts and names match if attachments are in scope.

**Level 3 — periodic assurance**
- Random sample comparison of complete normalized rows.
- Report output comparison against a known production report snapshot.
- Performance trend and throttle-rate review.

### 9.2 Refresh control record

| Field | Example |
|---|---|
| Dataset | `ClaimsReportingTest` |
| RunId | GUID |
| State | `Loading`, `Validating`, `Ready`, `Failed` |
| StartedUtc | Timestamp |
| CompletedUtc | Timestamp |
| SourceCount | Integer |
| TargetCount | Integer |
| Creates/Updates/Deletes | Integers |
| ValidationSummary | Text or JSON |
| ErrorReference | Link/correlation ID |

Reports should use only a snapshot whose control row is `Ready`. If report technology cannot enforce this directly, make readiness part of the report-validation operating procedure.

### 9.3 Recovery rules

- A rerun with the same normalization and key rules must be idempotent.
- Checkpoint only after a batch is confirmed successful.
- If a batch partially fails, retry failed operations only.
- If the process stops, resume from the last committed checkpoint or safely restart the current list.
- Never set `Ready` unless all mandatory validations pass.
- Keep the previous successful target state when feasible. For strict isolation, use blue/green test lists and switch the report connection only after validation; defer this unless partial visibility is a material risk.

---

## 10. Security and governance

1. Use managed identity when hosted in Azure; otherwise use a certificate-backed Entra application.
2. Prefer selected permissions scoped to the source and target resources instead of tenant-wide SharePoint access.[10]
3. Grant production read and test write/delete only.
4. Do not use SharePoint Add-in/ACS authentication. Microsoft retired that model; use Entra ID application authentication.[11]
5. Store configuration in source control and secrets/certificates in approved secure storage.
6. Do not log sensitive field values. Log keys, counts, hashes, and correlation identifiers.
7. Record code/module version, configuration version, identity, and run initiator.
8. Establish separate owners for source schema, synchronization code, and report validation.
9. Include the job in change management because a schema change can silently affect report equivalence.

---

## 11. Third-party evaluation plan

### Shortlist

#### Layer2 Cloud Connector

Evaluate first among commercial tools because it is positioned specifically for one-way data synchronization with key matching, scheduling, mapping, and logging.[12]

Proof-of-concept questions:
- Can it reconcile 100,000+ items within the required window?
- Does it propagate deletions reliably?
- How does it map person, lookup, taxonomy, and multi-value fields?
- Can it copy list-item attachments?
- How is state/metabase recovery performed?
- Can identity be limited to selected SharePoint resources?
- What infrastructure, service account/identity, and licensing are required?

#### ShareGate Migrate

Evaluate if already licensed or if schema, permissions, versions, and broader SharePoint content are also in scope.[13]

Validate recurring same-tenant execution, deletion propagation, target replacement behavior, scheduling host requirements, and performance on representative lists.

#### AvePoint Fly

Evaluate if enterprise migration governance, dashboards, and vendor support are desired.[14]

Validate same-tenant scheduled list refresh, complex fields, attachment behavior, incremental/reconciliation semantics, target deletion, and licensing.

### Product acceptance gate

Do not select a vendor based on capability sheets alone. Require a trial using:

- one simple 100,000-item list;
- one parent/child pair;
- person, lookup, managed metadata, multi-value, and attachment examples;
- creates, updates, and source deletions;
- an interrupted run followed by recovery;
- count/hash validation and exported error logs.

---

## 12. Phased implementation plan

### Phase 0 — decisions and prerequisites

**Decide now**

- Exact source and target sites/lists.
- Which columns affect report results.
- Whether attachments, folders, versions, and system metadata are in scope.
- Whether `SourceID` is sufficient or an immutable business key exists.
- Parent/child dependency order.
- Maximum acceptable refresh window and outage/partial-visibility tolerance.
- Azure Automation availability and identity approval path.
- Required log retention and operational owner.

**Can be deferred**

- Graph delta/incremental processing.
- Parallel list processing.
- Blue/green target lists.
- Vendor product purchase.
- Generic multi-tenant synchronization framework.

### Phase 1 — inventory and schema assessment

- Export source and target list schemas using internal names.
- Classify every field using the fidelity table in Section 5.
- Identify indexed columns and query filters.
- Diagram list dependencies.
- Capture representative values for every complex field.
- Record source counts, attachment counts, and approximate data volume.
- Confirm how the reports join lists and whether they rely on SharePoint `ID`, `SourceID`, lookup IDs, or business keys.

**Deliverables:** schema inventory, mapping matrix, dependency graph, test-data profile.

### Phase 2 — identity and target preparation

- Create/enable the managed identity or certificate-based Entra application.
- Grant production read and test contribute/delete permissions at the narrowest approved scope.
- Create or align target schemas.
- Add and index `SourceID`; add `SourceHash`, `SourceModifiedUtc`, `SyncRunId`, and `SyncUpdatedUtc`.
- Create the refresh-control and run-log stores.
- Create test views filtered by indexed columns.

**Exit criterion:** identity can page source rows and create/update/delete a disposable target item without broader access.

### Phase 3 — proof of concept

Start with one uncomplicated list and then one parent/child pair.

- Implement paging and field projection.
- Implement normalization and SHA-256 hash.
- Implement target index and reconciliation plan.
- Implement bounded batches and `Retry-After` handling.
- Implement parent lookup translation.
- Implement checkpoints and structured logs.
- Implement Level 1 and Level 2 validation.
- Run at representative scale, not only with a small sample.

**Exit criterion:** two consecutive runs are idempotent; the second run performs no business-data writes when source data is unchanged.

### Phase 4 — production hardening

- Add remaining lists and complex fields.
- Add attachment handling if required.
- Add Pester unit tests and disposable-list integration tests.
- Configure monitoring and alerts.
- Document support ownership and escalation.
- Conduct failure injection: throttle, connection loss, malformed field value, missing parent, and interrupted run.
- Complete security review and change record.

### Phase 5 — operational rollout

- Schedule the job outside report validation activity.
- Run the first controlled refresh.
- Review logs and validation results with report owners.
- Execute report comparison.
- Mark the data `Ready` only after acceptance.
- Retain run evidence and record issues for the next iteration.

### Phase 6 — optimization, only if measurements justify it

- Tune batch/concurrency settings.
- Add delta processing with periodic full reconciliation.
- Move to .NET if maintainability or reliability requirements exceed the script design.
- Introduce blue/green target lists if users must never see an in-progress snapshot.

---

## 13. Test plan

| Test | Expected result |
|---|---|
| Initial empty-target load | All in-scope source items appear once in target |
| No-change rerun | Zero creates, updates, and deletes; validation passes |
| Source create | One target create with correct `SourceID` and fields |
| Source update | One target update; unchanged rows are not rewritten |
| Source delete | Corresponding target row is removed |
| Parent/child load | Every child points to the translated target parent |
| Missing parent | Run fails validation and remains not `Ready` |
| Person field | Target resolves the intended account |
| Managed metadata | Target contains intended term GUID/value |
| Multi-value field | Complete value set is preserved |
| Attachment copy | Names and selected validation attributes match |
| Throttling response | Job honors `Retry-After`, retries, and records the event |
| Mid-batch interruption | Restart does not duplicate items and resumes/reconciles safely |
| Schema mismatch | Job stops with a clear field/list error before `Ready` |
| Duplicate `SourceID` in target | Validation fails and identifies duplicate keys |
| Report validation | Test report output matches the approved production comparison |

---

## 14. Operational runbook

### Before a run

- Confirm no target schema changes are pending.
- Confirm the report team knows the planned refresh.
- Verify the most recent run is `Ready` or document why it is not.
- Confirm the identity and module/certificate are valid.
- Record source counts and start a new `RunId`.

### During a run

- Monitor run state, pages processed, write counts, throttles, and failures.
- Do not allow report sign-off while state is `Loading` or `Validating`.
- Do not manually edit synchronized target rows.

### After a successful run

- Review count and hash validation.
- Review unresolved field, lookup, user, taxonomy, and attachment errors.
- Confirm state is `Ready`.
- Notify report validators with run ID and completion timestamp.
- Retain logs according to policy.

### After a failed run

- Keep status `Failed`.
- Capture the failing list, page/batch, request/correlation ID, and exception.
- Correct configuration/data/schema or transient service issue.
- Resume from the committed checkpoint or restart the affected list.
- Run all validation again; never manually force `Ready`.

---

## 15. Monitoring and acceptance criteria

### Metrics

- Run duration by list.
- Source/target counts.
- Pages read and operations planned.
- Creates, updates, deletes, and no-change rows.
- Batch success/failure counts.
- `429`/`503` count and retry delay.
- Missing lookup/person/taxonomy mappings.
- Attachment failures.
- Validation mismatches.
- Last successful `Ready` timestamp.

### Alerts

- Job terminates unsuccessfully.
- State remains `Loading`/`Validating` beyond the approved window.
- Validation fails.
- Duplicate/missing `SourceID` is detected.
- Throttling exceeds the baseline threshold established during testing.
- No successful refresh occurs within the expected business cadence.

### Acceptance criteria

The solution is ready for operational use when:

1. Representative-volume refresh completes within the agreed window.
2. Source and target counts match, subject only to documented exclusions.
3. Every in-scope `SourceID` is unique and present in target.
4. Key/hash validation passes with zero unexplained mismatches.
5. All required parent/child relationships resolve correctly.
6. Two unchanged consecutive runs are idempotent.
7. An interrupted run recovers without duplicates or an invalid `Ready` state.
8. The identity has only the approved source-read and target-write permissions.
9. Logs and alerts identify failures sufficiently for support staff to act.
10. Report owners confirm test output is equivalent to the approved production comparison.

---

## 16. Immediate work backlog

### Priority 1

- [ ] Identify the first source/target list pair and report owner.
- [ ] Inventory field internal names, types, indexes, dependencies, and attachments.
- [ ] Confirm the report's use of `ID`, `SourceID`, lookups, and business keys.
- [ ] Approve Azure Automation and managed identity as the hosting model.
- [ ] Create the target technical columns and refresh-control list.
- [ ] Create the PowerShell repository and configuration skeleton.

### Priority 2

- [ ] Implement paged source read and target index.
- [ ] Implement normalization and `SourceHash`.
- [ ] Implement create/update/delete reconciliation.
- [ ] Implement retry, batch inspection, checkpointing, and logs.
- [ ] Implement count/key/hash validation.
- [ ] Prove idempotency on one large simple list.

### Priority 3

- [ ] Add parent/child lookup remapping.
- [ ] Add complex field and attachment support where required.
- [ ] Add failure-injection and recovery tests.
- [ ] Configure schedule, monitoring, alerts, and operational documentation.
- [ ] Run report-equivalence acceptance testing.

---

## 17. Source notes

The technical recommendations above combine cited platform capabilities with implementation recommendations. Batch sizes, concurrency, checkpoints, hashing, validation levels, and phase gates are proposed design choices and should be confirmed through representative testing.

1. Microsoft Support, *Manage large lists and libraries in SharePoint*: https://support.microsoft.com/office/manage-large-lists-and-libraries-in-sharepoint-0b7d9c8b-b6d4-4041-b9b9-9420c228e6ec
2. Microsoft Learn, *SharePoint limits*: https://learn.microsoft.com/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits
3. Microsoft Learn, *Working with lists and list items with REST*: https://learn.microsoft.com/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
4. Microsoft Learn, *Working with folders and files with REST*: https://learn.microsoft.com/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
5. Microsoft Learn, *Paging Microsoft Graph data in your app*: https://learn.microsoft.com/graph/paging
6. Microsoft Learn, *Avoid getting throttled or blocked in SharePoint Online*: https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
7. Microsoft Learn, *Combine multiple HTTP requests using JSON batching*: https://learn.microsoft.com/graph/json-batching
8. PnP PowerShell, *Get-PnPListItem*: https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html
9. PnP PowerShell, *Connect-PnPOnline*: https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
10. Microsoft Learn, *Selected permissions overview*: https://learn.microsoft.com/graph/permissions-selected-overview
11. Microsoft Learn, *SharePoint Add-In retirement in Microsoft 365*: https://learn.microsoft.com/sharepoint/dev/sp-add-ins/retirement-announcement-for-add-ins
12. Layer2 Solutions, *Layer2 Cloud Connector documentation/product information*: https://www.layer2solutions.com/products/layer2-cloud-connector
13. ShareGate, *Migrate documentation*: https://help.sharegate.com/hc/en-us/categories/360001449011-Migrate
14. AvePoint, *Fly migration platform*: https://www.avepoint.com/products/fly
15. Microsoft Learn, *Power Platform request limits and allocations*: https://learn.microsoft.com/power-platform/admin/api-request-limits-allocations
16. Microsoft Learn, *Power Automate limits and configuration*: https://learn.microsoft.com/power-automate/limits-and-config
17. PnP Core SDK documentation: https://pnp.github.io/pnpcore/
18. Microsoft Learn, *List items - Microsoft Graph*: https://learn.microsoft.com/graph/api/resources/listitem

