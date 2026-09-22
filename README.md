# SharePoint List Sync

PowerShell tooling that keeps one or more SharePoint Online **test** lists in sync with their
**production** counterparts, so reports and testing can run against realistic data without
touching production.

## How it works

Each run reads the complete production list, compares it to the test list by a stable
`SourceID` column, and reconciles the difference:

1. Items missing from the test list are created.
2. Items whose content changed (tracked via a `SourceHash` of the report-relevant fields) are
   updated.
3. Items no longer present in production are deleted from the test list - only when both the
   list's configuration and the `-AllowDelete` switch explicitly allow it.
4. Parent lists are processed before dependent child lists, and SharePoint Lookup fields are
   automatically remapped from production item IDs to the corresponding test item IDs.

This is a full reconciliation, not a clear-and-reload: unchanged items are never rewritten,
and test item IDs stay stable across routine runs. See
[`Script/SharePoint-List-Sync-Implementation-Plan.md`](Script/SharePoint-List-Sync-Implementation-Plan.md)
for the full design rationale and phased rollout plan.

## Which lists get synced

There are two ways to tell the script which production lists map to which test lists:

- **SharePoint-driven (recommended)** - a small SharePoint list, "List Sync Configuration"
  (provisioned by `Script/New-ListSyncMappingList.ps1`), holds one row per list pair
  (`SourceListName`, `TargetListName`, `Enabled`, ...). Field mapping is discovered
  automatically each run by intersecting the production and test list schemas by internal
  field name. Add or remove a row in SharePoint to change what gets synced - no script or
  config file changes needed. See `listMapping` in
  [`Script/sharepoint-list-sync.sample.json`](Script/sharepoint-list-sync.sample.json).
- **Manual JSON** - an explicit `lists` array in the JSON configuration, for cases needing
  field-by-field control (renamed target fields, a subset of source fields, etc.). See
  [`Script/sharepoint-list-sync.manual-fields.sample.json`](Script/sharepoint-list-sync.manual-fields.sample.json).
  An explicit `lists` array always takes precedence over `listMapping` when both are present.

## Repository layout

```
Script/
  Invoke-SharePointListSync.ps1              Main reconciliation script
  New-ListSyncMappingList.ps1                Provisions the "List Sync Configuration" list
  sharepoint-list-sync.sample.json           Sample config (SharePoint-driven mode)
  sharepoint-list-sync.manual-fields.sample.json   Sample config (manual field mapping)
  SharePoint-List-Sync-Implementation-Plan.md/.docx   Full architecture and rollout plan
```

## Requirements

- PowerShell 7.2 or later
- The [PnP.PowerShell](https://pnp.github.io/powershell/) module
- An Entra ID app registration with access to the source (read) and target (read/write) sites

## Quick start

1. **Provision the mapping list** on the site that will hold it (usually the test site):

   ```powershell
   ./Script/New-ListSyncMappingList.ps1 `
       -SiteUrl "https://tenant.sharepoint.com/sites/Test" `
       -ClientId "<entra-app-client-id>"
   ```

2. **Add a row per list pair** you want synced (Title, `SourceListName`, `TargetListName`,
   `Enabled` = Yes) directly in SharePoint.

3. **Copy a sample configuration** and fill in your site URLs and authentication:

   ```powershell
   cp Script/sharepoint-list-sync.sample.json my-sync.json
   ```

4. **Dry run** to see what would happen without writing anything:

   ```powershell
   ./Script/Invoke-SharePointListSync.ps1 -ConfigurationPath ./my-sync.json -WhatIf
   ```

5. **Run for real**, adding `-AllowDelete` only once you're ready for target-only items to be
   removed (and only for lists with `deleteTargetOnly` set):

   ```powershell
   ./Script/Invoke-SharePointListSync.ps1 -ConfigurationPath ./my-sync.json -AllowDelete
   ```

## Authentication

Set `authentication.mode` in the configuration file:

| Mode | Use when |
|---|---|
| `Interactive` | Running manually as yourself; a browser sign-in prompt appears |
| `Certificate` | Unattended/scheduled runs (e.g. Azure Automation); app-only auth via a certificate thumbprint or file |
| `ManagedIdentity` | Hosted somewhere with a managed identity assigned |

## Supported fields

Text, Note, Number, Currency, Boolean, DateTime, Choice, MultiChoice, Url, Person, MultiPerson,
Lookup, and MultiLookup. Taxonomy, attachments, folders, and version history are intentionally
unsupported baseline - add and test a custom transformer before relying on them.
