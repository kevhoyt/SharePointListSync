#requires -Version 7.2
#requires -Modules PnP.PowerShell
<#!
.SYNOPSIS
  Reconciles one or more SharePoint Online production lists to test lists.
.DESCRIPTION
  Source is authoritative. Items are matched by a dedicated target SourceID column.
  Missing items are created, changed items are updated using a deterministic SHA-256
  SourceHash, and target-only items can be deleted only when both configuration and
  the -AllowDelete switch enable deletion.

  Authentication supports interactive Entra ID app sign-in (matches
  Provision-MARFoundationLists.ps1 - browser sign-in as your account) and
  certificate-based app-only Entra application auth for unattended runs (e.g. Azure
  Automation). Managed identity is also supported when hosted with one. It is
  idempotent: rerunning converges the target to the source.

  Which lists to sync can come from two places:
   - a "listMapping" section in the JSON configuration, which points at a
     SharePoint list (see New-ListSyncMappingList.ps1) holding one row per
     source/target list pair. Field mapping is discovered automatically by
     intersecting the source and target list schemas by internal field name, so
     the client can add or remove synced lists from SharePoint without editing
     the script or the JSON file.
   - an explicit "lists" array in the JSON configuration, for cases needing
     manual field-by-field control (renamed target fields, a subset of source
     fields, etc.). If "lists" is present and non-empty it takes precedence over
     "listMapping".

  Supported baseline field kinds: Text, Note, Number, Currency, Boolean, DateTime,
  Choice, MultiChoice, Url, Person, MultiPerson, Lookup, MultiLookup.
  Taxonomy, folders, attachments, version history, and system-field preservation are
  intentionally blocked unless a tenant-tested custom transformer is added.

.PARAMETER ConfigurationPath
  Path to the JSON configuration file.
.PARAMETER AllowDelete
  Required in addition to deleteTargetOnly=true before target-only items are deleted.
.PARAMETER Resume
  Resumes reconciliation after the last committed source item ID in the checkpoint.
  The delete and validation passes always rescan source keys.
.PARAMETER WhatIf
  Produces a plan without writing list items, status rows, or checkpoints.
.EXAMPLE
  ./Invoke-SharePointListSync.ps1 -ConfigurationPath ./sharepoint-list-sync.sample.json -WhatIf
.EXAMPLE
  ./Invoke-SharePointListSync.ps1 -ConfigurationPath ./listsync.prod.json -AllowDelete
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ConfigurationPath,

    [switch]$AllowDelete,
    [switch]$Resume
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:RunId = [guid]::NewGuid().ToString('D')
$script:LogPath = $null
$script:CheckpointPath = $null
$script:Configuration = $null
$script:SourceConnection = $null
$script:TargetConnection = $null
$script:IdMaps = @{}
$script:RunSummary = [ordered]@{
    RunId = $script:RunId; StartedUtc = [DateTime]::UtcNow.ToString('o')
    CompletedUtc = $null; State = 'Starting'; Lists = @(); Errors = 0
}

function Write-SyncLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Debug','Information','Warning','Error')][string]$Level,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Data = @{}
    )
    $entry = [ordered]@{
        timestampUtc = [DateTime]::UtcNow.ToString('o')
        level = $Level
        runId = $script:RunId
        message = $Message
        data = $Data
    }
    $json = $entry | ConvertTo-Json -Compress -Depth 12
    if ($script:LogPath) { Add-Content -LiteralPath $script:LogPath -Value $json -Encoding utf8 }
    switch ($Level) {
        'Debug'       { Write-Verbose $Message }
        'Information' { Write-Host $Message }
        'Warning'     { Write-Warning $Message }
        'Error'       { Write-Error $Message -ErrorAction Continue }
    }
}

function Assert-Configuration {
    param([Parameter(Mandatory)]$Config)
    foreach ($name in @('sourceSiteUrl','targetSiteUrl','authentication')) {
        if (-not $Config.PSObject.Properties.Name.Contains($name)) {
            throw "Configuration is missing required property '$name'."
        }
    }
    $hasManualLists = $Config.PSObject.Properties.Name.Contains('lists') -and @($Config.lists).Count -gt 0
    $hasListMapping = $Config.PSObject.Properties.Name.Contains('listMapping') -and
        $Config.listMapping.PSObject.Properties.Name.Contains('enabled') -and [bool]$Config.listMapping.enabled
    if (-not $hasManualLists -and -not $hasListMapping) {
        throw 'Configuration must define a non-empty "lists" array or an enabled "listMapping" section.'
    }
    if ($hasListMapping -and (-not $Config.listMapping.PSObject.Properties.Name.Contains('listName') -or -not $Config.listMapping.listName)) {
        throw "'listMapping' configuration is missing required property 'listName'."
    }
    if (-not $hasManualLists) { return }

    $names = @{}
    foreach ($list in $Config.lists) {
        foreach ($name in @('name','sourceList','targetList','fields')) {
            if (-not $list.PSObject.Properties.Name.Contains($name)) {
                throw "List configuration is missing required property '$name'."
            }
        }
        if ($names.ContainsKey($list.name)) { throw "Duplicate list configuration name '$($list.name)'." }
        $names[$list.name] = $true
        if (-not $list.PSObject.Properties.Name.Contains('sourceIdField')) {
            $list | Add-Member -NotePropertyName sourceIdField -NotePropertyValue 'SourceID'
        }
        if (-not $list.PSObject.Properties.Name.Contains('sourceHashField')) {
            $list | Add-Member -NotePropertyName sourceHashField -NotePropertyValue 'SourceHash'
        }
        if (-not $list.PSObject.Properties.Name.Contains('deleteTargetOnly')) {
            $list | Add-Member -NotePropertyName deleteTargetOnly -NotePropertyValue $false
        }
        if (-not $list.PSObject.Properties.Name.Contains('dependsOn')) {
            $list | Add-Member -NotePropertyName dependsOn -NotePropertyValue @()
        }
        if (-not $list.PSObject.Properties.Name.Contains('lookups')) {
            $list | Add-Member -NotePropertyName lookups -NotePropertyValue @()
        }
        foreach ($field in $list.fields) {
            foreach ($required in @('source','target','type')) {
                if (-not $field.PSObject.Properties.Name.Contains($required)) {
                    throw "Field in list '$($list.name)' is missing '$required'."
                }
            }
            if ($field.type -in @('Taxonomy','MultiTaxonomy','Attachment','Folder','System','VersionHistory')) {
                throw "Field '$($field.source)' in '$($list.name)' uses unsupported baseline type '$($field.type)'. Add and test a custom transformer before production use."
            }
        }
    }
    foreach ($list in $Config.lists) {
        foreach ($dependency in @($list.dependsOn)) {
            if (-not $names.ContainsKey([string]$dependency)) {
                throw "List '$($list.name)' depends on unknown list '$dependency'."
            }
        }
    }
}

function Assert-ListPairsDistinct {
    param([Parameter(Mandatory)][array]$Lists)
    # Source and target sites are allowed to be the same site (e.g. a same-site trial
    # with two differently-named lists) - what must never happen is a list reconciling
    # against itself, which would self-overwrite using its own technical columns.
    $sameSite = [string]$script:Configuration.sourceSiteUrl -eq [string]$script:Configuration.targetSiteUrl
    if (-not $sameSite) { return }
    foreach ($list in $Lists) {
        if ([string]::Equals([string]$list.sourceList, [string]$list.targetList, [StringComparison]::OrdinalIgnoreCase)) {
            throw "List '$($list.name)' has the same source and target list ('$($list.sourceList)') on the same site. Source and target must be different lists when sourceSiteUrl equals targetSiteUrl."
        }
    }
}

function Get-DependencyOrderedLists {
    param([Parameter(Mandatory)][array]$Lists)
    $remaining = [System.Collections.Generic.List[object]]::new()
    foreach ($list in $Lists) { $remaining.Add($list) }
    $ordered = [System.Collections.Generic.List[object]]::new()
    $done = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    while ($remaining.Count -gt 0) {
        $progress = $false
        foreach ($list in @($remaining)) {
            $ready = $true
            foreach ($dependency in @($list.dependsOn)) {
                if (-not $done.Contains([string]$dependency)) { $ready = $false; break }
            }
            if ($ready) {
                $ordered.Add($list); [void]$done.Add([string]$list.name)
                [void]$remaining.Remove($list); $progress = $true
            }
        }
        if (-not $progress) { throw 'List dependency graph contains a cycle.' }
    }
    return $ordered.ToArray()
}

function Connect-SyncSite {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)]$Authentication)
    switch ([string]$Authentication.mode) {
        'ManagedIdentity' {
            if ($Authentication.userAssignedManagedIdentityClientId) {
                return Connect-PnPOnline -Url $Url -ManagedIdentity `
                    -UserAssignedManagedIdentityClientId $Authentication.userAssignedManagedIdentityClientId `
                    -ReturnConnection
            }
            return Connect-PnPOnline -Url $Url -ManagedIdentity -ReturnConnection
        }
        'Certificate' {
            foreach ($required in @('clientId','tenant')) {
                if (-not $Authentication.$required) { throw "Certificate authentication requires '$required'." }
            }
            $params = @{
                Url = $Url; ClientId = [string]$Authentication.clientId
                Tenant = [string]$Authentication.tenant
                ReturnConnection = $true
            }
            if ($Authentication.certificateThumbprint) {
                # Matches Provision-MARFoundationLists.ps1: app-only auth via a certificate
                # already installed on this machine, referenced by thumbprint.
                $params.Thumbprint = [string]$Authentication.certificateThumbprint
            }
            elseif ($Authentication.certificatePath) {
                $params.CertificatePath = [string]$Authentication.certificatePath
                if ($Authentication.certificatePasswordEnvironmentVariable) {
                    $plain = [Environment]::GetEnvironmentVariable(
                        [string]$Authentication.certificatePasswordEnvironmentVariable)
                    if ([string]::IsNullOrWhiteSpace($plain)) {
                        throw 'Certificate password environment variable is empty.'
                    }
                    $params.CertificatePassword = ConvertTo-SecureString $plain -AsPlainText -Force
                }
            }
            else { throw "Certificate authentication requires either 'certificateThumbprint' or 'certificatePath'." }
            return Connect-PnPOnline @params
        }
        'Interactive' {
            # Matches Provision-MARFoundationLists.ps1: browser sign-in as your own account
            # against a registered Entra ID app. Not suitable for unattended/scheduled runs.
            if (-not $Authentication.clientId) { throw "Interactive authentication requires 'clientId'." }
            return Connect-PnPOnline -Url $Url -Interactive -ClientId ([string]$Authentication.clientId) -ReturnConnection
        }
        default { throw "Unsupported authentication mode '$($Authentication.mode)'." }
    }
}

function ConvertTo-CanonicalScalar {
    param($Value, [string]$Type)
    if ($null -eq $Value) { return $null }
    switch ($Type) {
        { $_ -in @('Text','Note','Choice') } { return ([string]$Value).Trim() }
        { $_ -in @('Number','Currency') } {
            return ([Convert]::ToDecimal($Value, [Globalization.CultureInfo]::InvariantCulture)).ToString(
                'G29', [Globalization.CultureInfo]::InvariantCulture)
        }
        'Boolean' { return [bool]$Value }
        'DateTime' { return ([DateTime]$Value).ToUniversalTime().ToString('o') }
        'Url' {
            if ($Value.PSObject.Properties.Name -contains 'Url') {
                return [ordered]@{ url = [string]$Value.Url; description = [string]$Value.Description }
            }
            return [string]$Value
        }
        'Person' {
            if ($Value.PSObject.Properties.Name -contains 'Email' -and $Value.Email) {
                return ([string]$Value.Email).Trim().ToLowerInvariant()
            }
            if ($Value.PSObject.Properties.Name -contains 'LookupValue') {
                return ([string]$Value.LookupValue).Trim().ToLowerInvariant()
            }
            return ([string]$Value).Trim().ToLowerInvariant()
        }
        default { return $Value }
    }
}

function ConvertTo-CanonicalValue {
    param($Value, [string]$Type)
    if ($null -eq $Value) { return $null }
    if ($Type -in @('MultiChoice','MultiPerson')) {
        $scalarType = if ($Type -eq 'MultiChoice') { 'Choice' } else { 'Person' }
        return @($Value | ForEach-Object { ConvertTo-CanonicalScalar $_ $scalarType } | Sort-Object)
    }
    if ($Type -in @('Lookup','MultiLookup')) {
        $values = @($Value)
        $ids = foreach ($v in $values) {
            if ($v.PSObject.Properties.Name -contains 'LookupId') { [int]$v.LookupId }
            else { [int]$v }
        }
        if ($Type -eq 'Lookup') { return ($ids | Select-Object -First 1) }
        return @($ids | Sort-Object)
    }
    return ConvertTo-CanonicalScalar $Value $Type
}

function Get-RecordHash {
    param([Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$CanonicalRecord)
    $json = $CanonicalRecord | ConvertTo-Json -Compress -Depth 20
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $hash = [Security.Cryptography.SHA256]::HashData($bytes)
    return [Convert]::ToHexString($hash).ToLowerInvariant()
}

function Resolve-PersonWriteValue {
    param($Value, [bool]$Multiple)
    $values = @($Value)
    $resolved = foreach ($v in $values) {
        if ($null -eq $v) { continue }
        if ($v.PSObject.Properties.Name -contains 'Email' -and $v.Email) { [string]$v.Email }
        elseif ($v.PSObject.Properties.Name -contains 'LookupValue' -and $v.LookupValue) { [string]$v.LookupValue }
        else { [string]$v }
    }
    if ($Multiple) { return @($resolved) }
    return ($resolved | Select-Object -First 1)
}

function Convert-SourceItem {
    param([Parameter(Mandatory)]$Item, [Parameter(Mandatory)]$ListConfig)
    $writeValues = @{}
    $canonical = [ordered]@{}

    foreach ($field in $ListConfig.fields) {
        $sourceValue = $Item.FieldValues[[string]$field.source]
        $canonicalValue = ConvertTo-CanonicalValue $sourceValue ([string]$field.type)
        $canonical[[string]$field.target] = $canonicalValue

        switch ([string]$field.type) {
            { $_ -in @('Text','Note','Choice','Boolean') } { $writeValues[[string]$field.target] = $sourceValue }
            { $_ -in @('Number','Currency') } {
                # [decimal] is used for hashing (ConvertTo-CanonicalScalar) but must NOT be used
                # here: Add-PnPListItem -Batch serializes values through a REST path that throws
                # "Cannot convert a primitive value to the expected type 'Edm.String'" on
                # [decimal] specifically, even though the identical value as [double] succeeds.
                # Confirmed 2026-09-22 against PnP.PowerShell 3.3.0 / SharePoint Online.
                $writeValues[[string]$field.target] = if ($null -eq $sourceValue) { $null } else { [double]$sourceValue }
            }
            'DateTime' {
                $writeValues[[string]$field.target] = if ($null -eq $sourceValue) { $null } else { ([DateTime]$sourceValue).ToUniversalTime() }
            }
            'MultiChoice' { $writeValues[[string]$field.target] = @($sourceValue) }
            'Url' { $writeValues[[string]$field.target] = $sourceValue }
            'Person' { $writeValues[[string]$field.target] = Resolve-PersonWriteValue $sourceValue $false }
            'MultiPerson' { $writeValues[[string]$field.target] = Resolve-PersonWriteValue $sourceValue $true }
            { $_ -in @('Lookup','MultiLookup') } {
                $lookupConfig = @($ListConfig.lookups | Where-Object { $_.field -eq $field.source })
                if ($lookupConfig.Count -ne 1) {
                    throw "Lookup field '$($field.source)' in '$($ListConfig.name)' requires exactly one lookups configuration entry."
                }
                $parentName = [string]$lookupConfig[0].parentList
                if (-not $script:IdMaps.ContainsKey($parentName)) {
                    throw "No ID map is available for parent list '$parentName'."
                }
                $sourceIds = @($canonicalValue)
                $targetIds = foreach ($sourceId in $sourceIds) {
                    if (-not $script:IdMaps[$parentName].ContainsKey([int]$sourceId)) {
                        if ($WhatIfPreference) {
                            # A dry run against an empty/partial target never actually writes the
                            # parent list, so its real target IDs cannot be known yet. Substitute a
                            # placeholder so WhatIf can still preview dependent lists instead of
                            # aborting the whole run.
                            continue
                        }
                        throw "Missing target lookup mapping for '$parentName' source ID '$sourceId'."
                    }
                    [int]$script:IdMaps[$parentName][[int]$sourceId]
                }
                $writeValues[[string]$field.target] = if ($field.type -eq 'Lookup') {
                    $targetIds | Select-Object -First 1
                } else { @($targetIds) }
            }
            default { throw "Unsupported field type '$($field.type)'." }
        }
    }

    $sourceId = [int]$Item.Id
    $hash = Get-RecordHash $canonical
    $writeValues[[string]$ListConfig.sourceIdField] = $sourceId
    $writeValues[[string]$ListConfig.sourceHashField] = $hash
    if ($ListConfig.PSObject.Properties.Name -contains 'sourceModifiedField' -and $ListConfig.sourceModifiedField) {
        $writeValues[[string]$ListConfig.sourceModifiedField] = ([DateTime]$Item.FieldValues.Modified).ToUniversalTime()
    }
    if ($ListConfig.PSObject.Properties.Name -contains 'syncRunIdField' -and $ListConfig.syncRunIdField) {
        $writeValues[[string]$ListConfig.syncRunIdField] = $script:RunId
    }
    if ($ListConfig.PSObject.Properties.Name -contains 'syncUpdatedField' -and $ListConfig.syncUpdatedField) {
        $writeValues[[string]$ListConfig.syncUpdatedField] = [DateTime]::UtcNow
    }
    [pscustomobject]@{ SourceID = $sourceId; SourceHash = $hash; Values = $writeValues }
}

function Get-TargetIndex {
    param([Parameter(Mandatory)]$ListConfig)
    $fields = @('ID', [string]$ListConfig.sourceIdField, [string]$ListConfig.sourceHashField)
    $index = @{}
    $duplicates = [System.Collections.Generic.List[int]]::new()
    Get-PnPListItem -List $ListConfig.targetList -Fields $fields `
        -PageSize $script:Configuration.readPageSize -Connection $script:TargetConnection |
        ForEach-Object {
            $raw = $_.FieldValues[[string]$ListConfig.sourceIdField]
            if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) { return }
            $sourceId = [int]$raw
            if ($index.ContainsKey($sourceId)) { $duplicates.Add($sourceId); return }
            $index[$sourceId] = [pscustomobject]@{
                TargetID = [int]$_.Id
                SourceHash = [string]$_.FieldValues[[string]$ListConfig.sourceHashField]
            }
        }
    if ($duplicates.Count -gt 0) {
        throw "Target list '$($ListConfig.targetList)' contains duplicate SourceID values: $($duplicates -join ', ')."
    }
    return $index
}

# Fields never eligible for auto field-mapping: SharePoint system/read-only columns
# plus this script's own technical columns. Title is deliberately NOT excluded - it is
# an ordinary Text field on almost every list and target items would otherwise be
# created blank-titled.
$script:AutoMappingExcludedFields = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'ID','Attachments','ContentType','ContentTypeId','Author','Editor',
        'Created','Modified','Owshiddenversion','Version','_UIVersionString','GUID',
        'FileRef','FileLeafRef','FSObjType','FileSystemObjectType','Order','UniqueId',
        'WorkflowVersion','ProgId','ScopeId','MetaInfo','InstanceID','AppAuthor','AppEditor',
        'SourceID','SourceHash','SourceModifiedUtc','SyncRunId','SyncUpdatedUtc'
    ), [StringComparer]::OrdinalIgnoreCase)

function ConvertTo-BaselineFieldType {
    param([Parameter(Mandatory)][string]$TypeAsString)
    switch ($TypeAsString) {
        'Text'        { return 'Text' }
        'Note'        { return 'Note' }
        'Number'      { return 'Number' }
        'Currency'    { return 'Currency' }
        'Boolean'     { return 'Boolean' }
        'DateTime'    { return 'DateTime' }
        'Choice'      { return 'Choice' }
        'MultiChoice' { return 'MultiChoice' }
        'URL'         { return 'Url' }
        'User'        { return 'Person' }
        'UserMulti'   { return 'MultiPerson' }
        'Lookup'      { return 'Lookup' }
        'LookupMulti' { return 'MultiLookup' }
        default       { return $null }
    }
}

function Get-ListPairsFromSharePoint {
    $mapping = $script:Configuration.listMapping
    $mappingConnection = $script:TargetConnection
    if ($mapping.PSObject.Properties.Name -contains 'siteUrl' -and $mapping.siteUrl -and
        [string]$mapping.siteUrl -ne [string]$script:Configuration.targetSiteUrl) {
        $mappingConnection = Connect-SyncSite ([string]$mapping.siteUrl) $script:Configuration.authentication
    }

    $mappingPropertyNames = $mapping.PSObject.Properties.Name
    function Get-MappingFieldName {
        param([string]$PropertyName, [string]$Default)
        if ($mappingPropertyNames -contains $PropertyName -and $mapping.$PropertyName) { return [string]$mapping.$PropertyName }
        return $Default
    }
    $field = @{
        SourceListName = Get-MappingFieldName 'sourceListNameField' 'SourceListName'
        TargetListName = Get-MappingFieldName 'targetListNameField' 'TargetListName'
        Enabled = Get-MappingFieldName 'enabledField' 'Enabled'
        DeleteTargetOnly = Get-MappingFieldName 'deleteTargetOnlyField' 'DeleteTargetOnly'
        ProcessOrder = Get-MappingFieldName 'processOrderField' 'ProcessOrder'
    }

    $rows = Get-PnPListItem -List ([string]$mapping.listName) `
        -Fields @('Title', $field.SourceListName, $field.TargetListName, $field.Enabled, $field.DeleteTargetOnly, $field.ProcessOrder) `
        -PageSize 500 -Connection $mappingConnection

    $pairs = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $rows) {
        $enabledValue = $row.FieldValues[$field.Enabled]
        $isEnabled = if ($null -eq $enabledValue) { $true } else { [bool]$enabledValue }
        if (-not $isEnabled) { continue }

        $sourceListName = [string]$row.FieldValues[$field.SourceListName]
        $targetListName = [string]$row.FieldValues[$field.TargetListName]
        if ([string]::IsNullOrWhiteSpace($sourceListName) -or [string]::IsNullOrWhiteSpace($targetListName)) {
            throw "List-mapping item $($row.Id) in '$($mapping.listName)' is missing '$($field.SourceListName)' or '$($field.TargetListName)'."
        }

        $pairs.Add([pscustomobject]@{
            Name = if ($row.FieldValues['Title']) { [string]$row.FieldValues['Title'] } else { $sourceListName }
            SourceListName = $sourceListName
            TargetListName = $targetListName
            DeleteTargetOnly = if ($null -eq $row.FieldValues[$field.DeleteTargetOnly]) { $false } else { [bool]$row.FieldValues[$field.DeleteTargetOnly] }
            ProcessOrder = if ($null -eq $row.FieldValues[$field.ProcessOrder]) { 0 } else { [int]$row.FieldValues[$field.ProcessOrder] }
        })
    }
    if ($pairs.Count -eq 0) { throw "No enabled rows found in list-mapping list '$($mapping.listName)'." }

    $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($pair in $pairs) {
        if (-not $names.Add($pair.Name)) { throw "Duplicate list-mapping name '$($pair.Name)' in '$($mapping.listName)'." }
    }
    return @($pairs | Sort-Object ProcessOrder, Name)
}

function Get-AutoListConfigurations {
    $pairs = Get-ListPairsFromSharePoint
    $sourceListMeta = @{}
    foreach ($pair in $pairs) {
        $sourceListMeta[$pair.Name] = Get-PnPList -Identity $pair.SourceListName -Connection $script:SourceConnection
    }

    $configs = foreach ($pair in $pairs) {
        $sourceFields = Get-PnPField -List $pair.SourceListName -Connection $script:SourceConnection
        $targetFields = Get-PnPField -List $pair.TargetListName -Connection $script:TargetConnection
        $targetByName = @{}
        foreach ($f in $targetFields) { $targetByName[[string]$f.InternalName] = $f }

        $fields = [System.Collections.Generic.List[object]]::new()
        $lookups = [System.Collections.Generic.List[object]]::new()
        $dependsOn = [System.Collections.Generic.List[string]]::new()

        foreach ($sourceField in $sourceFields) {
            $internalName = [string]$sourceField.InternalName
            if ($script:AutoMappingExcludedFields.Contains($internalName)) { continue }
            if ($sourceField.Hidden -or $sourceField.ReadOnlyField) { continue }
            if (-not $targetByName.ContainsKey($internalName)) { continue }

            $mappedType = ConvertTo-BaselineFieldType ([string]$sourceField.TypeAsString)
            if (-not $mappedType) {
                Write-SyncLog Warning "Auto field-mapping skipped unsupported field '$internalName' in '$($pair.SourceListName)' ($($pair.Name))." `
                    @{ type = [string]$sourceField.TypeAsString }
                continue
            }
            $fields.Add([pscustomobject]@{ source = $internalName; target = $internalName; type = $mappedType })

            if ($mappedType -in @('Lookup','MultiLookup')) {
                $lookupListId = ([string]$sourceField.LookupList).Trim('{','}')
                $parentName = $null
                foreach ($candidate in $pairs) {
                    if ([string]$sourceListMeta[$candidate.Name].Id -eq $lookupListId) { $parentName = $candidate.Name; break }
                }
                if (-not $parentName) {
                    throw "Field '$internalName' in '$($pair.SourceListName)' ($($pair.Name)) looks up a list that is not configured as a sync pair in '$([string]$script:Configuration.listMapping.listName)'. Add that list's row or remove the field from the target schema."
                }
                $lookups.Add([pscustomobject]@{ field = $internalName; parentList = $parentName })
                if ($dependsOn -notcontains $parentName -and $parentName -ne $pair.Name) { [void]$dependsOn.Add($parentName) }
            }
        }

        [pscustomobject]@{
            name = $pair.Name
            sourceList = $pair.SourceListName
            targetList = $pair.TargetListName
            sourceIdField = 'SourceID'
            sourceHashField = 'SourceHash'
            sourceModifiedField = 'SourceModifiedUtc'
            syncRunIdField = 'SyncRunId'
            syncUpdatedField = 'SyncUpdatedUtc'
            deleteTargetOnly = $pair.DeleteTargetOnly
            dependsOn = @($dependsOn)
            lookups = @($lookups)
            fields = @($fields)
        }
    }
    return @($configs)
}

function Get-SourceQuery {
    param([int]$LastProcessedSourceID)
    if ($LastProcessedSourceID -le 0) { return $null }
    return @"
<View Scope='RecursiveAll'>
  <Query>
    <Where><Gt><FieldRef Name='ID'/><Value Type='Counter'>$LastProcessedSourceID</Value></Gt></Where>
    <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
  </Query>
</View>
"@
}

function Get-Checkpoint {
    param([string]$ListName)
    if (-not (Test-Path -LiteralPath $script:CheckpointPath)) { return $null }
    $all = Get-Content -LiteralPath $script:CheckpointPath -Raw | ConvertFrom-Json -AsHashtable
    if ($all.ContainsKey($ListName)) { return $all[$ListName] }
    return $null
}

function Save-Checkpoint {
    param([string]$ListName, [hashtable]$State)
    if ($WhatIfPreference) { return }
    $all = @{}
    if (Test-Path -LiteralPath $script:CheckpointPath) {
        $all = Get-Content -LiteralPath $script:CheckpointPath -Raw | ConvertFrom-Json -AsHashtable
    }
    $State.runId = $script:RunId
    $State.updatedUtc = [DateTime]::UtcNow.ToString('o')
    $all[$ListName] = $State
    $temp = "$($script:CheckpointPath).tmp"
    $all | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $temp -Encoding utf8
    Move-Item -LiteralPath $temp -Destination $script:CheckpointPath -Force
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [Parameter(Mandatory)][string]$Description
    )
    $max = [int]$script:Configuration.retry.maxAttempts
    $base = [int]$script:Configuration.retry.baseDelaySeconds
    for ($attempt = 1; $attempt -le $max; $attempt++) {
        try { return & $Operation }
        catch {
            $message = $_.Exception.Message
            $retryable = $message -match '(429|503|throttl|temporar|server busy|timeout)'
            if (-not $retryable -or $attempt -eq $max) { throw }
            $retryAfter = $null
            if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
                try { $retryAfter = [int]$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch { }
                if (-not $retryAfter) {
                    try { $retryAfter = [int]$_.Exception.Response.Headers.GetValues('Retry-After')[0] } catch { }
                }
            }
            $delay = if ($retryAfter -and $retryAfter -gt 0) { $retryAfter } else {
                [Math]::Min(300, ($base * [Math]::Pow(2, $attempt - 1)) + (Get-Random -Minimum 0 -Maximum 4))
            }
            Write-SyncLog -Level Warning -Message "$Description throttled/transient failure; retrying." `
                -Data @{ attempt = $attempt; delaySeconds = $delay; error = $message }
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-WriteBatch {
    param(
        [Parameter(Mandatory)]$ListConfig,
        # AllowEmptyCollection is required: PowerShell's implicit "Mandatory parameters
        # reject an empty collection" binder check runs before the function body, so without
        # it, every no-change reconciliation (the normal steady-state rerun) would fail here
        # rather than reaching the Count -eq 0 check below. Confirmed 2026-09-22.
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Operations,
        [Parameter(Mandatory)][hashtable]$Counters
    )
    if ($Operations.Count -eq 0) { return }
    if ($WhatIfPreference) {
        foreach ($op in $Operations) { Write-SyncLog Information "WhatIf: $($op.Type) SourceID $($op.SourceID)" @{ list=$ListConfig.name } }
        $Operations.Clear(); return
    }

    $batch = New-PnPBatch -Connection $script:TargetConnection
    foreach ($op in $Operations) {
        if ($op.Type -eq 'Create') {
            Add-PnPListItem -List $ListConfig.targetList -Values $op.Values `
                -Batch $batch -Connection $script:TargetConnection | Out-Null
        }
        elseif ($op.Type -eq 'Update') {
            Set-PnPListItem -List $ListConfig.targetList -Identity $op.TargetID -Values $op.Values `
                -Batch $batch -Connection $script:TargetConnection | Out-Null
        }
        elseif ($op.Type -eq 'Delete') {
            Remove-PnPListItem -List $ListConfig.targetList -Identity $op.TargetID -Force `
                -Batch $batch -Connection $script:TargetConnection | Out-Null
        }
    }
    $batchResult = Invoke-WithRetry -Description "Batch for $($ListConfig.name)" -Operation {
        Invoke-PnPBatch -Batch $batch -Connection $script:TargetConnection -Details
    }
    # A successful outer Invoke-PnPBatch call does not prove every operation inside it
    # succeeded - each result entry can carry its own ErrorMessage (e.g. a value that fails
    # server-side validation). Surface those loudly instead of reporting false success.
    $failures = @($batchResult | Where-Object {
        $_ -and ($_.PSObject.Properties.Name -contains 'ErrorMessage') -and $_.ErrorMessage
    })
    if ($failures.Count -gt 0) {
        $summary = ($failures | ForEach-Object { $_.ErrorMessage } | Select-Object -Unique) -join '; '
        throw "Batch write for '$($ListConfig.name)' had $($failures.Count) failed operation(s) of $($Operations.Count): $summary"
    }
    $Operations.Clear()
}

function Set-RefreshStatus {
    param([ValidateSet('Loading','Validating','Ready','Failed')][string]$State, [hashtable]$Summary = @{})
    $control = $script:Configuration.control
    if ($WhatIfPreference -or -not $control.enabled) { return }
    $fields = @([string]$control.datasetField, [string]$control.stateField, [string]$control.runIdField)
    $item = Get-PnPListItem -List $control.list -Fields $fields -PageSize 100 `
        -Connection $script:TargetConnection | Where-Object {
            $_.FieldValues[[string]$control.datasetField] -eq [string]$control.dataset
        } | Select-Object -First 1
    $values = @{
        ([string]$control.datasetField) = [string]$control.dataset
        ([string]$control.stateField) = $State
        ([string]$control.runIdField) = $script:RunId
    }
    if ($control.startedUtcField) { $values[[string]$control.startedUtcField] = [DateTime]$script:RunSummary.StartedUtc }
    if ($control.completedUtcField -and $State -in @('Ready','Failed')) { $values[[string]$control.completedUtcField] = [DateTime]::UtcNow }
    if ($control.summaryField) { $values[[string]$control.summaryField] = ($Summary | ConvertTo-Json -Compress -Depth 10) }
    if ($item) {
        Set-PnPListItem -List $control.list -Identity $item.Id -Values $values -Connection $script:TargetConnection | Out-Null
    } else {
        Add-PnPListItem -List $control.list -Values $values -Connection $script:TargetConnection | Out-Null
    }
}

function Invoke-ListReconciliation {
    param([Parameter(Mandatory)]$ListConfig)
    Write-SyncLog Information "Starting list '$($ListConfig.name)'." @{ source=$ListConfig.sourceList; target=$ListConfig.targetList }
    $targetIndex = Get-TargetIndex $ListConfig
    $checkpoint = if ($Resume) { Get-Checkpoint $ListConfig.name } else { $null }
    $lastProcessed = if ($checkpoint -and $checkpoint.phase -eq 'Reconcile') { [int]$checkpoint.lastSourceId } else { 0 }
    $query = Get-SourceQuery $lastProcessed

    $sourceFields = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    [void]$sourceFields.Add('ID'); [void]$sourceFields.Add('Modified')
    foreach ($field in $ListConfig.fields) { [void]$sourceFields.Add([string]$field.source) }

    $operations = [System.Collections.Generic.List[object]]::new()
    $counters = @{ Read=0; Create=0; Update=0; Delete=0; Unchanged=0; LastSourceId=$lastProcessed }
    $params = @{
        List=$ListConfig.sourceList; Fields=@($sourceFields)
        PageSize=[int]$script:Configuration.readPageSize; Connection=$script:SourceConnection
    }
    if ($query) { $params.Query = $query }

    Get-PnPListItem @params | ForEach-Object {
        $record = Convert-SourceItem -Item $_ -ListConfig $ListConfig
        $counters.Read++; $counters.LastSourceId = $record.SourceID
        if (-not $targetIndex.ContainsKey($record.SourceID)) {
            $counters.Create++
            $operations.Add([pscustomobject]@{ Type='Create'; SourceID=$record.SourceID; Values=$record.Values })
        }
        elseif ($targetIndex[$record.SourceID].SourceHash -ne $record.SourceHash) {
            $counters.Update++
            $operations.Add([pscustomobject]@{ Type='Update'; SourceID=$record.SourceID; TargetID=$targetIndex[$record.SourceID].TargetID; Values=$record.Values })
        }
        else { $counters.Unchanged++ }

        if ($operations.Count -ge [int]$script:Configuration.writeBatchSize) {
            Invoke-WriteBatch $ListConfig $operations $counters
            Save-Checkpoint $ListConfig.name @{ phase='Reconcile'; lastSourceId=$counters.LastSourceId; counters=$counters }
        }
        elseif (($counters.Read % [int]$script:Configuration.checkpointEveryItems) -eq 0) {
            Save-Checkpoint $ListConfig.name @{ phase='Reconcile'; lastSourceId=$counters.LastSourceId; counters=$counters }
        }
    }
    Invoke-WriteBatch $ListConfig $operations $counters
    Save-Checkpoint $ListConfig.name @{ phase='Reconciled'; lastSourceId=$counters.LastSourceId; counters=$counters }

    # Reload after batched creates so IDs are available to dependent lists.
    $targetIndex = Get-TargetIndex $ListConfig
    $idMap = @{}
    foreach ($key in $targetIndex.Keys) { $idMap[[int]$key] = [int]$targetIndex[$key].TargetID }
    $script:IdMaps[[string]$ListConfig.name] = $idMap

    # Deletion uses a complete source-key rescan even after a resumed reconciliation.
    if ([bool]$ListConfig.deleteTargetOnly) {
        if (-not $AllowDelete) {
            Write-SyncLog Warning "Deletion configured for '$($ListConfig.name)' but -AllowDelete was not supplied; no items deleted." @{}
        } else {
            $sourceIds = [System.Collections.Generic.HashSet[int]]::new()
            Get-PnPListItem -List $ListConfig.sourceList -Fields @('ID') `
                -PageSize $script:Configuration.readPageSize -Connection $script:SourceConnection |
                ForEach-Object { [void]$sourceIds.Add([int]$_.Id) }
            $deleteOps = [System.Collections.Generic.List[object]]::new()
            foreach ($sourceId in @($targetIndex.Keys)) {
                if (-not $sourceIds.Contains([int]$sourceId)) {
                    $counters.Delete++
                    if ($PSCmdlet.ShouldProcess("$($ListConfig.targetList) item $($targetIndex[$sourceId].TargetID)", "Delete target-only SourceID $sourceId")) {
                        $deleteOps.Add([pscustomobject]@{ Type='Delete'; SourceID=$sourceId; TargetID=$targetIndex[$sourceId].TargetID })
                    }
                    if ($deleteOps.Count -ge [int]$script:Configuration.writeBatchSize) {
                        Invoke-WriteBatch $ListConfig $deleteOps $counters
                    }
                }
            }
            Invoke-WriteBatch $ListConfig $deleteOps $counters
        }
    }
    Save-Checkpoint $ListConfig.name @{ phase='Complete'; lastSourceId=$counters.LastSourceId; counters=$counters }
    return $counters
}

function Test-ListSnapshot {
    param([Parameter(Mandatory)]$ListConfig)
    $sourceHashes = @{}
    $sourceFields = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    [void]$sourceFields.Add('ID'); [void]$sourceFields.Add('Modified')
    foreach ($field in $ListConfig.fields) { [void]$sourceFields.Add([string]$field.source) }

    Get-PnPListItem -List $ListConfig.sourceList -Fields @($sourceFields) `
        -PageSize $script:Configuration.readPageSize -Connection $script:SourceConnection |
        ForEach-Object {
            $record = Convert-SourceItem $_ $ListConfig
            if ($sourceHashes.ContainsKey($record.SourceID)) { throw "Duplicate source ID '$($record.SourceID)'." }
            $sourceHashes[$record.SourceID] = $record.SourceHash
        }

    $targetHashes = @{}
    Get-PnPListItem -List $ListConfig.targetList `
        -Fields @('ID',[string]$ListConfig.sourceIdField,[string]$ListConfig.sourceHashField) `
        -PageSize $script:Configuration.readPageSize -Connection $script:TargetConnection |
        ForEach-Object {
            $raw = $_.FieldValues[[string]$ListConfig.sourceIdField]
            if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) { return }
            $id = [int]$raw
            if ($targetHashes.ContainsKey($id)) { throw "Duplicate target SourceID '$id'." }
            $targetHashes[$id] = [string]$_.FieldValues[[string]$ListConfig.sourceHashField]
        }

    $missing = @($sourceHashes.Keys | Where-Object { -not $targetHashes.ContainsKey($_) })
    $extra = @($targetHashes.Keys | Where-Object { -not $sourceHashes.ContainsKey($_) })
    $mismatch = @($sourceHashes.Keys | Where-Object { $targetHashes.ContainsKey($_) -and $sourceHashes[$_] -ne $targetHashes[$_] })
    $valid = ($missing.Count -eq 0 -and $extra.Count -eq 0 -and $mismatch.Count -eq 0)
    $result = [ordered]@{
        List=$ListConfig.name; SourceCount=$sourceHashes.Count; TargetCount=$targetHashes.Count
        MissingCount=$missing.Count; ExtraCount=$extra.Count; HashMismatchCount=$mismatch.Count
        MissingSample=@($missing | Select-Object -First 20)
        ExtraSample=@($extra | Select-Object -First 20)
        HashMismatchSample=@($mismatch | Select-Object -First 20)
        Valid=$valid
    }
    Write-SyncLog -Level $(if ($valid) {'Information'} else {'Error'}) `
        -Message "Validation for '$($ListConfig.name)': valid=$valid source=$($sourceHashes.Count) target=$($targetHashes.Count) missing=$($missing.Count) extra=$($extra.Count) hashMismatch=$($mismatch.Count)" `
        -Data $result
    return [pscustomobject]$result
}

try {
    $resolvedConfig = (Resolve-Path -LiteralPath $ConfigurationPath).Path
    $script:Configuration = Get-Content -LiteralPath $resolvedConfig -Raw | ConvertFrom-Json
    Assert-Configuration $script:Configuration

    $basePath = Split-Path -Parent $resolvedConfig
    $logDirectory = if ($script:Configuration.logDirectory) { [string]$script:Configuration.logDirectory } else { Join-Path $basePath 'logs' }
    $checkpointDirectory = if ($script:Configuration.checkpointDirectory) { [string]$script:Configuration.checkpointDirectory } else { Join-Path $basePath 'checkpoints' }
    New-Item -ItemType Directory -Force -Path $logDirectory, $checkpointDirectory | Out-Null
    $script:LogPath = Join-Path $logDirectory "SharePointListSync-$($script:RunId).jsonl"
    $script:CheckpointPath = Join-Path $checkpointDirectory 'SharePointListSync.checkpoint.json'

    Write-SyncLog Information 'Starting SharePoint list synchronization.' @{ config=$resolvedConfig; whatIf=[bool]$WhatIfPreference; resume=[bool]$Resume }
    $script:SourceConnection = Connect-SyncSite $script:Configuration.sourceSiteUrl $script:Configuration.authentication
    $script:TargetConnection = Connect-SyncSite $script:Configuration.targetSiteUrl $script:Configuration.authentication
    $script:RunSummary.State = 'Loading'
    Set-RefreshStatus -State Loading

    $hasManualLists = $script:Configuration.PSObject.Properties.Name.Contains('lists') -and @($script:Configuration.lists).Count -gt 0
    if (-not $hasManualLists) {
        Write-SyncLog Information "Resolving list pairs from SharePoint mapping list '$($script:Configuration.listMapping.listName)'." @{}
        $autoLists = Get-AutoListConfigurations
        $script:Configuration | Add-Member -NotePropertyName lists -NotePropertyValue $autoLists -Force
        Write-SyncLog Information "Resolved $($autoLists.Count) list pair(s) from SharePoint." @{ lists = @($autoLists | ForEach-Object { $_.name }) }
    }

    Assert-ListPairsDistinct @($script:Configuration.lists)
    $orderedLists = Get-DependencyOrderedLists @($script:Configuration.lists)
    foreach ($list in $orderedLists) {
        $counters = Invoke-ListReconciliation $list
        $script:RunSummary.Lists += [pscustomobject]@{ Name=$list.name; Counters=$counters }
    }

    if ($WhatIfPreference) {
        $script:RunSummary.State = 'WhatIfComplete'
        $script:RunSummary | Add-Member -NotePropertyName Validation -NotePropertyValue @(
            [pscustomobject]@{ Valid=$null; Note='Skipped because WhatIf does not modify the target.' }
        ) -Force
    } else {
        $script:RunSummary.State = 'Validating'
        Set-RefreshStatus -State Validating -Summary $script:RunSummary
        $validation = foreach ($list in $orderedLists) { Test-ListSnapshot $list }
        $script:RunSummary | Add-Member -NotePropertyName Validation -NotePropertyValue $validation -Force
        if (@($validation | Where-Object { -not $_.Valid }).Count -gt 0) { throw 'One or more list validations failed.' }
        $script:RunSummary.State = 'Ready'
    }
    $script:RunSummary.CompletedUtc = [DateTime]::UtcNow.ToString('o')
    if (-not $WhatIfPreference) { Set-RefreshStatus -State Ready -Summary $script:RunSummary }
    Write-SyncLog Information "Synchronization completed with state '$($script:RunSummary.State)'." $script:RunSummary
    $script:RunSummary | ConvertTo-Json -Depth 15
    exit 0
}
catch {
    $script:RunSummary.State = 'Failed'; $script:RunSummary.Errors++
    $script:RunSummary.CompletedUtc = [DateTime]::UtcNow.ToString('o')
    $script:RunSummary | Add-Member -NotePropertyName Error -NotePropertyValue $_.Exception.Message -Force
    Write-SyncLog Error 'Synchronization failed.' @{ error=$_.Exception.Message; stack=$_.ScriptStackTrace }
    try { if ($script:TargetConnection) { Set-RefreshStatus -State Failed -Summary $script:RunSummary } } catch { }
    $script:RunSummary | ConvertTo-Json -Depth 15
    exit 1
}
