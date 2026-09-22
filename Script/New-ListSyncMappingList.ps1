<#
.SYNOPSIS
    Creates the "List Sync Configuration" SharePoint list used by
    Invoke-SharePointListSync.ps1 to discover which production lists get
    copied to which test lists, so the client can add or remove synced
    list pairs from SharePoint without editing the script or its JSON
    configuration.

.DESCRIPTION
    One item = one source/test list pair. Columns:
      Title             - friendly sync name (also the default "lists[].name")
      SourceListName    - production list Title/internal identity to read from
      TargetListName    - test list Title/internal identity to write to
      Enabled           - Yes/No; disabled rows are skipped (default Yes)
      DeleteTargetOnly  - Yes/No; mirrors the per-list "deleteTargetOnly"
                          config flag. Still requires the script's
                          -AllowDelete switch before anything is deleted.
      ProcessOrder      - Number; optional manual tie-breaker for run order.
                          Lookup-field dependencies between configured list
                          pairs are detected automatically and do not need
                          this column set.
      Notes             - free-text, for the client's own documentation.

    Idempotent: safe to re-run, will not recreate the list or fields that
    already exist.

.PARAMETER SiteUrl
    The SharePoint site where this configuration list should live -
    normally the test site, since that is where Invoke-SharePointListSync.ps1
    already holds a connection while resolving list pairs.

.PARAMETER ClientId
    Entra ID app registration Client ID used for interactive sign-in, same
    app registration used by Invoke-SharePointListSync.ps1 and by
    Provision-MARFoundationLists.ps1.

.PARAMETER WhatIfOnly
    If set, connects and reports what would be created without making any
    changes.

.EXAMPLE
    .\New-ListSyncMappingList.ps1 `
        -SiteUrl "https://tenant.sharepoint.com/sites/Test" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -WhatIfOnly

.NOTES
    Requires the PnP.PowerShell module. Run interactively - this script
    does not store or accept credentials.
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$ListTitle = 'List Sync Configuration',

    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
    [string]$Tenant,

    [switch]$WhatIfOnly
)

$ErrorActionPreference = 'Stop'

Import-Module PnP.PowerShell -ErrorAction Stop

if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
    Write-Host "Connecting to $SiteUrl app-only (as the app, no sign-in prompt) ..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant $Tenant
}
else {
    Write-Host "Connecting to $SiteUrl (browser sign-in as your account) ..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
}

function Ensure-Field {
    param(
        [Parameter()][AllowNull()]$List,
        [Parameter(Mandatory = $true)][string]$InternalName,
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$Type,
        [switch]$Required,
        [switch]$AddToDefaultView,
        [string]$DefaultValue
    )

    if (-not $List) {
        Write-Host "[WhatIf] Would add field '$DisplayName' ($Type) to list." -ForegroundColor Green
        return
    }

    $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Field '$DisplayName' already exists - skipping." -ForegroundColor Yellow
        return
    }

    Write-Host "  Adding field '$DisplayName' ($Type) ..." -ForegroundColor Green

    switch ($Type) {
        'Text' {
            Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName `
                -Type Text -Required:$Required -AddToDefaultView:$AddToDefaultView | Out-Null
        }
        'Note' {
            Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName `
                -Type Note -Required:$Required -AddToDefaultView:$AddToDefaultView | Out-Null
        }
        'Number' {
            Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName `
                -Type Number -Required:$Required -AddToDefaultView:$AddToDefaultView | Out-Null
        }
        'Boolean' {
            $field = Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName `
                -Type Boolean -Required:$Required -AddToDefaultView:$AddToDefaultView
            if ($PSBoundParameters.ContainsKey('DefaultValue')) {
                Set-PnPField -List $List -Identity $InternalName -Values @{ DefaultValue = $DefaultValue } | Out-Null
            }
        }
        default {
            throw "Unsupported field type: $Type"
        }
    }
}

$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if ($list) {
    Write-Host "List '$ListTitle' already exists - skipping creation." -ForegroundColor Yellow
}
elseif ($WhatIfOnly) {
    Write-Host "[WhatIf] Would create list '$ListTitle'." -ForegroundColor Green
}
else {
    Write-Host "Creating list '$ListTitle' ..." -ForegroundColor Green
    $list = New-PnPList -Title $ListTitle -Template GenericList -EnableVersioning -OnQuickLaunch
    Set-PnPList -Identity $ListTitle `
        -Description 'Drives Invoke-SharePointListSync.ps1: one row per production list that should be copied to a test list. Add or remove rows here to change what the sync job processes - no script or config file changes needed.' `
        -MajorVersions 50 | Out-Null
}

Ensure-Field -List $list -InternalName 'SourceListName' -DisplayName 'Source List Name' -Type Text -Required -AddToDefaultView
Ensure-Field -List $list -InternalName 'TargetListName' -DisplayName 'Target List Name' -Type Text -Required -AddToDefaultView
Ensure-Field -List $list -InternalName 'Enabled' -DisplayName 'Enabled' -Type Boolean -DefaultValue '1' -AddToDefaultView
Ensure-Field -List $list -InternalName 'DeleteTargetOnly' -DisplayName 'Delete Target-Only Items' -Type Boolean -DefaultValue '0' -AddToDefaultView
Ensure-Field -List $list -InternalName 'ProcessOrder' -DisplayName 'Process Order (optional)' -Type Number
Ensure-Field -List $list -InternalName 'Notes' -DisplayName 'Notes' -Type Note

Write-Host ""
Write-Host "Done. Add one item per list pair, e.g. Title='Parents', SourceListName='Parents', TargetListName='Parents-Test', Enabled=Yes." -ForegroundColor Cyan
