#requires -Version 7.0
# WAF Reliability analyzer — ported 1:1 from analyzers/reliability.py.

$script:AzMonDefaultAzRegions = @(
    'australiaeast', 'brazilsouth', 'canadacentral', 'centralindia', 'centralus', 'eastasia', 'eastus', 'eastus2',
    'francecentral', 'germanywestcentral', 'italynorth', 'japaneast', 'koreacentral', 'northeurope', 'norwayeast',
    'polandcentral', 'qatarcentral', 'southafricanorth', 'southcentralus', 'southeastasia', 'spaincentral',
    'swedencentral', 'switzerlandnorth', 'uaenorth', 'uksouth', 'westeurope', 'westus2', 'westus3'
)

function Get-AzMonAzRegionSet {
    [CmdletBinding()]
    param()
    if ($env:AZMON_AZ_REGIONS) {
        $fromEnv = @($env:AZMON_AZ_REGIONS -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
        if ($fromEnv.Count -gt 0) { return $fromEnv }
    }
    return $script:AzMonDefaultAzRegions
}

function Find-AzMonReliabilityFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight,
        [array] $AlertRule = @(),
        [array] $DiagnosticSetting = @(),
        [array] $ResourceRef = @()
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()
    $azRegions = Get-AzMonAzRegionSet

    $tagsById = @{}
    foreach ($r in $ResourceRef) { $tagsById[([string]$r['Id']).ToLowerInvariant()] = $r['Tags'] }

    foreach ($ws in $Workspace) {
        $loc = ([string]$ws['Location']).ToLowerInvariant()
        if ($loc -and ($azRegions -notcontains $loc)) {
            $findings.Add((New-AzMonFinding -Category 'reliability' -Severity $(if (Test-AzMonProdTag $ws['Tags']) { 'high' } else { 'medium' }) `
                -Title "$($ws['Name']): region '$($ws['Location'])' does not support workspace availability zones" `
                -Detail ('Workspaces in non-AZ regions have no in-region redundancy. A single datacenter fault ' +
                    'can make log ingestion and query unavailable, breaking incident response for every ' +
                    'downstream service.') `
                -ResourceIds @($ws['Id']) `
                -Recommendation ('Plan a migration to an AZ-enabled region (e.g., eastus2, westus2, westeurope). ' +
                    'If moving the workspace is not feasible, add a secondary workspace in an AZ region and ' +
                    'dual-write diagnostic settings for mission-critical resources.') `
                -Evidence @{ location = $ws['Location']; az_capable = $false }))
        }
    }

    $dsByResource = @{}
    foreach ($ds in $DiagnosticSetting) {
        $key = ([string]$ds['ResourceId']).ToLowerInvariant()
        if (-not $dsByResource.ContainsKey($key)) { $dsByResource[$key] = [System.Collections.Generic.List[hashtable]]::new() }
        $dsByResource[$key].Add($ds)
    }
    foreach ($resourceId in $dsByResource.Keys) {
        $dsList = $dsByResource[$resourceId]
        $hasWorkspace = @($dsList | Where-Object { $_['WorkspaceId'] }).Count -gt 0
        $hasStorage = @($dsList | Where-Object { $_['StorageId'] }).Count -gt 0
        $hasEventHub = @($dsList | Where-Object { $_['EventHubId'] }).Count -gt 0
        $prod = Test-AzMonProdTag ($tagsById[$resourceId])
        if ($prod -and $hasWorkspace -and -not ($hasStorage -or $hasEventHub)) {
            $shortName = ($resourceId -split '/')[-1]
            $findings.Add((New-AzMonFinding -Category 'reliability' -Severity 'medium' `
                -Title "Prod resource ${shortName}: single log destination (workspace only)" `
                -Detail ('For business-critical workloads, a single workspace destination leaves no evidence ' +
                    'trail if the workspace becomes unavailable. WAF recommends dual-writing to a storage ' +
                    'account or Event Hub for at least 90-day retention off the analytics path.') `
                -ResourceIds @($resourceId) `
                -Recommendation ('Add a second destination (storage account for cheap long-term retention or a ' +
                    'paired workspace in another region) to the diagnostic setting.')))
        }
    }

    $wsIdSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($ws in $Workspace) { [void]$wsIdSet.Add(([string]$ws['Id']).ToLowerInvariant()) }
    $alertedWorkspaces = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($rule in $AlertRule) {
        foreach ($scope in @($rule['Scopes'])) {
            $s = ([string]$scope).ToLowerInvariant()
            if ($wsIdSet.Contains($s)) { [void]$alertedWorkspaces.Add($s) }
        }
    }
    foreach ($ws in $Workspace) {
        if ($alertedWorkspaces.Contains(([string]$ws['Id']).ToLowerInvariant())) { continue }
        if (([double]($ws['IngestionGb30d'] ?? 0)) -lt 1) { continue }
        $findings.Add((New-AzMonFinding -Category 'reliability' -Severity 'medium' `
            -Title "$($ws['Name']): no health / ingestion alert configured on the workspace" `
            -Detail ('There is no alert rule scoped to this workspace. Ingestion outages, quota hits, or query ' +
                'failures will go undetected until users notice missing data or broken dashboards.') `
            -ResourceIds @($ws['Id']) `
            -Recommendation ('Deploy a metric alert on Heartbeat (or a log query alert on Operation table) to ' +
                'fire when ingestion stops for >30 min. Link it to the primary action group.')))
    }

    $wsById = @{}
    foreach ($ws in $Workspace) { $wsById[([string]$ws['Id']).ToLowerInvariant()] = $ws }
    foreach ($ai in $AppInsight) {
        if (-not $ai['WorkspaceResourceId']) { continue }
        $parent = $wsById[([string]$ai['WorkspaceResourceId']).ToLowerInvariant()]
        if (-not $parent) { continue }
        $aiLoc = ([string]$ai['Location']).ToLowerInvariant()
        if ($aiLoc -and ($azRegions -notcontains $aiLoc)) {
            $findings.Add((New-AzMonFinding -Category 'reliability' -Severity 'medium' `
                -Title "App Insights '$($ai['Name'])' in non-AZ region '$($ai['Location'])'" `
                -Detail ('Application Insights ingestion in a non-AZ region shares fate with a single datacenter. ' +
                    'During a zonal fault, telemetry from the monitored app is lost - exactly when it is most needed.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Re-create the AI resource in an AZ-capable region colocated with the backing ' +
                    'workspace, then migrate SDK connection strings during the next release.') `
                -Evidence @{ location = $ai['Location']; workspace_location = $parent['Location'] }))
        }
    }

    return $findings.ToArray()
}
