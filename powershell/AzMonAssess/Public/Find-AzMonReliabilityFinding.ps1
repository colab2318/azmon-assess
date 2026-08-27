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
        [array] $ResourceRef = @(),
        [string[]] $SubscriptionId = @()
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()
    $compliance = [System.Collections.Generic.List[hashtable]]::new()
    $azRegions = Get-AzMonAzRegionSet

    $tagsById = @{}
    foreach ($r in $ResourceRef) { $tagsById[([string]$r['Id']).ToLowerInvariant()] = $r['Tags'] }

    foreach ($ws in $Workspace) {
        $loc = ([string]$ws['Location']).ToLowerInvariant()
        if ($loc -and ($azRegions -notcontains $loc)) {
            $findings.Add((New-AzMonFinding -Category 'reliability' -Severity $(if (Test-AzMonProdTag $ws['Tags']) { 'high' } else { 'medium' }) `
                -Title "$($ws['Name']): region '$($ws['Location'])' does not support workspace availability zones" `
                -CheckId 'reliability.workspace-non-az-region' `
                -Detail ('Workspaces in non-AZ regions have no in-region redundancy. A single datacenter fault ' +
                    'can make log ingestion and query unavailable, breaking incident response for every ' +
                    'downstream service.') `
                -ResourceIds @($ws['Id']) `
                -Recommendation ('Plan a migration to an AZ-enabled region (e.g., eastus2, westus2, westeurope). ' +
                    'If moving the workspace is not feasible, add a secondary workspace in an AZ region and ' +
                    'dual-write diagnostic settings for mission-critical resources.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/availability-zones' `
                -Evidence @{ location = $ws['Location']; az_capable = $false }))
        } elseif ($loc) {
            $compliance.Add((New-AzMonComplianceItem -Category 'reliability' -CheckId 'reliability.workspace-non-az-region' `
                -Title "$($ws['Name']): region '$($ws['Location'])' supports workspace availability zones" `
                -ResourceIds @($ws['Id'])))
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
                -CheckId 'reliability.single-log-destination' `
                -Detail ('For business-critical workloads, a single workspace destination leaves no evidence ' +
                    'trail if the workspace becomes unavailable. WAF recommends dual-writing to a storage ' +
                    'account or Event Hub for at least 90-day retention off the analytics path.') `
                -ResourceIds @($resourceId) `
                -Recommendation ('Add a second destination (storage account for cheap long-term retention or a ' +
                    'paired workspace in another region) to the diagnostic setting.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/essentials/diagnostic-settings'))
        } elseif ($prod -and $hasWorkspace -and ($hasStorage -or $hasEventHub)) {
            $compliance.Add((New-AzMonComplianceItem -Category 'reliability' -CheckId 'reliability.single-log-destination' `
                -Title "Prod resource $(($resourceId -split '/')[-1]): has more than one log destination" `
                -ResourceIds @($resourceId)))
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
            -CheckId 'reliability.workspace-no-health-alert' `
            -Detail ('There is no alert rule scoped to this workspace. Ingestion outages, quota hits, or query ' +
                'failures will go undetected until users notice missing data or broken dashboards.') `
            -ResourceIds @($ws['Id']) `
            -Recommendation ('Deploy a metric alert on Heartbeat (or a log query alert on Operation table) to ' +
                'fire when ingestion stops for >30 min. Link it to the primary action group.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/log-analytics-workspace-health'))
    }
    $healthAlertedWorkspaces = @($Workspace | Where-Object { $alertedWorkspaces.Contains(([string]$_['Id']).ToLowerInvariant()) -and (([double]($_['IngestionGb30d'] ?? 0)) -ge 1) })
    if ($healthAlertedWorkspaces.Count -gt 0) {
        $compliance.Add((New-AzMonComplianceItem -Category 'reliability' -CheckId 'reliability.workspace-no-health-alert' `
            -Title "$($healthAlertedWorkspaces.Count) active workspace(s) have a health / ingestion alert configured" `
            -ResourceIds @($healthAlertedWorkspaces | ForEach-Object { $_['Id'] })))
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
                -CheckId 'reliability.ai-non-az-region' `
                -Detail ('Application Insights ingestion in a non-AZ region shares fate with a single datacenter. ' +
                    'During a zonal fault, telemetry from the monitored app is lost - exactly when it is most needed.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Re-create the AI resource in an AZ-capable region colocated with the backing ' +
                    'workspace, then migrate SDK connection strings during the next release.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/availability-zones' `
                -Evidence @{ location = $ai['Location']; workspace_location = $parent['Location'] }))
        } elseif ($aiLoc) {
            $compliance.Add((New-AzMonComplianceItem -Category 'reliability' -CheckId 'reliability.ai-non-az-region' `
                -Title "App Insights '$($ai['Name'])' is in an AZ-capable region" `
                -ResourceIds @($ai['Id'])))
        }
    }

    # ---- Azure Service Health alert coverage ----------------------------
    # Standard baseline: an activityLogAlert whose condition matches
    # category=ServiceHealth, enabled, scoped to the subscription. Without
    # one, Azure outage/maintenance notices are never proactively surfaced.
    $serviceHealthCoveredSubs = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($rule in $AlertRule) {
        if ($rule['AlertKind'] -ne 'activityLog' -or -not $rule['Enabled'] -or -not $rule['IsServiceHealthAlert']) { continue }
        foreach ($scope in @($rule['Scopes'])) {
            $s = ([string]$scope).ToLowerInvariant()
            if ($s -match '^/subscriptions/([^/]+)\s*$') { [void]$serviceHealthCoveredSubs.Add($matches[1]) }
        }
    }
    $uncoveredSubs = @($SubscriptionId | Where-Object { -not $serviceHealthCoveredSubs.Contains(([string]$_).ToLowerInvariant()) })
    if ($uncoveredSubs.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'reliability' -Severity 'medium' `
            -Title "$($uncoveredSubs.Count) subscription(s) have no Azure Service Health alert configured" `
            -CheckId 'reliability.service-health-alert-missing' `
            -Detail ('Without a Service Health alert, Azure service outages, planned maintenance, and health ' +
                'advisories affecting your resources are only visible by checking the portal manually - they are ' +
                'never proactively pushed to your team.') `
            -ResourceIds @($uncoveredSubs | ForEach-Object { "/subscriptions/$_" }) `
            -Recommendation ('Create an activity log alert scoped to the subscription with condition ' +
                'category=ServiceHealth, covering Service issues, Planned maintenance, and Security advisories, ' +
                'routed to the primary action group.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/service-health/alerts-activity-log-service-notifications-arm' `
            -Evidence @{ subscription_ids = $uncoveredSubs }))
    }
    $coveredSubs = @($SubscriptionId | Where-Object { $serviceHealthCoveredSubs.Contains(([string]$_).ToLowerInvariant()) })
    if ($coveredSubs.Count -gt 0) {
        $compliance.Add((New-AzMonComplianceItem -Category 'reliability' -CheckId 'reliability.service-health-alert-missing' `
            -Title "$($coveredSubs.Count) subscription(s) have an Azure Service Health alert configured" `
            -ResourceIds @($coveredSubs | ForEach-Object { "/subscriptions/$_" })))
    }

    return @{ Findings = $findings.ToArray(); ComplianceItems = $compliance.ToArray() }
}
