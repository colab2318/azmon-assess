#requires -Version 7.0
# Fixture-data generator — lets you preview reports without Azure access.
# Ported 1:1 from cli.py's `demo` command; values are chosen to exercise
# every analyzer rule at least once.

function New-AzMonDemoSnapshot {
    [CmdletBinding()]
    param(
        [string] $CustomerName = 'Your Organization',
        [string] $AoaiEndpoint,
        [string] $AoaiDeployment,
        [string] $AoaiApiVersion
    )

    $envRegionPairs = @(
        @('prod', 'eastus'), @('prod', 'eastus'), @('prod', 'eastus'),
        @('prod', 'westus2'), @('prod', 'westus2'),
        @('dev', 'eastus'), @('dev', 'eastus'),
        @('qa', 'eastus'), @('qa', 'eastus'),
        @('core', 'eastus'), @('test', 'eastus')
    )

    $workspaces = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = 0; $i -lt $envRegionPairs.Count; $i++) {
        $envName = $envRegionPairs[$i][0]
        $region = $envRegionPairs[$i][1]
        $nameSuffix = '{0:d2}' -f $i
        $name = "la-$envName-$region-$nameSuffix"
        $workspaces.Add((New-AzMonWorkspace `
                -Id "/subscriptions/demo/rg/rg/providers/Microsoft.OperationalInsights/workspaces/$name" `
                -Name $name -SubscriptionId 'demo' -ResourceGroup 'rg-monitoring' -Location $region `
                -Sku 'PerGB2018' -RetentionDays 90 -Tags @{ env = $envName } `
                -IngestionGb30d ([double](80 + $i * 25)) `
                -IngestionByTable @{ AppTraces = 20.0; ContainerLogV2 = 40.0; AzureDiagnostics = 25.0; Perf = 15.0 } `
                -PublicNetworkAccessForIngestion $(if ($i % 2 -eq 0) { 'Enabled' } else { 'Disabled' }) `
                -PublicNetworkAccessForQuery $(if ($i % 3 -eq 0) { 'Enabled' } else { 'Disabled' }) `
                -DisableLocalAuth ($i % 2 -eq 1)))
    }
    $workspaces.Add((New-AzMonWorkspace `
            -Id '/subscriptions/demo/rg/rg/providers/Microsoft.OperationalInsights/workspaces/la-prod-westcentralus-99' `
            -Name 'la-prod-westcentralus-99' -SubscriptionId 'demo' -ResourceGroup 'rg-monitoring' -Location 'westcentralus' `
            -Sku 'PerGB2018' -RetentionDays 180 -Tags @{ env = 'prod' } -IngestionGb30d 16000.0 `
            -IngestionByTable @{ AzureDiagnostics = 8000.0; ContainerLogV2 = 5000.0; SecurityEvent = 1500.0; AppTraces = 1500.0 } `
            -PublicNetworkAccessForIngestion 'Enabled' -PublicNetworkAccessForQuery 'Enabled' -DisableLocalAuth $false))
    $workspaceArr = $workspaces.ToArray()

    $appInsights = @(
        (New-AzMonAppInsight -Id '/subscriptions/demo/rg/rg/providers/microsoft.insights/components/webapp-primary' `
                -Name 'webapp-primary' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'westus2' `
                -WorkspaceResourceId $workspaceArr[0]['Id'] -SamplingPercentage 100 `
                -PublicNetworkAccessForIngestion 'Enabled' -PublicNetworkAccessForQuery 'Enabled' -DisableLocalAuth $false)
        (New-AzMonAppInsight -Id '/subscriptions/demo/rg/rg/providers/microsoft.insights/components/legacy' `
                -Name 'legacy' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus')
    )

    $rules = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = 0; $i -lt 20; $i++) {
        $agIds = if ($i % 4 -eq 0) { @() } else { @('/subscriptions/demo/ag/ag-1') }
        $fireCount = if ($i -eq 0) { 180 } elseif ($i -eq 1) { 60 } else { $null }
        $rules.Add((New-AzMonAlertRule -Id "/subscriptions/demo/rule-$i" -Name "rule-$i" -SubscriptionId 'demo' `
                -ResourceGroup 'rg' -AlertKind 'metric' -Enabled ($i % 3 -ne 0) -Severity 2 `
                -Scopes @($workspaceArr[0]['Id']) -ActionGroupIds $agIds -FireCount30d $fireCount))
    }
    # Exercises the broad-metric-scope cost finding (many resources on one rule).
    $rules.Add((New-AzMonAlertRule -Id '/subscriptions/demo/rule-broad-scope' -Name 'rule-broad-scope' -SubscriptionId 'demo' `
            -ResourceGroup 'rg' -AlertKind 'metric' -Enabled $true -Severity 3 `
            -Scopes (1..25 | ForEach-Object { "/subscriptions/demo/rg/rg/providers/Microsoft.Compute/virtualMachines/fleet$_" }) `
            -ActionGroupIds @('/subscriptions/demo/ag/ag-1')))
    # Exercises the high-frequency log-search-alert cost finding.
    $rules.Add((New-AzMonAlertRule -Id '/subscriptions/demo/rule-high-freq-log' -Name 'rule-high-freq-log' -SubscriptionId 'demo' `
            -ResourceGroup 'rg' -AlertKind 'log' -Enabled $true -Severity 3 `
            -Scopes @($workspaceArr[0]['Id']) -ActionGroupIds @('/subscriptions/demo/ag/ag-1') -EvaluationFrequencyMinutes 1))
    # Good example: uses a dynamic threshold, so it's NOT flagged by the
    # static-threshold-only finding (rule-0..19 above all default to static).
    $rules.Add((New-AzMonAlertRule -Id '/subscriptions/demo/rule-dynamic-threshold' -Name 'rule-dynamic-threshold' -SubscriptionId 'demo' `
            -ResourceGroup 'rg' -AlertKind 'metric' -Enabled $true -Severity 2 `
            -Scopes @($workspaceArr[0]['Id']) -ActionGroupIds @('/subscriptions/demo/ag/ag-1') -HasDynamicThreshold $true))
    # No Service Health activity-log alert is added for the 'demo'
    # subscription, intentionally exercising the missing-coverage finding.

    $actionGroups = @(
        (New-AzMonActionGroup -Id '/subscriptions/demo/ag/ag-1' -Name 'ag-1' -SubscriptionId 'demo' -ResourceGroup 'rg' -UsedByRules 15)
        (New-AzMonActionGroup -Id '/subscriptions/demo/ag/ag-orphan' -Name 'ag-orphan' -SubscriptionId 'demo' -ResourceGroup 'rg' -UsedByRules 0)
    )

    $resources = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($t in @('Microsoft.KeyVault/vaults', 'Microsoft.Web/sites', 'Microsoft.Sql/servers/databases', 'Microsoft.ContainerService/managedClusters')) {
        for ($i = 0; $i -lt 3; $i++) {
            $resources.Add((New-AzMonResourceRef -Id "/subscriptions/demo/rg/rg/providers/$t/x$i" -Name "x$i" -Type $t `
                    -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus'))
        }
    }

    $vms = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = 0; $i -lt 5; $i++) {
        $vms.Add((New-AzMonResourceRef -Id "/subscriptions/demo/rg/rg/providers/Microsoft.Compute/virtualMachines/vm$i" -Name "vm$i" `
                -Type 'Microsoft.Compute/virtualMachines' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus'))
    }
    foreach ($vm in $vms) { $resources.Add($vm) }
    # vm0-vm2 have a recent heartbeat; vm3/vm4 do not (exercises the coverage-gap finding).
    $heartbeatResourceIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($vm in ($vms | Select-Object -First 3)) { [void]$heartbeatResourceIds.Add(([string]$vm['Id']).ToLowerInvariant()) }

    # vm0: legacy agent only (exercises the retired-MMA-agent finding).
    # vm1: Azure Monitor Agent only (already migrated, no finding).
    # vm2: both agents present (mid-migration, no finding). vm3/vm4: neither.
    $vmAgentExtensions = @(
        @{ Kind = 'VmAgentExtension'; VmId = $vms[0]['Id']; AgentKind = 'legacy'; ExtensionType = 'MicrosoftMonitoringAgent'; SubscriptionId = 'demo'; ResourceGroup = 'rg' }
        @{ Kind = 'VmAgentExtension'; VmId = $vms[1]['Id']; AgentKind = 'ama'; ExtensionType = 'AzureMonitorWindowsAgent'; SubscriptionId = 'demo'; ResourceGroup = 'rg' }
        @{ Kind = 'VmAgentExtension'; VmId = $vms[2]['Id']; AgentKind = 'legacy'; ExtensionType = 'OmsAgentForLinux'; SubscriptionId = 'demo'; ResourceGroup = 'rg' }
        @{ Kind = 'VmAgentExtension'; VmId = $vms[2]['Id']; AgentKind = 'ama'; ExtensionType = 'AzureMonitorLinuxAgent'; SubscriptionId = 'demo'; ResourceGroup = 'rg' }
    )

    # Exercises the Azure Advisor cost-recommendation finding.
    $advisorRecommendations = @(
        @{
            Kind             = 'AdvisorRecommendation'
            Id               = '/subscriptions/demo/providers/Microsoft.Advisor/recommendations/demo-1'
            ResourceId       = $workspaceArr[0]['Id']
            ImpactedField    = 'microsoft.operationalinsights/workspaces'
            Impact           = 'Medium'
            Problem          = 'Consider configuring the cost effective Basic logs plan on selected tables'
            Solution         = 'One or more tables are eligible for the low-cost Basic log data plan, which still supports query for debugging and troubleshooting.'
            PotentialBenefit = 'Lower ingestion cost for eligible tables'
            LearnMoreLink    = 'https://learn.microsoft.com/azure/azure-monitor/logs/logs-table-plans'
        }
    )

    $dcrs = @(
        (New-AzMonDataCollectionRule -Id '/subscriptions/demo/rg/rg/providers/Microsoft.Insights/dataCollectionRules/dcr-vm-perf' `
                -Name 'dcr-vm-perf' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus' `
                -DcrKind 'Linux' -DataFlowCount 2 -WorkspaceResourceId $workspaceArr[0]['Id'])
        (New-AzMonDataCollectionRule -Id '/subscriptions/demo/rg/rg/providers/Microsoft.Insights/dataCollectionRules/dcr-container-insights' `
                -Name 'dcr-container-insights' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus' `
                -DcrKind 'Linux' -DataFlowCount 1 -WorkspaceResourceId $workspaceArr[1]['Id'])
        # Zero associations - exercises the orphaned-DCR finding.
        (New-AzMonDataCollectionRule -Id '/subscriptions/demo/rg/rg/providers/Microsoft.Insights/dataCollectionRules/dcr-orphaned' `
                -Name 'dcr-orphaned' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus' `
                -DcrKind 'Linux' -DataFlowCount 1 -WorkspaceResourceId $workspaceArr[0]['Id'])
    )
    $dcrAssociatedIds = [System.Collections.Generic.HashSet[string]]::new()
    [void]$dcrAssociatedIds.Add(([string]$dcrs[0]['Id']).ToLowerInvariant())
    [void]$dcrAssociatedIds.Add(([string]$dcrs[1]['Id']).ToLowerInvariant())

    $snapshot = New-AzMonSnapshotObject -SubscriptionId @('demo') -CustomerName $CustomerName `
        -Workspaces $workspaceArr -AppInsights $appInsights -AlertRules $rules.ToArray() `
        -ActionGroups $actionGroups -Resources $resources.ToArray() -DataCollectionRules $dcrs

    $snapshot['Findings'] = @()
    $snapshot['Findings'] += Find-AzMonConsolidationFinding -Workspace $workspaceArr
    $snapshot['Findings'] += Find-AzMonCoverageGapFinding -ResourceRef $resources.ToArray() -DiagnosticSetting @() -Workspace $workspaceArr -AppInsight $appInsights -HeartbeatResourceId $heartbeatResourceIds -DataCollectionRule $dcrs -DcrAssociatedId $dcrAssociatedIds
    $snapshot['Findings'] += Find-AzMonAlertQualityFinding -AlertRule $rules.ToArray() -ActionGroup $actionGroups -ResourceRef $resources.ToArray()
    $snapshot['Findings'] += Find-AzMonCostOptimizationFinding -Workspace $workspaceArr -AppInsight $appInsights -VmAgentExtension $vmAgentExtensions -AdvisorRecommendation $advisorRecommendations
    $snapshot['Findings'] += Find-AzMonTracingFinding -ResourceRef $resources.ToArray() -AppInsight $appInsights
    $snapshot['Findings'] += Find-AzMonReliabilityFinding -Workspace $workspaceArr -AppInsight $appInsights -AlertRule $rules.ToArray() -DiagnosticSetting @() -ResourceRef $resources.ToArray() -SubscriptionId @('demo')
    $snapshot['Findings'] += Find-AzMonSecurityFinding -Workspace $workspaceArr -AppInsight $appInsights
    $snapshot['Findings'] += Find-AzMonPerformanceFinding -Workspace $workspaceArr -AppInsight $appInsights

    $snapshot['AiSummary'] = New-AzMonAiSummary -Snapshot $snapshot -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion
    return $snapshot
}
