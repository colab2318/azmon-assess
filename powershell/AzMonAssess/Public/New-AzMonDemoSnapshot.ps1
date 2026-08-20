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

    $dcrs = @(
        (New-AzMonDataCollectionRule -Id '/subscriptions/demo/rg/rg/providers/Microsoft.Insights/dataCollectionRules/dcr-vm-perf' `
                -Name 'dcr-vm-perf' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus' `
                -DcrKind 'Linux' -DataFlowCount 2 -WorkspaceResourceId $workspaceArr[0]['Id'])
        (New-AzMonDataCollectionRule -Id '/subscriptions/demo/rg/rg/providers/Microsoft.Insights/dataCollectionRules/dcr-container-insights' `
                -Name 'dcr-container-insights' -SubscriptionId 'demo' -ResourceGroup 'rg' -Location 'eastus' `
                -DcrKind 'Linux' -DataFlowCount 1 -WorkspaceResourceId $workspaceArr[1]['Id'])
    )

    $snapshot = New-AzMonSnapshotObject -SubscriptionId @('demo') -CustomerName $CustomerName `
        -Workspaces $workspaceArr -AppInsights $appInsights -AlertRules $rules.ToArray() `
        -ActionGroups $actionGroups -Resources $resources.ToArray() -DataCollectionRules $dcrs

    $snapshot['Findings'] = @()
    $snapshot['Findings'] += Find-AzMonConsolidationFinding -Workspace $workspaceArr
    $snapshot['Findings'] += Find-AzMonCoverageGapFinding -ResourceRef $resources.ToArray() -DiagnosticSetting @() -Workspace $workspaceArr -AppInsight $appInsights -HeartbeatResourceId $heartbeatResourceIds
    $snapshot['Findings'] += Find-AzMonAlertQualityFinding -AlertRule $rules.ToArray() -ActionGroup $actionGroups -ResourceRef $resources.ToArray()
    $snapshot['Findings'] += Find-AzMonCostOptimizationFinding -Workspace $workspaceArr -AppInsight $appInsights
    $snapshot['Findings'] += Find-AzMonTracingFinding -ResourceRef $resources.ToArray() -AppInsight $appInsights
    $snapshot['Findings'] += Find-AzMonReliabilityFinding -Workspace $workspaceArr -AppInsight $appInsights -AlertRule $rules.ToArray() -DiagnosticSetting @() -ResourceRef $resources.ToArray()
    $snapshot['Findings'] += Find-AzMonSecurityFinding -Workspace $workspaceArr -AppInsight $appInsights
    $snapshot['Findings'] += Find-AzMonPerformanceFinding -Workspace $workspaceArr -AppInsight $appInsights

    $snapshot['AiSummary'] = New-AzMonAiSummary -Snapshot $snapshot -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion
    return $snapshot
}
