#requires -Version 7.0
# Coverage-gap analyzer — ported 1:1 from analyzers/coverage_gaps.py.

$script:AzMonCriticalDiagnosticTypes = @(
    'microsoft.keyvault/vaults', 'microsoft.sql/servers/databases', 'microsoft.dbforpostgresql/flexibleservers',
    'microsoft.dbformysql/flexibleservers', 'microsoft.storage/storageaccounts', 'microsoft.network/networksecuritygroups',
    'microsoft.network/applicationgateways', 'microsoft.containerservice/managedclusters', 'microsoft.apimanagement/service',
    'microsoft.web/sites', 'microsoft.app/containerapps', 'microsoft.servicebus/namespaces', 'microsoft.eventhub/namespaces',
    'microsoft.cache/redis'
)
$script:AzMonWebTypes = @('microsoft.web/sites', 'microsoft.app/containerapps')

function Find-AzMonCoverageGapFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ResourceRef,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $DiagnosticSetting,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight,
        [object] $HeartbeatResourceId = $null
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()

    $diagByTarget = @{}
    foreach ($d in $DiagnosticSetting) {
        $key = ([string]$d['ResourceId']).ToLowerInvariant()
        if (-not $diagByTarget.ContainsKey($key)) { $diagByTarget[$key] = [System.Collections.Generic.List[hashtable]]::new() }
        $diagByTarget[$key].Add($d)
    }

    $wsIds = @{}
    foreach ($w in $Workspace) { $wsIds[([string]$w['Id']).ToLowerInvariant()] = $true }

    $aiNamesByRg = @{}
    foreach ($ai in $AppInsight) {
        $rgKey = "$($ai['SubscriptionId'])/$($ai['ResourceGroup'])".ToLowerInvariant()
        if (-not $aiNamesByRg.ContainsKey($rgKey)) { $aiNamesByRg[$rgKey] = [System.Collections.Generic.List[string]]::new() }
        $aiNamesByRg[$rgKey].Add(([string]$ai['Name']).ToLowerInvariant())
    }

    $missingDiag = [System.Collections.Generic.List[hashtable]]::new()
    $notToWs = [System.Collections.Generic.List[hashtable]]::new()
    $webWithoutAi = [System.Collections.Generic.List[hashtable]]::new()
    $vmNoHeartbeat = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($r in $ResourceRef) {
        $rtype = ([string]$r['Type']).ToLowerInvariant()
        $rid = ([string]$r['Id']).ToLowerInvariant()
        $settings = if ($diagByTarget.ContainsKey($rid)) { $diagByTarget[$rid] } else { @() }

        if ($script:AzMonCriticalDiagnosticTypes -contains $rtype) {
            if (@($settings).Count -eq 0) {
                $missingDiag.Add($r)
            } else {
                $hasWsDest = $false
                foreach ($s in $settings) {
                    if ($s['WorkspaceId'] -and $wsIds.ContainsKey(([string]$s['WorkspaceId']).ToLowerInvariant())) { $hasWsDest = $true; break }
                }
                if (-not $hasWsDest) { $notToWs.Add($r) }
            }
        }

        if ($script:AzMonWebTypes -contains $rtype) {
            $rgKey = "$($r['SubscriptionId'])/$($r['ResourceGroup'])".ToLowerInvariant()
            $candidates = if ($aiNamesByRg.ContainsKey($rgKey)) { $aiNamesByRg[$rgKey] } else { @() }
            $rNameLower = ([string]$r['Name']).ToLowerInvariant()
            $match = $false
            foreach ($c in $candidates) {
                if ($rNameLower.Contains($c) -or $c.Contains($rNameLower)) { $match = $true; break }
            }
            if (-not $match) { $webWithoutAi.Add($r) }
        }

        if ($rtype -eq 'microsoft.compute/virtualmachines' -and $null -ne $HeartbeatResourceId) {
            if (-not $HeartbeatResourceId.Contains($rid)) { $vmNoHeartbeat.Add($r) }
        }
    }

    if ($missingDiag.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'coverage' -Severity 'high' `
            -Title "$($missingDiag.Count) critical resources have no diagnostic settings" `
            -Detail ('These resource types should always route platform logs to a Log Analytics workspace for ' +
                'security detection and operational alerting. Without diagnostic settings, you cannot alert on ' +
                'failures or investigate incidents.') `
            -ResourceIds @($missingDiag | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Deploy the included Bicep policy (bicep/diag-settings-baseline.bicep) or an Azure ' +
                'Policy assignment deployIfNotExists to auto-enable diagnostic settings on new resources.') `
            -Evidence @{ by_type = Get-AzMonCountByType -ResourceRef $missingDiag.ToArray() }))
    }

    if ($notToWs.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'coverage' -Severity 'medium' `
            -Title "$($notToWs.Count) resources have diagnostic settings but do NOT send to a Log Analytics workspace" `
            -Detail ('Storage-account-only or event-hub-only routing prevents correlation and alerting inside ' +
                'Azure Monitor. Add a workspace destination.') `
            -ResourceIds @($notToWs | ForEach-Object { $_['Id'] }) `
            -Recommendation 'Add a workspace destination to each diagnostic setting.' `
            -Evidence @{ by_type = Get-AzMonCountByType -ResourceRef $notToWs.ToArray() }))
    }

    if ($webWithoutAi.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'coverage' -Severity 'medium' `
            -Title "$($webWithoutAi.Count) web apps / container apps have no obvious Application Insights component" `
            -Detail ('Web workloads should adopt workspace-based Application Insights + OpenTelemetry ' +
                'auto-instrumentation to enable end-to-end distributed tracing, request-rate / latency / error ' +
                'metrics, and live metrics.') `
            -ResourceIds @($webWithoutAi | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Enable App Insights via the site config APPLICATIONINSIGHTS_CONNECTION_STRING and ' +
                'turn on the platform auto-instrumentation (App Service -> Application Insights blade -> Enable). ' +
                'For containerized apps, deploy the OTel collector sidecar and point it at the workspace-based ' +
                'AI resource.') `
            -Evidence @{ by_type = Get-AzMonCountByType -ResourceRef $webWithoutAi.ToArray() }))
    }

    if ($vmNoHeartbeat.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'coverage' -Severity 'high' `
            -Title "$($vmNoHeartbeat.Count) VMs have no Heartbeat in the last 2 days" `
            -Detail ('No monitoring agent (Azure Monitor Agent / legacy MMA) is reporting Heartbeat for these VMs ' +
                'in any connected workspace. They are effectively invisible to alerting, Update Management, and ' +
                'VM insights.') `
            -ResourceIds @($vmNoHeartbeat | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Install/repair the Azure Monitor Agent via a DCR association ' +
                '(Microsoft.Insights/dataCollectionRuleAssociations) pointed at a workspace, or investigate agent ' +
                'health if it was previously connected.')))
    }

    $classicAi = @($AppInsight | Where-Object { -not $_['WorkspaceResourceId'] })
    if ($classicAi.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'coverage' -Severity 'high' `
            -Title "$($classicAi.Count) classic Application Insights components (not workspace-based)" `
            -Detail ('Classic App Insights is deprecated (retired Feb 2024). It cannot be used with commitment ' +
                'tiers, unified alerting, or cross-workspace queries.') `
            -ResourceIds @($classicAi | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Migrate to workspace-based Application Insights. See ' +
                'https://learn.microsoft.com/azure/azure-monitor/app/convert-classic-resource') `
            -Evidence @{ components = @($classicAi | ForEach-Object { $_['Name'] }) }))
    }

    return $findings.ToArray()
}
