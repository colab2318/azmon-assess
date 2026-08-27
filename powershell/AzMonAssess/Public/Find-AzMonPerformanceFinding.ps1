#requires -Version 7.0
# WAF Performance Efficiency analyzer — ported 1:1 from analyzers/performance.py.

$script:AzMonDedicatedClusterThresholdGb30d = 15000.0
$script:AzMonSearchJobCandidates = @(
    'AzureDiagnostics', 'AppTraces', 'ContainerLogV2', 'ContainerLog', 'AppServiceHTTPLogs', 'AKSAudit',
    'AKSAuditAdmin', 'AGWAccessLogs', 'APIMGatewayLogs', 'SecurityEvent', 'WindowsEvent', 'Syslog'
)

function Find-AzMonPerformanceFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()
    $compliance = [System.Collections.Generic.List[hashtable]]::new()
    $wsById = @{}
    foreach ($ws in $Workspace) { $wsById[([string]$ws['Id']).ToLowerInvariant()] = $ws }

    foreach ($ai in $AppInsight) {
        if (-not $ai['WorkspaceResourceId']) { continue }
        $ws = $wsById[([string]$ai['WorkspaceResourceId']).ToLowerInvariant()]
        if (-not $ws) { continue }
        if ($ai['Location'] -and $ws['Location'] -and (([string]$ai['Location']).ToLowerInvariant() -ne ([string]$ws['Location']).ToLowerInvariant())) {
            $findings.Add((New-AzMonFinding -Category 'performance' -Severity 'medium' `
                -Title "App Insights '$($ai['Name'])' region '$($ai['Location'])' differs from workspace region '$($ws['Location'])'" `
                -CheckId 'performance.region-mismatch' `
                -Detail ('Cross-region telemetry adds ingest latency (P95 typically +100-300 ms), cross-region ' +
                    'egress cost, and creates a second regional failure domain for the same telemetry pipeline.') `
                -ResourceIds @($ai['Id'], $ws['Id']) `
                -Recommendation ('Recreate App Insights in the same region as the backing workspace (ideally ' +
                    'colocated with the monitored workload), migrate connection strings on next release, then ' +
                    'retire the mis-regioned resource.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/workspace-design' `
                -Evidence @{ ai_location = $ai['Location']; workspace_location = $ws['Location'] }))
        } elseif ($ai['Location'] -and $ws['Location']) {
            $compliance.Add((New-AzMonComplianceItem -Category 'performance' -CheckId 'performance.region-mismatch' `
                -Title "App Insights '$($ai['Name'])' is colocated with workspace '$($ws['Name'])'" `
                -ResourceIds @($ai['Id'], $ws['Id'])))
        }
    }

    foreach ($ws in $Workspace) {
        $gb30 = [double]($ws['IngestionGb30d'] ?? 0)
        if ($gb30 -ge $script:AzMonDedicatedClusterThresholdGb30d -and -not $ws['ClusterResourceId']) {
            $findings.Add((New-AzMonFinding -Category 'performance' -Severity 'high' `
                -Title "$($ws['Name']): $([Math]::Round($gb30,0)) GB/30d - candidate for dedicated cluster tier" `
                -CheckId 'performance.dedicated-cluster-candidate' `
                -Detail ('Workspaces above ~500 GB/day benefit from a dedicated cluster: guaranteed query ' +
                    'capacity, customer-managed keys (CMK), cross-workspace cost pooling, and up to ~25% ' +
                    'commit-tier discount.') `
                -ResourceIds @($ws['Id']) `
                -Recommendation ('Create a dedicated cluster with a 500 GB/day commit tier and link this ' +
                    'workspace (plus any peer workspaces in the same region). Model with Azure Pricing Calculator ' +
                    'vs. current pay-as-you-go cost first.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/logs-dedicated-clusters' `
                -Evidence @{ ingestion_gb_30d = $gb30 }))
        } elseif ($gb30 -ge $script:AzMonDedicatedClusterThresholdGb30d -and $ws['ClusterResourceId']) {
            $compliance.Add((New-AzMonComplianceItem -Category 'performance' -CheckId 'performance.dedicated-cluster-candidate' `
                -Title "$($ws['Name']): $([Math]::Round($gb30,0)) GB/30d - already on a dedicated cluster" `
                -ResourceIds @($ws['Id'])))
        }
    }

    foreach ($ws in $Workspace) {
        $byTable = $ws['IngestionByTable']
        if (-not $byTable -or $byTable.Count -eq 0) { continue }
        if (([int]($ws['RetentionDays'] ?? 0)) -le 30) { continue }
        $candidates = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($t in $byTable.Keys) {
            $gb = [double]$byTable[$t]
            if (($script:AzMonSearchJobCandidates -contains $t) -and $gb -ge 100.0) {
                $candidates.Add(@{ Table = $t; Gb = [Math]::Round($gb, 2) })
            }
        }
        if ($candidates.Count -eq 0) {
            $compliance.Add((New-AzMonComplianceItem -Category 'performance' -CheckId 'performance.search-jobs-candidate' `
                -Title "$($ws['Name']): no high-volume tables need Search Jobs / Basic tier despite $($ws['RetentionDays'])d retention" `
                -ResourceIds @($ws['Id'])))
            continue
        }
        $findings.Add((New-AzMonFinding -Category 'performance' -Severity 'medium' `
            -Title "$($ws['Name']): $($candidates.Count) high-volume tables with >30d retention - consider Search Jobs / Basic tier + long-term archive" `
            -CheckId 'performance.search-jobs-candidate' `
            -Detail ('Keeping these tables at Analytics-tier retention is expensive and slows interactive ' +
                'queries. Search Jobs let you run one-shot investigations against months of cold data at ~5% of ' +
                'the ingestion cost, while Basic tier avoids the cross-query cost hit.') `
            -ResourceIds @($ws['Id']) `
            -Recommendation ('Move these tables to Basic tier with retention capped at 30d, then attach a ' +
                'long-term storage account. Rehydrate via Search Jobs on demand.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/search-jobs' `
            -Evidence @{ tables = @($candidates | ForEach-Object { , @($_.Table, $_.Gb) }); retention_days = $ws['RetentionDays'] }))
    }

    return @{ Findings = $findings.ToArray(); ComplianceItems = $compliance.ToArray() }
}
