#requires -Version 7.0
# Cost optimization analyzer — ported 1:1 from analyzers/cost_optimization.py.

$script:AzMonBasicTierCandidates = @(
    'ContainerLogV2', 'ContainerLog', 'AppTraces', 'AzureDiagnostics', 'AppServiceHTTPLogs',
    'AppServiceConsoleLogs', 'AGWAccessLogs', 'APIMGatewayLogs', 'AKSAudit', 'AKSAuditAdmin'
)

# Auxiliary/Lake tier: even cheaper than Basic ("minimal" vs "reduced"
# ingestion cost per Microsoft's table-plan comparison), purpose-built for
# low-touch, compliance/audit-oriented raw event tables.
$script:AzMonAuxiliaryTierCandidates = @('SecurityEvent', 'WindowsEvent', 'Syslog')

function Find-AzMonCostOptimizationFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight,
        [array] $VmAgentExtension = @(),
        [array] $AdvisorRecommendation = @()
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($ws in $Workspace) {
        $byTable = $ws['IngestionByTable']
        if (-not $byTable -or $byTable.Count -eq 0) { continue }
        $candidates = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($table in $byTable.Keys) {
            $gb = [double]$byTable[$table]
            if (($script:AzMonBasicTierCandidates -contains $table) -and $gb -ge 5.0) {
                $candidates.Add(@{ Table = $table; Gb = [Math]::Round($gb, 2) })
            }
        }
        if ($candidates.Count -eq 0) { continue }
        $sumGb = ($candidates | Measure-Object -Property Gb -Sum).Sum
        $savings = [Math]::Round($sumGb * 2.0, 2)
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity $(if ($savings -gt 200) { 'high' } else { 'medium' }) `
            -Title "$($ws['Name']): $($candidates.Count) verbose tables good candidates for Basic Logs tier" `
            -Detail ('Basic Logs is ~5x cheaper than Analytics Logs but supports only KQL search (no alerts, no ' +
                'cross-table joins). Ideal for high-volume tables queried only for incident forensics.') `
            -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $savings `
            -Recommendation ('Convert candidate tables to Basic tier via az monitor log-analytics workspace table ' +
                'update --plan Basic. Verify no alert rules depend on the table first.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/logs-table-plans' `
            -Evidence @{ tables = @($candidates | ForEach-Object { , @($_.Table, $_.Gb) }); gb_30d_total = [Math]::Round($sumGb, 2) }))
    }

    foreach ($ws in $Workspace) {
        $byTable = $ws['IngestionByTable']
        if (-not $byTable -or $byTable.Count -eq 0) { continue }
        $auxCandidates = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($table in $byTable.Keys) {
            $gb = [double]$byTable[$table]
            if (($script:AzMonAuxiliaryTierCandidates -contains $table) -and $gb -ge 5.0) {
                $auxCandidates.Add(@{ Table = $table; Gb = [Math]::Round($gb, 2) })
            }
        }
        if ($auxCandidates.Count -eq 0) { continue }
        $auxSumGb = ($auxCandidates | Measure-Object -Property Gb -Sum).Sum
        $auxSavings = [Math]::Round($auxSumGb * 2.3, 2)
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity $(if ($auxSavings -gt 200) { 'high' } else { 'medium' }) `
            -Title "$($ws['Name']): $($auxCandidates.Count) audit/security tables good candidates for Auxiliary (Lake) tier" `
            -Detail ('Auxiliary tier ingestion is even cheaper than Basic - built for low-touch, verbose audit ' +
                'and compliance data like security/OS event logs that are rarely queried interactively and can ' +
                'tolerate slower query performance.') `
            -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $auxSavings `
            -Recommendation ('Convert candidate tables to the Auxiliary (Lake) plan. Confirm no alert rules or ' +
                'dashboards need real-time query performance on these tables first, since Auxiliary does not ' +
                'support alerts.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/data-platform-logs#table-plans' `
            -Evidence @{ tables = @($auxCandidates | ForEach-Object { , @($_.Table, $_.Gb) }); gb_30d_total = [Math]::Round($auxSumGb, 2) }))
    }

    foreach ($ws in $Workspace) {
        $gb30 = [double]($ws['IngestionGb30d'] ?? 0)
        if ($gb30 -gt 100 -and -not $ws['DailyQuotaGb']) {
            $findings.Add((New-AzMonFinding -Category 'cost' -Severity 'medium' `
                -Title "$($ws['Name']): no daily quota (dailyQuotaGb) set" `
                -Detail ('A daily cap protects against runaway ingestion from misconfigured agents or noisy ' +
                    'applications, preventing surprise bills.') `
                -ResourceIds @($ws['Id']) `
                -Recommendation ('Set dailyQuotaGb to ~120% of the P95 daily ingestion. Log an alert when the cap ' +
                    'is hit so the team is notified before data is dropped.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/daily-cap' `
                -Evidence @{ ingestion_gb_30d = $gb30 }))
        }
    }

    foreach ($ai in $AppInsight) {
        $pct = $ai['SamplingPercentage']
        if ($null -eq $pct -or $pct -ge 100) {
            $findings.Add((New-AzMonFinding -Category 'cost' -Severity 'low' `
                -Title "App Insights '$($ai['Name'])': sampling not configured" `
                -Detail ('Adaptive sampling reduces telemetry volume by 50-90% while preserving trends and error ' +
                    'visibility.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Enable adaptive sampling in the SDK (default 5 items/sec target) or set an ' +
                    'ingestion sampling percentage on the AI resource.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/app/sampling'))
        }
    }

    foreach ($ai in $AppInsight) {
        $cap = $ai['DailyCapGb']
        if ($null -eq $cap -or $cap -le 0) {
            $findings.Add((New-AzMonFinding -Category 'cost' -Severity 'medium' `
                -Title "App Insights '$($ai['Name'])': no daily volume cap configured" `
                -Detail ('Without a daily cap, a stack trace loop or a debug-verbose deployment can 10x monthly ' +
                    'cost overnight before anyone notices. WAF recommends a cap sized ~120% of the P95 daily volume.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Set the daily volume cap via az monitor app-insights component update ' +
                    '--daily-cap <gb> and enable the "90% of cap reached" alert.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/daily-cap'))
        }
    }

    $securityTables = @('SecurityEvent', 'SecurityAlert', 'SecurityIncident', 'SentinelHealth')
    $opsTables = @('AppTraces', 'ContainerLogV2', 'ContainerLog', 'AppServiceHTTPLogs', 'AppServiceConsoleLogs', 'AzureDiagnostics')
    foreach ($ws in $Workspace) {
        $byTable = $ws['IngestionByTable']
        if (-not $byTable -or $byTable.Count -eq 0) { continue }
        $tables = @($byTable.Keys)
        $hasSecurity = @($tables | Where-Object { $securityTables -contains $_ }).Count -gt 0
        $overlappingOps = @($tables | Where-Object { $opsTables -contains $_ })
        if (-not $hasSecurity -or $overlappingOps.Count -eq 0) { continue }
        $opsGb = [Math]::Round((($overlappingOps | ForEach-Object { [double]$byTable[$_] } | Measure-Object -Sum).Sum), 2)
        if ($opsGb -lt 50) { continue }
        $savings = [Math]::Round($opsGb * 2.30, 2)
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity $(if ($savings -gt 500) { 'high' } else { 'medium' }) `
            -Title "$($ws['Name']): Sentinel-enabled workspace commingled with $([Math]::Round($opsGb,0)) GB/30d of operational data" `
            -Detail ('When Microsoft Sentinel is enabled on a workspace, all ingested data incurs the Sentinel ' +
                'surcharge - even AppTraces / ContainerLogV2 the SOC never queries. Splitting security and ' +
                'operational data into separate workspaces removes that surcharge from ops volume.') `
            -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $savings `
            -Recommendation ('Create a dedicated ops workspace in the same region, redirect diagnostic settings ' +
                'and DCRs for operational sources, then keep Sentinel scoped to security telemetry only.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/cost-logs#workspaces-with-microsoft-sentinel' `
            -Evidence @{
                ops_tables              = @($overlappingOps | Sort-Object)
                ops_gb_30d              = $opsGb
                security_tables_present = @($tables | Where-Object { $securityTables -contains $_ } | Sort-Object)
            }))
    }

    foreach ($ws in $Workspace) {
        $byTable = $ws['IngestionByTable']
        if (-not $byTable -or $byTable.Count -eq 0) { continue }
        if (([int]($ws['RetentionDays'] ?? 0)) -le 30) { continue }
        $summaryCandidates = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($t in $byTable.Keys) {
            $gb = [double]$byTable[$t]
            if ($gb -ge 100.0) { $summaryCandidates.Add(@{ Table = $t; Gb = [Math]::Round($gb, 2) }) }
        }
        if ($summaryCandidates.Count -eq 0) { continue }
        $volumeSaved = (($summaryCandidates | Measure-Object -Property Gb -Sum).Sum) * 0.9
        $savings = [Math]::Round($volumeSaved * 2.30, 2)
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity 'medium' `
            -Title "$($ws['Name']): $($summaryCandidates.Count) tables >100 GB/30d - candidates for Summary Rules" `
            -Detail ('Summary Rules run a scheduled KQL query and write compact rollup rows to a destination ' +
                'table. For dashboards and long-retention trends you can drop the raw table to Basic (or delete ' +
                'it) while keeping the aggregates hot.') `
            -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $savings `
            -Recommendation ('Author Summary Rules for the noisiest tables (e.g., 5-min aggregates of ' +
                'AppServiceHTTPLogs per status/route). Retain rollups in Analytics tier and raw data in Basic + ' +
                'storage archive.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/summary-rules' `
            -Evidence @{ tables = @($summaryCandidates | ForEach-Object { , @($_.Table, $_.Gb) }) }))
    }

    # ---- Retired Log Analytics agent (MMA/OMS) still installed ---------
    # Retired 2024-08-31; cloud ingestion for it is being shut down and can
    # stop at any time without notice after 2026-03-02. Flag any VM that has
    # the legacy extension without also having Azure Monitor Agent installed
    # (both present briefly during a migration is expected, not a finding).
    $agentKindsByVm = @{}
    foreach ($ext in $VmAgentExtension) {
        $vid = ([string]$ext['VmId']).ToLowerInvariant()
        if (-not $agentKindsByVm.ContainsKey($vid)) { $agentKindsByVm[$vid] = [System.Collections.Generic.HashSet[string]]::new() }
        [void]$agentKindsByVm[$vid].Add($ext['AgentKind'])
    }
    $legacyOnlyVmIds = @($agentKindsByVm.Keys | Where-Object { $agentKindsByVm[$_].Contains('legacy') -and -not $agentKindsByVm[$_].Contains('ama') })
    if ($legacyOnlyVmIds.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity 'critical' `
            -Title "$($legacyOnlyVmIds.Count) VMs still run the retired Log Analytics agent (MMA/OMS) with no Azure Monitor Agent" `
            -Detail ('The Log Analytics agent (MMA/OMS) was retired on 2024-08-31. Microsoft no longer supports it, ' +
                'it receives no new distros/service packs, and cloud ingestion for it is being shut down - after ' +
                '2026-03-02, data upload from this agent can stop at any time without further notice. Every VM ' +
                'still on it alone is one silent outage away from a monitoring blind spot, and running two agent ' +
                'generations side by side (during migration) doubles ingestion cost until the legacy agent is removed.') `
            -ResourceIds $legacyOnlyVmIds `
            -Recommendation ('Migrate to Azure Monitor Agent: use the Migration Helper workbook to inventory ' +
                'remaining agents, generate data collection rules with the DCR Config Generator, deploy AMA via ' +
                'Azure Policy, validate ingestion, then remove the legacy agent with the MMA Discovery and Removal ' +
                'tool. See https://learn.microsoft.com/azure/azure-monitor/agents/azure-monitor-agent-migration') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/agents/azure-monitor-agent-migration' `
            -Evidence @{ vm_count = $legacyOnlyVmIds.Count }))
    }

    # ---- Standalone workspace - commitment tier opportunity -------------
    # The consolidation analyzer only prices a commitment tier for groups of
    # 2+ workspaces sharing a region/environment; a single large workspace
    # with no consolidation peers would otherwise never get this check.
    $groupSizeByKey = @{}
    foreach ($ws in $Workspace) {
        $key = "$(([string]$ws['Location']).ToLowerInvariant())|$(Get-AzMonEnvironmentTag -Workspace $ws)"
        $groupSizeByKey[$key] = ($groupSizeByKey[$key] ?? 0) + 1
    }
    foreach ($ws in $Workspace) {
        $key = "$(([string]$ws['Location']).ToLowerInvariant())|$(Get-AzMonEnvironmentTag -Workspace $ws)"
        if ($groupSizeByKey[$key] -gt 1) { continue }
        $dailyGb = [double]($ws['IngestionGb30d'] ?? 0) / 30.0
        if ($dailyGb -lt 100) { continue }
        $paygMo = $dailyGb * 30.0 * $script:AzMonPaygPricePerGb
        $best = Get-AzMonBestCommitment -DailyGb $dailyGb
        $savings = [Math]::Round([Math]::Max($paygMo - $best.MonthlyCost, 0), 2)
        if ($savings -lt 100) { continue }
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity $(if ($savings -gt 500) { 'high' } else { 'medium' }) `
            -Title "$($ws['Name']): $([Math]::Round($dailyGb,0)) GB/day on pay-as-you-go - commitment tier available" `
            -Detail ('This workspace has no consolidation peers in its region/environment, so it would not ' +
                'otherwise be flagged for a pricing-tier change. At its current volume, a commitment tier reduces ' +
                "list price from $(Format-AzMonUsd $paygMo)/mo (PAYG) to $(Format-AzMonUsd $best.MonthlyCost)/mo.") `
            -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $savings `
            -Recommendation ('Configure a commitment tier sized to this daily volume via az monitor log-analytics ' +
                'workspace update --sku CapacityReservation --capacity-reservation-level <tier>. Use the Azure ' +
                'Pricing Calculator to pick the closest tier at or below current daily GB, then re-check after ' +
                '30 days of billing to confirm the fit.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/cost-logs#commitment-tiers' `
            -Evidence @{ daily_gb = [Math]::Round($dailyGb, 2); price_per_gb = $best.PricePerGb }))
    }

    # ---- Azure Advisor cost recommendations (Monitor-scoped) ------------
    # Surfaces the actual current Advisor recommendations rather than just
    # checking whether an alert exists for future ones.
    foreach ($rec in $AdvisorRecommendation) {
        $impact = ([string]$rec['Impact'])
        $sev = switch ($impact) { 'High' { 'high' } 'Medium' { 'medium' } default { 'low' } }
        $benefitSuffix = if ($rec['PotentialBenefit']) { " Potential benefit: $($rec['PotentialBenefit'])" } else { '' }
        $recommendation = if ($rec['LearnMoreLink']) { "See $($rec['LearnMoreLink'])" } else { 'Review this recommendation in Azure Advisor.' }
        $findings.Add((New-AzMonFinding -Category 'cost' -Severity $sev `
            -Title "Azure Advisor: $($rec['Problem'])" `
            -Detail "$($rec['Solution'])$benefitSuffix" `
            -ResourceIds @($rec['ResourceId']) `
            -Recommendation $recommendation `
            -LearnMoreLink $rec['LearnMoreLink'] `
            -Evidence @{ source = 'Azure Advisor'; impacted_field = $rec['ImpactedField'] }))
    }

    return $findings.ToArray()
}
