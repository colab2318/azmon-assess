#requires -Version 7.0
# Cost optimization analyzer — ported 1:1 from analyzers/cost_optimization.py.

$script:AzMonBasicTierCandidates = @(
    'ContainerLogV2', 'ContainerLog', 'AppTraces', 'AzureDiagnostics', 'AppServiceHTTPLogs',
    'AppServiceConsoleLogs', 'AGWAccessLogs', 'APIMGatewayLogs', 'AKSAudit', 'AKSAuditAdmin'
)

function Find-AzMonCostOptimizationFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight
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
            -Evidence @{ tables = @($candidates | ForEach-Object { , @($_.Table, $_.Gb) }); gb_30d_total = [Math]::Round($sumGb, 2) }))
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
                    'ingestion sampling percentage on the AI resource.')))
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
                    '--daily-cap <gb> and enable the "90% of cap reached" alert.')))
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
            -Evidence @{ tables = @($summaryCandidates | ForEach-Object { , @($_.Table, $_.Gb) }) }))
    }

    return $findings.ToArray()
}
