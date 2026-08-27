#requires -Version 7.0
# Log Analytics workspace consolidation planner — ported 1:1 from
# analyzers/consolidation.py. Pricing is directional (validate against the
# customer's EA/MCA rate card before quoting savings to leadership).

function Find-AzMonConsolidationFinding {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace)

    $findings = [System.Collections.Generic.List[hashtable]]::new()
    if (-not $Workspace -or $Workspace.Count -eq 0) { return @{ Findings = @(); ComplianceItems = @() } }

    $groups = @{}
    foreach ($ws in $Workspace) {
        $envName = Get-AzMonEnvironmentTag -Workspace $ws
        $key = "$(([string]$ws['Location']).ToLowerInvariant())|$envName"
        if (-not $groups.Contains($key)) { $groups[$key] = [System.Collections.Generic.List[hashtable]]::new() }
        $groups[$key].Add($ws)
    }

    $total = $Workspace.Count
    $regionCount = @($Workspace | ForEach-Object { $_['Location'] } | Sort-Object -Unique).Count
    $groupSummary = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($key in $groups.Keys) {
        $parts = $key -split '\|', 2
        $groupSummary.Add(@{ region = $parts[0]; environment = $parts[1]; workspaces = $groups[$key].Count })
    }

    $findings.Add((New-AzMonFinding -Category 'consolidation' `
        -Severity $(if ($total -ge 10) { 'high' } else { 'medium' }) `
        -Title "$total Log Analytics workspaces detected — evaluate consolidation" `
        -CheckId 'consolidation.workspace-count' `
        -Detail ("Detected $total workspaces across $regionCount regions. Consolidating to per-region-per-environment " +
            '(prod / non-prod / core) typically unlocks commitment-tier discounts of 20-40% and simplifies ' +
            'cross-service correlation for App Insights and infrastructure telemetry.') `
        -ResourceIds @($Workspace | ForEach-Object { $_['Id'] }) `
        -Recommendation ('Adopt a hub-and-spoke workspace design: 1 workspace per region per sensitivity tier ' +
            '(prod, non-prod, security/core). Migrate using diagnostic-setting re-pointing. See docs: ' +
            'https://learn.microsoft.com/azure/azure-monitor/logs/workspace-design') `
        -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/workspace-design' `
        -Evidence @{ group_summary = $groupSummary.ToArray() }))

    foreach ($key in $groups.Keys) {
        $wsList = $groups[$key]
        if ($wsList.Count -lt 2) { continue }
        $parts = $key -split '\|', 2
        $region = $parts[0]; $envName = $parts[1]
        $sumGb = 0.0
        foreach ($w in $wsList) { $sumGb += [double]($w['IngestionGb30d'] ?? 0) }
        $dailyGb = $sumGb / 30.0
        if ($dailyGb -le 0) { continue }
        $paygMonthly = $dailyGb * 30.0 * $script:AzMonPaygPricePerGb
        $best = Get-AzMonBestCommitment -DailyGb $dailyGb
        $savings = [Math]::Round([Math]::Max($paygMonthly - $best.MonthlyCost, 0), 2)
        $dailyGbRounded = [Math]::Round($dailyGb, 1)

        $findings.Add((New-AzMonFinding -Category 'consolidation' `
            -Severity $(if ($savings -gt 500) { 'high' } else { 'medium' }) `
            -Title "Consolidation candidate — $($wsList.Count) workspaces in $region / env=$envName" `
            -CheckId 'consolidation.candidate-group' `
            -Detail ("Combined ingestion approx $dailyGbRounded GB/day. Merging into a single workspace and " +
                "applying the best-fit commitment tier reduces list price from $(Format-AzMonUsd $paygMonthly)/mo " +
                "(PAYG) to $(Format-AzMonUsd $best.MonthlyCost)/mo.") `
            -ResourceIds @($wsList | ForEach-Object { $_['Id'] }) `
            -EstimatedMonthlySavingsUsd $savings `
            -Recommendation ('Create a single target workspace, re-point diagnostic settings, migrate legacy ' +
                'alerts to scheduled-query rules against the target workspace, then decommission source ' +
                'workspaces after a 30-day parallel-run.') `
            -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/workspace-design' `
            -Evidence @{
                daily_gb   = [Math]::Round($dailyGb, 2)
                workspaces = @($wsList | ForEach-Object { @{ name = $_['Name']; id = $_['Id']; gb_30d = $_['IngestionGb30d'] } })
            }))
    }

    foreach ($ws in $Workspace) {
        $retention = $ws['RetentionDays']
        $gb30 = [double]($ws['IngestionGb30d'] ?? 0)
        if ($retention -and $retention -gt 90 -and $gb30 -gt 30) {
            $extraDays = $retention - 30
            $extraCost = [Math]::Round($gb30 * 0.10 * ($extraDays / 30.0), 2)
            $findings.Add((New-AzMonFinding -Category 'retention' -Severity 'medium' `
                -Title "Workspace $($ws['Name']): interactive retention ${retention}d" `
                -CheckId 'retention.interactive-retention-high' `
                -Detail ('Interactive retention above 30 days incurs additional cost. If long-term retention is ' +
                    'satisfied by another system (SIEM, cold storage, MSSP), reduce interactive retention.') `
                -ResourceIds @($ws['Id']) -EstimatedMonthlySavingsUsd $extraCost `
                -Recommendation ('Reduce interactive retention to 30 days; enable Archive tier for tables that ' +
                    'still require long-term access (~10x cheaper than interactive).') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/data-retention-configure' `
                -Evidence @{ retention_days = $retention; gb_30d = $ws['IngestionGb30d'] }))
        }
    }

    return @{ Findings = $findings.ToArray(); ComplianceItems = @() }
}
