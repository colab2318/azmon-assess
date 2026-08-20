#requires -Version 7.0
# Alert quality analyzer — ported 1:1 from analyzers/alert_quality.py.

$script:AzMonExpectedAlertsByType = @{
    'microsoft.compute/virtualmachines'          = @('availability', 'cpu', 'memory', 'disk')
    'microsoft.web/sites'                        = @('http_5xx', 'response_time', 'availability')
    'microsoft.app/containerapps'                = @('restart_count', 'http_5xx', 'replica_count')
    'microsoft.sql/servers/databases'            = @('dtu_high', 'connection_failed', 'deadlocks')
    'microsoft.dbforpostgresql/flexibleservers'  = @('cpu', 'connections', 'storage')
    'microsoft.containerservice/managedclusters' = @('node_not_ready', 'pod_restart', 'api_server_latency')
    'microsoft.storage/storageaccounts'          = @('availability', 'throttling')
    'microsoft.keyvault/vaults'                  = @('service_api_result_failure')
    'microsoft.servicebus/namespaces'            = @('dead_letter_messages', 'server_errors')
    'microsoft.eventhub/namespaces'              = @('throttled_requests', 'user_errors')
}

function Find-AzMonAlertQualityFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AlertRule,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ActionGroup,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ResourceRef
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()

    $disabled = @($AlertRule | Where-Object { -not $_['Enabled'] })
    if ($disabled.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'medium' `
            -Title "$($disabled.Count) disabled alert rules" `
            -Detail ('Disabled rules are common sources of drift - either delete them or re-enable and tune. ' +
                'Disabled rules are a frequent cause of missing critical alerts.') `
            -ResourceIds @($disabled | ForEach-Object { $_['Id'] }) `
            -Recommendation 'Review each disabled rule; either delete or re-enable with tuned thresholds.' `
            -Evidence @{ names = @($disabled | Select-Object -First 50 | ForEach-Object { $_['Name'] }) }))
    }

    $silent = @($AlertRule | Where-Object { $_['Enabled'] -and @($_['ActionGroupIds']).Count -eq 0 })
    if ($silent.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'high' `
            -Title "$($silent.Count) enabled alert rules have NO action group" `
            -Detail 'These rules fire but notify nobody - a common cause of missed critical incidents.' `
            -ResourceIds @($silent | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Attach a standard action group (see bicep/action-group.bicep). For business-critical ' +
                'rules, route to the paging channel; for hygiene rules, route to a low-priority email or Teams webhook.') `
            -Evidence @{ names = @($silent | Select-Object -First 50 | ForEach-Object { $_['Name'] }) }))
    }

    $orphans = @($ActionGroup | Where-Object { $_['UsedByRules'] -eq 0 })
    if ($orphans.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'low' `
            -Title "$($orphans.Count) action groups are not referenced by any rule" `
            -Detail 'Unused action groups add clutter and confuse on-call rotation.' `
            -ResourceIds @($orphans | ForEach-Object { $_['Id'] }) `
            -Recommendation 'Delete or repurpose these action groups.' `
            -Evidence @{ names = @($orphans | ForEach-Object { $_['Name'] }) }))
    }

    $unsevered = @($AlertRule | Where-Object { $_['Enabled'] -and $null -eq $_['Severity'] })
    if ($unsevered.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'low' `
            -Title "$($unsevered.Count) rules do not set an explicit severity" `
            -Detail 'Without severity, correlation and routing become inconsistent.' `
            -ResourceIds @($unsevered | Select-Object -First 100 | ForEach-Object { $_['Id'] }) `
            -Recommendation 'Set severity 0-4 explicitly on all rules; standardize per runbook.'))
    }

    # ---- Noisy rules: high fire-rate over the last 30 days -------------
    # FireCount30d is matched by rule *name* against fired alert instances
    # (Resource Graph AlertsManagementResources) — best-effort, since
    # same-named rules across resource groups can't be disambiguated.
    $noisyThreshold = 50
    $noisy = @($AlertRule | Where-Object { ($_['FireCount30d'] ?? 0) -ge $noisyThreshold } | Sort-Object -Property { - [int]($_['FireCount30d'] ?? 0) })
    if ($noisy.Count -gt 0) {
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'medium' `
            -Title "$($noisy.Count) rules fired $noisyThreshold+ times in the last 30 days (noisy)" `
            -Detail ('High-frequency rules dominate on-call attention and train responders to ignore alerts. ' +
                'Fire counts are matched to rules by name via Resource Graph''s AlertsManagementResources table.') `
            -ResourceIds @($noisy | Select-Object -First 50 | ForEach-Object { $_['Id'] }) `
            -Recommendation ('Review the noisiest rules first: raise the threshold, add an aggregation window, ' +
                'suppress during known maintenance, or replace with a smarter dynamic-threshold condition.') `
            -Evidence @{ top_noisy_rules = @($noisy | Select-Object -First 10 | ForEach-Object { @{ name = $_['Name']; fire_count_30d = $_['FireCount30d'] } }) }))
    }

    $typesWithRules = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in $AlertRule) {
        foreach ($s in @($r['Scopes'])) {
            $sLower = ([string]$s).ToLowerInvariant()
            $idx = $sLower.IndexOf('/providers/')
            if ($idx -ge 0) {
                $rest = $sLower.Substring($idx + '/providers/'.Length)
                $segments = $rest -split '/'
                if ($segments.Count -ge 2) { [void]$typesWithRules.Add("$($segments[0])/$($segments[1])") }
            }
        }
    }

    $resourcesByType = @{}
    foreach ($r in $ResourceRef) {
        $t = ([string]$r['Type']).ToLowerInvariant()
        $resourcesByType[$t] = ($resourcesByType[$t] ?? 0) + 1
    }

    $uncovered = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($rtype in $resourcesByType.Keys) {
        if ($script:AzMonExpectedAlertsByType.Contains($rtype)) {
            $covered = $false
            foreach ($t in $typesWithRules) { if ($t.Contains($rtype)) { $covered = $true; break } }
            if (-not $covered) { $uncovered.Add([pscustomobject]@{ Type = $rtype; Count = $resourcesByType[$rtype] }) }
        }
    }

    if ($uncovered.Count -gt 0) {
        $totalUncovered = ($uncovered | Measure-Object -Property Count -Sum).Sum
        $findings.Add((New-AzMonFinding -Category 'alerting' -Severity 'high' `
            -Title "$totalUncovered resources across $($uncovered.Count) critical types have NO alert rule scoped to them" `
            -Detail ('Recommended baseline alerts (availability, error-rate, throttling) are missing. Combined ' +
                "with the 'silent rules' finding, this explains under-alerting.") `
            -Recommendation ('Deploy the included Bicep alert-baseline (bicep/alert-baseline.bicep) which creates ' +
                'recommended metric alerts per resource type. Use Azure Policy deployIfNotExists to auto-provision ' +
                'alerts on new resources.') `
            -Evidence @{ uncovered = @($uncovered | ForEach-Object { @{ type = $_.Type; count = $_.Count; expected_alerts = $script:AzMonExpectedAlertsByType[$_.Type] } }) }))
    }

    return $findings.ToArray()
}
