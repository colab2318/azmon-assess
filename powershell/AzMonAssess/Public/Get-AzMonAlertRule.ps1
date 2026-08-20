#requires -Version 7.0

function ConvertTo-AzMonActionGroupId {
    <#
    .SYNOPSIS
        Normalizes the various shapes of an alert rule's 'actions' JSON
        (metricAlerts / scheduledQueryRules / activityLogAlerts) to a flat
        list of action group resource IDs.
    #>
    [CmdletBinding()]
    param($ActionsField)
    $ids = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $ActionsField) { return @() }

    if ($ActionsField -is [System.Collections.IEnumerable] -and -not ($ActionsField -is [string]) -and -not ($ActionsField -is [hashtable])) {
        foreach ($a in $ActionsField) {
            $ah = ConvertTo-AzMonHashtable $a
            if ($ah -is [hashtable] -and $ah['actionGroupId']) { $ids.Add([string]$ah['actionGroupId']) }
        }
        return $ids.ToArray()
    }

    $hash = ConvertTo-AzMonHashtable $ActionsField
    if ($hash -is [hashtable] -and $hash.ContainsKey('actionGroups')) {
        foreach ($a in @($hash['actionGroups'])) {
            if ($a -is [string]) { $ids.Add($a) }
            elseif ($a -is [hashtable] -and $a['actionGroupId']) { $ids.Add([string]$a['actionGroupId']) }
        }
    }
    return $ids.ToArray()
}

function Get-AzMonAlertRule {
    <#
    .SYNOPSIS
        Collects metric, log (scheduled query) and activity-log alert rules.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rules = [System.Collections.Generic.List[hashtable]]::new()
    $queries = @(
        @{ Kind = 'metric'; Query = $script:AzMonQMetricAlerts }
        @{ Kind = 'log'; Query = $script:AzMonQLogAlerts }
        @{ Kind = 'activityLog'; Query = $script:AzMonQActivityAlerts }
    )
    foreach ($q in $queries) {
        $rows = Invoke-AzMonGraphQuery -Query $q.Query -SubscriptionId $SubscriptionId
        foreach ($r in $rows) {
            $scopes = ConvertTo-AzMonHashtable $r.scopes
            if ($scopes -isnot [array]) { $scopes = if ($scopes) { @($scopes) } else { @() } }
            $evalFreq = if ($q.Kind -eq 'log') { ConvertFrom-AzMonIso8601Duration -Duration $r.evaluationFrequency } else { $null }
            $rules.Add((New-AzMonAlertRule -Id $r.id -Name $r.name -SubscriptionId $r.subscriptionId -ResourceGroup $r.resourceGroup `
                -AlertKind $q.Kind -Enabled ([bool]$r.enabled) -Severity $r.severity -Scopes $scopes `
                -ActionGroupIds (ConvertTo-AzMonActionGroupId $r.actions) -Description $r.description -EvaluationFrequencyMinutes $evalFreq))
        }
    }
    return $rules.ToArray()
}

function Add-AzMonAlertFireRate {
    <#
    .SYNOPSIS
        Populates FireCount30d on alert rules by matching fired alert
        instances (Resource Graph AlertsManagementResources) back to rules
        by name (best-effort — the essentials payload does not expose the
        source rule's resource id). Mirrors kql/alert_fire_rate.kql.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $SubscriptionId,
        [Parameter(Mandatory)] [array] $AlertRule
    )
    if (-not $AlertRule -or $AlertRule.Count -eq 0) { return $AlertRule }

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQAlertFireRate -SubscriptionId $SubscriptionId
    $countsByName = @{}
    foreach ($r in $rows) {
        $name = ([string]$r.ruleName).ToLowerInvariant()
        if ($name) { $countsByName[$name] = [int]($r.FireCount ?? 0) }
    }
    foreach ($rule in $AlertRule) {
        $key = ([string]$rule['Name']).ToLowerInvariant()
        if ($countsByName.ContainsKey($key)) {
            $rule['FireCount30d'] = $countsByName[$key]
        }
    }
    return $AlertRule
}

function Get-AzMonActionGroup {
    <#
    .SYNOPSIS
        Collects action groups and cross-references usage against alert rules.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $SubscriptionId,
        [array] $AlertRule = @()
    )

    $usage = @{}
    foreach ($rule in $AlertRule) {
        foreach ($agId in @($rule['ActionGroupIds'])) {
            $key = ([string]$agId).ToLowerInvariant()
            if (-not $key) { continue }
            $usage[$key] = ($usage[$key] ?? 0) + 1
        }
    }

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQActionGroups -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            $used = $usage[$r.id.ToLowerInvariant()] ?? 0
            New-AzMonActionGroup -Id $r.id -Name $r.name -SubscriptionId $r.subscriptionId -ResourceGroup $r.resourceGroup `
                -ShortName $r.groupShortName -EmailReceivers ([int]($r.emailReceivers ?? 0)) `
                -SmsReceivers ([int]($r.smsReceivers ?? 0)) -WebhookReceivers ([int]($r.webhookReceivers ?? 0)) `
                -LogicAppReceivers ([int]($r.logicAppReceivers ?? 0)) -ItsmReceivers ([int]($r.itsmReceivers ?? 0)) `
                -UsedByRules $used
        }
    )
}
