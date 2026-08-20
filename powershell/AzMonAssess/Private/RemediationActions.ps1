#requires -Version 7.0
# Remediation actions — every action defaults to dry-run, returns an
# ActionResult, and never deletes data (only modifies config or attaches
# references). Uses ARM REST (Invoke-AzRestMethod) uniformly so behaviour
# doesn't depend on which typed Az cmdlets happen to expose a given
# property. Ported from remediation/actions/*.py.

function New-AzMonActionResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Action,
        [Parameter(Mandatory)] [bool] $Ok,
        [Parameter(Mandatory)] [bool] $DryRun,
        [Parameter(Mandatory)] [string] $Message,
        [string] $FindingId,
        [string] $TargetId,
        [hashtable] $Changes = @{},
        [string] $ErrorMessage
    )
    return @{
        Action    = $Action
        Ok        = $Ok
        DryRun    = $DryRun
        Message   = $Message
        FindingId = $FindingId
        TargetId  = $TargetId
        Changes   = $Changes
        Error     = $ErrorMessage
        At        = (Get-Date).ToUniversalTime().ToString('o')
    }
}

function Get-AzMonFindingTarget {
    param([hashtable] $Finding)
    return @($Finding['ResourceIds']) | Select-Object -First 1
}

# ---- workspace actions ---------------------------------------------------

function Invoke-AzMonSetWorkspaceDailyQuotaAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'workspace.set_daily_quota'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No workspace resource_id on finding.' -FindingId $Finding['Id'] }

    $cap = $Finding['Evidence']['recommended_daily_cap_gb']
    if (-not $cap) { $cap = $Context['DefaultDailyQuotaGb'] }
    $changes = @{ 'workspace_capping.dailyQuotaGb' = @{ to = $cap } }
    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would set daily cap to $cap GB." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01'
        $fromCap = $null
        if ($current['properties']['workspaceCapping']) { $fromCap = $current['properties']['workspaceCapping']['dailyQuotaGb'] }
        $changes['workspace_capping.dailyQuotaGb']['from'] = $fromCap
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01' -Body @{ properties = @{ workspaceCapping = @{ dailyQuotaGb = [double]$cap } } } | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Set daily cap on $(($target -split '/')[-1]) to $cap GB." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to set daily cap.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

function Invoke-AzMonSetWorkspaceRetentionAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'workspace.set_retention'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No workspace resource_id on finding.' -FindingId $Finding['Id'] }

    $days = $Finding['Evidence']['recommended_retention_days']
    if (-not $days) { $days = $Context['DefaultRetentionDays'] }
    $changes = @{ retentionInDays = @{ to = $days } }
    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would set retention to $days days." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01'
        $changes['retentionInDays']['from'] = $current['properties']['retentionInDays']
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01' -Body @{ properties = @{ retentionInDays = [int]$days } } | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Set retention on $(($target -split '/')[-1]) to ${days}d." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to set retention.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

function Invoke-AzMonSetTablePlanBasicAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'workspace.set_table_plan_basic'
    $target = Get-AzMonFindingTarget -Finding $Finding
    $rawTables = @($Finding['Evidence']['tables'])
    $tables = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $rawTables) {
        if ($t -is [string]) { $tables.Add($t) }
        elseif ($t -is [array] -and $t.Count -gt 0) { $tables.Add([string]$t[0]) }
    }
    if (-not $target -or $tables.Count -eq 0) {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No workspace / tables specified on finding.' -FindingId $Finding['Id'] -TargetId $target
    }
    $changes = @{ tables = @{} }
    foreach ($t in $tables) { $changes['tables'][$t] = @{ plan = @{ to = 'Basic' } } }

    if (-not $Apply) {
        $listed = ($tables | Select-Object -First 5) -join ', '
        $suffix = if ($tables.Count -gt 5) { ' ...' } else { '' }
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would set plan=Basic on $($tables.Count) tables: $listed$suffix" -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    $done = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $tables) {
        try {
            $tableResourceId = "$target/tables/$t"
            $currentTable = Get-AzMonArmResource -ResourceId $tableResourceId -ApiVersion '2022-10-01'
            $changes['tables'][$t]['plan']['from'] = $currentTable['properties']['plan']
            Set-AzMonArmResource -ResourceId $tableResourceId -ApiVersion '2022-10-01' -Body @{ properties = @{ plan = 'Basic' } } | Out-Null
            $done.Add($t)
        } catch {
            $changes['tables'][$t]['error'] = $_.Exception.Message
        }
    }
    return New-AzMonActionResult -Action $name -Ok ($done.Count -gt 0) -DryRun $false -Message "Converted $($done.Count)/$($tables.Count) tables to Basic." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
}

# ---- alert actions --------------------------------------------------------

function Get-AzMonAlertKind {
    param([string] $ResourceId)
    $rid = $ResourceId.ToLowerInvariant()
    if ($rid.Contains('/metricalerts/')) { return 'metric' }
    if ($rid.Contains('/scheduledqueryrules/')) { return 'log' }
    if ($rid.Contains('/activitylogalerts/')) { return 'activityLog' }
    return 'unknown'
}

function Invoke-AzMonAttachActionGroupAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'alerts.attach_action_group'
    $target = Get-AzMonFindingTarget -Finding $Finding
    $agId = $Context['DefaultActionGroupId']
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No alert rule on finding.' -FindingId $Finding['Id'] }
    if (-not $agId) {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No DefaultActionGroupId configured. Pass -ActionGroupId to Invoke-AzMonRemediation.' -FindingId $Finding['Id'] -TargetId $target
    }
    $kind = $Finding['Evidence']['kind']
    if (-not $kind) { $kind = Get-AzMonAlertKind -ResourceId $target }
    $changes = @{ 'actions.actionGroups' = @{ add = $agId; kind = $kind } }

    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would attach action group to $kind rule." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        switch ($kind) {
            'metric' {
                $apiVersion = '2018-03-01'
                $current = Get-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion
                $actions = @($current['properties']['actions'])
                if (-not ($actions | Where-Object { ([string]$_['actionGroupId']).ToLowerInvariant() -eq $agId.ToLowerInvariant() })) {
                    $actions += @{ actionGroupId = $agId }
                }
                Set-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion -Body @{ properties = @{ actions = $actions } } | Out-Null
            }
            'log' {
                $apiVersion = '2021-08-01'
                $current = Get-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion
                $existingActions = $current['properties']['actions']
                $groups = if ($existingActions -and $existingActions['actionGroups']) { @($existingActions['actionGroups']) } else { @() }
                if ($groups -notcontains $agId) { $groups += $agId }
                Set-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion -Body @{ properties = @{ actions = @{ actionGroups = $groups } } } | Out-Null
            }
            'activityLog' {
                $apiVersion = '2020-10-01'
                $current = Get-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion
                $existing = @()
                if ($current['properties']['actions'] -and $current['properties']['actions']['actionGroups']) { $existing = @($current['properties']['actions']['actionGroups']) }
                if (-not ($existing | Where-Object { ([string]$_['actionGroupId']).ToLowerInvariant() -eq $agId.ToLowerInvariant() })) {
                    $existing += @{ actionGroupId = $agId }
                }
                Set-AzMonArmResource -ResourceId $target -ApiVersion $apiVersion -Body @{ properties = @{ actions = @{ actionGroups = $existing } } } | Out-Null
            }
            default {
                return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message "Unsupported alert kind: $kind" -FindingId $Finding['Id'] -TargetId $target
            }
        }
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Attached action group to $kind rule $(($target -split '/')[-1])." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to attach action group.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

# ---- App Insights actions -------------------------------------------------

function Invoke-AzMonEnableAppInsightsSamplingAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'appinsights.enable_sampling'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No App Insights resource_id on finding.' -FindingId $Finding['Id'] }
    $pct = $Finding['Evidence']['recommended_sampling_percentage']
    if (-not $pct) { $pct = $Context['DefaultSamplingPercentage'] }
    $changes = @{ SamplingPercentage = @{ to = $pct } }
    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would set sampling to $pct%." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2020-02-02'
        $changes['SamplingPercentage']['from'] = $current['properties']['SamplingPercentage']
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2020-02-02' -Body @{ properties = @{ SamplingPercentage = [double]$pct } } | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Set sampling on $(($target -split '/')[-1]) to $pct%." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to set sampling.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

function Invoke-AzMonDisableAppInsightsLocalAuthAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'appinsights.disable_local_auth'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No App Insights resource_id on finding.' -FindingId $Finding['Id'] }
    $changes = @{ DisableLocalAuth = @{ to = $true } }
    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message 'Would disable instrumentation-key auth. WARNING: apps still using iKey will stop sending telemetry until migrated to Entra ID.' -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2020-02-02'
        $changes['DisableLocalAuth']['from'] = $current['properties']['DisableLocalAuth']
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2020-02-02' -Body @{ properties = @{ DisableLocalAuth = $true } } | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Disabled instrumentation-key auth on $(($target -split '/')[-1])." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to disable App Insights local auth.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

# ---- security actions -----------------------------------------------------

function Get-AzMonFindingDirection {
    param([hashtable] $Finding)
    $ev = $Finding['Evidence']['direction']
    if ($ev -in @('ingestion', 'query')) { return $ev }
    $title = ([string]$Finding['Title']).ToLowerInvariant()
    if ($title.Contains('ingestion')) { return 'ingestion' }
    if ($title.Contains('query')) { return 'query' }
    return 'both'
}

function Invoke-AzMonDisableWorkspacePublicNetworkAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'workspace.disable_public_network'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No workspace resource_id on finding.' -FindingId $Finding['Id'] }
    $direction = Get-AzMonFindingDirection -Finding $Finding
    $changes = @{}
    if ($direction -in @('ingestion', 'both')) { $changes['publicNetworkAccessForIngestion'] = @{ to = 'Disabled' } }
    if ($direction -in @('query', 'both')) { $changes['publicNetworkAccessForQuery'] = @{ to = 'Disabled' } }

    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message "Would disable public network access ($direction). WARNING: clients not on Private Link will lose access." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01'
        $body = @{ properties = @{} }
        if ($direction -in @('ingestion', 'both')) {
            $changes['publicNetworkAccessForIngestion']['from'] = $current['properties']['publicNetworkAccessForIngestion']
            $body['properties']['publicNetworkAccessForIngestion'] = 'Disabled'
        }
        if ($direction -in @('query', 'both')) {
            $changes['publicNetworkAccessForQuery']['from'] = $current['properties']['publicNetworkAccessForQuery']
            $body['properties']['publicNetworkAccessForQuery'] = 'Disabled'
        }
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01' -Body $body | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Disabled public network access ($direction) on $(($target -split '/')[-1])." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to disable public network access.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

function Invoke-AzMonDisableWorkspaceLocalAuthAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [Parameter(Mandatory)] [hashtable] $Context, [Parameter(Mandatory)] [bool] $Apply)
    $name = 'workspace.disable_local_auth'
    $target = Get-AzMonFindingTarget -Finding $Finding
    if (-not $target) { return New-AzMonActionResult -Action $name -Ok $false -DryRun (-not $Apply) -Message 'No workspace resource_id on finding.' -FindingId $Finding['Id'] }
    $changes = @{ 'features.disableLocalAuth' = @{ to = $true } }
    if (-not $Apply) {
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $true -Message 'Would disable shared-key (local) auth. WARNING: any agent still using workspace keys will stop ingesting.' -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    }
    try {
        $current = Get-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01'
        $features = $current['properties']['features']
        $prior = $null
        if ($features) { $prior = $features['disableLocalAuth'] } else { $features = @{} }
        $changes['features.disableLocalAuth']['from'] = $prior
        $features['disableLocalAuth'] = $true
        Set-AzMonArmResource -ResourceId $target -ApiVersion '2022-10-01' -Body @{ properties = @{ features = $features } } | Out-Null
        return New-AzMonActionResult -Action $name -Ok $true -DryRun $false -Message "Disabled shared-key auth on $(($target -split '/')[-1])." -FindingId $Finding['Id'] -TargetId $target -Changes $changes
    } catch {
        return New-AzMonActionResult -Action $name -Ok $false -DryRun $false -Message 'Failed to disable local auth.' -FindingId $Finding['Id'] -TargetId $target -ErrorMessage $_.Exception.Message
    }
}

# ---- manual runbooks -------------------------------------------------------

$script:AzMonRunbookTemplates = @{
    ClassicAiMigration    = @'
# Classic Application Insights -> Workspace-based migration

Classic AI components can be migrated in-place. The workspace-based mode
preserves the ingestion key and telemetry pipeline but stores data in a
Log Analytics workspace.

## Portal
Application Insights blade -> Properties -> "Migrate to Workspace-based" -> pick a workspace.

## CLI
```bash
az resource update --ids "{ResourceId}" --set properties.WorkspaceResourceId="/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.OperationalInsights/workspaces/<workspace>"
```

**Caveats**
- Ingestion keys unchanged; SDK reconfig NOT required.
- Migration is one-way. Test on a non-prod AI component first.
'@
    WorkspaceConsolidation = @'
# Workspace consolidation runbook

Target design: one workspace per (region, sensitivity tier).

## Phase 1 - Parallel dual-shipping (30 days)
1. Create target workspaces (see `bicep/` starter pack).
2. Update diagnostic settings on high-volume resources to ship to BOTH old + new workspace.
3. Reconstruct workbooks / dashboards against the new workspace.
4. Rewrite scheduled query alerts against the new workspace.

## Phase 2 - Cutover
5. Remove old workspace from diagnostic settings.
6. Disable / delete alert rules pointing at the old workspace.

## Phase 3 - Decommission
7. Delete the old workspace after retention window expires.
'@
    AzRegionMigration      = @'
# Move workspace to an availability-zone region

Workspaces in non-AZ regions have no in-region redundancy. Migration is
disruptive: create a new workspace in a paired AZ-capable region and
dual-ship for 30 days before cutover.

## Steps
1. Provision a new workspace in an AZ region.
2. Update diagnostic settings on high-volume resources to add the new workspace as a second destination.
3. Reconstruct alert rules against the new workspace.
4. After validation, remove the old workspace destination.
5. Delete the old workspace after retention lapses.
'@
    WorkspaceHealthAlert   = @'
# Workspace health / ingestion outage alert

## KQL
```kusto
Heartbeat
| where TimeGenerated > ago(30m)
| summarize LastHeartbeat = max(TimeGenerated) by Computer
| where LastHeartbeat < ago(15m)
```

Use `bicep/alert-baseline.bicep` as a starting point for a scheduled query alert scoped to the workspace with a 30-minute window and severity 2.
'@
    DualDestinationLogs    = @'
# Dual-destination diagnostic settings for prod

For business-critical workloads, WAF recommends shipping logs to BOTH:
1. The primary Log Analytics workspace (fast query).
2. A storage account (long-term retention, cheap forensic backup).

```bash
az monitor diagnostic-settings create --resource "{ResourceId}" --name "ds-dual" --workspace "<workspace-id>" --storage-account "<storage-id>" --logs '[{"category":"AuditEvent","enabled":true}]' --metrics '[{"category":"AllMetrics","enabled":true}]'
```
'@
    DedicatedCluster       = @'
# Move to a Log Analytics dedicated cluster

Dedicated clusters unlock 500 GB/day commit-tier pricing, CMK, and
cross-workspace cost pooling. Recommended for workspaces > 500 GB/day.

## Steps
1. Create the cluster and choose a commit tier close to current daily ingest.
2. Associate the target workspace(s) via `az monitor log-analytics cluster update`.
3. Validate ingestion continues normally, then enable CMK if required.
'@
    SearchJobsTier         = @'
# Move high-volume / long-retention tables to Basic + Search Jobs

Basic logs are ~5x cheaper than Analytics for ingestion but only support
KQL search (no alerts, no cross-table joins).

## Steps
1. Identify Analytics alerts on the target tables and refactor to a summary/rollup table.
2. Switch each table to Basic:
   ```bash
   az monitor log-analytics workspace table update --resource-group <rg> --workspace-name <ws> --name <TableName> --plan Basic
   ```
3. Configure a Search Job template for forensic queries.
'@
    CoverageDiagBaseline   = @'
# Enable diagnostic settings

Deploy the `bicep/diag-settings-baseline.bicep` starter across the affected
resources to route logs to your central Log Analytics workspace:

{ResourceIdList}

## Example
```bash
az deployment sub create --location eastus --template-file bicep/diag-settings-baseline.bicep --parameters workspaceId=<target-workspace-id>
```
'@
}

function Write-AzMonManualRunbookAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Finding,
        [Parameter(Mandatory)] [string] $TemplateKey,
        [Parameter(Mandatory)] [string] $RunbookDir,
        [Parameter(Mandatory)] [bool] $Apply
    )
    $name = 'manual.runbook'
    $target = (Get-AzMonFindingTarget -Finding $Finding)
    if (-not $target) { $target = 'n/a' }
    $template = $script:AzMonRunbookTemplates[$TemplateKey]
    $idList = (@($Finding['ResourceIds']) | Select-Object -First 20 | ForEach-Object { "- $_" }) -join "`n"
    if (-not $idList) { $idList = '- (none)' }
    $body = $template.Replace('{ResourceId}', $target).Replace('{ResourceIdList}', $idList)
    $fileName = "runbook-$($Finding['Id'].Substring(0, 8)).md"
    $outPath = Join-Path $RunbookDir $fileName

    if ($Apply) {
        New-AzMonOutputDirectory -Path $RunbookDir | Out-Null
        "# $($Finding['Title'])`n`n$body`n" | Set-Content -LiteralPath $outPath -Encoding utf8NoBOM
        $msg = "Wrote runbook to $fileName."
    } else {
        $msg = "Would emit runbook $fileName for manual action."
    }
    return New-AzMonActionResult -Action $name -Ok $true -DryRun (-not $Apply) -Message $msg -FindingId $Finding['Id'] -TargetId $target -Changes @{ runbook = $outPath }
}
