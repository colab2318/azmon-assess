#requires -Version 7.0
# Azure Resource Graph KQL queries + paging wrapper. Curated 1:1 from the
# Python collectors/resource_graph.py so results line up with the same model.

$script:AzMonQWorkspaces = @'
resources
| where type =~ 'microsoft.operationalinsights/workspaces'
| project id, name, subscriptionId, resourceGroup, location, tags,
          sku = tostring(properties.sku.name),
          retentionInDays = toint(properties.retentionInDays),
          dailyQuotaGb = todouble(properties.workspaceCapping.dailyQuotaGb),
          customerId = tostring(properties.customerId),
          publicNetworkAccessForIngestion = tostring(properties.publicNetworkAccessForIngestion),
          publicNetworkAccessForQuery = tostring(properties.publicNetworkAccessForQuery),
          disableLocalAuth = tobool(properties.features.disableLocalAuth),
          capacityReservationLevel = toint(properties.sku.capacityReservationLevel),
          clusterResourceId = tostring(properties.clusterResourceId)
'@

$script:AzMonQAppInsights = @'
resources
| where type =~ 'microsoft.insights/components'
| project id, name, subscriptionId, resourceGroup, location, kind, tags,
          applicationType = tostring(properties.Application_Type),
          workspaceResourceId = tostring(properties.WorkspaceResourceId),
          samplingPercentage = todouble(properties.SamplingPercentage),
          retentionInDays = toint(properties.RetentionInDays),
          publicNetworkAccessForIngestion = tostring(properties.publicNetworkAccessForIngestion),
          publicNetworkAccessForQuery = tostring(properties.publicNetworkAccessForQuery),
          disableLocalAuth = tobool(properties.DisableLocalAuth),
          dailyCapGb = todouble(properties.DailyVolumeCap)
'@

$script:AzMonQMetricAlerts = @'
resources
| where type =~ 'microsoft.insights/metricalerts'
| project id, name, subscriptionId, resourceGroup,
          enabled = tobool(properties.enabled),
          severity = toint(properties.severity),
          scopes = properties.scopes,
          actions = properties.actions,
          description = tostring(properties.description)
'@

$script:AzMonQLogAlerts = @'
resources
| where type =~ 'microsoft.insights/scheduledqueryrules'
| project id, name, subscriptionId, resourceGroup,
          enabled = tobool(properties.enabled),
          severity = toint(properties.severity),
          scopes = properties.scopes,
          actions = properties.actions,
          description = tostring(properties.description),
          evaluationFrequency = tostring(properties.evaluationFrequency)
'@

$script:AzMonQActivityAlerts = @'
resources
| where type =~ 'microsoft.insights/activitylogalerts'
| project id, name, subscriptionId, resourceGroup,
          enabled = tobool(properties.enabled),
          scopes = properties.scopes,
          actions = properties.actions,
          description = tostring(properties.description)
'@

$script:AzMonQActionGroups = @'
resources
| where type =~ 'microsoft.insights/actiongroups'
| project id, name, subscriptionId, resourceGroup,
          groupShortName = tostring(properties.groupShortName),
          emailReceivers = array_length(properties.emailReceivers),
          smsReceivers = array_length(properties.smsReceivers),
          webhookReceivers = array_length(properties.webhookReceivers),
          logicAppReceivers = array_length(properties.logicAppReceivers),
          itsmReceivers = array_length(properties.itsmReceivers)
'@

$script:AzMonQAllMonitorable = @'
resources
| where type in~ (
    'microsoft.compute/virtualmachines',
    'microsoft.compute/virtualmachinescalesets',
    'microsoft.web/sites',
    'microsoft.web/serverfarms',
    'microsoft.storage/storageaccounts',
    'microsoft.keyvault/vaults',
    'microsoft.sql/servers/databases',
    'microsoft.dbforpostgresql/flexibleservers',
    'microsoft.dbformysql/flexibleservers',
    'microsoft.documentdb/databaseaccounts',
    'microsoft.network/applicationgateways',
    'microsoft.network/loadbalancers',
    'microsoft.network/publicipaddresses',
    'microsoft.network/networksecuritygroups',
    'microsoft.containerservice/managedclusters',
    'microsoft.app/containerapps',
    'microsoft.servicebus/namespaces',
    'microsoft.eventhub/namespaces',
    'microsoft.cache/redis',
    'microsoft.apimanagement/service',
    'microsoft.network/azurefirewalls',
    'microsoft.containerregistry/registries',
    'microsoft.network/vpngateways',
    'microsoft.network/virtualnetworkgateways',
    'microsoft.network/expressroutecircuits',
    'microsoft.recoveryservices/vaults',
    'microsoft.automation/automationaccounts',
    'microsoft.logic/workflows',
    'microsoft.datafactory/factories',
    'microsoft.network/bastionhosts',
    'microsoft.cognitiveservices/accounts',
    'microsoft.cdn/profiles'
  )
| project id, name, type, subscriptionId, resourceGroup, location, tags
'@

$script:AzMonQDiagnosticSettings = @'
resources
| where type =~ 'microsoft.insights/diagnosticsettings'
| extend targetResourceId = tolower(tostring(properties.targetResourceId))
| project id, name, targetResourceId,
          workspaceId = tostring(properties.workspaceId),
          storageAccountId = tostring(properties.storageAccountId),
          eventHubAuthorizationRuleId = tostring(properties.eventHubAuthorizationRuleId),
          logs = properties.logs, metrics = properties.metrics
'@

# VM extensions — used to detect the retired Log Analytics agent (MMA/OMS:
# MicrosoftMonitoringAgent / OmsAgentForLinux) vs. Azure Monitor Agent
# (AzureMonitorWindowsAgent / AzureMonitorLinuxAgent). Extensions are child
# resources; vmId is derived by trimming the '/extensions/<name>' suffix.
$script:AzMonQVmExtensions = @'
resources
| where type =~ 'microsoft.compute/virtualmachines/extensions'
| extend vmId = tostring(split(id, '/extensions/')[0])
| project id, vmId, name, subscriptionId, resourceGroup,
          extensionType = tostring(properties.type),
          publisher = tostring(properties.publisher)
'@

$script:AzMonQDataCollectionRules = @'
resources
| where type =~ 'microsoft.insights/datacollectionrules'
| project id, name, subscriptionId, resourceGroup, location, kind, tags,
          dataFlowCount = array_length(properties.dataFlows),
          workspaceResourceId = tostring(properties.destinations.logAnalytics[0].workspaceResourceId)
'@

# Fired alert instances, grouped by source rule name — Resource Graph's
# dedicated alerts-management table (not the generic `resources` table).
# Matches by rule *name* (best-effort — the essentials payload does not
# expose the source rule's resource id), mirroring kql/alert_fire_rate.kql.
$script:AzMonQAlertFireRate = @'
AlertsManagementResources
| where type =~ 'microsoft.alertsmanagement/alerts'
| where todatetime(properties.essentials.startDateTime) > ago(30d)
| extend ruleName = tostring(properties.essentials.alertRule)
| where isnotempty(ruleName)
| summarize FireCount = count() by ruleName
'@

function Invoke-AzMonGraphQuery {
    <#
    .SYNOPSIS
        Runs an Azure Resource Graph KQL query across subscriptions via the
        raw ARM REST API (POST /providers/Microsoft.ResourceGraph/resources),
        paging through all results via $skipToken. Uses Invoke-AzRestMethod
        (Az.Accounts only) instead of Search-AzGraph, so the Az.ResourceGraph
        module is never required.
    .NOTES
        Request/response contract verified against the Resource Graph REST
        reference (api-version 2024-04-01): request body is
        { query, subscriptions, options: { $top, $skipToken } }; response is
        { data, $skipToken, resultTruncated, totalRecords }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Query,
        [Parameter(Mandatory)] [string[]] $SubscriptionId
    )
    if (-not $SubscriptionId -or $SubscriptionId.Count -eq 0) {
        Write-Warning 'Invoke-AzMonGraphQuery: no subscriptions in scope.'
        return @()
    }
    $all = [System.Collections.Generic.List[object]]::new()
    $skipToken = $null
    $guard = 0
    do {
        $options = @{ '$top' = 1000 }
        if ($skipToken) { $options['$skipToken'] = $skipToken }
        $body = @{ query = $Query; subscriptions = @($SubscriptionId); options = $options }
        $json = $body | ConvertTo-Json -Depth 10
        $resp = Invoke-AzRestMethod -Path '/providers/Microsoft.ResourceGraph/resources?api-version=2024-04-01' -Method POST -Payload $json -ErrorAction Stop
        if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
            throw "Azure Resource Graph query failed with status $($resp.StatusCode): $($resp.Content)"
        }
        $parsed = $resp.Content | ConvertFrom-Json
        # ARG returns data:null (not data:[]) when a page has zero rows —
        # @($null) wraps that into a 1-element array *containing* null, so
        # guard on truthiness before wrapping/adding rather than on Count.
        # Also strip any individual null rows ARG may include within an
        # otherwise non-empty data array (seen in production, not just the
        # whole-page-null case) — a null propagating into a collector's
        # result silently breaks every downstream consumer that dereferences
        # each row (e.g. "Cannot bind argument ... because it is null").
        if ($parsed.data) {
            $rows = @($parsed.data | Where-Object { $null -ne $_ })
            if ($rows.Count -gt 0) { $all.AddRange($rows) }
        }
        $skipToken = $parsed.'$skipToken'
        $guard++
    } while ($skipToken -and $guard -lt 1000)
    return $all.ToArray()
}
