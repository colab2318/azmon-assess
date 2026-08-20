// Baseline metric alerts for common Azure resource types.
// Deploy at subscription scope. Idempotent — safe to redeploy.

targetScope = 'resourceGroup'

@description('Action group resource ID to notify.')
param actionGroupId string

@description('Optional prefix for alert names.')
param namePrefix string = 'azmon'

@description('Resource IDs of VMs to monitor for availability.')
param vmResourceIds array = []

@description('Resource IDs of App Services to monitor.')
param appServiceResourceIds array = []

@description('Resource IDs of SQL DBs to monitor.')
param sqlDbResourceIds array = []

@description('Resource IDs of Key Vaults to monitor.')
param keyVaultResourceIds array = []

// ---- VM availability (heartbeat) ---------------------------------------
resource vmHeartbeat 'Microsoft.Insights/metricAlerts@2018-03-01' = [for (vmId, i) in vmResourceIds: {
  name: '${namePrefix}-vm-heartbeat-${i}'
  location: 'global'
  properties: {
    severity: 1
    enabled: true
    scopes: [ vmId ]
    evaluationFrequency: 'PT1M'
    windowSize: 'PT5M'
    criteria: {
      'odata.type': 'Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria'
      allOf: [{
        name: 'VMAvailable'
        metricName: 'VmAvailabilityMetric'
        metricNamespace: 'Microsoft.Compute/virtualMachines'
        operator: 'LessThan'
        threshold: 1
        timeAggregation: 'Average'
        criterionType: 'StaticThresholdCriterion'
      }]
    }
    actions: [{ actionGroupId: actionGroupId }]
    description: 'VM heartbeat / availability dropped'
  }
}]

// ---- App Service HTTP 5xx ----------------------------------------------
resource appHttp5xx 'Microsoft.Insights/metricAlerts@2018-03-01' = [for (siteId, i) in appServiceResourceIds: {
  name: '${namePrefix}-app-http5xx-${i}'
  location: 'global'
  properties: {
    severity: 2
    enabled: true
    scopes: [ siteId ]
    evaluationFrequency: 'PT1M'
    windowSize: 'PT5M'
    criteria: {
      'odata.type': 'Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria'
      allOf: [{
        name: 'Http5xx'
        metricName: 'Http5xx'
        metricNamespace: 'Microsoft.Web/sites'
        operator: 'GreaterThan'
        threshold: 5
        timeAggregation: 'Total'
        criterionType: 'StaticThresholdCriterion'
      }]
    }
    actions: [{ actionGroupId: actionGroupId }]
    description: 'App Service returning 5xx responses'
  }
}]

// ---- SQL DTU High ------------------------------------------------------
resource sqlDtu 'Microsoft.Insights/metricAlerts@2018-03-01' = [for (dbId, i) in sqlDbResourceIds: {
  name: '${namePrefix}-sql-dtu-${i}'
  location: 'global'
  properties: {
    severity: 3
    enabled: true
    scopes: [ dbId ]
    evaluationFrequency: 'PT5M'
    windowSize: 'PT15M'
    criteria: {
      'odata.type': 'Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria'
      allOf: [{
        name: 'DtuHigh'
        metricName: 'dtu_consumption_percent'
        metricNamespace: 'Microsoft.Sql/servers/databases'
        operator: 'GreaterThan'
        threshold: 85
        timeAggregation: 'Average'
        criterionType: 'StaticThresholdCriterion'
      }]
    }
    actions: [{ actionGroupId: actionGroupId }]
    description: 'SQL Database DTU consumption > 85%'
  }
}]

// ---- Key Vault API failures --------------------------------------------
resource kvFailures 'Microsoft.Insights/metricAlerts@2018-03-01' = [for (kvId, i) in keyVaultResourceIds: {
  name: '${namePrefix}-kv-apifail-${i}'
  location: 'global'
  properties: {
    severity: 2
    enabled: true
    scopes: [ kvId ]
    evaluationFrequency: 'PT5M'
    windowSize: 'PT15M'
    criteria: {
      'odata.type': 'Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria'
      allOf: [{
        name: 'ApiFailures'
        metricName: 'ServiceApiResult'
        metricNamespace: 'Microsoft.KeyVault/vaults'
        operator: 'GreaterThan'
        threshold: 5
        timeAggregation: 'Total'
        criterionType: 'StaticThresholdCriterion'
      }]
    }
    actions: [{ actionGroupId: actionGroupId }]
    description: 'Key Vault API failures'
  }
}]
