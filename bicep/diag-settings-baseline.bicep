// Azure Policy assignment: enforce diagnostic settings on all Key Vaults / SQL / App Services.
// Uses the built-in DINE (DeployIfNotExists) policy definitions.
//
// Deploy at management-group or subscription scope, then run remediation.

targetScope = 'subscription'

@description('Log Analytics workspace resource ID to route diagnostics to.')
param workspaceId string

@description('Policy assignment name prefix.')
param namePrefix string = 'azmon-diag-baseline'

@description('Location for the managed identity used by policy remediation.')
param identityLocation string = 'eastus'

@description('Prefix used in policy assignment display names (e.g. organization short name).')
param displayNamePrefix string = 'AzMon'

// Built-in DINE policy definitions (well-known IDs)
var vaultPolicyId = '/providers/Microsoft.Authorization/policyDefinitions/951af2fa-529b-416e-ab6e-066fd85ac459'  // Deploy diag settings for Key Vaults
var sqlPolicyId   = '/providers/Microsoft.Authorization/policyDefinitions/32e6bbec-16b6-44c2-be37-c5b672d103cf'  // SQL DB
var webPolicyId   = '/providers/Microsoft.Authorization/policyDefinitions/b79fa14e-238a-4c2d-b376-442ce508fc84'  // App Service

resource kvAssign 'Microsoft.Authorization/policyAssignments@2023-04-01' = {
  name: '${namePrefix}-kv'
  location: identityLocation
  identity: { type: 'SystemAssigned' }
  properties: {
    displayName: '${displayNamePrefix} — Deploy diagnostic settings for Key Vaults'
    policyDefinitionId: vaultPolicyId
    parameters: {
      logAnalytics: { value: workspaceId }
    }
  }
}

resource sqlAssign 'Microsoft.Authorization/policyAssignments@2023-04-01' = {
  name: '${namePrefix}-sql'
  location: identityLocation
  identity: { type: 'SystemAssigned' }
  properties: {
    displayName: '${displayNamePrefix} — Deploy diagnostic settings for SQL Databases'
    policyDefinitionId: sqlPolicyId
    parameters: {
      logAnalytics: { value: workspaceId }
    }
  }
}

resource webAssign 'Microsoft.Authorization/policyAssignments@2023-04-01' = {
  name: '${namePrefix}-web'
  location: identityLocation
  identity: { type: 'SystemAssigned' }
  properties: {
    displayName: '${displayNamePrefix} — Deploy diagnostic settings for App Services'
    policyDefinitionId: webPolicyId
    parameters: {
      logAnalytics: { value: workspaceId }
    }
  }
}

output policyAssignmentIds array = [
  kvAssign.id
  sqlAssign.id
  webAssign.id
]
