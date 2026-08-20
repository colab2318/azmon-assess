# Bicep starter pack

Reusable Bicep templates that implement the top recommendations `azmon-assess`
emits. Everything is idempotent — safe to redeploy.

Deploy in order:

## 1. Action group(s)

One per environment (e.g. `prod-pager`, `non-prod-email`, `security-siem`):

```powershell
az deployment group create `
  -g <rg-monitoring> `
  -f ./action-group.bicep `
  -p name='org-prod-pager' `
     shortName='orgprod' `
     emails='["oncall@example.com"]' `
     pagerDutyWebhookUrl='https://events.pagerduty.com/...'
```

## 2. Diagnostic-settings policy (subscription scope)

Auto-deploys `Microsoft.Insights/diagnosticSettings` on all Key Vaults, SQL DBs,
and App Services in the subscription and routes them to your workspace.

```powershell
az deployment sub create `
  -l eastus `
  -f ./diag-settings-baseline.bicep `
  -p workspaceId='/subscriptions/<sub>/resourceGroups/<rg>/providers/Microsoft.OperationalInsights/workspaces/<ws>' `
     displayNamePrefix='<OrgShortName>'

# Then remediate existing resources:
az policy remediation create `
  --name azmon-diag-baseline-remediate-kv `
  --policy-assignment azmon-diag-baseline-kv
```

## 3. Alert baseline (per resource group)

Uses the action-group ID from step 1:

```powershell
az deployment group create `
  -g <rg-workloads> `
  -f ./alert-baseline.bicep `
  -p actionGroupId='<actionGroupResourceId>' `
     vmResourceIds='["<vmResourceId1>", "<vmResourceId2>"]' `
     appServiceResourceIds='[]' `
     sqlDbResourceIds='[]' `
     keyVaultResourceIds='[]'
```

## Parameters at a glance

| Template | Key parameter | Purpose |
|---|---|---|
| `action-group.bicep` | `shortName` (≤12 chars) | On-call routing identity |
| `diag-settings-baseline.bicep` | `workspaceId`, `displayNamePrefix` | Central sink + friendly labels |
| `alert-baseline.bicep` | `actionGroupId`, `*ResourceIds` | Which resources to protect |

All name prefixes default to `azmon` / `azmon-diag-baseline` — override with
`namePrefix=` when deploying if you want your own naming convention.
