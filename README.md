# azmon-assess

**AI-powered Azure Monitoring & Observability Assessment.**

`azmon-assess` scans an Azure tenant and produces a prioritized, cost-aware
roadmap to consolidate Log Analytics workspaces, close observability gaps,
reduce alert noise, and adopt distributed tracing.

PowerShell 7+ only, with a single dependency: `Az.Accounts` (already present
anywhere `Connect-AzAccount` works, including Azure Cloud Shell). No admin
rights, no `Install-Module`, no Python, no Office install — see
[`powershell/README.md`](powershell/README.md) for full usage, or jump
straight in:

```powershell
cd powershell
./azmon-assess.ps1 demo -Output ./out   # preview with fixture data, no Azure access required
./azmon-assess.ps1 run  -Output ./out   # full assessment across every subscription you can see
```

## What it does

| Capability | Concern addressed |
|---|---|
| Inventory of Log Analytics workspaces, App Insights, alerts, action groups, DCRs | Workspace / component sprawl |
| Per-workspace ingestion & cost analysis (billable GB by table, retention, commitment tier fit) | Unpredictable monitoring bills |
| Coverage-gap detection (resources without diagnostics, VMs without heartbeat, web apps without App Insights, orphaned Data Collection Rules) | Missed critical incidents |
| Alert quality scoring (noisy rules, missing critical rules, orphaned action groups, broad-scope metric alerts, high-frequency log search alerts, static-threshold-only metric alerts) | Alert noise + under-alerting + alert evaluation cost |
| Consolidation planner (groups workspaces by region/env, projects commitment-tier savings) | Central monitoring strategy |
| Distributed-tracing readiness check (OpenTelemetry / App Insights auto-instrumentation coverage) | End-to-end tracing gap |
| **Reliability (WAF)** — workspace AZ-region check, prod dual-destination logs, workspace health alert coverage, Azure Service Health alert coverage | Zonal outages, silent ingestion failures, unnoticed Azure service incidents |
| **Security (WAF)** — public network access, Entra-only (disable local auth), classic AI detection, workspace/AI coherence | Data exfiltration, key sprawl, missing Private Link |
| **Performance Efficiency (WAF)** — AI/workspace region colocation, Search Jobs / Basic-tier candidates, dedicated cluster fit | Cross-region latency, hot-storage cost, throttled queries |
| **Cost Optimization (WAF, extended)** — App Insights daily cap, Sentinel commingled with ops data, Summary Rules candidates, standalone-workspace commitment tiers, retired Log Analytics agent (MMA) detection, live Azure Advisor cost recommendations, Auxiliary (Lake) tier candidates | Runaway telemetry cost, Sentinel surcharge waste, silent monitoring blind spots from the retired agent |
| AI-generated executive summary + prioritized recommendations (Azure OpenAI) | Leadership-ready briefing |
| **Multi-format reports** — HTML, Markdown, **Excel workbook** (incl. resource-centric Impacted Resources Analysis sheet), **PowerPoint deck** | Analyst + executive audience |
| **Interactive / batch triage** — decisions per finding, resumable sessions | Structured follow-through |
| **Safe remediation** — dry-run by default; applies daily-quota, retention, table-plan, sampling, silent-alert, public-network, local-auth fixes via ARM REST | Close the loop from finding to fix |
| Bicep alerting starter pack + generated runbooks for manual actions (AZ migration, health alerts, dedicated cluster, Search Jobs) | Actionable deliverable |

## Outputs

Files land in the `-Output` directory (default `./out`):

- `snapshot.json` — raw collected data + findings (input to `report`, `triage`, `remediate`, `summarize`)
- `report.html` — interactive assessment report (open in any browser)
- `report.md` — markdown version (great for PRs / wiki)
- `report.xlsx` — Excel workbook: Summary, Findings, Impacted Resources Analysis, Workspaces, App Insights, Alert Rules, Action Groups, Diagnostic Settings, Data Collection Rules, Consolidation Plan
- `report.pptx` — 11-slide executive PowerPoint deck (KPIs, charts, impacted resources, top-10, 30/60/90-day roadmap)
- `triage.json` — decisions per finding (from `triage`)
- `triage.template.yaml` / `.json` — skeleton for batch triage (from `triage -EmitTemplate`)
- `remediation-log.json` — audit trail of every dry-run / applied change (from `remediate`)
- `runbooks/runbook-*.md` — generated markdown runbooks for manual-only actions

## What gets collected

**Read-only.** The tool never modifies resources. It queries:

- Azure Resource Graph for resource inventory
- `Microsoft.OperationalInsights/workspaces` REST for workspace config
- `Microsoft.Insights/components` REST for App Insights
- `Microsoft.Insights/metricAlerts` + `scheduledQueryRules` + `activityLogAlerts`
- `Usage` + `Operation` + `Heartbeat` tables in each LA workspace for ingestion
  / cost / VM-agent coverage analysis
- `AppRequests` / `AppDependencies` for tracing coverage

## Repository layout

```
powershell/
  azmon-assess.ps1               # CLI wrapper (run / demo / report / triage / remediate / ...)
  AzMonAssess/
    AzMonAssess.psd1             # module manifest
    AzMonAssess.psm1             # dot-sources Public/Private, exports the public surface
    Public/                      # collectors, analyzers, reports, orchestrator, triage, remediation
    Private/                     # ARG/KQL REST queries, pricing tables, model constructors, PPTX/XLSX/REST helpers
kql/                    # Curated KQL library (reusable)
bicep/                  # Standardized alert + diagnostic baseline
```

## Safety

- Assessment commands (`run`, `consolidate`, `gaps`, `alerts`, `cost`, `tracing`, `reliability`, `security`, `performance`, `demo`, `report`, `triage`) are **read-only**
- `remediate` **defaults to dry-run** — no changes are made without `-Apply`
- Every applied change is logged to `remediation-log.json` with before/after values
- No secrets logged or written to disk
- KQL queries use aggregations, not raw log dumps, to stay well under any daily cap
- No paths, subscription IDs, or organization names are hardcoded — everything is passed via parameters/environment variables, resolved from Azure, or defaults to a generic placeholder
- Unsafe / complex actions (workspace consolidation, classic-AI migration, coverage rollouts) emit a **manual runbook** instead of auto-executing

