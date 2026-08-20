# azmon-assess

**AI-powered Azure Monitoring & Observability Assessment — zero installs, zero admin rights. Runs entirely in Azure Cloud Shell, or any local PowerShell 7+ session that already has `Connect-AzAccount` working.**

Purpose-built for restricted customer environments where you **cannot
install any dependency** — not even in `-Scope CurrentUser`. The
only requirement is `Az.Accounts` (already present anywhere `Connect-AzAccount`
works, and preinstalled in Azure Cloud Shell). Every Azure Resource Graph,
Log Analytics, and ARM call is made with plain REST via `Invoke-AzRestMethod`
/ `Get-AzAccessToken` — no `Az.ResourceGraph`, no `Az.OperationalInsights`,
no `Az.Resources`. The Excel and PowerPoint deliverables are both generated
as raw OOXML (zip + XML) using only `System.IO.Compression` — no `ImportExcel`,
no EPPlus, no Office install, no COM.

It produces a read-only assessment: Log Analytics
workspace consolidation opportunities, coverage gaps (including VM Heartbeat
coverage), alert quality (including noisy/high-fire-rate rule scoring), cost
optimization, distributed-tracing readiness, a Data Collection Rule (DCR)
inventory, and WAF Reliability / Security / Performance findings — then
renders them as **Excel**, **PowerPoint**, and **HTML** deliverables you can
hand to a customer, plus Markdown and a JSON snapshot. The Excel workbook
includes a resource-centric **`4.ImpactedResourcesAnalysis`** sheet (one row
per finding × impacted resource, styled after the standard WARA/APRL expert-
analysis workbook layout) alongside the finding-centric `Findings` sheet.

## Why zero-install

| Constraint | How this tool handles it |
|---|---|
| No admin rights, **no module installs at all** | The only dependency is `Az.Accounts`, which is preinstalled in Azure Cloud Shell and ships with any Az PowerShell install — `Initialize-AzMonPrerequisite` only checks for it, it never runs `Install-Module` |
| No Python / pip | Pure PowerShell 7+, zero external modules |
| Excel output, no Excel/ImportExcel install | `.xlsx` is generated as raw OOXML (zip + XML) using only built-in .NET (`Private/XlsxBuilder.ps1`) — no EPPlus, no ImportExcel |
| PowerPoint output, no Office install | `.pptx` is generated as raw OOXML (zip + XML) using only built-in .NET — no COM, no Office, works on Cloud Shell's Linux host |
| Auth | Uses your existing `Connect-AzAccount` / Cloud Shell session via `Az.Accounts`; every Resource Graph / Log Analytics / ARM call goes through `Invoke-AzRestMethod` and `Get-AzAccessToken` |

## Prerequisites

- PowerShell 7.0+ (Azure Cloud Shell and the `pwsh` you already have both qualify)
- Signed in via `Connect-AzAccount` (Cloud Shell: already signed in) with at least **Reader** + **Log Analytics Reader** on the subscriptions you want assessed
- `Az.Accounts` — the **only** module this tool needs. It's preinstalled in
  Azure Cloud Shell and ships with any Az PowerShell install; nothing is ever
  installed automatically. If it's genuinely missing (rare, local-only
  scenario), install it yourself once with
  `Install-Module -Name Az.Accounts -Scope CurrentUser`.
- (Optional) An Azure OpenAI deployment for the AI executive summary — the tool falls back to a deterministic rule-based summary if not configured
- (Optional) `powershell-yaml` — only if you want YAML triage plans instead of JSON; JSON works with zero extra modules and is used automatically if `powershell-yaml` isn't present

No Python, no Office, no admin rights, no Docker, no `Install-Module` at runtime.

## Quick start (Azure Cloud Shell)

```powershell
git clone <this-repo-url>
cd azmon-assess/powershell

# Preview the reports with fixture data — no Azure access required.
./azmon-assess.ps1 demo -Output ./out

# Full assessment across every subscription your identity can see.
./azmon-assess.ps1 run -Output ./out
```

Open `./out/report.xlsx` and `./out/report.pptx` — that's the customer deliverable.

## Quick start (local PowerShell 7+, no admin)

```powershell
cd azmon-assess/powershell
./azmon-assess.ps1 run -Output ./out
```

The first run only checks that `Az.Accounts` is available (it is, on Cloud
Shell) and prompts you to `Connect-AzAccount` if you aren't already signed in.
Nothing is ever installed automatically.

### Assessing a different tenant (e.g., a customer tenant as a guest/consultant)

Pass `-TenantId` to sign into that tenant specifically:

```powershell
./azmon-assess.ps1 run -Output ./out -TenantId <customer-tenant-id-or-domain>
```

If you're already signed into the right tenant (e.g., via `Connect-AzAccount
-TenantId ...` run beforehand, or a cached session), the tool reuses that
session as-is and `-TenantId` isn't needed. To be certain which tenant/
subscriptions will be assessed before running a full scan, check first:

```powershell
Get-AzContext                                    # confirms signed-in tenant
Get-AzSubscription -TenantId <tenant-id>          # lists subscriptions in that tenant
./azmon-assess.ps1 run -Output ./out -TenantId <tenant-id> -SubscriptionId <sub-id-1>,<sub-id-2>
```

`-SubscriptionId` (or `-ManagementGroupId`) scopes the assessment explicitly;
omitting both auto-discovers every subscription the signed-in identity can
see in that tenant.

## Commands

All commands go through `./azmon-assess.ps1 <command> [options]`.

```powershell
# Full assessment (collect + analyze + AI summary + reports)
./azmon-assess.ps1 run -Output ./out

# Focused single-topic scans (still collect everything, just filter findings)
./azmon-assess.ps1 consolidate -Output ./out   # Log Analytics workspace consolidation
./azmon-assess.ps1 gaps        -Output ./out   # coverage gaps
./azmon-assess.ps1 alerts      -Output ./out   # alert quality
./azmon-assess.ps1 cost        -Output ./out   # cost optimization
./azmon-assess.ps1 tracing     -Output ./out   # OpenTelemetry readiness
./azmon-assess.ps1 reliability -Output ./out   # WAF reliability
./azmon-assess.ps1 security    -Output ./out   # WAF security
./azmon-assess.ps1 performance -Output ./out   # WAF performance efficiency

# Preview with fixture data — no Azure access required
./azmon-assess.ps1 demo -Output ./out

# Regenerate one or more report formats from an existing snapshot
./azmon-assess.ps1 report -Snapshot ./out/snapshot.json -Output ./out -Format all
./azmon-assess.ps1 report -Snapshot ./out/snapshot.json -Output ./out -Format excel,pptx,html

# Regenerate just the AI executive summary
./azmon-assess.ps1 summarize -Snapshot ./out/snapshot.json -Output ./out

# Triage findings (interactive, resumable)
./azmon-assess.ps1 triage -Snapshot ./out/snapshot.json -TriageOutput ./out/triage.json

# Emit a YAML/JSON template pre-populated with every finding as "snooze"
./azmon-assess.ps1 triage -Snapshot ./out/snapshot.json -TriageOutput ./out/triage.json -EmitTemplate
# ...edit ./out/triage.template.yaml (or .json), set decisions to accept|reject|snooze|defer|manual...
./azmon-assess.ps1 triage -Snapshot ./out/snapshot.json -TriageOutput ./out/triage.json -Plan ./out/triage.template.yaml

# Remediate — dry-run by default, add -Apply to actually change Azure
./azmon-assess.ps1 remediate -Snapshot ./out/snapshot.json -TriagePath ./out/triage.json -Output ./out
./azmon-assess.ps1 remediate -Snapshot ./out/snapshot.json -TriagePath ./out/triage.json -Output ./out -Apply -ActionGroup "/subscriptions/<sub>/resourceGroups/<rg>/providers/microsoft.insights/actionGroups/<name>"
```

Common options: `-SubscriptionId <id,id,...>` / `-ManagementGroupId <id>` (defaults to
auto-discovering every subscription your identity can see), `-CustomerName`,
`-AoaiEndpoint` / `-AoaiDeployment` / `-AoaiApiVersion` (or set
`AZURE_OPENAI_ENDPOINT` / `AZURE_OPENAI_DEPLOYMENT` / `AZURE_OPENAI_API_VERSION`
environment variables), `-LookbackDays` (default 30), `-SkipIngestionEnrichment`
(faster, skips the per-workspace KQL calls), `-SkipAiSummary`.

## Using the module directly (no wrapper script)

Everything is also available as a normal PowerShell module if you'd rather
script it yourself:

```powershell
Import-Module ./AzMonAssess/AzMonAssess.psd1
Initialize-AzMonPrerequisite      # verifies Az.Accounts is available (never installs anything)
Connect-AzMonSession              # Connect-AzAccount if not already signed in

$snapshot = Invoke-AzMonAssessment -OutputPath ./out
# ...or drive it piece by piece:
$subs        = Resolve-AzMonSubscription
$workspaces  = Get-AzMonWorkspace -SubscriptionId $subs
$appInsights = Get-AzMonAppInsight -SubscriptionId $subs
$findings    = Find-AzMonCostOptimizationFinding -Workspace $workspaces -AppInsight $appInsights
```

## Outputs

Written to `-Output` (default `./out`):

- `snapshot.json` — raw collected data + findings (input to `report`, `triage`, `remediate`, `summarize`)
- `report.xlsx` — Excel workbook: Summary, Findings (severity-colored), **`4.ImpactedResourcesAnalysis`** (resource-centric, one row per finding × impacted resource), Workspaces, App Insights, Alert Rules (with 30d fire count), Action Groups, Diagnostic Settings, Data Collection Rules, Consolidation Plan
- `report.pptx` — 11-slide executive PowerPoint deck (KPIs, severity/category breakdowns, top-10, impacted resources, savings, roadmap) — generated as plain OOXML, opens in PowerPoint / LibreOffice / Google Slides
- `report.html` — self-contained dark-theme HTML report (same KPIs + findings as the Excel/Markdown reports), generated as plain string-built HTML/CSS — no browser or Jinja2 dependency
- `report.md` — Markdown version
- `triage.json` / `triage.template.yaml` — triage decisions / template
- `remediation-log.json` — audit trail of every dry-run / applied change
- `runbooks/runbook-*.md` — generated runbooks for manual-only actions

## Safety

- `run`, `consolidate`, `gaps`, `alerts`, `cost`, `tracing`, `reliability`, `security`, `performance`, `demo`, `report`, `triage` are **read-only**
- `remediate` **defaults to dry-run** — no changes are made without `-Apply`
- Every applied change is logged to `remediation-log.json` with before/after values
- Remediation uses `Invoke-AzRestMethod` (ARM REST, get-then-patch) so it doesn't depend on every Az module exposing every property as a typed cmdlet parameter
- No secrets are logged or written to disk

## Differences from the Python edition (by design)

- PowerPoint charts (severity/category/savings breakdowns) are rendered as
  proportional bars/segments built from plain shapes instead of native Office
  chart objects — this keeps the file generator dependency-free (no embedded
  workbook parts) at a small cost to "this is an editable native chart" — the
  numbers and labels are identical.
- Excel is generated the same way — raw OOXML via `Private/XlsxBuilder.ps1`
  instead of a module like `ImportExcel`/EPPlus — so there's no dependency to
  install in a restricted environment. It includes an
  additional `4.ImpactedResourcesAnalysis` sheet (styled after the standard
  WARA/APRL expert-analysis workbook layout: one row per finding × impacted
  resource).
- The HTML report is built with plain PowerShell string-building/here-strings
  (no Jinja2 available in PowerShell) instead of templating.
- Config is environment variables + script parameters rather than a `.env` file
  (`AZURE_SUBSCRIPTION_IDS`,
  `AZURE_MANAGEMENT_GROUP_ID`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT`,
  `AZURE_OPENAI_API_VERSION`, `AZMON_LOOKBACK_DAYS`, `AZMON_MAX_PARALLEL`,
  `AZMON_CUSTOMER_NAME`, `AZMON_DEFAULT_ACTION_GROUP_ID`, `AZMON_AZ_REGIONS`).

## Repository layout

```
powershell/
  azmon-assess.ps1               # CLI wrapper (run / demo / report / triage / remediate / ...)
  AzMonAssess/
    AzMonAssess.psd1             # module manifest
    AzMonAssess.psm1             # dot-sources Public/Private, exports the public surface
    Public/                      # collectors, analyzers, reports, orchestrator, triage, remediation
    Private/                     # ARG/KQL queries, pricing tables, model constructors, PPTX/REST helpers
```

The `../bicep/` and `../kql/` folders (alerting starter pack, curated KQL
library) are referenced directly from generated runbooks.
