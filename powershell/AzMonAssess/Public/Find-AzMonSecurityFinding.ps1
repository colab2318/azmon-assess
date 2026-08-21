#requires -Version 7.0
# WAF Security analyzer — ported 1:1 from analyzers/security.py.

function Find-AzMonSecurityFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Workspace,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($ws in $Workspace) {
        foreach ($pair in @(
                @{ Direction = 'ingestion'; Value = $ws['PublicNetworkAccessForIngestion'] },
                @{ Direction = 'query'; Value = $ws['PublicNetworkAccessForQuery'] })) {
            if ((([string]$pair.Value).ToLowerInvariant()) -eq 'enabled') {
                $dirTitle = (Get-Culture).TextInfo.ToTitleCase($pair.Direction)
                $findings.Add((New-AzMonFinding -Category 'security' -Severity $(if (Test-AzMonProdTag $ws['Tags']) { 'high' } else { 'medium' }) `
                    -Title "$($ws['Name']): public network access for $($pair.Direction) is Enabled" `
                    -Detail ('Any client with a valid workspace key or Entra token can reach the workspace over ' +
                        'the public internet. WAF requires network isolation for prod workloads via a Private ' +
                        'Link Scope (AMPLS) or firewall.') `
                    -ResourceIds @($ws['Id']) `
                    -Recommendation ("Set publicNetworkAccessFor$dirTitle to 'Disabled' and attach the workspace " +
                        'to an Azure Monitor Private Link Scope. Update agents and app-side connection strings ' +
                        'before cutover to avoid ingestion outage.') `
                    -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/fundamentals/private-link-security' `
                    -Evidence @{ direction = $pair.Direction; value = $pair.Value }))
            }
        }
    }

    foreach ($ws in $Workspace) {
        if ($ws['DisableLocalAuth'] -eq $false) {
            $findings.Add((New-AzMonFinding -Category 'security' -Severity 'high' `
                -Title "$($ws['Name']): workspace shared-key (local) authentication is enabled" `
                -Detail ('Shared-key auth bypasses Entra ID conditional access and MFA. Rotating compromised keys ' +
                    'requires updating every agent/app manually. WAF and Defender for Cloud both recommend ' +
                    'Entra-only ingestion.') `
                -ResourceIds @($ws['Id']) `
                -Recommendation ('Roll out DCR-based agents with managed identity, then set ' +
                    'features.disableLocalAuth = true on the workspace. Retire legacy Log Analytics Agent (MMA) ' +
                    'as part of the change.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/logs/api/overview#microsoft-entra-authentication-for-workspace-data'))
        }
    }

    foreach ($ai in $AppInsight) {
        foreach ($pair in @(
                @{ Direction = 'ingestion'; Value = $ai['PublicNetworkAccessForIngestion'] },
                @{ Direction = 'query'; Value = $ai['PublicNetworkAccessForQuery'] })) {
            if ((([string]$pair.Value).ToLowerInvariant()) -eq 'enabled') {
                $dirTitle = (Get-Culture).TextInfo.ToTitleCase($pair.Direction)
                $findings.Add((New-AzMonFinding -Category 'security' -Severity 'medium' `
                    -Title "App Insights '$($ai['Name'])': public network access for $($pair.Direction) is Enabled" `
                    -Detail ('Telemetry ingestion and query are reachable from the public internet. For prod apps ' +
                        'behind Private Link, telemetry becomes the weakest link for data exfiltration.') `
                    -ResourceIds @($ai['Id']) `
                    -Recommendation ("Attach the App Insights resource to an AMPLS and set " +
                        "publicNetworkAccessFor$dirTitle to 'Disabled'. Ensure app hosts have a private endpoint " +
                        'route to the AMPLS.') `
                    -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/fundamentals/private-link-security'))
            }
        }
    }

    foreach ($ai in $AppInsight) {
        if ($ai['DisableLocalAuth'] -eq $false) {
            $findings.Add((New-AzMonFinding -Category 'security' -Severity 'medium' `
                -Title "App Insights '$($ai['Name'])': instrumentation-key auth (local auth) is enabled" `
                -Detail ('When local auth is enabled, any leaked instrumentation key allows telemetry injection ' +
                    'or exfiltration. WAF recommends Entra-only ingestion for prod telemetry pipelines.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Set DisableLocalAuth = true on the AI resource, roll app-side to the Azure ' +
                    'Monitor OpenTelemetry Distro with managed identity, then remove any embedded instrumentation ' +
                    'keys from config/secrets.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/app/azure-ad-authentication'))
        }
    }

    foreach ($ai in $AppInsight) {
        if (-not $ai['WorkspaceResourceId']) {
            $findings.Add((New-AzMonFinding -Category 'security' -Severity 'high' `
                -Title "App Insights '$($ai['Name'])': classic (non-workspace-based) resource" `
                -Detail ('Classic App Insights is deprecated (retirement Feb 2024 completed). It cannot ' +
                    'participate in RBAC-scoped queries, cross-workspace joins, or workspace-level security ' +
                    'controls like Private Link and CMK.') `
                -ResourceIds @($ai['Id']) `
                -Recommendation ('Migrate to workspace-based Application Insights (in-place, no data loss). Use ' +
                    'az monitor app-insights component update --workspace ...') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/app/convert-classic-resource'))
        }
    }

    $wsById = @{}
    foreach ($ws in $Workspace) { $wsById[([string]$ws['Id']).ToLowerInvariant()] = $ws }
    foreach ($ai in $AppInsight) {
        if (-not $ai['WorkspaceResourceId']) { continue }
        $ws = $wsById[([string]$ai['WorkspaceResourceId']).ToLowerInvariant()]
        if (-not $ws) { continue }
        $aiPub = ([string]$ai['PublicNetworkAccessForIngestion']).ToLowerInvariant()
        $wsPub = ([string]$ws['PublicNetworkAccessForIngestion']).ToLowerInvariant()
        if ($aiPub -and $wsPub -and $aiPub -ne $wsPub) {
            $findings.Add((New-AzMonFinding -Category 'security' -Severity 'medium' `
                -Title "App Insights '$($ai['Name'])' network policy diverges from parent workspace '$($ws['Name'])'" `
                -Detail ('Ingestion policy is inconsistent between the AI resource and the underlying workspace. ' +
                    'Attackers can pivot through the more permissive endpoint. Both must be aligned for effective ' +
                    'network isolation.') `
                -ResourceIds @($ai['Id'], $ws['Id']) `
                -Recommendation ('Align publicNetworkAccessForIngestion on both resources; if using an AMPLS, add ' +
                    'both to the same private link scope.') `
                -LearnMoreLink 'https://learn.microsoft.com/en-us/azure/azure-monitor/fundamentals/private-link-security' `
                -Evidence @{ ai = $aiPub; workspace = $wsPub }))
        }
    }

    return $findings.ToArray()
}
