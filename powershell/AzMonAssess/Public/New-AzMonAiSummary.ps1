#requires -Version 7.0
# AI-powered executive summary — Azure OpenAI via REST (Entra ID bearer
# token, no SDK dependency), with a rule-based fallback when AOAI isn't
# configured or the call fails. Ported from ai/recommender.py + prompts.py.

function Get-AzMonAiSystemPrompt {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $CustomerName)
    return @"
You are a Principal Azure Monitoring & Observability architect advising the $CustomerName platform team.

Common concerns this assessment targets:
1. Sprawl of Log Analytics workspaces driving unpredictable cost.
2. Unexpectedly high monitoring bills from misconfigured features.
3. Gaps in alerting - critical issues not surfacing to on-call.
4. Weak signal correlation across infra, network, and application layers.
5. Transition from legacy monitoring (SCOM, Dynatrace, classic App Insights) to Azure-native.
6. Distributed tracing / OpenTelemetry adoption.
7. Retention & long-term archival strategy.

Your job:
- Read the JSON findings (produced by static analyzers) and turn them into a prioritized executive briefing.
- Emphasize DIRECT cost impact and the fewest-actions path to the biggest gains.
- Group recommendations under: "Quick Wins (this sprint)", "30-day roadmap", "90-day roadmap".
- Be specific: reference finding IDs, resource IDs where useful, and Azure CLI or Bicep patterns.
- Never invent facts not present in the findings JSON. If a number is missing, say so; do not make one up.
- Output in Markdown, no preamble, no closing pleasantries.
"@
}

function Get-AzMonAiUserPrompt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $FindingsJson,
        [Parameter(Mandatory)] [string] $CustomerName,
        [int] $SubscriptionCount,
        [int] $WorkspaceCount,
        [int] $AppInsightCount,
        [int] $AlertCount,
        [double] $TotalSavings
    )
    $savingsText = Format-AzMonUsd $TotalSavings
    return @"
Findings JSON (may be large; focus on the top 30 by severity + savings):
``````json
$FindingsJson
``````

Environment context:
- Customer: $CustomerName
- Subscriptions assessed: $SubscriptionCount
- Log Analytics workspaces: $WorkspaceCount
- Application Insights components: $AppInsightCount
- Alert rules: $AlertCount
- Total estimated monthly savings (sum of finding estimates): $savingsText

Produce the briefing now.
"@
}

function Get-AzMonRuleBasedSummary {
    <#
    .SYNOPSIS
        Deterministic Markdown executive summary — used whenever Azure
        OpenAI isn't configured or the call fails.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Snapshot)

    $findings = @($Snapshot['Findings'])
    $byCat = @{}
    foreach ($f in $findings) { $byCat[$f['Category']] = ($byCat[$f['Category']] ?? 0) + 1 }
    $totalSavings = (@($findings | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum
    $top = @(Sort-AzMonFinding -Finding $findings | Select-Object -First 10)

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add("# Azure Monitoring Assessment — $($Snapshot['CustomerName'])")
    [void]$lines.Add('')
    [void]$lines.Add('## Executive Summary')
    [void]$lines.Add('')
    [void]$lines.Add("- Subscriptions assessed: **$(@($Snapshot['SubscriptionIds']).Count)**")
    [void]$lines.Add("- Log Analytics workspaces: **$(@($Snapshot['Workspaces']).Count)**")
    [void]$lines.Add("- Application Insights components: **$(@($Snapshot['AppInsights']).Count)**")
    [void]$lines.Add("- Alert rules: **$(@($Snapshot['AlertRules']).Count)**")
    [void]$lines.Add("- Total findings: **$($findings.Count)**")
    [void]$lines.Add("- Estimated monthly savings (directional): **$(Format-AzMonUsd $totalSavings)**")
    [void]$lines.Add('')
    [void]$lines.Add('## Findings by category')
    [void]$lines.Add('')
    foreach ($cat in $byCat.Keys) { [void]$lines.Add("- **$cat**: $($byCat[$cat]) findings") }
    [void]$lines.Add('')
    [void]$lines.Add('## Top 10 actionable items')
    [void]$lines.Add('')
    $i = 1
    foreach ($f in $top) {
        $savingsText = if ($f['EstimatedMonthlySavingsUsd']) { " — est. $(Format-AzMonUsd $f['EstimatedMonthlySavingsUsd'])/mo" } else { '' }
        [void]$lines.Add("$i. **[$(([string]$f['Severity']).ToUpperInvariant())]** $($f['Title'])$savingsText")
        if ($f['Recommendation']) { [void]$lines.Add("   - $($f['Recommendation'])") }
        $i++
    }
    [void]$lines.Add('')
    [void]$lines.Add('_AI-powered summary was not generated (Azure OpenAI not configured). Set AZURE_OPENAI_ENDPOINT ' +
        'and pass -AoaiEndpoint/-AoaiDeployment to enable._')
    return ($lines -join "`n")
}

function Get-AzMonAoaiSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $Endpoint,
        [Parameter(Mandatory)] [string] $Deployment,
        [Parameter(Mandatory)] [string] $ApiVersion
    )
    $findings = @($Snapshot['Findings'])
    $top = @(Sort-AzMonFinding -Finding $findings | Select-Object -First 30)
    $findingsJson = $top | ConvertTo-Json -Depth 10
    $totalSavings = (@($findings | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum
    $customerName = $Snapshot['CustomerName']

    $systemPrompt = Get-AzMonAiSystemPrompt -CustomerName $customerName
    $userPrompt = Get-AzMonAiUserPrompt -FindingsJson $findingsJson -CustomerName $customerName `
        -SubscriptionCount (@($Snapshot['SubscriptionIds']).Count) -WorkspaceCount (@($Snapshot['Workspaces']).Count) `
        -AppInsightCount (@($Snapshot['AppInsights']).Count) -AlertCount (@($Snapshot['AlertRules']).Count) `
        -TotalSavings $totalSavings

    $token = Get-AzMonBearerToken -ResourceUrl 'https://cognitiveservices.azure.com'
    $uri = "$($Endpoint.TrimEnd('/'))/openai/deployments/$Deployment/chat/completions?api-version=$ApiVersion"
    $body = @{
        messages    = @(
            @{ role = 'system'; content = $systemPrompt }
            @{ role = 'user'; content = $userPrompt }
        )
        temperature = 0.2
        max_tokens  = 2500
    } | ConvertTo-Json -Depth 10

    $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers @{ Authorization = "Bearer $token" } -ContentType 'application/json; charset=utf-8' -Body $body -ErrorAction Stop
    $content = $resp.choices[0].message.content
    if (-not $content) { return Get-AzMonRuleBasedSummary -Snapshot $Snapshot }
    return $content
}

function New-AzMonAiSummary {
    <#
    .SYNOPSIS
        Generates the executive summary — Azure OpenAI when configured
        (endpoint reachable via param or AZURE_OPENAI_ENDPOINT), else a
        deterministic rule-based Markdown summary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [string] $AoaiEndpoint,
        [string] $AoaiDeployment,
        [string] $AoaiApiVersion
    )
    if (-not $AoaiEndpoint) { $AoaiEndpoint = $env:AZURE_OPENAI_ENDPOINT }
    if (-not $AoaiDeployment) { $AoaiDeployment = if ($env:AZURE_OPENAI_DEPLOYMENT) { $env:AZURE_OPENAI_DEPLOYMENT } else { 'gpt-4o' } }
    if (-not $AoaiApiVersion) { $AoaiApiVersion = if ($env:AZURE_OPENAI_API_VERSION) { $env:AZURE_OPENAI_API_VERSION } else { '2024-10-21' } }

    if (-not $AoaiEndpoint) {
        Write-Verbose 'Azure OpenAI not configured - using rule-based summary.'
        return Get-AzMonRuleBasedSummary -Snapshot $Snapshot
    }
    try {
        return Get-AzMonAoaiSummary -Snapshot $Snapshot -Endpoint $AoaiEndpoint -Deployment $AoaiDeployment -ApiVersion $AoaiApiVersion
    } catch {
        Write-Warning "AI summary failed ($($_.Exception.Message)) - falling back to rule-based."
        return Get-AzMonRuleBasedSummary -Snapshot $Snapshot
    }
}
