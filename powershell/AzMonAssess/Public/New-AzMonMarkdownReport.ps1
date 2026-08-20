#requires -Version 7.0
# Markdown report generator — ported 1:1 from reports/markdown_report.py.

function New-AzMonMarkdownReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $Path
    )

    $findingsRaw = @($Snapshot['Findings'])
    $totalSavings = (@($findingsRaw | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum
    $findings = Sort-AzMonFinding -Finding $findingsRaw

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add("# Azure Monitoring & Observability Assessment — $($Snapshot['CustomerName'])")
    [void]$lines.Add('')
    $generatedAt = [datetime]$Snapshot['GeneratedAt']
    [void]$lines.Add("_Generated $($generatedAt.ToString('yyyy-MM-dd HH:mm')) UTC_")
    [void]$lines.Add('')
    [void]$lines.Add('## Environment')
    [void]$lines.Add('')
    [void]$lines.Add('| Metric | Value |')
    [void]$lines.Add('|---|---|')
    [void]$lines.Add("| Subscriptions | $(@($Snapshot['SubscriptionIds']).Count) |")
    [void]$lines.Add("| Log Analytics workspaces | $(@($Snapshot['Workspaces']).Count) |")
    [void]$lines.Add("| App Insights components | $(@($Snapshot['AppInsights']).Count) |")
    [void]$lines.Add("| Alert rules | $(@($Snapshot['AlertRules']).Count) |")
    [void]$lines.Add("| Action groups | $(@($Snapshot['ActionGroups']).Count) |")
    [void]$lines.Add("| Data collection rules | $(@($Snapshot['DataCollectionRules']).Count) |")
    [void]$lines.Add("| Findings | $($findings.Count) |")
    [void]$lines.Add("| Est. monthly savings (directional) | **$(Format-AzMonUsd $totalSavings)** |")
    [void]$lines.Add('')

    if ($Snapshot['AiSummary']) {
        [void]$lines.Add('## AI Executive Summary')
        [void]$lines.Add('')
        [void]$lines.Add([string]$Snapshot['AiSummary'])
        [void]$lines.Add('')
    }

    [void]$lines.Add('## Findings')
    [void]$lines.Add('')
    foreach ($f in $findings) {
        $savingsText = if ($f['EstimatedMonthlySavingsUsd']) { " — est. $(Format-AzMonUsd $f['EstimatedMonthlySavingsUsd'])/mo" } else { '' }
        [void]$lines.Add("### [$(([string]$f['Severity']).ToUpperInvariant())] [$($f['Category'])] $($f['Title'])$savingsText")
        [void]$lines.Add('')
        [void]$lines.Add([string]$f['Detail'])
        [void]$lines.Add('')
        if ($f['Recommendation']) {
            [void]$lines.Add("**Recommendation:** $($f['Recommendation'])")
            [void]$lines.Add('')
        }
        if (@($f['ResourceIds']).Count -gt 0) {
            [void]$lines.Add("_Affects $(@($f['ResourceIds']).Count) resource(s)._")
            [void]$lines.Add('')
        }
    }

    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    ($lines -join "`n") | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
}
