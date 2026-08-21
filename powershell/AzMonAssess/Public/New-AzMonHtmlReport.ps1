#requires -Version 7.0
# HTML report generator — dependency-free port of reports/templates/report.html.j2
# (no Jinja2 available in PowerShell; pure string-building instead, with
# explicit HTML output-encoding on every dynamic value since finding text,
# resource names, and evidence ultimately derive from Azure resource
# names/tags that could contain untrusted characters).

function ConvertTo-AzMonHtmlText {
    <#
    .SYNOPSIS
        HTML-encodes a string for safe embedding as text content (used for
        every dynamic value in the HTML report — defensive output encoding
        against markup/script injection from Azure resource names, tags,
        or descriptions).
    #>
    [CmdletBinding()]
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;')
}

$script:AzMonHtmlCss = @'
  :root {
    --bg: #0f1420; --panel: #1a2130; --text: #e6ecf5; --muted: #94a3b8;
    --critical: #ef4444; --high: #f59e0b; --medium: #3b82f6; --low: #10b981; --info: #6b7280;
    --accent: #38bdf8;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: 'Segoe UI', Roboto, sans-serif; background: var(--bg); color: var(--text); }
  header { padding: 32px 48px; background: linear-gradient(135deg, #1e3a8a 0%, #0f1420 100%); }
  header h1 { margin: 0; font-size: 28px; }
  header .sub { color: var(--muted); margin-top: 6px; }
  .container { max-width: 1200px; margin: 0 auto; padding: 24px 48px; }
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin: 24px 0; }
  .kpi { background: var(--panel); border-radius: 8px; padding: 20px; border-left: 4px solid var(--accent); }
  .kpi .label { color: var(--muted); font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi .value { font-size: 28px; font-weight: 600; margin-top: 8px; }
  .kpi.savings { border-left-color: var(--low); }
  .ai-summary { background: var(--panel); border-radius: 8px; padding: 24px; margin: 24px 0; border-left: 4px solid #a855f7; }
  .ai-summary h2 { margin-top: 0; }
  .ai-summary pre { white-space: pre-wrap; word-wrap: break-word; font-family: 'Segoe UI', Roboto, sans-serif; }
  .section { margin: 32px 0; }
  .section h2 { border-bottom: 1px solid #334155; padding-bottom: 8px; }
  .finding { background: var(--panel); border-radius: 8px; padding: 18px 22px; margin: 12px 0; border-left: 4px solid var(--medium); }
  .finding.critical { border-left-color: var(--critical); }
  .finding.high { border-left-color: var(--high); }
  .finding.medium { border-left-color: var(--medium); }
  .finding.low { border-left-color: var(--low); }
  .finding.info { border-left-color: var(--info); }
  .finding-head { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
  .badge.critical { background: var(--critical); color: white; }
  .badge.high { background: var(--high); color: white; }
  .badge.medium { background: var(--medium); color: white; }
  .badge.low { background: var(--low); color: white; }
  .badge.info { background: var(--info); color: white; }
  .badge.cat { background: #334155; color: var(--text); }
  .savings-tag { background: #065f46; color: #d1fae5; padding: 3px 10px; border-radius: 12px; font-size: 12px; }
  .finding h3 { margin: 8px 0; font-size: 16px; }
  .finding .detail { color: var(--muted); font-size: 14px; line-height: 1.6; }
  .finding .rec { margin-top: 12px; padding: 10px 14px; background: rgba(56, 189, 248, 0.08); border-left: 3px solid var(--accent); font-size: 14px; }
  .finding .learn-more { margin-top: 8px; font-size: 13px; }
  .finding .learn-more a { color: var(--accent); text-decoration: none; }
  .finding .learn-more a:hover { text-decoration: underline; }
  details { margin-top: 10px; }
  summary { cursor: pointer; color: var(--accent); font-size: 13px; }
  pre.evidence { background: #0b1220; padding: 12px; border-radius: 6px; overflow-x: auto; font-size: 12px; white-space: pre-wrap; word-wrap: break-word; }
  footer { padding: 24px 48px; color: var(--muted); font-size: 12px; text-align: center; }
'@

function New-AzMonHtmlReport {
    <#
    .SYNOPSIS
        Renders a self-contained dark-theme HTML report from a snapshot —
        a dependency-free port of reports/templates/report.html.j2 (no
        Jinja2 available in PowerShell).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $Path
    )

    $findingsRaw = @($Snapshot['Findings'])
    $totalSavings = (@($findingsRaw | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum
    $findings = Sort-AzMonFinding -Finding $findingsRaw
    $customerName = ConvertTo-AzMonHtmlText ([string]$Snapshot['CustomerName'])
    $generatedAt = [datetime]$Snapshot['GeneratedAt']

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="utf-8">')
    [void]$sb.AppendLine("<title>Azure Monitoring Assessment — $customerName</title>")
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine($script:AzMonHtmlCss)
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')
    [void]$sb.AppendLine('<header>')
    [void]$sb.AppendLine('  <h1>Azure Monitoring &amp; Observability Assessment</h1>')
    [void]$sb.AppendLine("  <div class=""sub"">$customerName — generated $($generatedAt.ToString('yyyy-MM-dd HH:mm')) UTC</div>")
    [void]$sb.AppendLine('</header>')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('<div class="container">')
    [void]$sb.AppendLine('')

    # ---- KPI grid ----------------------------------------------------
    [void]$sb.AppendLine('  <div class="kpi-grid">')
    $kpis = @(
        @{ Label = 'Subscriptions'; Value = @($Snapshot['SubscriptionIds']).Count }
        @{ Label = 'Log Analytics workspaces'; Value = @($Snapshot['Workspaces']).Count }
        @{ Label = 'App Insights components'; Value = @($Snapshot['AppInsights']).Count }
        @{ Label = 'Alert rules'; Value = @($Snapshot['AlertRules']).Count }
        @{ Label = 'Data collection rules'; Value = @($Snapshot['DataCollectionRules']).Count }
        @{ Label = 'Findings'; Value = $findings.Count }
    )
    foreach ($k in $kpis) {
        [void]$sb.AppendLine("    <div class=""kpi""><div class=""label"">$($k.Label)</div><div class=""value"">$($k.Value)</div></div>")
    }
    [void]$sb.AppendLine("    <div class=""kpi savings""><div class=""label"">Est. monthly savings</div><div class=""value"">$(Format-AzMonUsd $totalSavings)</div></div>")
    [void]$sb.AppendLine('  </div>')
    [void]$sb.AppendLine('')

    # ---- AI summary ----------------------------------------------------
    if ($Snapshot['AiSummary']) {
        [void]$sb.AppendLine('  <div class="ai-summary">')
        [void]$sb.AppendLine('    <h2>Executive Summary</h2>')
        [void]$sb.AppendLine("    <pre>$(ConvertTo-AzMonHtmlText ([string]$Snapshot['AiSummary']))</pre>")
        [void]$sb.AppendLine('  </div>')
        [void]$sb.AppendLine('')
    }

    # ---- Findings ----------------------------------------------------
    [void]$sb.AppendLine('  <div class="section">')
    [void]$sb.AppendLine('    <h2>All Findings</h2>')
    foreach ($f in $findings) {
        $severity = ([string]$f['Severity']).ToLowerInvariant()
        $category = ConvertTo-AzMonHtmlText ([string]$f['Category'])
        [void]$sb.AppendLine("    <div class=""finding $severity"">")
        [void]$sb.AppendLine('      <div class="finding-head">')
        [void]$sb.AppendLine("        <span class=""badge $severity"">$severity</span>")
        [void]$sb.AppendLine("        <span class=""badge cat"">$category</span>")
        $savings = [double]($f['EstimatedMonthlySavingsUsd'] ?? 0)
        if ($savings -gt 0) {
            [void]$sb.AppendLine("        <span class=""savings-tag"">est. $(Format-AzMonUsd $savings)/mo saved</span>")
        }
        [void]$sb.AppendLine('      </div>')
        [void]$sb.AppendLine("      <h3>$(ConvertTo-AzMonHtmlText ([string]$f['Title']))</h3>")
        [void]$sb.AppendLine("      <div class=""detail"">$(ConvertTo-AzMonHtmlText ([string]$f['Detail']))</div>")
        if ($f['Recommendation']) {
            [void]$sb.AppendLine("      <div class=""rec""><strong>Recommendation:</strong> $(ConvertTo-AzMonHtmlText ([string]$f['Recommendation']))</div>")
        }
        if ($f['LearnMoreLink']) {
            $learnUrl = ConvertTo-AzMonHtmlText ([string]$f['LearnMoreLink'])
            [void]$sb.AppendLine("      <div class=""learn-more""><a href=""$learnUrl"" target=""_blank"" rel=""noopener noreferrer"">Learn more (Microsoft Learn)</a></div>")
        }

        $resourceIds = @($f['ResourceIds'])
        if ($resourceIds.Count -gt 0) {
            [void]$sb.AppendLine('      <details>')
            [void]$sb.AppendLine("        <summary>$($resourceIds.Count) resource(s)</summary>")
            [void]$sb.Append('        <pre class="evidence">')
            foreach ($rid in ($resourceIds | Select-Object -First 200)) {
                [void]$sb.AppendLine((ConvertTo-AzMonHtmlText ([string]$rid)))
            }
            if ($resourceIds.Count -gt 200) {
                [void]$sb.AppendLine("... ($($resourceIds.Count - 200) more)")
            }
            [void]$sb.AppendLine('</pre>')
            [void]$sb.AppendLine('      </details>')
        }

        $evidence = $f['Evidence']
        if ($evidence -and $evidence.Count -gt 0) {
            [void]$sb.AppendLine('      <details>')
            [void]$sb.AppendLine('        <summary>Evidence</summary>')
            $json = $evidence | ConvertTo-Json -Depth 10
            [void]$sb.AppendLine("        <pre class=""evidence"">$(ConvertTo-AzMonHtmlText $json)</pre>")
            [void]$sb.AppendLine('      </details>')
        }

        [void]$sb.AppendLine('    </div>')
    }
    [void]$sb.AppendLine('  </div>')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('<footer>')
    [void]$sb.AppendLine('  Generated by <code>azmon-assess</code> — read-only assessment.')
    [void]$sb.AppendLine('</footer>')
    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')

    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    $sb.ToString() | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
}
