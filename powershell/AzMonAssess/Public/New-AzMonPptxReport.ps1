#requires -Version 7.0
# Executive PowerPoint deck — ported from reports/pptx_report.py, using the
# dependency-free Private/PptxBuilder.ps1 (rectangles/text instead of native
# chart parts, so no Office install / COM / external library is required).

$script:AzMonPptxNavy = '1F3A5F'
$script:AzMonPptxBlue = '388BFD'
$script:AzMonPptxCritical = 'EF4444'
$script:AzMonPptxHigh = 'F59E0B'
$script:AzMonPptxMedium = '3B82F6'
$script:AzMonPptxLow = '10B981'
$script:AzMonPptxInfo = '6B7280'
$script:AzMonPptxWhite = 'FFFFFF'
$script:AzMonPptxGrey = '64748B'
$script:AzMonPptxSevColor = @{
    critical = $script:AzMonPptxCritical
    high     = $script:AzMonPptxHigh
    medium   = $script:AzMonPptxMedium
    low      = $script:AzMonPptxLow
    info     = $script:AzMonPptxInfo
}

function New-AzMonPptxReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $Path
    )

    $deck = New-AzMonPptxDeck
    Add-AzMonPptxTitleSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxEnvironmentSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxSeveritySlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxCategorySlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxTopFindingsSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxImpactedResourcesSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxSavingsSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxConsolidationSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxTracingSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxRoadmapSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxNextStepsSlide -Deck $deck -Snapshot $Snapshot

    Save-AzMonPptxDeck -Deck $deck -Path $Path -Title "Azure Monitoring Assessment — $($Snapshot['CustomerName'])"
    return (Resolve-Path -LiteralPath $Path).Path
}

function Add-AzMonPptxTitleSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxRect -Slide $slide -X 0 -Y 0 -Width 13.333 -Height 7.5 -ColorHex $script:AzMonPptxNavy
    Add-AzMonPptxRect -Slide $slide -X 0 -Y 3.0 -Width 13.333 -Height 0.06 -ColorHex $script:AzMonPptxBlue
    Add-AzMonPptxText -Slide $slide -X 0.7 -Y 2.0 -Width 12 -Height 1.2 -Text 'Azure Monitoring & Observability' -Size 44 -Bold -ColorHex $script:AzMonPptxWhite
    Add-AzMonPptxText -Slide $slide -X 0.7 -Y 3.2 -Width 12 -Height 0.8 -Text "Assessment Report — $($Snapshot['CustomerName'])" -Size 28 -ColorHex $script:AzMonPptxWhite
    $generatedAt = [datetime]$Snapshot['GeneratedAt']
    Add-AzMonPptxText -Slide $slide -X 0.7 -Y 5.5 -Width 12 -Height 0.4 -Text "Generated $($generatedAt.ToString('MMMM dd, yyyy'))" -Size 14 -ColorHex $script:AzMonPptxBlue
}

function Add-AzMonPptxEnvironmentSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Environment at a glance' -Subtitle 'Read-only inventory across all in-scope subscriptions'

    $kpis = @(
        @{ Label = 'Subscriptions'; Value = @($Snapshot['SubscriptionIds']).Count }
        @{ Label = 'Log Analytics workspaces'; Value = @($Snapshot['Workspaces']).Count }
        @{ Label = 'App Insights'; Value = @($Snapshot['AppInsights']).Count }
        @{ Label = 'Alert rules'; Value = @($Snapshot['AlertRules']).Count }
        @{ Label = 'Action groups'; Value = @($Snapshot['ActionGroups']).Count }
        @{ Label = 'Monitorable resources'; Value = @($Snapshot['Resources']).Count }
    )
    $totalSavings = (@(@($Snapshot['Findings']) | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum

    for ($i = 0; $i -lt $kpis.Count; $i++) {
        $col = $i % 3; $row = [Math]::Floor($i / 3)
        $cx = 0.5 + ($col * 4.3); $cy = 1.5 + ($row * 1.9)
        Add-AzMonPptxKpiCard -Slide $slide -X $cx -Y $cy -Width 4.0 -Height 1.6 -Label $kpis[$i].Label -Value ([string]$kpis[$i].Value) -AccentColorHex $script:AzMonPptxBlue
    }

    $dcrCount = @($Snapshot['DataCollectionRules']).Count
    Add-AzMonPptxText -Slide $slide -X 0.5 -Y 5.05 -Width 12.3 -Height 0.35 `
        -Text "+ $dcrCount data collection rules (Azure Monitor Agent) in scope" -Size 12 -ColorHex $script:AzMonPptxGrey

    Add-AzMonPptxRect -Slide $slide -X 0.5 -Y 5.5 -Width 12.3 -Height 1.5 -ColorHex $script:AzMonPptxNavy
    Add-AzMonPptxText -Slide $slide -X 0.9 -Y 5.65 -Width 6 -Height 0.4 -Text 'TOTAL FINDINGS' -Size 11 -Bold -ColorHex $script:AzMonPptxBlue
    Add-AzMonPptxText -Slide $slide -X 0.9 -Y 5.95 -Width 6 -Height 1.0 -Text ([string]@($Snapshot['Findings']).Count) -Size 36 -Bold -ColorHex $script:AzMonPptxWhite
    Add-AzMonPptxText -Slide $slide -X 6.9 -Y 5.65 -Width 6 -Height 0.4 -Text 'EST. MONTHLY SAVINGS (DIRECTIONAL)' -Size 11 -Bold -ColorHex $script:AzMonPptxBlue
    Add-AzMonPptxText -Slide $slide -X 6.9 -Y 5.95 -Width 6 -Height 1.0 -Text (Format-AzMonUsd $totalSavings) -Size 36 -Bold -ColorHex $script:AzMonPptxWhite
}

function Add-AzMonPptxSeveritySlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Findings by severity'

    $counts = @{ critical = 0; high = 0; medium = 0; low = 0; info = 0 }
    foreach ($f in @($Snapshot['Findings'])) { $counts[$f['Severity']] = ($counts[$f['Severity']] ?? 0) + 1 }
    $order = @('critical', 'high', 'medium', 'low', 'info')

    Add-AzMonPptxHBarChart -Slide $slide -Category @($order | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_) }) `
        -Value @($order | ForEach-Object { [double]$counts[$_] }) -ColorHex @($order | ForEach-Object { $script:AzMonPptxSevColor[$_] }) `
        -X 0.7 -Y 1.5 -Width 8.3 -RowHeight 0.9

    Add-AzMonPptxBullets -Slide $slide -X 9.2 -Y 1.6 -Width 3.6 -Height 5.0 -Bullet @(
        "Critical: $($counts['critical'])"
        "High: $($counts['high'])"
        "Medium: $($counts['medium'])"
        "Low: $($counts['low'])"
        "Info: $($counts['info'])"
        ''
        'Focus this sprint on Critical + High.'
    ) -Size 16
}

function Add-AzMonPptxCategorySlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Findings by category'

    $cats = @{}
    foreach ($f in @($Snapshot['Findings'])) { $cats[$f['Category']] = ($cats[$f['Category']] ?? 0) + 1 }
    if ($cats.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text 'No findings.' -Size 20 -ColorHex $script:AzMonPptxGrey
        return
    }
    $palette = @($script:AzMonPptxBlue, $script:AzMonPptxCritical, $script:AzMonPptxHigh, $script:AzMonPptxMedium, $script:AzMonPptxLow, $script:AzMonPptxInfo, $script:AzMonPptxNavy, $script:AzMonPptxGrey)
    $keys = @($cats.Keys | Sort-Object { -$cats[$_] })
    $colors = @()
    for ($i = 0; $i -lt $keys.Count; $i++) { $colors += $palette[$i % $palette.Count] }

    Add-AzMonPptxShareBar -Slide $slide -Category $keys -Value @($keys | ForEach-Object { [double]$cats[$_] }) -ColorHex $colors -X 0.7 -Y 1.6 -Width 12 -BarHeight 0.9
}

function Add-AzMonPptxTopFindingsSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Top 10 findings' -Subtitle 'Sorted by severity, then estimated monthly savings'

    $top = @(Sort-AzMonFinding -Finding @($Snapshot['Findings']) | Select-Object -First 10)
    $y = 1.3
    foreach ($f in $top) {
        $color = $script:AzMonPptxSevColor[$f['Severity']]
        if (-not $color) { $color = $script:AzMonPptxInfo }
        Add-AzMonPptxRect -Slide $slide -X 0.5 -Y $y -Width 0.15 -Height 0.5 -ColorHex $color
        $savingsText = if ($f['EstimatedMonthlySavingsUsd']) { "  —  $(Format-AzMonUsd $f['EstimatedMonthlySavingsUsd'])/mo" } else { '' }
        Add-AzMonPptxText -Slide $slide -X 0.75 -Y $y -Width 12 -Height 0.5 `
            -Text "[$(([string]$f['Severity']).ToUpperInvariant())]  $($f['Title'])$savingsText" -Size 13 -Bold -ColorHex $script:AzMonPptxNavy
        $y += 0.55
    }
}

function Add-AzMonPptxImpactedResourcesSlide {
    <#
    .SYNOPSIS
        Resource-centric complement to the Excel "4.ImpactedResourcesAnalysis"
        sheet: which resource types carry the most findings, and the top
        individual resources by finding count.
    #>
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Impacted resources' -Subtitle 'Resources referenced by one or more findings'

    $lookup = New-AzMonResourceLookup -Snapshot $Snapshot
    $byResource = @{}
    $byType = @{}
    foreach ($f in @($Snapshot['Findings'])) {
        foreach ($rid in @($f['ResourceIds'])) {
            if (-not $rid) { continue }
            $key = ([string]$rid).ToLowerInvariant()
            if (-not $byResource.ContainsKey($key)) {
                $desc = Resolve-AzMonResourceDescriptor -ResourceId $rid -Lookup $lookup
                $byResource[$key] = @{ Count = 0; Name = $desc.Name; Type = $desc.Type }
            }
            $byResource[$key].Count++
        }
    }
    foreach ($v in $byResource.Values) {
        $t = if ($v.Type) { $v.Type } else { 'unknown' }
        $byType[$t] = ($byType[$t] ?? 0) + 1
    }

    if ($byResource.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text 'No resources referenced by findings.' -Size 20 -ColorHex $script:AzMonPptxGrey
        return
    }

    $topTypes = @($byType.Keys | Sort-Object { -$byType[$_] } | Select-Object -First 8)
    Add-AzMonPptxHBarChart -Slide $slide -Category $topTypes -Value @($topTypes | ForEach-Object { [double]$byType[$_] }) `
        -X 0.7 -Y 1.5 -Width 7.6 -RowHeight 0.55 -LabelWidth 3.2

    Add-AzMonPptxKpiCard -Slide $slide -X 8.7 -Y 1.5 -Width 3.9 -Height 1.5 -Label 'Impacted resources' -Value ([string]$byResource.Count) -AccentColorHex $script:AzMonPptxCritical
    $totalMonitorable = @($Snapshot['Resources']).Count + @($Snapshot['Workspaces']).Count + @($Snapshot['AppInsights']).Count
    if ($totalMonitorable -gt 0) {
        $pct = [Math]::Round(($byResource.Count / $totalMonitorable) * 100, 0)
        Add-AzMonPptxText -Slide $slide -X 8.7 -Y 3.15 -Width 3.9 -Height 0.5 -Text "$pct% of $totalMonitorable inventoried resources" -Size 12 -ColorHex $script:AzMonPptxGrey
    }

    $top = @($byResource.GetEnumerator() | Sort-Object { -$_.Value.Count } | Select-Object -First 6)
    $bullets = [System.Collections.Generic.List[string]]::new()
    [void]$bullets.Add('Top impacted resources:')
    foreach ($kv in $top) {
        $nm = if ($kv.Value.Name) { $kv.Value.Name } else { $kv.Key }
        $plural = if ($kv.Value.Count -ne 1) { 's' } else { '' }
        [void]$bullets.Add("$nm — $($kv.Value.Count) finding$plural")
    }
    Add-AzMonPptxBullets -Slide $slide -X 8.7 -Y 3.75 -Width 3.9 -Height 3.2 -Bullet $bullets.ToArray() -Size 12
}

function Add-AzMonPptxSavingsSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Estimated monthly savings by category'

    $catSavings = @{}
    foreach ($f in @($Snapshot['Findings'])) {
        if ($f['EstimatedMonthlySavingsUsd']) { $catSavings[$f['Category']] = ($catSavings[$f['Category']] ?? 0) + [double]$f['EstimatedMonthlySavingsUsd'] }
    }
    if ($catSavings.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text 'No quantified savings identified.' -Size 20 -ColorHex $script:AzMonPptxGrey
        return
    }
    $keys = @($catSavings.Keys)
    Add-AzMonPptxVBarChart -Slide $slide -Category $keys -Value @($keys | ForEach-Object { [Math]::Round($catSavings[$_], 2) }) `
        -ValueLabel @($keys | ForEach-Object { Format-AzMonUsd $catSavings[$_] }) -ColorHex $script:AzMonPptxBlue -X 0.7 -Y 1.5 -Width 11.9 -Height 4.3

    $total = ($catSavings.Values | Measure-Object -Sum).Sum
    Add-AzMonPptxText -Slide $slide -X 0.7 -Y 6.3 -Width 12 -Height 0.5 `
        -Text "Total: $(Format-AzMonUsd $total)/mo. Figures directional — validate against EA/MCA pricing." -Size 12 -ColorHex $script:AzMonPptxGrey
}

function Add-AzMonPptxConsolidationSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Log Analytics workspace consolidation'

    $cons = @(@($Snapshot['Findings']) | Where-Object { $_['Category'] -eq 'consolidation' })
    if ($cons.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text 'No consolidation candidates.' -Size 20 -ColorHex $script:AzMonPptxGrey
        return
    }
    $bullets = [System.Collections.Generic.List[string]]::new()
    [void]$bullets.Add("$(@($Snapshot['Workspaces']).Count) workspaces detected.")
    foreach ($f in ($cons | Select-Object -Skip 1 -First 5)) {
        $s = if ($f['EstimatedMonthlySavingsUsd']) { " (est. $(Format-AzMonUsd $f['EstimatedMonthlySavingsUsd'])/mo)" } else { '' }
        [void]$bullets.Add("$($f['Title'])$s")
    }
    [void]$bullets.Add('')
    [void]$bullets.Add('Target design: 1 workspace per region x sensitivity tier (prod / non-prod / security).')
    [void]$bullets.Add('Migration approach: dual-shipping via diagnostic settings during a 30-day parallel run.')
    Add-AzMonPptxBullets -Slide $slide -X 0.7 -Y 1.5 -Width 12 -Height 5.5 -Bullet $bullets.ToArray() -Size 15
}

function Add-AzMonPptxTracingSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Distributed tracing readiness'

    $tracing = @(@($Snapshot['Findings']) | Where-Object { $_['Category'] -eq 'tracing' })
    $aiList = @($Snapshot['AppInsights'])
    $aiWsBased = @($aiList | Where-Object { $_['WorkspaceResourceId'] }).Count
    $aiClassic = $aiList.Count - $aiWsBased
    $traceTypes = @('microsoft.web/sites', 'microsoft.app/containerapps', 'microsoft.containerservice/managedclusters', 'microsoft.apimanagement/service', 'microsoft.eventhub/namespaces', 'microsoft.servicebus/namespaces')
    $traceWorthy = @(@($Snapshot['Resources']) | Where-Object { $traceTypes -contains ([string]$_['Type']).ToLowerInvariant() }).Count

    $bullets = [System.Collections.Generic.List[string]]::new()
    [void]$bullets.Add("Application Insights components — workspace-based: $aiWsBased, classic: $aiClassic")
    [void]$bullets.Add("Trace-worthy workloads (web / APIM / AKS / eventing): $traceWorthy")
    if ($tracing.Count -gt 0) {
        [void]$bullets.Add('')
        [void]$bullets.Add('Recommended: Azure Monitor OpenTelemetry Distro rollout')
        [void]$bullets.Add('.NET / Java / Node / Python: install the AOT SDK, set APPLICATIONINSIGHTS_CONNECTION_STRING')
        [void]$bullets.Add('App Service / Container Apps: enable platform auto-instrumentation')
        [void]$bullets.Add('AKS: deploy the OTel Collector as a DaemonSet / sidecar')
        [void]$bullets.Add('Sample at 5-10% trace-based to control cost')
    }
    Add-AzMonPptxBullets -Slide $slide -X 0.7 -Y 1.4 -Width 12 -Height 5.5 -Bullet $bullets.ToArray() -Size 15
}

function Add-AzMonPptxRoadmapSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Recommended roadmap'

    $top = Sort-AzMonFinding -Finding @($Snapshot['Findings'])
    $quick = @($top | Where-Object { $_['Severity'] -in @('critical', 'high') } | Select-Object -First 5)
    $thirty = @($top | Where-Object { $_['Severity'] -eq 'medium' } | Select-Object -First 5)
    $ninety = @($top | Where-Object { $_['Severity'] -in @('low', 'info') } | Select-Object -First 5)
    $columns = @(
        @{ Title = 'This sprint (quick wins)'; Items = $quick; Color = $script:AzMonPptxBlue }
        @{ Title = '30-60 days'; Items = $thirty; Color = $script:AzMonPptxMedium }
        @{ Title = '60-90 days'; Items = $ninety; Color = $script:AzMonPptxLow }
    )
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $x = 0.5 + ($i * 4.3)
        Add-AzMonPptxRect -Slide $slide -X $x -Y 1.4 -Width 4.0 -Height 0.5 -ColorHex $columns[$i].Color
        Add-AzMonPptxText -Slide $slide -X ($x + 0.2) -Y 1.45 -Width 3.8 -Height 0.4 -Text $columns[$i].Title -Size 14 -Bold -ColorHex $script:AzMonPptxWhite
        $items = @($columns[$i].Items | ForEach-Object { ([string]$_['Title']).Substring(0, [Math]::Min(80, ([string]$_['Title']).Length)) })
        if ($items.Count -eq 0) { $items = @('(no items)') }
        Add-AzMonPptxBullets -Slide $slide -X $x -Y 2.0 -Width 4.1 -Height 4.8 -Bullet $items -Size 11
    }
}

function Add-AzMonPptxNextStepsSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Next steps' -Subtitle 'How this assessment transitions into action'
    Add-AzMonPptxBullets -Slide $slide -X 0.7 -Y 1.5 -Width 12 -Height 5.5 -Size 16 -Bullet @(
        '1. Triage - review findings interactively (Invoke-AzMonTriage), assign owners & target dates.'
        '2. Remediate safe items - dry-run first (Invoke-AzMonRemediation), then apply with -Apply.'
        '3. Deploy the Bicep starter pack - action groups, diagnostic-settings policy, alert baseline.'
        '4. Rerun the assessment monthly - track delta and quantify realized savings.'
        '5. Adopt the Azure Monitor OpenTelemetry Distro for end-to-end distributed tracing.'
        '6. Consolidate workspaces on the design above during the next quarter.'
    )
}
