#requires -Version 7.0
# Executive PowerPoint deck — ported from reports/pptx_report.py, using the
# dependency-free Private/PptxBuilder.ps1 (rectangles/text instead of native
# chart parts, so no Office install / COM / external library is required).

# Brand palette matches the customer's real exec-summary template theme
# (Segoe Sans Text, Microsoft brand colors) - see PptxBuilder.ps1's theme XML.
$script:AzMonPptxNavy = '091F2C'
$script:AzMonPptxBlue = '0078D4'
$script:AzMonPptxCritical = 'F4364F'
$script:AzMonPptxHigh = 'FF5C39'
$script:AzMonPptxMedium = '49C5B1'
$script:AzMonPptxLow = '07641D'
$script:AzMonPptxInfo = '454142'
$script:AzMonPptxWhite = 'FFFFFF'
$script:AzMonPptxGrey = '454142'
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
    Add-AzMonPptxWhatsGoingWellSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxTopFindingsSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxImpactRecommendationsSlide -Deck $deck -Snapshot $Snapshot -Tier 'High' -Severity @('critical', 'high') -ColorHex $script:AzMonPptxCritical
    Add-AzMonPptxImpactRecommendationsSlide -Deck $deck -Snapshot $Snapshot -Tier 'Medium' -Severity @('medium') -ColorHex $script:AzMonPptxHigh
    Add-AzMonPptxImpactRecommendationsSlide -Deck $deck -Snapshot $Snapshot -Tier 'Low' -Severity @('low', 'info') -ColorHex $script:AzMonPptxLow
    Add-AzMonPptxServiceHealthSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxImpactedResourcesSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxSavingsSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxConsolidationSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxTracingSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxRoadmapSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxResourcesSlide -Deck $deck -Snapshot $Snapshot
    Add-AzMonPptxNextStepsSlide -Deck $deck -Snapshot $Snapshot

    Save-AzMonPptxDeck -Deck $deck -Path $Path -Title "Azure Monitoring Assessment — $($Snapshot['CustomerName'])"
    return (Resolve-Path -LiteralPath $Path).Path
}

function Add-AzMonPptxTitleSlide {
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxRect -Slide $slide -X 0 -Y 0 -Width 13.333 -Height 7.5 -ColorHex $script:AzMonPptxNavy
    # Aspect ratio (~4.67:1) matches the source EMF's own native size.
    Add-AzMonPptxPicture -Slide $slide -X 0.6 -Y 0.5 -Width 1.8 -Height 0.3854 -RelId 'rId2' -Name 'Microsoft logo' -Description 'Microsoft logo'
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

function Add-AzMonPptxWhatsGoingWellSlide {
    <#
    .SYNOPSIS
        Positive-findings counterpart to the severity/category slides:
        what's already configured correctly, from ComplianceItems - the
        same "good state" data now in the Excel Compliant Resources sheet.
    #>
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'What is going well' -Subtitle 'Monitoring & Observability capabilities already in place'

    $items = @($Snapshot['ComplianceItems'] | Sort-Object -Property @{ Expression = { @($_['ResourceIds']).Count } } -Descending | Select-Object -First 9)
    if ($items.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 11 -Height 1 -Text 'No compliance data available for this assessment.' -Size 18 -ColorHex $script:AzMonPptxGrey
        return
    }
    $bullets = @($items | ForEach-Object { ([string]$_['Title']).Substring(0, [Math]::Min(110, ([string]$_['Title']).Length)) })
    Add-AzMonPptxBullets -Slide $slide -X 0.7 -Y 1.5 -Width 12 -Height 5.5 -Size 16 -Bullet $bullets
}

function Add-AzMonPptxImpactRecommendationsSlide {
    <#
    .SYNOPSIS
        Full recommendations table for one impact tier (High/Medium/Low),
        mapping azmon-assess's 5 severities into the 3-tier structure used
        by the reference exec-summary template (High=critical+high,
        Medium=medium, Low=low+info).
    #>
    param($Deck, [hashtable] $Snapshot, [string] $Tier, [string[]] $Severity, [string] $ColorHex)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title "$Tier impact issues - Recommendations"
    Add-AzMonPptxRect -Slide $slide -X 0.5 -Y 1.15 -Width 2.0 -Height 0.35 -ColorHex $ColorHex
    Add-AzMonPptxText -Slide $slide -X 0.5 -Y 1.2 -Width 2.0 -Height 0.3 -Text $Tier.ToUpperInvariant() -Size 13 -Bold -ColorHex $script:AzMonPptxWhite -Align 'ctr'

    $matching = @(Sort-AzMonFinding -Finding @($Snapshot['Findings'] | Where-Object { $Severity -contains $_['Severity'] }))
    # Stable-sort on top of Sort-AzMonFinding's severity/savings order, so ties in
    # resource count still fall back to that order instead of an arbitrary one.
    $matching = @($matching | Sort-Object -Descending -Property @{ Expression = { @($_['ResourceIds'] | Where-Object { $_ }).Count } })
    $maxRows = 9
    $shown = @($matching | Select-Object -First $maxRows)
    $header = @('#', 'Recommendation', 'Category', 'Impacted Resources')
    $colWidth = @(0.5, 7.3, 2.5, 2.0)
    $rows = for ($i = 0; $i -lt $shown.Count; $i++) {
        $f = $shown[$i]
        ,@(
            ($i + 1),
            (([string]$f['Title']).Substring(0, [Math]::Min(95, ([string]$f['Title']).Length))),
            (Get-AzMonCategoryLabel $f['Category']),
            ([Math]::Max(1, @($f['ResourceIds'] | Where-Object { $_ }).Count))
        )
    }
    if ($shown.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text "No $Tier-impact findings." -Size 18 -ColorHex $script:AzMonPptxGrey
        return
    }
    Add-AzMonPptxTable -Slide $slide -X 0.5 -Y 1.65 -Header $header -ColumnWidth $colWidth -Row $rows -RowHeight 0.5 | Out-Null
    if ($matching.Count -gt $maxRows) {
        Add-AzMonPptxText -Slide $slide -X 0.5 -Y 7.0 -Width 12 -Height 0.35 -Text "+ $($matching.Count - $maxRows) more $Tier-impact finding(s) — see the Excel Action Plan / Impacted Resources sheets for the full list." -Size 11 -ColorHex $script:AzMonPptxGrey
    }
}

function Add-AzMonPptxServiceHealthSlide {
    <#
    .SYNOPSIS
        Per-subscription Service Health alert coverage table, combining the
        reliability.service-health-alert-missing finding (not covered) and
        its ComplianceItem counterpart (covered) into one Yes/No table.
    #>
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Service Health Alerts for Resiliency' -Subtitle 'Coverage per subscription for Azure service outages, planned maintenance, and advisories'

    $uncovered = @($Snapshot['Findings'] | Where-Object { $_['CheckId'] -eq 'reliability.service-health-alert-missing' } | ForEach-Object { $_['ResourceIds'] })
    $covered = @($Snapshot['ComplianceItems'] | Where-Object { $_['CheckId'] -eq 'reliability.service-health-alert-missing' } | ForEach-Object { $_['ResourceIds'] })
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($subId in $uncovered) { $rows.Add(@(($subId -replace '^/subscriptions/', ''), 'No')) }
    foreach ($subId in $covered) { $rows.Add(@(($subId -replace '^/subscriptions/', ''), 'Yes')) }

    if ($rows.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 10 -Height 1 -Text 'No subscription data available for this check.' -Size 18 -ColorHex $script:AzMonPptxGrey
        return
    }
    $header = @('Subscription Id', 'Service Health Alert Configured?')
    $colWidth = @(8.8, 3.5)
    $shown = @($rows | Sort-Object { $_[1] } | Select-Object -First 12)
    Add-AzMonPptxTable -Slide $slide -X 0.5 -Y 1.7 -Header $header -ColumnWidth $colWidth -Row $shown -RowHeight 0.42 | Out-Null
    if ($rows.Count -gt 12) {
        Add-AzMonPptxText -Slide $slide -X 0.5 -Y 7.0 -Width 12 -Height 0.35 -Text "+ $($rows.Count - 12) more subscription(s) — see report.xlsx for the full list." -Size 11 -ColorHex $script:AzMonPptxGrey
    }
}

function Add-AzMonPptxResourcesSlide {
    <#
    .SYNOPSIS
        Q&A / further-reading slide built from the verified Learn More
        links already attached to findings — one representative link per
        category, so nothing here is a guessed/fabricated URL.
    #>
    param($Deck, [hashtable] $Snapshot)
    $slide = New-AzMonPptxSlide -Deck $Deck
    Add-AzMonPptxSlideHeader -Slide $slide -Title 'Q&A and Resources' -Subtitle 'Official Microsoft Learn guidance referenced by this assessment''s findings'

    $byCategory = @{}
    foreach ($f in @($Snapshot['Findings'])) {
        if (-not $f['LearnMoreLink']) { continue }
        $cat = $f['Category'] ?? 'other'
        if (-not $byCategory.ContainsKey($cat)) { $byCategory[$cat] = $f['LearnMoreLink'] }
    }
    if ($byCategory.Count -eq 0) {
        Add-AzMonPptxText -Slide $slide -X 1 -Y 3 -Width 11 -Height 1 -Text 'No Learn More links available for this assessment.' -Size 18 -ColorHex $script:AzMonPptxGrey
        return
    }
    $y = 1.5
    foreach ($cat in ($byCategory.Keys | Sort-Object)) {
        Add-AzMonPptxText -Slide $slide -X 0.7 -Y $y -Width 3.2 -Height 0.45 -Text (Get-AzMonCategoryLabel $cat) -Size 14 -Bold -ColorHex $script:AzMonPptxNavy
        Add-AzMonPptxText -Slide $slide -X 4.0 -Y $y -Width 8.6 -Height 0.45 -Text $byCategory[$cat] -Size 12 -ColorHex $script:AzMonPptxBlue
        $y += 0.55
    }
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
    # Must match every collection New-AzMonResourceLookup draws from, since that's
    # the universe $byResource's finding-ResourceIds can resolve into - excluding
    # any of these previously under-counted the denominator and pushed pct over 100.
    $totalMonitorable = @($Snapshot['Resources']).Count + @($Snapshot['Workspaces']).Count + @($Snapshot['AppInsights']).Count `
        + @($Snapshot['AlertRules']).Count + @($Snapshot['ActionGroups']).Count + @($Snapshot['DataCollectionRules']).Count
    if ($totalMonitorable -gt 0) {
        # Still cap defensively: a finding can reference a subscription-level scope
        # (e.g. missing service health alert) that isn't part of any inventory count.
        $pct = [Math]::Min(100, [Math]::Round(($byResource.Count / $totalMonitorable) * 100, 0))
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
