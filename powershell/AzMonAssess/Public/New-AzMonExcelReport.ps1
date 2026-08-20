#requires -Version 7.0
# Excel workbook generator — zero-dependency, built on Private/XlsxBuilder.ps1
# (plain OOXML via System.IO.Compression). Ported from reports/excel_report.py,
# matching its severity color palette. No ImportExcel/EPPlus/Excel install
# required — works anywhere PowerShell 7+ runs, including Azure Cloud Shell.

$script:AzMonSevFillHex = @{
    critical = 'FFC7CE'
    high     = 'FFEB9C'
    medium   = 'FCE4D6'
    low      = 'C6EFCE'
    info     = 'D9E1F2'
}

function New-AzMonExcelReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $Path
    )

    $wb = New-AzMonXlsxWorkbook
    $lookup = New-AzMonResourceLookup -Snapshot $Snapshot

    Add-AzMonExcelSummarySheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelFindingsSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelImpactedResourcesSheet -Workbook $wb -Snapshot $Snapshot -Lookup $lookup
    Add-AzMonExcelWorkspacesSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelAppInsightsSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelAlertsSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelActionGroupsSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelDiagnosticsSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelDataCollectionRulesSheet -Workbook $wb -Snapshot $Snapshot
    Add-AzMonExcelConsolidationSheet -Workbook $wb -Snapshot $Snapshot

    return (Save-AzMonXlsxWorkbook -Workbook $wb -Path $Path -Title "Azure Monitoring Assessment — $($Snapshot['CustomerName'])")
}

# ---- label helpers (Findings + Impacted Resources Analysis sheets) -------

function Get-AzMonCategoryLabel {
    param([string] $Category)
    switch ($Category) {
        'consolidation' { 'Consolidation' }
        'coverage' { 'Coverage Gaps' }
        'alerting' { 'Alert Quality' }
        'cost' { 'Cost Optimization' }
        'tracing' { 'Tracing Readiness' }
        'reliability' { 'Reliability' }
        'security' { 'Security' }
        'performance' { 'Performance Efficiency' }
        default { $Category }
    }
}

function Get-AzMonWafPillar {
    param([string] $Category)
    switch ($Category) {
        'cost' { 'Cost Optimization' }
        'consolidation' { 'Cost Optimization' }
        'security' { 'Security' }
        'reliability' { 'Reliability' }
        'performance' { 'Performance Efficiency' }
        default { 'Operational Excellence' }
    }
}

# ---- sheet builders ------------------------------------------------------

function Add-AzMonExcelSummarySheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $findings = @($Snapshot['Findings'])
    $totalSavings = (@($findings | ForEach-Object { [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) }) | Measure-Object -Sum).Sum
    $sevCounts = @{ critical = 0; high = 0; medium = 0; low = 0; info = 0 }
    $catCounts = @{}
    foreach ($f in $findings) {
        $sevCounts[$f['Severity']] = ($sevCounts[$f['Severity']] ?? 0) + 1
        $catCounts[$f['Category']] = ($catCounts[$f['Category']] ?? 0) + 1
    }
    $generatedAt = [datetime]$Snapshot['GeneratedAt']

    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Summary'

    Add-AzMonXlsxRow -Sheet $ws -Cell @((New-AzMonXlsxCell -Value "Azure Monitoring Assessment — $($Snapshot['CustomerName'])" -AsString -Bold -FontColorHex '1F4E78' -FontSizePt 16)) | Out-Null
    Add-AzMonXlsxRow -Sheet $ws -Cell @((New-AzMonXlsxCell -Value "Generated $($generatedAt.ToString('yyyy-MM-dd HH:mm')) UTC" -AsString)) | Out-Null
    Add-AzMonXlsxRow -Sheet $ws -Cell @() | Out-Null
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Metric', 'Value') | Out-Null

    $metricRows = [System.Collections.Generic.List[object[]]]::new()
    $metricRows.Add(@('Environment', ''))
    $metricRows.Add(@('Subscriptions', [double]@($Snapshot['SubscriptionIds']).Count))
    $metricRows.Add(@('Log Analytics workspaces', [double]@($Snapshot['Workspaces']).Count))
    $metricRows.Add(@('App Insights components', [double]@($Snapshot['AppInsights']).Count))
    $metricRows.Add(@('Alert rules', [double]@($Snapshot['AlertRules']).Count))
    $metricRows.Add(@('Action groups', [double]@($Snapshot['ActionGroups']).Count))
    $metricRows.Add(@('Monitorable resources', [double]@($Snapshot['Resources']).Count))
    $metricRows.Add(@('Diagnostic settings', [double]@($Snapshot['DiagnosticSettings']).Count))
    $metricRows.Add(@('Data collection rules', [double]@($Snapshot['DataCollectionRules']).Count))
    $metricRows.Add(@('', ''))
    $metricRows.Add(@('Findings', ''))
    $metricRows.Add(@('Total findings', [double]$findings.Count))
    $metricRows.Add(@('Estimated monthly savings (USD, directional)', [Math]::Round($totalSavings, 2)))
    $metricRows.Add(@('', ''))
    $metricRows.Add(@('Findings by severity', ''))
    foreach ($s in @('critical', 'high', 'medium', 'low', 'info')) {
        $metricRows.Add(@((Get-Culture).TextInfo.ToTitleCase($s), [double]$sevCounts[$s]))
    }
    $metricRows.Add(@('', ''))
    $metricRows.Add(@('Findings by category', ''))
    foreach ($cat in ($catCounts.Keys | Sort-Object { -$catCounts[$_] })) {
        $metricRows.Add(@((Get-AzMonCategoryLabel $cat), [double]$catCounts[$cat]))
    }
    foreach ($r in $metricRows) { Add-AzMonXlsxRow -Sheet $ws -Cell $r | Out-Null }

    if ($Snapshot['AiSummary']) {
        Add-AzMonXlsxRow -Sheet $ws -Cell @() | Out-Null
        Add-AzMonXlsxRow -Sheet $ws -Cell @((New-AzMonXlsxCell -Value 'AI Executive Summary' -AsString -Bold -FontColorHex '6B21A8' -FontSizePt 13)) | Out-Null
        $textRow = Add-AzMonXlsxRow -Sheet $ws -Cell @((New-AzMonXlsxCell -Value ([string]$Snapshot['AiSummary']) -AsString -WrapText -VerticalTop)) -HeightPt 300
        Merge-AzMonXlsxCells -Sheet $ws -Range "A$textRow`:F$textRow"
    }

    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(50, 20)
}

function Add-AzMonExcelFindingsSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $findings = Sort-AzMonFinding -Finding @($Snapshot['Findings'])
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Findings'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Severity', 'Category', 'Title', 'Detail', 'Recommendation', 'Est. $/mo Savings', 'Affected Resources', 'Finding ID') | Out-Null

    if ($findings.Count -eq 0) {
        Add-AzMonXlsxRow -Sheet $ws -Cell @('', '', 'No findings', '', '', 0.0, 0.0, '') | Out-Null
    }
    foreach ($f in $findings) {
        $cells = @(
            ([string]$f['Severity']).ToUpperInvariant(), (Get-AzMonCategoryLabel $f['Category']), $f['Title'], $f['Detail'],
            ($f['Recommendation'] ?? ''), [double]($f['EstimatedMonthlySavingsUsd'] ?? 0), [double]@($f['ResourceIds']).Count, $f['Id']
        )
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxRowFill -Sheet $ws -Row $rowNum -ColumnCount 8 -Hex $script:AzMonSevFillHex[$f['Severity']]
        Set-AzMonXlsxCellWrap -Sheet $ws -Row $rowNum -Column 4
        Set-AzMonXlsxCellWrap -Sheet $ws -Row $rowNum -Column 5
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(10, 16, 55, 70, 70, 15, 12, 38)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelImpactedResourcesSheet {
    <#
    .SYNOPSIS
        Resource-centric findings view: one row per (finding, resource)
        pair. Column layout mirrors the "4.ImpactedResourcesAnalysis" sheet
        of the reference WARA-style expert-analysis workbook, adapted to
        azmon-assess's own data model. Columns azmon-assess has no source
        for (custom1-5, Learn More Link, Platform Issue / Retirement /
        Support Request tracking IDs) are left blank rather than guessed.
    #>
    param([hashtable] $Workbook, [hashtable] $Snapshot, [hashtable] $Lookup)

    $findings = Sort-AzMonFinding -Finding @($Snapshot['Findings'])
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name '4.ImpactedResourcesAnalysis'
    $header = @(
        'REQUIRED ACTIONS / REVIEW STATUS', 'ValidationCategory', 'Resource Type', 'subscriptionId', 'resourceGroup',
        'location', 'name', 'id', 'custom1', 'custom2', 'custom3', 'custom4', 'custom5',
        'Recommendation Title', 'Impact', 'Recommendation Control', 'Potential Benefit', 'Learn More Link',
        'Long Description', 'Category', 'Source', 'WAF Pillar', 'Platform Issue TrackingId', 'Retirement TrackingId',
        'Support Request Number', 'Notes', 'checkName'
    )
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header $header | Out-Null

    if ($findings.Count -eq 0) {
        Add-AzMonXlsxRow -Sheet $ws -Cell (@('No findings') + (1..26 | ForEach-Object { '' })) | Out-Null
    }
    foreach ($f in $findings) {
        $resourceIds = @($f['ResourceIds'])
        if ($resourceIds.Count -eq 0) { $resourceIds = @('') }
        $impact = switch ($f['Severity']) { 'critical' { 'Critical' } 'high' { 'High' } 'medium' { 'Medium' } default { 'Low' } }
        $benefit = if ($f['EstimatedMonthlySavingsUsd']) { "$(Format-AzMonUsd $f['EstimatedMonthlySavingsUsd'])/mo estimated savings" } else { ($f['Recommendation'] ?? '') }
        foreach ($rid in $resourceIds) {
            $desc = if ($rid) { Resolve-AzMonResourceDescriptor -ResourceId $rid -Lookup $Lookup } else { @{ Name = ''; Type = ''; ResourceGroup = ''; Location = ''; SubscriptionId = '' } }
            $cells = @(
                'Not Reviewed', 'Resource', $desc.Type, $desc.SubscriptionId, $desc.ResourceGroup, $desc.Location, $desc.Name, $rid,
                '', '', '', '', '',
                $f['Title'], $impact, (Get-AzMonCategoryLabel $f['Category']), $benefit, '',
                $f['Detail'], 'Azure Monitor', 'azmon-assess', (Get-AzMonWafPillar $f['Category']), '', '', '', '', $f['Id']
            )
            $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
            Set-AzMonXlsxCellWrap -Sheet $ws -Row $rowNum -Column 14
            Set-AzMonXlsxCellWrap -Sheet $ws -Row $rowNum -Column 17
            Set-AzMonXlsxCellWrap -Sheet $ws -Row $rowNum -Column 19
            $impactFill = $script:AzMonSevFillHex[$f['Severity']]
            if ($impactFill) { $ws.Rows[$rowNum - 1].Cells[14].FillHex = $impactFill }
        }
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(18, 14, 26, 36, 20, 14, 26, 60, 10, 10, 10, 10, 10, 40, 10, 20, 30, 24, 60, 14, 12, 20, 16, 16, 16, 20, 36)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelWorkspacesSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $workspaces = @($Snapshot['Workspaces'])
    if ($workspaces.Count -eq 0) { return }
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Workspaces'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Name', 'Subscription', 'Resource Group', 'Region', 'SKU', 'Retention (d)', 'Daily Cap (GB)', 'Ingest 30d (GB)', 'Connected Sources', 'Portal', 'Resource ID') | Out-Null
    foreach ($w in $workspaces) {
        $cells = @($w['Name'], $w['SubscriptionId'], $w['ResourceGroup'], $w['Location'], $w['Sku'], $w['RetentionDays'], $w['DailyQuotaGb'], $w['IngestionGb30d'], $w['ConnectedSources'], '', $w['Id'])
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxHyperlink -Sheet $ws -Row $rowNum -Column 10 -Url (ConvertTo-AzMonPortalUrl -ResourceId $w['Id'])
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(30, 24, 20, 14, 14, 12, 12, 14, 14, 10, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelAppInsightsSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $ais = @($Snapshot['AppInsights'])
    if ($ais.Count -eq 0) { return }
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'App Insights'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Name', 'Subscription', 'RG', 'Region', 'Kind', 'App Type', 'Workspace-based?', 'Sampling %', 'Retention (d)', 'Portal', 'Resource ID') | Out-Null
    foreach ($ai in $ais) {
        $wsBased = if ($ai['WorkspaceResourceId']) { 'Yes' } else { 'No (Classic)' }
        $cells = @($ai['Name'], $ai['SubscriptionId'], $ai['ResourceGroup'], $ai['Location'], $ai['AiKind'], $ai['ApplicationType'], $wsBased, $ai['SamplingPercentage'], $ai['RetentionDays'], '', $ai['Id'])
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxHyperlink -Sheet $ws -Row $rowNum -Column 10 -Url (ConvertTo-AzMonPortalUrl -ResourceId $ai['Id'])
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(30, 24, 20, 14, 12, 14, 16, 12, 12, 10, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelAlertsSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $rules = @($Snapshot['AlertRules'])
    if ($rules.Count -eq 0) { return }
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Alert Rules'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Name', 'Kind', 'Enabled', 'Severity', 'Fire Count (30d)', 'Scopes', 'Action Groups', 'Description', 'Portal', 'Resource ID') | Out-Null
    foreach ($r in $rules) {
        $enabled = if ($r['Enabled']) { 'Yes' } else { 'No' }
        $cells = @($r['Name'], $r['AlertKind'], $enabled, $r['Severity'], $r['FireCount30d'], [double]@($r['Scopes']).Count, [double]@($r['ActionGroupIds']).Count, $r['Description'], '', $r['Id'])
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxHyperlink -Sheet $ws -Row $rowNum -Column 9 -Url (ConvertTo-AzMonPortalUrl -ResourceId $r['Id'])
        if ($r['Enabled'] -and @($r['ActionGroupIds']).Count -eq 0) {
            Set-AzMonXlsxRowFill -Sheet $ws -Row $rowNum -ColumnCount 10 -Hex $script:AzMonSevFillHex['high']
        } elseif (([int]($r['FireCount30d'] ?? 0)) -ge 50) {
            Set-AzMonXlsxRowFill -Sheet $ws -Row $rowNum -ColumnCount 10 -Hex $script:AzMonSevFillHex['medium']
        }
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(36, 12, 10, 10, 16, 10, 14, 50, 10, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelActionGroupsSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $groups = @($Snapshot['ActionGroups'])
    if ($groups.Count -eq 0) { return }
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Action Groups'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Name', 'Short Name', 'Email', 'SMS', 'Webhook', 'Logic App', 'ITSM', 'Used By Rules', 'Portal', 'Resource ID') | Out-Null
    foreach ($g in $groups) {
        $cells = @($g['Name'], $g['ShortName'], [double]$g['EmailReceivers'], [double]$g['SmsReceivers'], [double]$g['WebhookReceivers'], [double]$g['LogicAppReceivers'], [double]$g['ItsmReceivers'], [double]$g['UsedByRules'], '', $g['Id'])
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxHyperlink -Sheet $ws -Row $rowNum -Column 9 -Url (ConvertTo-AzMonPortalUrl -ResourceId $g['Id'])
        if ($g['UsedByRules'] -eq 0) {
            Set-AzMonXlsxRowFill -Sheet $ws -Row $rowNum -ColumnCount 10 -Hex $script:AzMonSevFillHex['medium']
        }
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(30, 16, 10, 10, 10, 10, 10, 14, 10, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelDiagnosticsSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $settings = @($Snapshot['DiagnosticSettings'])
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Diagnostic Settings'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Setting Name', 'Target Resource ID', 'Workspace ID', 'Storage ID', 'Event Hub Rule', 'Logs?', 'Metrics?') | Out-Null
    if ($settings.Count -eq 0) {
        Add-AzMonXlsxRow -Sheet $ws -Cell @('', '', '', '', '', '', '') | Out-Null
    }
    foreach ($s in $settings) {
        $logs = if ($s['LogsEnabled']) { 'Yes' } else { 'No' }
        $metrics = if ($s['MetricsEnabled']) { 'Yes' } else { 'No' }
        Add-AzMonXlsxRow -Sheet $ws -Cell @($s['Name'], $s['ResourceId'], $s['WorkspaceId'], $s['StorageId'], $s['EventHubId'], $logs, $metrics) | Out-Null
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(30, 70, 60, 60, 60, 10, 10)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelDataCollectionRulesSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $dcrs = @($Snapshot['DataCollectionRules'])
    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Data Collection Rules'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Name', 'Resource Group', 'Location', 'Kind', 'Data Flows', 'Log Analytics Destination', 'Portal', 'Resource ID') | Out-Null
    if ($dcrs.Count -eq 0) {
        Add-AzMonXlsxRow -Sheet $ws -Cell @('', '', '', '', '', '', '', '') | Out-Null
    }
    foreach ($d in $dcrs) {
        $cells = @($d['Name'], $d['ResourceGroup'], $d['Location'], $d['DcrKind'], [double]$d['DataFlowCount'], $d['WorkspaceResourceId'], '', $d['Id'])
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        Set-AzMonXlsxHyperlink -Sheet $ws -Row $rowNum -Column 7 -Url (ConvertTo-AzMonPortalUrl -ResourceId $d['Id'])
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(30, 20, 14, 14, 12, 60, 10, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

function Add-AzMonExcelConsolidationSheet {
    param([hashtable] $Workbook, [hashtable] $Snapshot)

    $workspaces = @($Snapshot['Workspaces'])
    $groups = @{}
    foreach ($w in $workspaces) {
        $envName = Get-AzMonEnvironmentTag -Workspace $w
        $key = "$(([string]$w['Location']).ToLowerInvariant())|$envName"
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [System.Collections.Generic.List[hashtable]]::new() }
        $groups[$key].Add($w)
    }
    if ($groups.Count -eq 0) { return }

    $ws = Add-AzMonXlsxSheet -Workbook $Workbook -Name 'Consolidation Plan'
    Add-AzMonXlsxHeaderRow -Sheet $ws -Header @('Region', 'Environment', '# Workspaces', 'Sum GB (30d)', 'Daily GB', 'Current PAYG $/mo', 'Best Commitment $/mo', 'Est. Savings $/mo', 'Workspace Names') | Out-Null

    foreach ($key in ($groups.Keys | Sort-Object)) {
        $parts = $key -split '\|', 2
        $items = $groups[$key]
        $gb30 = (@($items | ForEach-Object { [double]($_['IngestionGb30d'] ?? 0) }) | Measure-Object -Sum).Sum
        $daily = $gb30 / 30.0
        $paygMo = $daily * 30.0 * $script:AzMonPaygPricePerGb
        $best = Get-AzMonBestCommitment -DailyGb $daily
        $savings = [Math]::Max($paygMo - $best.MonthlyCost, 0)
        $cells = @(
            $parts[0], $parts[1], [double]$items.Count, [Math]::Round($gb30, 2), [Math]::Round($daily, 2),
            [Math]::Round($paygMo, 2), [Math]::Round($best.MonthlyCost, 2), [Math]::Round($savings, 2),
            (($items | ForEach-Object { $_['Name'] }) -join ', ')
        )
        $rowNum = Add-AzMonXlsxRow -Sheet $ws -Cell $cells
        if ($savings -gt 200) {
            Set-AzMonXlsxRowFill -Sheet $ws -Row $rowNum -ColumnCount 9 -Hex $script:AzMonSevFillHex['high']
        }
    }
    Set-AzMonXlsxColumnWidths -Sheet $ws -Width @(16, 14, 12, 14, 12, 16, 18, 16, 60)
    Set-AzMonXlsxFreezeHeader -Sheet $ws
    Set-AzMonXlsxAutoFilter -Sheet $ws
}

