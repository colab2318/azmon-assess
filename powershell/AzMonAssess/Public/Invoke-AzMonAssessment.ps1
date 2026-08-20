#requires -Version 7.0

function Invoke-AzMonAssessment {
    <#
    .SYNOPSIS
        End-to-end assessment: collect -> analyze -> AI summary -> reports.
        Equivalent to the Python tool's `run` command, plus the focused
        single-topic scans via -Only.
    .EXAMPLE
        Invoke-AzMonAssessment -OutputPath ./out
    .EXAMPLE
        Invoke-AzMonAssessment -OutputPath ./out -Only cost,alerting -Report excel,pptx
    #>
    [CmdletBinding()]
    param(
        [string[]] $SubscriptionId,
        [string] $ManagementGroupId,
        [string] $CustomerName,
        [int] $LookbackDays,
        [int] $ThrottleLimit,
        [Parameter(Mandatory)] [string] $OutputPath,
        [ValidateSet('all', 'consolidation', 'coverage', 'alerting', 'cost', 'tracing', 'reliability', 'security', 'performance')]
        [string[]] $Only = @('all'),
        [switch] $SkipIngestionEnrichment,
        [switch] $SkipAiSummary,
        [ValidateSet('all', 'markdown', 'excel', 'pptx', 'html', 'none')]
        [string[]] $Report = @('all'),
        [string] $AoaiEndpoint,
        [string] $AoaiDeployment,
        [string] $AoaiApiVersion
    )

    if (-not $CustomerName) { $CustomerName = if ($env:AZMON_CUSTOMER_NAME) { $env:AZMON_CUSTOMER_NAME } else { 'Your Organization' } }
    if (-not $LookbackDays) { $LookbackDays = if ($env:AZMON_LOOKBACK_DAYS) { [int]$env:AZMON_LOOKBACK_DAYS } else { 30 } }
    if (-not $ThrottleLimit) { $ThrottleLimit = if ($env:AZMON_MAX_PARALLEL) { [int]$env:AZMON_MAX_PARALLEL } else { 8 } }

    $subs = Resolve-AzMonSubscription -SubscriptionId $SubscriptionId -ManagementGroupId $ManagementGroupId
    if (@($subs).Count -eq 0) { throw 'No subscriptions resolved - check access or pass -SubscriptionId explicitly.' }
    Write-Host "[azmon-assess] Assessing $(@($subs).Count) subscription(s)..." -ForegroundColor Cyan

    Write-Host '[azmon-assess] Collecting Log Analytics workspaces...' -ForegroundColor DarkCyan
    $workspaces = Get-AzMonWorkspace -SubscriptionId $subs -SkipIngestionEnrichment:$SkipIngestionEnrichment -LookbackDays $LookbackDays -ThrottleLimit $ThrottleLimit
    Write-Host "[azmon-assess] Found $(@($workspaces).Count) workspace(s)." -ForegroundColor DarkCyan

    Write-Host '[azmon-assess] Collecting Application Insights...' -ForegroundColor DarkCyan
    $appInsights = Get-AzMonAppInsight -SubscriptionId $subs
    Write-Host "[azmon-assess] Found $(@($appInsights).Count) App Insights component(s)." -ForegroundColor DarkCyan

    Write-Host '[azmon-assess] Collecting alert rules + action groups...' -ForegroundColor DarkCyan
    $alertRules = Get-AzMonAlertRule -SubscriptionId $subs
    $alertRules = Add-AzMonAlertFireRate -SubscriptionId $subs -AlertRule $alertRules
    $actionGroups = Get-AzMonActionGroup -SubscriptionId $subs -AlertRule $alertRules
    Write-Host "[azmon-assess] Found $(@($alertRules).Count) alert rule(s), $(@($actionGroups).Count) action group(s)." -ForegroundColor DarkCyan

    Write-Host '[azmon-assess] Collecting data collection rules (DCRs)...' -ForegroundColor DarkCyan
    $dataCollectionRules = Get-AzMonDataCollectionRule -SubscriptionId $subs
    Write-Host "[azmon-assess] Found $(@($dataCollectionRules).Count) DCR(s)." -ForegroundColor DarkCyan

    Write-Host '[azmon-assess] Collecting resources + diagnostic settings...' -ForegroundColor DarkCyan
    $resources = Get-AzMonResource -SubscriptionId $subs
    $diagSettings = Get-AzMonDiagnosticSetting -SubscriptionId $subs
    Write-Host "[azmon-assess] Found $(@($resources).Count) monitorable resource(s), $(@($diagSettings).Count) diagnostic setting(s)." -ForegroundColor DarkCyan

    Write-Host '[azmon-assess] Collecting VM heartbeat coverage...' -ForegroundColor DarkCyan
    $heartbeatResourceIds = Get-AzMonHeartbeatResourceId -Workspace $workspaces -ThrottleLimit $ThrottleLimit
    Write-Host "[azmon-assess] Found heartbeat for $($heartbeatResourceIds.Count) resource(s)." -ForegroundColor DarkCyan

    $snapshot = New-AzMonSnapshotObject -SubscriptionId $subs -CustomerName $CustomerName `
        -Workspaces $workspaces -AppInsights $appInsights -AlertRules $alertRules `
        -ActionGroups $actionGroups -DiagnosticSettings $diagSettings -Resources $resources `
        -DataCollectionRules $dataCollectionRules

    $runAll = $Only -contains 'all'
    $findings = [System.Collections.Generic.List[hashtable]]::new()

    if ($runAll -or $Only -contains 'consolidation') {
        Write-Host '[azmon-assess] Analyzing consolidation opportunities...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonConsolidationFinding -Workspace $workspaces))
    }
    if ($runAll -or $Only -contains 'coverage') {
        Write-Host '[azmon-assess] Analyzing coverage gaps...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonCoverageGapFinding -ResourceRef $resources -DiagnosticSetting $diagSettings -Workspace $workspaces -AppInsight $appInsights -HeartbeatResourceId $heartbeatResourceIds))
    }
    if ($runAll -or $Only -contains 'alerting') {
        Write-Host '[azmon-assess] Analyzing alert quality...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonAlertQualityFinding -AlertRule $alertRules -ActionGroup $actionGroups -ResourceRef $resources))
    }
    if ($runAll -or $Only -contains 'cost') {
        Write-Host '[azmon-assess] Analyzing cost optimization...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonCostOptimizationFinding -Workspace $workspaces -AppInsight $appInsights))
    }
    if ($runAll -or $Only -contains 'tracing') {
        Write-Host '[azmon-assess] Analyzing tracing readiness...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonTracingFinding -ResourceRef $resources -AppInsight $appInsights))
    }
    if ($runAll -or $Only -contains 'reliability') {
        Write-Host '[azmon-assess] Analyzing reliability posture (WAF)...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonReliabilityFinding -Workspace $workspaces -AppInsight $appInsights -AlertRule $alertRules -DiagnosticSetting $diagSettings -ResourceRef $resources))
    }
    if ($runAll -or $Only -contains 'security') {
        Write-Host '[azmon-assess] Analyzing security posture (WAF)...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonSecurityFinding -Workspace $workspaces -AppInsight $appInsights))
    }
    if ($runAll -or $Only -contains 'performance') {
        Write-Host '[azmon-assess] Analyzing performance efficiency (WAF)...' -ForegroundColor DarkCyan
        $findings.AddRange((Find-AzMonPerformanceFinding -Workspace $workspaces -AppInsight $appInsights))
    }
    $snapshot['Findings'] = $findings.ToArray()

    if (-not $SkipAiSummary) {
        Write-Host '[azmon-assess] Generating executive summary...' -ForegroundColor DarkCyan
        $snapshot['AiSummary'] = New-AzMonAiSummary -Snapshot $snapshot -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion
    }

    $outDir = New-AzMonOutputDirectory -Path $OutputPath
    $snapshotPath = Save-AzMonSnapshot -Snapshot $snapshot -Path (Join-Path $outDir 'snapshot.json')
    Write-Host "[azmon-assess] Snapshot written to $snapshotPath" -ForegroundColor Green

    if ($Report -notcontains 'none') {
        Invoke-AzMonReportGeneration -Snapshot $snapshot -OutputPath $outDir -Report $Report
    }

    return $snapshot
}

function Invoke-AzMonReportGeneration {
    <#
    .SYNOPSIS
        Renders the requested report formats from a snapshot. Shared by the
        orchestrator and the standalone `report`/`demo` entry points.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $OutputPath,
        [ValidateSet('all', 'markdown', 'excel', 'pptx', 'html', 'none')]
        [string[]] $Report = @('all')
    )
    $outDir = New-AzMonOutputDirectory -Path $OutputPath
    $reportAll = $Report -contains 'all'

    if ($reportAll -or $Report -contains 'markdown') {
        New-AzMonMarkdownReport -Snapshot $Snapshot -Path (Join-Path $outDir 'report.md')
        Write-Host '[azmon-assess] report.md written.' -ForegroundColor Green
    }
    if ($reportAll -or $Report -contains 'excel') {
        try {
            New-AzMonExcelReport -Snapshot $Snapshot -Path (Join-Path $outDir 'report.xlsx')
            Write-Host '[azmon-assess] report.xlsx written.' -ForegroundColor Green
        } catch {
            Write-Warning "Excel report skipped: $($_.Exception.Message)"
        }
    }
    if ($reportAll -or $Report -contains 'pptx') {
        try {
            New-AzMonPptxReport -Snapshot $Snapshot -Path (Join-Path $outDir 'report.pptx')
            Write-Host '[azmon-assess] report.pptx written.' -ForegroundColor Green
        } catch {
            Write-Warning "PPTX report skipped: $($_.Exception.Message)"
        }
    }
    if ($reportAll -or $Report -contains 'html') {
        try {
            New-AzMonHtmlReport -Snapshot $Snapshot -Path (Join-Path $outDir 'report.html')
            Write-Host '[azmon-assess] report.html written.' -ForegroundColor Green
        } catch {
            Write-Warning "HTML report skipped: $($_.Exception.Message)"
        }
    }
}
