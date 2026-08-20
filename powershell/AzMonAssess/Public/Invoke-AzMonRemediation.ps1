#requires -Version 7.0
# Remediation dispatcher — maps findings to actions and executes (or
# dry-runs) them. Ported from remediation/dispatcher.py; matching is
# intentionally conservative (category + title keyword), same as the
# Python version.

function Get-AzMonRemediationActionKey {
    <#
    .SYNOPSIS
        Returns an action key string for a finding, or $null if there is no
        safe auto-remediation (in which case a manual runbook may still be
        emitted for categories that always produce one). 'Runbook:<Name>'
        keys map to a manual-runbook template instead of an Azure change.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding)

    $cat = $Finding['Category']
    $title = ([string]$Finding['Title']).ToLowerInvariant()
    $isAppInsightsTarget = [bool](@($Finding['ResourceIds']) | Where-Object { ([string]$_).ToLowerInvariant().Contains('/microsoft.insights/components/') })

    switch ($cat) {
        'cost' {
            if ($title -match 'daily cap|no daily cap|daily quota|\bcap\b|\bquota\b') {
                # The AI daily-cap finding shares the same "cap" wording but there is
                # no dedicated App Insights cap action (only the workspace one) -
                # skip rather than PATCH the wrong resource type.
                if ($isAppInsightsTarget) { return $null }
                return 'SetWorkspaceDailyQuota'
            }
            if ($title -match 'retention') { return 'SetWorkspaceRetention' }
            if ($title -match 'verbose table|basic logs|table plan') { return 'SetTablePlanBasic' }
            return $null
        }
        'retention' { return 'SetWorkspaceRetention' }
        'alerting' {
            if ($title -match 'no action group|silent|missing action') { return 'AttachActionGroup' }
            return $null
        }
        'tracing' {
            if ($title -match 'sampling|app insights') { return 'EnableAppInsightsSampling' }
            return $null
        }
        'coverage' { return 'Runbook:CoverageDiagBaseline' }
        'consolidation' { return 'Runbook:WorkspaceConsolidation' }
        'security' {
            $isAppInsights = [bool](@($Finding['ResourceIds']) | Where-Object { ([string]$_).ToLowerInvariant().Contains('/microsoft.insights/components/') })
            if ($title -match 'public network access|public network') {
                if ($isAppInsights) { return $null }
                return 'DisableWorkspacePublicNetwork'
            }
            if ($title -match 'local auth|shared-key|instrumentation-key auth') {
                if ($isAppInsights) { return 'DisableAppInsightsLocalAuth' }
                return 'DisableWorkspaceLocalAuth'
            }
            if ($title -match 'classic' -and $title -match 'insights') { return 'Runbook:ClassicAiMigration' }
            return $null
        }
        'reliability' {
            if ($title -match 'does not support workspace availability zones|non-az region') { return 'Runbook:AzRegionMigration' }
            if ($title -match 'no health|health / ingestion alert') { return 'Runbook:WorkspaceHealthAlert' }
            if ($title -match 'single log destination') { return 'Runbook:DualDestinationLogs' }
            return $null
        }
        'performance' {
            if ($title -match 'dedicated cluster') { return 'Runbook:DedicatedCluster' }
            if ($title -match 'search jobs|basic tier') { return 'Runbook:SearchJobsTier' }
            return $null
        }
        default {
            if ($title -match 'classic' -and $title -match 'insights') { return 'Runbook:ClassicAiMigration' }
            return $null
        }
    }
}

function Invoke-AzMonRemediationAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Finding,
        [Parameter(Mandatory)] [hashtable] $Context,
        [Parameter(Mandatory)] [string] $RunbookDir,
        [Parameter(Mandatory)] [bool] $Apply
    )
    $actionKey = Get-AzMonRemediationActionKey -Finding $Finding
    if (-not $actionKey) {
        return New-AzMonActionResult -Action 'unsupported' -Ok $false -DryRun (-not $Apply) -Message "No auto-remediation for '$($Finding['Title'])'" -FindingId $Finding['Id']
    }
    if ($actionKey.StartsWith('Runbook:')) {
        return Write-AzMonManualRunbookAction -Finding $Finding -TemplateKey $actionKey.Substring(8) -RunbookDir $RunbookDir -Apply $Apply
    }
    switch ($actionKey) {
        'SetWorkspaceDailyQuota' { return Invoke-AzMonSetWorkspaceDailyQuotaAction -Finding $Finding -Context $Context -Apply $Apply }
        'SetWorkspaceRetention' { return Invoke-AzMonSetWorkspaceRetentionAction -Finding $Finding -Context $Context -Apply $Apply }
        'SetTablePlanBasic' { return Invoke-AzMonSetTablePlanBasicAction -Finding $Finding -Context $Context -Apply $Apply }
        'AttachActionGroup' { return Invoke-AzMonAttachActionGroupAction -Finding $Finding -Context $Context -Apply $Apply }
        'EnableAppInsightsSampling' { return Invoke-AzMonEnableAppInsightsSamplingAction -Finding $Finding -Context $Context -Apply $Apply }
        'DisableWorkspacePublicNetwork' { return Invoke-AzMonDisableWorkspacePublicNetworkAction -Finding $Finding -Context $Context -Apply $Apply }
        'DisableWorkspaceLocalAuth' { return Invoke-AzMonDisableWorkspaceLocalAuthAction -Finding $Finding -Context $Context -Apply $Apply }
        'DisableAppInsightsLocalAuth' { return Invoke-AzMonDisableAppInsightsLocalAuthAction -Finding $Finding -Context $Context -Apply $Apply }
        default { return New-AzMonActionResult -Action 'unsupported' -Ok $false -DryRun (-not $Apply) -Message "Unmapped action key: $actionKey" -FindingId $Finding['Id'] }
    }
}

function Invoke-AzMonRemediation {
    <#
    .SYNOPSIS
        Applies (or dry-runs) remediations for findings triaged as 'accept'.
        Defaults to dry-run — pass -Apply to make real changes. Every run
        writes a full audit trail to remediation-log.json.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SnapshotPath,
        [Parameter(Mandatory)] [string] $TriagePath,
        [Parameter(Mandatory)] [string] $OutputPath,
        [switch] $Apply,
        [string] $ActionGroupId,
        [string[]] $OnlyCategory,
        [double] $DailyQuotaGb = 50.0,
        [int] $RetentionDays = 30,
        [double] $SamplingPercentage = 10.0
    )

    $snapshot = Import-AzMonSnapshot -Path $SnapshotPath
    $plan = Import-AzMonTriagePlan -Path $TriagePath
    $findingsById = @{}
    foreach ($f in @($snapshot['Findings'])) { $findingsById[$f['Id']] = $f }

    $context = @{
        DefaultActionGroupId      = if ($ActionGroupId) { $ActionGroupId } else { $env:AZMON_DEFAULT_ACTION_GROUP_ID }
        DefaultDailyQuotaGb       = $DailyQuotaGb
        DefaultRetentionDays      = $RetentionDays
        DefaultSamplingPercentage = $SamplingPercentage
    }

    $runbookDir = Join-Path $OutputPath 'runbooks'
    $accepted = @(@($plan['Entries']) | Where-Object { $_['Decision'] -eq 'accept' })
    Write-Host "Remediation mode=$(if ($Apply) { 'APPLY' } else { 'dry-run' }) accepted=$($accepted.Count)" -ForegroundColor Cyan

    $results = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($entry in $accepted) {
        $finding = $findingsById[$entry['FindingId']]
        if (-not $finding) { continue }
        if ($OnlyCategory -and ($OnlyCategory -notcontains $finding['Category'])) { continue }

        $result = Invoke-AzMonRemediationAction -Finding $finding -Context $context -RunbookDir $runbookDir -Apply $Apply.IsPresent
        $results.Add($result)
        $color = if ($result['Ok']) { 'Green' } else { 'Red' }
        Write-Host "[$($result['Action'])] $($result['Message'])" -ForegroundColor $color
    }

    $outDir = New-AzMonOutputDirectory -Path $OutputPath
    $logPath = Join-Path $outDir 'remediation-log.json'
    $logDoc = @{
        mode    = if ($Apply) { 'apply' } else { 'dry-run' }
        at      = (Get-Date).ToUniversalTime().ToString('o')
        results = $results.ToArray()
    }
    ($logDoc | ConvertTo-Json -Depth 15) | Set-Content -LiteralPath $logPath -Encoding utf8NoBOM
    Write-Host "[azmon-assess] Remediation log written to $logPath" -ForegroundColor Green

    $okCount = @($results | Where-Object { $_['Ok'] }).Count
    Write-Host "[azmon-assess] Summary: $okCount/$($results.Count) action(s) succeeded (or would, in dry-run)." -ForegroundColor Cyan
    return $results.ToArray()
}
