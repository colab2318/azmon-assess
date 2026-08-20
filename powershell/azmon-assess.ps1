<#
.SYNOPSIS
    azmon-assess (PowerShell edition) — AI-powered Azure Monitoring &
    Observability assessment. Runs entirely in Azure Cloud Shell or local
    PowerShell 7+ with no admin rights required.
.DESCRIPTION
    Thin command-style wrapper around the AzMonAssess module, mirroring the
    original Python CLI's commands: run, consolidate, gaps, alerts, cost,
    tracing, reliability, security, performance, demo, report, summarize,
    triage, remediate.
.PARAMETER Command
    Which action to run.
.EXAMPLE
    ./azmon-assess.ps1 demo -Output ./out
    Generate a full sample report from fixture data — no Azure access required.
.EXAMPLE
    ./azmon-assess.ps1 run -Output ./out
    Full assessment across every subscription the signed-in identity can see.
.EXAMPLE
    ./azmon-assess.ps1 run -Output ./out -SubscriptionId 00000000-0000-0000-0000-000000000000
    Full assessment scoped to one subscription.
.EXAMPLE
    ./azmon-assess.ps1 report -Snapshot ./out/snapshot.json -Output ./out -Format excel,pptx
    Regenerate just the Excel + PowerPoint deliverables from an existing snapshot.
.EXAMPLE
    ./azmon-assess.ps1 triage -Snapshot ./out/snapshot.json -TriageOutput ./out/triage.json -EmitTemplate
    Emit a YAML/JSON template pre-populated with every finding as "snooze".
.EXAMPLE
    ./azmon-assess.ps1 remediate -Snapshot ./out/snapshot.json -TriagePath ./out/triage.json -Output ./out
    Dry-run every finding triaged as "accept" (add -Apply to actually change Azure).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateSet('run', 'consolidate', 'gaps', 'alerts', 'cost', 'tracing', 'reliability', 'security', 'performance',
        'demo', 'report', 'summarize', 'triage', 'remediate')]
    [string] $Command,

    [Alias('o')] [string] $Output = './out',
    [string[]] $SubscriptionId,
    [string] $ManagementGroupId,
    [string] $TenantId,
    [string] $CustomerName,
    [string] $AoaiEndpoint,
    [string] $AoaiDeployment,
    [string] $AoaiApiVersion,
    [int] $LookbackDays,
    [int] $ThrottleLimit,
    [switch] $SkipIngestionEnrichment,
    [switch] $SkipAiSummary,

    [Alias('s')] [string] $Snapshot,
    [ValidateSet('all', 'markdown', 'excel', 'pptx', 'html', 'none')]
    [string[]] $Format = @('all'),

    [string] $TriageOutput = './out/triage.json',
    [string] $Plan,
    [switch] $EmitTemplate,
    [int] $Limit,

    [string] $TriagePath,
    [switch] $Apply,
    [string] $ActionGroup,
    [string[]] $Only,
    [double] $DailyQuotaGb = 50.0,
    [int] $RetentionDays = 30,
    [double] $SamplingPercentage = 10.0
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'AzMonAssess/AzMonAssess.psd1') -Force

$azureCommands = @('run', 'consolidate', 'gaps', 'alerts', 'cost', 'tracing', 'reliability', 'security', 'performance', 'remediate')
if ($Command -in $azureCommands) {
    Initialize-AzMonPrerequisite
    Connect-AzMonSession -TenantId $TenantId | Out-Null
}

function Invoke-LocalAssessment {
    param([string[]] $OnlyCategories)
    Invoke-AzMonAssessment -SubscriptionId $SubscriptionId -ManagementGroupId $ManagementGroupId -CustomerName $CustomerName `
        -LookbackDays $LookbackDays -ThrottleLimit $ThrottleLimit -OutputPath $Output -Only $OnlyCategories -Report $Format `
        -SkipIngestionEnrichment:$SkipIngestionEnrichment -SkipAiSummary:$SkipAiSummary `
        -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion | Out-Null
}

switch ($Command) {
    'run' { Invoke-LocalAssessment -OnlyCategories @('all') }
    'consolidate' { Invoke-LocalAssessment -OnlyCategories @('consolidation') }
    'gaps' { Invoke-LocalAssessment -OnlyCategories @('coverage') }
    'alerts' { Invoke-LocalAssessment -OnlyCategories @('alerting') }
    'cost' { Invoke-LocalAssessment -OnlyCategories @('cost') }
    'tracing' { Invoke-LocalAssessment -OnlyCategories @('tracing') }
    'reliability' { Invoke-LocalAssessment -OnlyCategories @('reliability') }
    'security' { Invoke-LocalAssessment -OnlyCategories @('security') }
    'performance' { Invoke-LocalAssessment -OnlyCategories @('performance') }

    'demo' {
        $name = if ($CustomerName) { $CustomerName } else { 'Your Organization' }
        $snap = New-AzMonDemoSnapshot -CustomerName $name -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion
        Save-AzMonSnapshot -Snapshot $snap -Path (Join-Path $Output 'snapshot.json') | Out-Null
        Invoke-AzMonReportGeneration -Snapshot $snap -OutputPath $Output -Report $Format
    }

    'report' {
        if (-not $Snapshot) { throw '-Snapshot is required for the report command.' }
        $snap = Import-AzMonSnapshot -Path $Snapshot
        Invoke-AzMonReportGeneration -Snapshot $snap -OutputPath $Output -Report $Format
    }

    'summarize' {
        if (-not $Snapshot) { throw '-Snapshot is required for the summarize command.' }
        Initialize-AzMonPrerequisite
        Connect-AzMonSession -TenantId $TenantId | Out-Null
        $snap = Import-AzMonSnapshot -Path $Snapshot
        $snap['AiSummary'] = New-AzMonAiSummary -Snapshot $snap -AoaiEndpoint $AoaiEndpoint -AoaiDeployment $AoaiDeployment -AoaiApiVersion $AoaiApiVersion
        Save-AzMonSnapshot -Snapshot $snap -Path (Join-Path $Output 'snapshot.json') | Out-Null
        Write-Host "[azmon-assess] snapshot.json updated with a new AI summary in $Output" -ForegroundColor Green
    }

    'triage' {
        if (-not $Snapshot) { throw '-Snapshot is required for the triage command.' }
        Invoke-AzMonTriage -SnapshotPath $Snapshot -OutputPath $TriageOutput -PlanPath $Plan -EmitTemplate:$EmitTemplate -Limit $Limit
    }

    'remediate' {
        if (-not $Snapshot) { throw '-Snapshot is required for the remediate command.' }
        if (-not $TriagePath) { throw '-TriagePath is required for the remediate command.' }
        Invoke-AzMonRemediation -SnapshotPath $Snapshot -TriagePath $TriagePath -OutputPath $Output -Apply:$Apply `
            -ActionGroupId $ActionGroup -OnlyCategory $Only -DailyQuotaGb $DailyQuotaGb -RetentionDays $RetentionDays `
            -SamplingPercentage $SamplingPercentage | Out-Null
    }
}
