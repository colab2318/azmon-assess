#requires -Version 7.0
# Distributed-tracing readiness analyzer — ported 1:1 from analyzers/tracing.py.

$script:AzMonTraceCandidateTypes = @(
    'microsoft.web/sites', 'microsoft.app/containerapps', 'microsoft.containerservice/managedclusters',
    'microsoft.eventhub/namespaces', 'microsoft.servicebus/namespaces', 'microsoft.apimanagement/service'
)

function Find-AzMonTracingFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ResourceRef,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $AppInsight
    )
    $findings = [System.Collections.Generic.List[hashtable]]::new()
    if (-not $ResourceRef -or $ResourceRef.Count -eq 0) { return @() }

    $candidates = @($ResourceRef | Where-Object { $script:AzMonTraceCandidateTypes -contains ([string]$_['Type']).ToLowerInvariant() })
    $aiWorkspaceBased = @($AppInsight | Where-Object { $_['WorkspaceResourceId'] })

    if ($candidates.Count -gt 0 -and $aiWorkspaceBased.Count -lt [Math]::Max(1, [Math]::Floor($candidates.Count / 5))) {
        $byType = Get-AzMonCountByType -ResourceRef $candidates
        $findings.Add((New-AzMonFinding -Category 'tracing' -Severity 'high' `
            -Title 'Low distributed-tracing coverage across web / integration workloads' `
            -Detail ('Only a fraction of candidate workloads are instrumented for distributed tracing. Adopting ' +
                "Azure Monitor's OpenTelemetry Distro provides vendor-neutral traces from front-end through APIM " +
                '-> Container Apps -> Service Bus -> DB, and lets Application Insights correlate requests ' +
                'end-to-end.') `
            -ResourceIds @($candidates | ForEach-Object { $_['Id'] }) `
            -Recommendation (
                "Roll out the Azure Monitor OpenTelemetry Distro:`n" +
                "1. For .NET / Java / Node / Python apps: install Azure.Monitor.OpenTelemetry.AspNetCore " +
                "(or language equivalent) and set APPLICATIONINSIGHTS_CONNECTION_STRING.`n" +
                "2. Enable auto-instrumentation on App Service and Container Apps.`n" +
                "3. Deploy the OpenTelemetry Collector as a sidecar in AKS for services that cannot embed the SDK.`n" +
                "4. Configure trace-based sampling at 5-10% to control cost."
            ) `
            -Evidence @{ candidate_count = $candidates.Count; workspace_based_ai_count = $aiWorkspaceBased.Count; candidate_types = $byType }))
    }

    return $findings.ToArray()
}
