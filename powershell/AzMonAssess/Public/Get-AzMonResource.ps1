#requires -Version 7.0

function Get-AzMonResource {
    <#
    .SYNOPSIS
        Collects the inventory of "monitorable" resource types used for
        coverage-gap and alert-quality analysis.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQAllMonitorable -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            New-AzMonResourceRef -Id $r.id -Name $r.name -Type $r.type -SubscriptionId $r.subscriptionId `
                -ResourceGroup $r.resourceGroup -Location $r.location -Tags (ConvertTo-AzMonHashtable $r.tags)
        }
    )
}

function Get-AzMonDiagnosticSetting {
    <#
    .SYNOPSIS
        Collects diagnostic settings via per-resource ARM REST calls.
    .NOTES
        microsoft.insights/diagnosticSettings is NOT a resource type the
        Azure Resource Graph 'resources' table supports (confirmed against
        the ARG supported-tables reference - only actiongroups,
        metricalerts, scheduledqueryrules, datacollectionrules,
        guestdiagnosticsettings, etc. are listed for microsoft.insights/*,
        diagnosticSettings itself is absent). A bulk ARG query for it
        returns unreliable/empty results regardless of what actually
        exists, so this queries each resource directly instead. Runspaces
        started by ForEach-Object -Parallel can't see module-scoped
        functions (only $using: variables), so the REST call and the
        DiagnosticSetting shape are inlined here rather than calling
        Invoke-AzRestMethod/New-AzMonDiagnosticSetting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ResourceRef,
        [int] $ThrottleLimit = 8
    )
    if (-not $ResourceRef -or $ResourceRef.Count -eq 0) { return @() }

    $context = Get-AzContext
    $results = $ResourceRef | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $r = $_
        $ctx = $using:context
        try {
            Import-Module Az.Accounts -ErrorAction Stop
            if ($ctx) { $null = Set-AzContext -Context $ctx -ErrorAction Stop }

            $resp = Invoke-AzRestMethod -Path "$($r['Id'])/providers/microsoft.insights/diagnosticSettings?api-version=2021-05-01-preview" -Method GET -ErrorAction Stop
            if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) { return @() }
            if ([string]::IsNullOrWhiteSpace($resp.Content)) { return @() }
            $parsed = $resp.Content | ConvertFrom-Json
            $settings = @($parsed.value)
            if ($settings.Count -eq 0) { return @() }

            return @(foreach ($s in $settings) {
                    $logsEnabled = [bool](@($s.properties.logs) | Where-Object { $_.enabled })
                    $metricsEnabled = [bool](@($s.properties.metrics) | Where-Object { $_.enabled })
                    @{
                        Kind           = 'DiagnosticSetting'
                        ResourceId     = ([string]$r['Id']).ToLowerInvariant()
                        Name           = $s.name
                        WorkspaceId    = $s.properties.workspaceId
                        StorageId      = $s.properties.storageAccountId
                        EventHubId     = $s.properties.eventHubAuthorizationRuleId
                        LogsEnabled    = $logsEnabled
                        MetricsEnabled = $metricsEnabled
                    }
                })
        } catch {
            return @()
        }
    }
    return @($results | Where-Object { $null -ne $_ })
}
