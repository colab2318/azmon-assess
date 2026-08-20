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
        Collects diagnostic settings (as extension resources) via Resource Graph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQDiagnosticSettings -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            $target = ($r.id -split '/providers/microsoft.insights/diagnosticSettings', 2)[0]
            New-AzMonDiagnosticSetting -ResourceId $target.ToLowerInvariant() -Name $r.name `
                -WorkspaceId $r.workspaceId -StorageId $r.storageAccountId -EventHubId $r.eventHubAuthorizationRuleId `
                -LogsEnabled (Test-AzMonAnyEnabled $r.logs) -MetricsEnabled (Test-AzMonAnyEnabled $r.metrics)
        }
    )
}

function Test-AzMonAnyEnabled {
    [CmdletBinding()]
    param($Items)
    $arr = ConvertTo-AzMonHashtable $Items
    if ($arr -isnot [array]) { return $false }
    foreach ($i in $arr) {
        if ($i -is [hashtable] -and $i['enabled']) { return $true }
    }
    return $false
}
