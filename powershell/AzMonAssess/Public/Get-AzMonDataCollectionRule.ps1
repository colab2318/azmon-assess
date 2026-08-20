#requires -Version 7.0

function Get-AzMonDataCollectionRule {
    <#
    .SYNOPSIS
        Collects Data Collection Rules (DCRs) via Azure Resource Graph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQDataCollectionRules -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            New-AzMonDataCollectionRule -Id $r.id -Name $r.name -SubscriptionId $r.subscriptionId -ResourceGroup $r.resourceGroup `
                -Location $r.location -DcrKind $r.kind -Tags (ConvertTo-AzMonHashtable $r.tags) `
                -DataFlowCount ([int]($r.dataFlowCount ?? 0)) `
                -WorkspaceResourceId $r.workspaceResourceId
        }
    )
}
