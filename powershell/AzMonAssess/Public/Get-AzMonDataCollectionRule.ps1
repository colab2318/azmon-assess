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

function Get-AzMonDcrAssociatedId {
    <#
    .SYNOPSIS
        Returns a HashSet of (lowercased) DCR resource IDs that have at
        least one association — used to detect orphaned DCRs. Associations
        are child resources of the TARGET (VM, etc.), not the DCR itself.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQDcrAssociations -SubscriptionId $SubscriptionId
    $ids = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in $rows) { if ($r.dcrId) { [void]$ids.Add([string]$r.dcrId) } }
    return $ids
}
