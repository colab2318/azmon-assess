#requires -Version 7.0

function Get-AzMonAppInsight {
    <#
    .SYNOPSIS
        Collects Application Insights components via Azure Resource Graph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQAppInsights -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            New-AzMonAppInsight -Id $r.id -Name $r.name -SubscriptionId $r.subscriptionId -ResourceGroup $r.resourceGroup `
                -Location $r.location -Kind $r.kind -ApplicationType $r.applicationType `
                -WorkspaceResourceId $r.workspaceResourceId -SamplingPercentage $r.samplingPercentage `
                -RetentionDays $r.retentionInDays -Tags (ConvertTo-AzMonHashtable $r.tags) `
                -PublicNetworkAccessForIngestion $r.publicNetworkAccessForIngestion `
                -PublicNetworkAccessForQuery $r.publicNetworkAccessForQuery `
                -DisableLocalAuth $r.disableLocalAuth -DailyCapGb $r.dailyCapGb
        }
    )
}
