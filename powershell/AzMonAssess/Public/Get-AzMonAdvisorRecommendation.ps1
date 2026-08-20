#requires -Version 7.0
# Azure Advisor cost recommendations relevant to Azure Monitor, fetched
# directly via ARM REST (Invoke-AzRestMethod / Az.Accounts only) — closes
# the "set up alerts on Azure Advisor cost recommendations" best practice
# by surfacing the recommendations themselves rather than just checking
# whether an alert exists.

function Get-AzMonAdvisorRecommendation {
    <#
    .SYNOPSIS
        Collects Azure Advisor Cost-category recommendations scoped to
        Log Analytics / Application Insights / alert-rule resources.
    .NOTES
        GET /subscriptions/{id}/providers/Microsoft.Advisor/recommendations
        (api-version 2025-01-01), filtered server-side to Category eq
        'Cost' and client-side to Monitor-related impactedField values.
        Paginated via nextLink.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $all = [System.Collections.Generic.List[hashtable]]::new()
    $filterValue = [System.Uri]::EscapeDataString("Category eq 'Cost'")

    foreach ($subId in $SubscriptionId) {
        $path = "/subscriptions/$subId/providers/Microsoft.Advisor/recommendations?api-version=2025-01-01&`$filter=$filterValue"
        $nextLink = $null
        $guard = 0
        do {
            try {
                if ($nextLink) { $resp = Invoke-AzRestMethod -Uri $nextLink -Method GET -ErrorAction Stop }
                else { $resp = Invoke-AzRestMethod -Path $path -Method GET -ErrorAction Stop }
            } catch {
                Write-Warning "Advisor recommendations query failed for subscription ${subId}: $($_.Exception.Message)"
                break
            }
            if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
                Write-Warning "Advisor recommendations query failed for subscription ${subId} with status $($resp.StatusCode)."
                break
            }
            $parsed = $resp.Content | ConvertFrom-Json
            foreach ($item in @($parsed.value)) {
                $impactedField = ([string]$item.properties.impactedField).ToLowerInvariant()
                if ($impactedField -notmatch 'operationalinsights|insights/components|insights/scheduledqueryrules|insights/metricalerts') { continue }
                $all.Add(@{
                        Kind             = 'AdvisorRecommendation'
                        Id               = $item.id
                        ResourceId       = $item.properties.resourceMetadata.resourceId
                        ImpactedField    = $item.properties.impactedField
                        Impact           = $item.properties.impact
                        Problem          = $item.properties.shortDescription.problem
                        Solution         = $item.properties.shortDescription.solution
                        PotentialBenefit = $item.properties.potentialBenefits
                        LearnMoreLink    = $item.properties.learnMoreLink
                    })
            }
            $nextLink = $parsed.nextLink
            $guard++
        } while ($nextLink -and $guard -lt 100)
    }
    return $all.ToArray()
}
