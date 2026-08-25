
$content = Get-Content -Raw -Path "C:\ghc-projects\azmon-assess\out-finalcheck\impacted-recommendations.json"
$json = ConvertFrom-Json $content
$allDetails = foreach ($item in $json) {
    [PSCustomObject]@{
        CheckName = $item.CheckName
        Title = $item.Title
        Category = $item.Category
        Impact = $item.Impact
        ImpactedResourceCount = $item.ImpactedResources.Count
        DistinctResourceTypes = (($item.ImpactedResources.ResourceType | Select-Object -Unique) -join ", ")
    }
}
$allDetails | Format-Table -Wrap
$totalResources = ($allDetails | Measure-Object -Property ImpactedResourceCount -Sum).Sum
Write-Host "Total Sum of ImpactedResourceCount: $totalResources"
