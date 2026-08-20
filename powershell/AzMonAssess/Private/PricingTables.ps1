#requires -Version 7.0
# Directional Azure Monitor Logs pricing — same figures as the Python
# consolidation analyzer. Validate against the customer's EA/MCA rate card
# before quoting savings to leadership.

$script:AzMonPaygPricePerGb = 2.76

# (minGbPerDay, pricePerGb)
$script:AzMonCommitmentTiers = @(
    [pscustomobject]@{ MinGb = 100;  Price = 2.208 }
    [pscustomobject]@{ MinGb = 200;  Price = 2.070 }
    [pscustomobject]@{ MinGb = 300;  Price = 2.001 }
    [pscustomobject]@{ MinGb = 400;  Price = 1.932 }
    [pscustomobject]@{ MinGb = 500;  Price = 1.863 }
    [pscustomobject]@{ MinGb = 1000; Price = 1.794 }
    [pscustomobject]@{ MinGb = 2000; Price = 1.725 }
    [pscustomobject]@{ MinGb = 5000; Price = 1.656 }
)

$script:AzMonEnvTagKeys = @('environment', 'env', 'Environment', 'Env', 'stage', 'tier')

function Get-AzMonEnvironmentTag {
    <#
    .SYNOPSIS
        Best-effort environment classification for a workspace (tag first,
        then a name-based guess), mirroring analyzers/consolidation.py::_env_of.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Workspace)

    $tags = $Workspace['Tags']
    if ($tags -is [hashtable]) {
        foreach ($k in $script:AzMonEnvTagKeys) {
            if ($tags.ContainsKey($k) -and $tags[$k]) { return ([string]$tags[$k]).ToLowerInvariant() }
        }
    }
    $name = ([string]$Workspace['Name']).ToLowerInvariant()
    foreach ($env in @('prod', 'production', 'qa', 'test', 'dev', 'staging', 'stage', 'core')) {
        if ($name -like "*$env*") { return $env }
    }
    return 'unknown'
}

function Get-AzMonBestCommitment {
    <#
    .SYNOPSIS
        Returns @{ PricePerGb; MonthlyCost } for the best-fit commitment
        tier vs. pay-as-you-go, given a daily GB volume.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [double] $DailyGb)

    $monthlyGb = $DailyGb * 30.0
    $paygCost = $monthlyGb * $script:AzMonPaygPricePerGb
    $bestCost = $paygCost
    $bestPrice = $script:AzMonPaygPricePerGb
    foreach ($tier in $script:AzMonCommitmentTiers) {
        if ($DailyGb -ge $tier.MinGb) {
            $commitCost = [Math]::Max($DailyGb, $tier.MinGb) * 30.0 * $tier.Price
            if ($commitCost -lt $bestCost) {
                $bestCost = $commitCost
                $bestPrice = $tier.Price
            }
        }
    }
    return [pscustomobject]@{ PricePerGb = $bestPrice; MonthlyCost = $bestCost }
}
