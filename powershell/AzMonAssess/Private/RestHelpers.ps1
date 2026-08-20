#requires -Version 7.0
# Thin ARM REST helpers used by remediation actions — Invoke-AzRestMethod is
# always available (part of Az.Accounts) and lets us PATCH properties that
# aren't exposed by every typed cmdlet, exactly mirroring what the Python
# tool's SDK calls did (get -> mutate -> create_or_update).

function Get-AzMonArmResource {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceId, [Parameter(Mandatory)] [string] $ApiVersion)
    $resp = Invoke-AzRestMethod -Path "$ResourceId`?api-version=$ApiVersion" -Method GET -ErrorAction Stop
    if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
        throw "GET $ResourceId failed with status $($resp.StatusCode): $($resp.Content)"
    }
    return ($resp.Content | ConvertFrom-Json -AsHashtable)
}

function Set-AzMonArmResource {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceId, [Parameter(Mandatory)] [string] $ApiVersion, [Parameter(Mandatory)] [hashtable] $Body)
    $json = $Body | ConvertTo-Json -Depth 20
    $resp = Invoke-AzRestMethod -Path "$ResourceId`?api-version=$ApiVersion" -Method PATCH -Payload $json -ErrorAction Stop
    if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
        throw "PATCH $ResourceId failed with status $($resp.StatusCode): $($resp.Content)"
    }
    if ([string]::IsNullOrWhiteSpace($resp.Content)) { return @{} }
    return ($resp.Content | ConvertFrom-Json -AsHashtable)
}

function ConvertFrom-AzMonResourceId {
    <#
    .SYNOPSIS
        Splits an ARM resource id into Subscription / ResourceGroup / Name.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceId)
    $parts = $ResourceId -split '/'
    $subIdx = [array]::IndexOf($parts, 'subscriptions')
    $rgIdx = [array]::IndexOf($parts, 'resourceGroups')
    return [pscustomobject]@{
        SubscriptionId = if ($subIdx -ge 0) { $parts[$subIdx + 1] } else { $null }
        ResourceGroup  = if ($rgIdx -ge 0) { $parts[$rgIdx + 1] } else { $null }
        Name           = $parts[-1]
    }
}
