#requires -Version 7.0

function Resolve-AzMonSubscription {
    <#
    .SYNOPSIS
        Resolves the list of subscription IDs to assess: explicit parameter
        > AZURE_SUBSCRIPTION_IDS env var > management group expansion >
        auto-discovery of every subscription the caller can see.
    #>
    [CmdletBinding()]
    param(
        [string[]] $SubscriptionId,
        [string] $ManagementGroupId
    )

    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        return @($SubscriptionId)
    }

    if ($env:AZURE_SUBSCRIPTION_IDS) {
        $fromEnv = @($env:AZURE_SUBSCRIPTION_IDS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($fromEnv.Count -gt 0) { return $fromEnv }
    }

    $mgId = if ($ManagementGroupId) { $ManagementGroupId } else { $env:AZURE_MANAGEMENT_GROUP_ID }
    if ($mgId) {
        Write-Host "[azmon-assess] Expanding management group '$mgId'..." -ForegroundColor Cyan
        try {
            $ids = @(Get-AzMonManagementGroupSubscription -GroupId $mgId)
            if ($ids.Count -gt 0) { return $ids }
        } catch {
            Write-Warning "Could not expand management group '$mgId' ($($_.Exception.Message)) — falling back to auto-discovery."
        }
    }

    Write-Host '[azmon-assess] No subscriptions specified — auto-discovering...' -ForegroundColor Cyan
    # Get-AzSubscription with no -TenantId tries every tenant the signed-in
    # identity has any access to (common for guest/B2B accounts), not just
    # the tenant that's actually signed in — scope explicitly to avoid
    # cross-tenant token failures and accidentally assessing other tenants.
    $currentTenantId = (Get-AzContext).Tenant.Id
    $subs = @(Get-AzSubscription -TenantId $currentTenantId -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id)
    Write-Host "[azmon-assess] Discovered $($subs.Count) subscription(s) in tenant $currentTenantId." -ForegroundColor Cyan
    return $subs
}

function Get-AzMonManagementGroupSubscription {
    <#
    .SYNOPSIS
        Lists every subscription ID under a management group (recursively),
        via the Management Groups "Get Descendants" ARM REST API
        (Invoke-AzRestMethod / Az.Accounts only — no Az.Resources module).
    .NOTES
        GET /providers/Microsoft.Management/managementGroups/{id}/descendants
        returns a flat list of descendant management groups AND
        subscriptions; subscriptions have type
        'Microsoft.Management/managementGroups/subscriptions' and a `name`
        equal to the subscription GUID. Paginated via `nextLink`.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $GroupId)

    $ids = [System.Collections.Generic.List[string]]::new()
    $path = "/providers/Microsoft.Management/managementGroups/$GroupId/descendants?api-version=2020-05-01"
    $nextLink = $null
    $guard = 0
    do {
        if ($nextLink) {
            $resp = Invoke-AzRestMethod -Uri $nextLink -Method GET -ErrorAction Stop
        } else {
            $resp = Invoke-AzRestMethod -Path $path -Method GET -ErrorAction Stop
        }
        if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
            throw "Management group descendants query failed with status $($resp.StatusCode): $($resp.Content)"
        }
        $parsed = $resp.Content | ConvertFrom-Json
        foreach ($item in @($parsed.value)) {
            if ($item.type -eq 'Microsoft.Management/managementGroups/subscriptions') { $ids.Add($item.name) }
        }
        $nextLink = $parsed.nextLink
        $guard++
    } while ($nextLink -and $guard -lt 1000)
    return $ids.ToArray()
}
