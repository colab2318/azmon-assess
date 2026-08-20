#requires -Version 7.0

function Connect-AzMonSession {
    <#
    .SYNOPSIS
        Ensures an authenticated Az context, signing in interactively if
        needed. Works unchanged in Azure Cloud Shell (already signed in).
    #>
    [CmdletBinding()]
    param([string] $TenantId)

    $ctx = Get-AzContext
    if (-not $ctx) {
        Write-Host '[azmon-assess] No Azure session found — signing in...' -ForegroundColor Cyan
        if ($TenantId) { Connect-AzAccount -Tenant $TenantId -ErrorAction Stop | Out-Null }
        else { Connect-AzAccount -ErrorAction Stop | Out-Null }
        $ctx = Get-AzContext
    }
    if (-not $ctx) { throw 'Unable to establish an Azure context. Run Connect-AzAccount manually and retry.' }
    Write-Host "[azmon-assess] Signed in as $($ctx.Account.Id) (tenant $($ctx.Tenant.Id))." -ForegroundColor Green
    return $ctx
}
