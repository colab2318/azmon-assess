#requires -Version 7.0

function Connect-AzMonSession {
    <#
    .SYNOPSIS
        Ensures an authenticated Az context in the requested tenant, signing
        in (or re-signing in) as needed. Works unchanged in Azure Cloud
        Shell (already signed in).
    .NOTES
        A pre-existing context in the WRONG tenant still requires a fresh
        Connect-AzAccount -Tenant call — an already-signed-in context is
        never assumed to match a caller-specified -TenantId. Note this is
        the PowerShell Az.Accounts session only; it's entirely separate
        from the Azure CLI's (`az login`) session/token cache.
    #>
    [CmdletBinding()]
    param([string] $TenantId)

    $ctx = Get-AzContext
    $wrongTenant = $ctx -and $TenantId -and $ctx.Tenant.Id -ne $TenantId
    if (-not $ctx -or $wrongTenant) {
        if ($wrongTenant) {
            Write-Host "[azmon-assess] Current session is tenant $($ctx.Tenant.Id), requested $TenantId — signing in again..." -ForegroundColor Cyan
        } else {
            Write-Host '[azmon-assess] No Azure session found — signing in...' -ForegroundColor Cyan
        }
        if ($TenantId) { Connect-AzAccount -Tenant $TenantId -ErrorAction Stop | Out-Null }
        else { Connect-AzAccount -ErrorAction Stop | Out-Null }
        $ctx = Get-AzContext
    }
    if (-not $ctx) { throw 'Unable to establish an Azure context. Run Connect-AzAccount manually and retry.' }
    Write-Host "[azmon-assess] Signed in as $($ctx.Account.Id) (tenant $($ctx.Tenant.Id))." -ForegroundColor Green
    return $ctx
}
