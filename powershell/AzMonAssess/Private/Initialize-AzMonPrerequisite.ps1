#requires -Version 7.0

function Initialize-AzMonPrerequisite {
    <#
    .SYNOPSIS
        Verifies Az.Accounts is available and never installs anything —
        designed for restricted customer environments (Azure Cloud Shell,
        or any PowerShell 7+ session) where installing modules isn't
        possible. Safe to call repeatedly.
    .DESCRIPTION
        Az.Accounts is the ONLY module this tool depends on (it ships
        Connect-AzAccount, Get-AzContext, Get-AzSubscription and, crucially,
        Invoke-AzRestMethod / Get-AzAccessToken — used to call every Azure
        Resource Graph, Log Analytics, and ARM REST endpoint directly).
        Az.Accounts is preinstalled in Azure Cloud Shell and ships with any
        Az PowerShell install, so this function only checks — it never runs
        Install-Module. YAML triage plans are optional and already degrade
        to JSON automatically if the (equally optional) powershell-yaml
        module isn't present — see Invoke-AzMonTriage.ps1.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "The 'Az.Accounts' module is required and was not found. " +
              'Azure Cloud Shell has it preinstalled; for a local PowerShell 7+ session ' +
              'install it yourself with: Install-Module -Name Az.Accounts -Scope CurrentUser'
    }
    Import-Module -Name Az.Accounts -ErrorAction Stop

    if (-not (Get-AzContext)) {
        Write-Host '[azmon-assess] No Azure context found — sign in with Connect-AzAccount.' -ForegroundColor Yellow
    }
}

function Test-AzMonModuleAvailable {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Name)
    return [bool](Get-Module -ListAvailable -Name $Name)
}
