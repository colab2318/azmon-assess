#requires -Version 7.0

$script:AzMonLegacyAgentExtensionTypes = @('MicrosoftMonitoringAgent', 'OmsAgentForLinux')
$script:AzMonAmaExtensionTypes = @('AzureMonitorWindowsAgent', 'AzureMonitorLinuxAgent')

function Get-AzMonVmAgentExtension {
    <#
    .SYNOPSIS
        Collects VM extensions that identify which monitoring agent (if any)
        is installed — the retired Log Analytics agent (MMA/OMS) or the
        Azure Monitor Agent (AMA) — via Azure Resource Graph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string[]] $SubscriptionId)

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQVmExtensions -SubscriptionId $SubscriptionId
    return @(
        foreach ($r in $rows) {
            $extType = [string]$r.extensionType
            $kind = if ($script:AzMonLegacyAgentExtensionTypes -contains $extType) { 'legacy' }
                elseif ($script:AzMonAmaExtensionTypes -contains $extType) { 'ama' }
                else { $null }
            if ($kind) {
                @{
                    Kind           = 'VmAgentExtension'
                    VmId           = $r.vmId
                    AgentKind      = $kind
                    ExtensionType  = $extType
                    SubscriptionId = $r.subscriptionId
                    ResourceGroup  = $r.resourceGroup
                }
            }
        }
    )
}
