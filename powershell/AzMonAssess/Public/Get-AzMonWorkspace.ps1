#requires -Version 7.0

function Get-AzMonWorkspace {
    <#
    .SYNOPSIS
        Collects Log Analytics workspaces via Azure Resource Graph.
    .PARAMETER SkipIngestionEnrichment
        Skip the per-workspace ingestion/heartbeat KQL enrichment (faster,
        but consolidation/cost/performance analyzers need it to fire).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $SubscriptionId,
        [switch] $SkipIngestionEnrichment,
        [int] $LookbackDays = 30,
        [int] $ThrottleLimit = 8
    )

    $rows = Invoke-AzMonGraphQuery -Query $script:AzMonQWorkspaces -SubscriptionId $SubscriptionId
    $workspaces = @(
        foreach ($r in $rows) {
            New-AzMonWorkspace -Id $r.id -Name $r.name -SubscriptionId $r.subscriptionId -ResourceGroup $r.resourceGroup `
                -Location $r.location -Sku $r.sku -RetentionDays $r.retentionInDays -DailyQuotaGb $r.dailyQuotaGb `
                -CustomerId $r.customerId -Tags (ConvertTo-AzMonHashtable $r.tags) `
                -PublicNetworkAccessForIngestion $r.publicNetworkAccessForIngestion `
                -PublicNetworkAccessForQuery $r.publicNetworkAccessForQuery `
                -DisableLocalAuth $r.disableLocalAuth -CapacityReservationLevel $r.capacityReservationLevel `
                -ClusterResourceId $r.clusterResourceId
        }
    )

    if (-not $SkipIngestionEnrichment -and $workspaces.Count -gt 0) {
        $workspaces = Add-AzMonWorkspaceIngestion -Workspace $workspaces -LookbackDays $LookbackDays -ThrottleLimit $ThrottleLimit
    }
    return $workspaces
}

function Add-AzMonWorkspaceIngestion {
    <#
    .SYNOPSIS
        Enriches workspaces with 30-day ingestion-by-table + connected-source
        (Heartbeat) counts via Log Analytics KQL, run with bounded
        concurrency (mirrors the Python ThreadPoolExecutor behaviour).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array] $Workspace,
        [int] $LookbackDays = 30,
        [int] $ThrottleLimit = 8
    )

    # NOTE: ForEach-Object -Parallel runspaces cannot see functions defined
    # in this module (only $using: variables), so the Log Analytics REST
    # call is inlined here rather than calling Invoke-AzMonLogAnalyticsQuery.
    $context = Get-AzContext
    $results = $Workspace | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $ws = $_
        $days = $using:LookbackDays
        $ctx = $using:context

        if (-not $ws['CustomerId']) { return $ws }
        try {
            Import-Module Az.Accounts -ErrorAction Stop
            if ($ctx) { $null = Set-AzContext -Context $ctx -ErrorAction Stop }

            function Invoke-LaQuery($WorkspaceId, $Query, $TimespanIso) {
                $body = @{ query = $Query; timespan = $TimespanIso } | ConvertTo-Json -Depth 5
                $uri = "https://api.loganalytics.io/v1/workspaces/$WorkspaceId/query"
                $resp = Invoke-AzRestMethod -Method POST -Uri $uri -ResourceId 'https://api.loganalytics.io' -Payload $body -ErrorAction Stop
                if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
                    throw "Log Analytics query failed with status $($resp.StatusCode): $($resp.Content)"
                }
                $parsed = $resp.Content | ConvertFrom-Json
                $table = $parsed.tables | Select-Object -First 1
                if (-not $table) { return @() }
                $colNames = @($table.columns | ForEach-Object { $_.name })
                return @(foreach ($row in @($table.rows)) {
                        $obj = [ordered]@{}
                        for ($i = 0; $i -lt $colNames.Count; $i++) { $obj[$colNames[$i]] = $row[$i] }
                        [pscustomobject]$obj
                    })
            }

            $ingestionQuery = @"
Usage
| where TimeGenerated > ago(${days}d)
| where IsBillable == true
| summarize BillableGB = sum(Quantity) / 1024.0 by DataType
| order by BillableGB desc
"@
            $heartbeatQuery = @'
Heartbeat
| where TimeGenerated > ago(1d)
| summarize ConnectedSources = dcount(Computer)
'@
            $ingestionRows = Invoke-LaQuery -WorkspaceId $ws['CustomerId'] -Query $ingestionQuery -TimespanIso "P${days}D"
            $byTable = @{}
            foreach ($row in $ingestionRows) {
                if ($row.DataType) { $byTable[[string]$row.DataType] = [double]$row.BillableGB }
            }
            $ws['IngestionByTable'] = $byTable
            $total = 0.0
            foreach ($v in $byTable.Values) { $total += $v }
            $ws['IngestionGb30d'] = [Math]::Round($total, 3)

            $heartbeatRows = Invoke-LaQuery -WorkspaceId $ws['CustomerId'] -Query $heartbeatQuery -TimespanIso 'P1D'
            if ($heartbeatRows.Count -gt 0 -and $heartbeatRows[0].ConnectedSources) {
                $ws['ConnectedSources'] = [int]$heartbeatRows[0].ConnectedSources
            }
        } catch {
            Write-Warning "Ingestion enrichment failed for $($ws['Name']): $($_.Exception.Message)"
        }
        return $ws
    }
    return @($results)
}

function Get-AzMonHeartbeatResourceId {
    <#
    .SYNOPSIS
        Returns the union of resource IDs (lowercased) with a Heartbeat row
        in the last 2 days across all given workspaces — used to flag VMs
        with no heartbeat in any connected workspace.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array] $Workspace,
        [int] $ThrottleLimit = 8
    )
    if (-not $Workspace -or $Workspace.Count -eq 0) { return [System.Collections.Generic.HashSet[string]]::new() }

    # NOTE: inlined REST call — see comment in Add-AzMonWorkspaceIngestion.
    $context = Get-AzContext
    $perWorkspaceIds = $Workspace | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $ws = $_
        $ctx = $using:context
        if (-not $ws['CustomerId']) { return @() }
        try {
            Import-Module Az.Accounts -ErrorAction Stop
            if ($ctx) { $null = Set-AzContext -Context $ctx -ErrorAction Stop }

            $query = @'
Heartbeat
| where TimeGenerated > ago(2d)
| where isnotempty(ResourceId)
| summarize by ResourceId
'@
            $body = @{ query = $query; timespan = 'P2D' } | ConvertTo-Json -Depth 5
            $uri = "https://api.loganalytics.io/v1/workspaces/$($ws['CustomerId'])/query"
            $resp = Invoke-AzRestMethod -Method POST -Uri $uri -ResourceId 'https://api.loganalytics.io' -Payload $body -ErrorAction Stop
            if ($resp.StatusCode -lt 200 -or $resp.StatusCode -ge 300) {
                throw "Log Analytics query failed with status $($resp.StatusCode): $($resp.Content)"
            }
            $parsed = $resp.Content | ConvertFrom-Json
            $table = $parsed.tables | Select-Object -First 1
            $colNames = @(if ($table) { $table.columns | ForEach-Object { $_.name } })
            $ridIdx = [array]::IndexOf($colNames, 'ResourceId')
            $ids = @()
            if ($table -and $ridIdx -ge 0) {
                $ids = @($table.rows | ForEach-Object { ([string]$_[$ridIdx]).ToLowerInvariant() } | Where-Object { $_ })
            }
            $ids
        } catch {
            Write-Warning "Heartbeat resource-id query failed for $($ws['Name']): $($_.Exception.Message)"
            @()
        }
    }

    $all = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($id in @($perWorkspaceIds)) { $null = $all.Add($id) }
    return $all
}
