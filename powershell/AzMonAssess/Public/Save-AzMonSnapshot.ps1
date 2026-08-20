#requires -Version 7.0

function Save-AzMonSnapshot {
    <#
    .SYNOPSIS
        Writes a snapshot hashtable to snapshot.json.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Snapshot, [Parameter(Mandatory)] [string] $Path)
    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    ($Snapshot | ConvertTo-Json -Depth 25) | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
    return (Resolve-Path -LiteralPath $Path).Path
}

function Import-AzMonSnapshot {
    <#
    .SYNOPSIS
        Reads snapshot.json back into hashtables (matches the in-memory
        shape produced by the collectors, so reports/triage/remediate work
        identically on a fresh run or a reloaded snapshot).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)
    $raw = Get-Content -LiteralPath $Path -Raw
    $snap = $raw | ConvertFrom-Json -AsHashtable -Depth 25
    # Heals snapshots saved before the Invoke-AzMonGraphQuery data:null fix
    # (and guards against any other source of stray nulls): a null element
    # in one of these arrays crashes any consumer that indexes it directly,
    # e.g. report generation doing $s['LogsEnabled'].
    foreach ($key in @('Workspaces', 'AppInsights', 'AlertRules', 'ActionGroups', 'DiagnosticSettings', 'Resources', 'DataCollectionRules', 'Findings')) {
        if ($snap.ContainsKey($key)) {
            $snap[$key] = @($snap[$key] | Where-Object { $null -ne $_ })
        }
    }
    return $snap
}
