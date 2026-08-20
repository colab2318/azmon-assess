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
    return ($raw | ConvertFrom-Json -AsHashtable -Depth 25)
}
