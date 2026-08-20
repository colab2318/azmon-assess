#requires -Version 7.0
# Small shared helpers used across collectors / analyzers / reports.

function ConvertTo-AzMonPortalUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceId)
    return "https://portal.azure.com/#@/resource$ResourceId/overview"
}

$script:AzMonSeverityRank = @{ critical = 0; high = 1; medium = 2; low = 3; info = 4 }

function Get-AzMonSeverityRank {
    [CmdletBinding()]
    param([string] $Severity)
    if ($script:AzMonSeverityRank.ContainsKey($Severity)) { return $script:AzMonSeverityRank[$Severity] }
    return 5
}

function Sort-AzMonFinding {
    <#
    .SYNOPSIS
        Sort findings by severity (critical..info) then by descending
        estimated savings — the same ordering used throughout the reports.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [array] $Finding)
    return @($Finding | Sort-Object -Property `
        @{ Expression = { Get-AzMonSeverityRank $_['Severity'] } }, `
        @{ Expression = { - [double]($_['EstimatedMonthlySavingsUsd'] ?? 0) } })
}

function ConvertTo-AzMonHashtable {
    <#
    .SYNOPSIS
        Recursively converts PSCustomObject/JSON output (as produced by
        ConvertFrom-Json without -AsHashtable) into nested hashtables/arrays.
        Defensive helper for callers that hand us mixed object graphs.
    #>
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline)] $InputObject)
    process {
        if ($null -eq $InputObject) { return $null }
        if ($InputObject -is [hashtable]) {
            $h = @{}
            foreach ($k in $InputObject.Keys) { $h[$k] = ConvertTo-AzMonHashtable $InputObject[$k] }
            return $h
        }
        if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
            $h = @{}
            foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-AzMonHashtable $p.Value }
            return $h
        }
        if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
            return @($InputObject | ForEach-Object { ConvertTo-AzMonHashtable $_ })
        }
        return $InputObject
    }
}

function New-AzMonOutputDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Format-AzMonUsd {
    [CmdletBinding()]
    param([double] $Value)
    return ('${0:N0}' -f $Value)
}

function Get-AzMonCountByType {
    <#
    .SYNOPSIS
        Counts ResourceRef records by Type, descending — used by several
        analyzers for the "by_type" evidence breakdown. Returns an array of
        @{ Type; Count } (not a dictionary) so descending order survives
        JSON round-tripping without depending on OrderedDictionary.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [array] $ResourceRef)
    $counts = @{}
    foreach ($r in $ResourceRef) {
        $t = $r['Type']
        $counts[$t] = ($counts[$t] ?? 0) + 1
    }
    return @($counts.Keys | Sort-Object { -$counts[$_] } | ForEach-Object { @{ type = $_; count = $counts[$_] } })
}

function Test-AzMonProdTag {
    <#
    .SYNOPSIS
        True if a tag bag marks the resource as production (env/environment
        tag in {prod, production, prd}) — shared by reliability/security.
    #>
    [CmdletBinding()]
    param([hashtable] $Tags)
    if (-not $Tags) { return $false }
    $envTag = $Tags['env']
    if (-not $envTag) { $envTag = $Tags['environment'] }
    if (-not $envTag) { $envTag = $Tags['Environment'] }
    if (-not $envTag) { return $false }
    return (([string]$envTag).ToLowerInvariant()) -in @('prod', 'production', 'prd')
}

function ConvertFrom-AzMonIso8601Duration {
    <#
    .SYNOPSIS
        Parses a simple ISO 8601 duration (as used by evaluationFrequency /
        windowSize on scheduled query rules, e.g. 'PT5M', 'PT1H') into total
        minutes. Returns $null if the input is empty or unparseable.
    #>
    [CmdletBinding()]
    param([string] $Duration)
    if ([string]::IsNullOrWhiteSpace($Duration)) { return $null }
    $m = [regex]::Match($Duration, '^P(?:(?<days>\d+)D)?(?:T(?:(?<hours>\d+)H)?(?:(?<minutes>\d+)M)?(?:(?<seconds>\d+)S)?)?$')
    if (-not $m.Success) { return $null }
    $days = if ($m.Groups['days'].Success) { [double]$m.Groups['days'].Value } else { 0 }
    $hours = if ($m.Groups['hours'].Success) { [double]$m.Groups['hours'].Value } else { 0 }
    $minutes = if ($m.Groups['minutes'].Success) { [double]$m.Groups['minutes'].Value } else { 0 }
    $seconds = if ($m.Groups['seconds'].Success) { [double]$m.Groups['seconds'].Value } else { 0 }
    return ($days * 1440) + ($hours * 60) + $minutes + ($seconds / 60.0)
}

$script:AzMonAlertKindType = @{
    metric      = 'microsoft.insights/metricalerts'
    log         = 'microsoft.insights/scheduledqueryrules'
    activityLog = 'microsoft.insights/activitylogalerts'
}

function New-AzMonResourceLookup {
    <#
    .SYNOPSIS
        Builds a lowercased-resource-id -> descriptor (@{Name;Type;
        ResourceGroup;Location;SubscriptionId}) lookup across every
        collected entity in a snapshot — used to enrich a Finding's bare
        ResourceIds for resource-centric report sections.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Snapshot)

    $lookup = @{}
    foreach ($ws in @($Snapshot['Workspaces'])) {
        $lookup[([string]$ws['Id']).ToLowerInvariant()] = @{
            Name = $ws['Name']; Type = 'microsoft.operationalinsights/workspaces'
            ResourceGroup = $ws['ResourceGroup']; Location = $ws['Location']; SubscriptionId = $ws['SubscriptionId']
        }
    }
    foreach ($ai in @($Snapshot['AppInsights'])) {
        $lookup[([string]$ai['Id']).ToLowerInvariant()] = @{
            Name = $ai['Name']; Type = 'microsoft.insights/components'
            ResourceGroup = $ai['ResourceGroup']; Location = $ai['Location']; SubscriptionId = $ai['SubscriptionId']
        }
    }
    foreach ($r in @($Snapshot['Resources'])) {
        $lookup[([string]$r['Id']).ToLowerInvariant()] = @{
            Name = $r['Name']; Type = $r['Type']
            ResourceGroup = $r['ResourceGroup']; Location = $r['Location']; SubscriptionId = $r['SubscriptionId']
        }
    }
    foreach ($rule in @($Snapshot['AlertRules'])) {
        $lookup[([string]$rule['Id']).ToLowerInvariant()] = @{
            Name = $rule['Name']; Type = $script:AzMonAlertKindType[$rule['AlertKind']]
            ResourceGroup = $rule['ResourceGroup']; Location = ''; SubscriptionId = $rule['SubscriptionId']
        }
    }
    foreach ($ag in @($Snapshot['ActionGroups'])) {
        $lookup[([string]$ag['Id']).ToLowerInvariant()] = @{
            Name = $ag['Name']; Type = 'microsoft.insights/actiongroups'
            ResourceGroup = $ag['ResourceGroup']; Location = ''; SubscriptionId = $ag['SubscriptionId']
        }
    }
    foreach ($dcr in @($Snapshot['DataCollectionRules'])) {
        $lookup[([string]$dcr['Id']).ToLowerInvariant()] = @{
            Name = $dcr['Name']; Type = 'microsoft.insights/datacollectionrules'
            ResourceGroup = $dcr['ResourceGroup']; Location = $dcr['Location']; SubscriptionId = $dcr['SubscriptionId']
        }
    }
    return $lookup
}

function Resolve-AzMonResourceDescriptor {
    <#
    .SYNOPSIS
        Resolves one resource ID to a descriptor via New-AzMonResourceLookup,
        falling back to parsing the bare ARM ID when the resource wasn't
        otherwise collected (e.g. an alert rule's scope resource).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceId, [Parameter(Mandatory)] [hashtable] $Lookup)

    $key = $ResourceId.ToLowerInvariant()
    if ($Lookup.ContainsKey($key)) { return $Lookup[$key] }

    $parts = ConvertFrom-AzMonResourceId -ResourceId $ResourceId
    $type = ''
    $m = [regex]::Match($ResourceId, '/providers/([^/]+)/([^/]+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($m.Success) { $type = "$($m.Groups[1].Value)/$($m.Groups[2].Value)".ToLowerInvariant() }
    return @{ Name = $parts.Name; Type = $type; ResourceGroup = $parts.ResourceGroup; Location = ''; SubscriptionId = $parts.SubscriptionId }
}
