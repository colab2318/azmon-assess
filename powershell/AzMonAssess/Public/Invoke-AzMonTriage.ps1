#requires -Version 7.0
# Interactive + batch triage engine — ported from triage/engine.py +
# triage/models.py. Decisions auto-save after every entry (resumable).
# Batch plans support YAML (if the optional powershell-yaml module is
# installed) or JSON.

function New-AzMonTriageEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $FindingId,
        [Parameter(Mandatory)] [ValidateSet('accept', 'reject', 'snooze', 'defer', 'manual')] [string] $Decision,
        [string] $Owner,
        [string] $TargetDate,
        [string] $Notes,
        [string] $TriagedBy
    )
    return @{
        FindingId  = $FindingId
        Decision   = $Decision
        Owner      = $Owner
        TargetDate = $TargetDate
        Notes      = $Notes
        TriagedAt  = (Get-Date).ToUniversalTime().ToString('o')
        TriagedBy  = $TriagedBy
    }
}

function New-AzMonTriagePlan {
    [CmdletBinding()]
    param([string] $CustomerName = 'Your Organization', [string] $SnapshotPath, [array] $Entries = @())
    return @{
        CustomerName = $CustomerName
        SnapshotPath = $SnapshotPath
        GeneratedAt  = (Get-Date).ToUniversalTime().ToString('o')
        Entries      = @($Entries)
    }
}

function Save-AzMonTriagePlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Plan, [Parameter(Mandatory)] [string] $Path)
    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }
    ($Plan | ConvertTo-Json -Depth 15) | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
}

function Import-AzMonTriagePlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) { return (New-AzMonTriagePlan) }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw
        $data = $raw | ConvertFrom-Json -AsHashtable -Depth 15
        if (-not $data.ContainsKey('Entries')) { $data['Entries'] = @() }
        return $data
    } catch {
        Write-Warning "Could not parse existing triage plan at $Path ($($_.Exception.Message)) - starting fresh."
        return (New-AzMonTriagePlan)
    }
}

function Show-AzMonTriageFinding {
    param([int] $Index, [int] $Total, [hashtable] $Finding)
    $sevColorMap = @{ critical = 'Red'; high = 'Yellow'; medium = 'Cyan'; low = 'Green'; info = 'White' }
    $sevColor = $sevColorMap[$Finding['Severity']]
    if (-not $sevColor) { $sevColor = 'White' }
    Write-Host ''
    Write-Host "Finding $Index/$Total  (id=$($Finding['Id']))" -ForegroundColor DarkGray
    Write-Host $Finding['Title'] -ForegroundColor $sevColor
    Write-Host $Finding['Detail']
    Write-Host "Category: $($Finding['Category'])   Severity: $(([string]$Finding['Severity']).ToUpperInvariant())" -ForegroundColor DarkGray
    $savings = if ($Finding['EstimatedMonthlySavingsUsd']) { "$(Format-AzMonUsd $Finding['EstimatedMonthlySavingsUsd'])/mo" } else { '-' }
    Write-Host "Est. savings: $savings   Resources: $(@($Finding['ResourceIds']).Count)" -ForegroundColor DarkGray
    if ($Finding['Recommendation']) { Write-Host "Recommendation: $($Finding['Recommendation'])" -ForegroundColor DarkCyan }
}

function Show-AzMonTriageSummary {
    param([hashtable] $Plan)
    $counts = @{}
    foreach ($e in @($Plan['Entries'])) { $counts[$e['Decision']] = ($counts[$e['Decision']] ?? 0) + 1 }
    Write-Host ''
    Write-Host 'Triage summary:' -ForegroundColor Cyan
    foreach ($d in @('accept', 'manual', 'defer', 'snooze', 'reject')) {
        Write-Host ('  {0,-8} {1}' -f $d, ($counts[$d] ?? 0))
    }
}

function Invoke-AzMonTriageInteractive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $OutputPath,
        [int] $Limit
    )
    $plan = Import-AzMonTriagePlan -Path $OutputPath
    if (-not $plan['CustomerName']) { $plan['CustomerName'] = $Snapshot['CustomerName'] }
    $already = @{}
    foreach ($e in @($plan['Entries'])) { $already[$e['FindingId']] = $true }

    $findings = Sort-AzMonFinding -Finding @($Snapshot['Findings'])
    $remaining = @($findings | Where-Object { -not $already.ContainsKey($_['Id']) })
    if ($Limit -and $Limit -gt 0) { $remaining = @($remaining | Select-Object -First $Limit) }

    Write-Host "Triage session  customer=$($Snapshot['CustomerName'])  findings=$($findings.Count)  already-triaged=$($already.Count)  remaining=$($remaining.Count)" -ForegroundColor Cyan
    if ($remaining.Count -eq 0) {
        Write-Host 'Nothing left to triage.' -ForegroundColor Green
        return $plan
    }

    $decisionMap = @{ a = 'accept'; r = 'reject'; s = 'snooze'; d = 'defer'; m = 'manual' }
    $triagedBy = [System.Environment]::UserName
    $idx = 0
    foreach ($f in $remaining) {
        $idx++
        Show-AzMonTriageFinding -Index $idx -Total $remaining.Count -Finding $f
        $answer = ''
        while ($answer -notin @('a', 'r', 's', 'd', 'm', 'q')) {
            $raw = Read-Host 'Decision ([a]ccept / [r]eject / [s]nooze / [d]efer / [m]anual / [q]uit) [s]'
            if ([string]::IsNullOrWhiteSpace($raw)) { $raw = 's' }
            $answer = $raw.Trim().ToLowerInvariant().Substring(0, 1)
        }
        if ($answer -eq 'q') {
            Write-Host 'Session paused. Re-run to resume.' -ForegroundColor Yellow
            break
        }
        $owner = Read-Host 'Owner (optional)'
        $targetDate = Read-Host 'Target date YYYY-MM-DD (optional)'
        $notes = Read-Host 'Notes (optional)'
        $plan['Entries'] += New-AzMonTriageEntry -FindingId $f['Id'] -Decision $decisionMap[$answer] -Owner $owner -TargetDate $targetDate -Notes $notes -TriagedBy $triagedBy
        Save-AzMonTriagePlan -Plan $plan -Path $OutputPath
    }

    Save-AzMonTriagePlan -Plan $plan -Path $OutputPath
    Show-AzMonTriageSummary -Plan $plan
    return $plan
}

function Invoke-AzMonTriageBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable] $Snapshot,
        [Parameter(Mandatory)] [string] $PlanFilePath,
        [Parameter(Mandatory)] [string] $OutputPath
    )
    $isYaml = $PlanFilePath -match '\.ya?ml$'
    $raw = Get-Content -LiteralPath $PlanFilePath -Raw
    if ($isYaml) {
        if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
            throw "Reading a YAML triage plan requires the 'powershell-yaml' module. Install it with: Install-Module powershell-yaml -Scope CurrentUser (or provide a .json plan instead)."
        }
        Import-Module powershell-yaml -ErrorAction Stop
        $doc = ConvertFrom-Yaml -Yaml $raw
    } else {
        $doc = $raw | ConvertFrom-Json -AsHashtable -Depth 15
    }
    if (-not $doc) { $doc = @{} }

    $validIds = @{}
    foreach ($f in @($Snapshot['Findings'])) { $validIds[$f['Id']] = $true }

    $plan = New-AzMonTriagePlan -CustomerName $Snapshot['CustomerName']
    $entriesSrc = $doc['entries']
    if (-not $entriesSrc) { $entriesSrc = $doc['decisions'] }
    foreach ($row in @($entriesSrc)) {
        $fid = $row['finding_id']
        if (-not $fid) { $fid = $row['id'] }
        if (-not $fid -or -not $validIds.ContainsKey($fid)) {
            Write-Warning "Skipping unknown finding_id $fid"
            continue
        }
        $decision = ([string]$row['decision']).ToLowerInvariant()
        if ($decision -notin @('accept', 'reject', 'snooze', 'defer', 'manual')) {
            Write-Warning "Skipping bad decision on $fid"
            continue
        }
        $triagedBy = if ($row['triaged_by']) { [string]$row['triaged_by'] } else { 'batch' }
        $plan['Entries'] += New-AzMonTriageEntry -FindingId $fid -Decision $decision -Owner $row['owner'] `
            -TargetDate ([string]$row['target_date']) -Notes $row['notes'] -TriagedBy $triagedBy
    }
    Save-AzMonTriagePlan -Plan $plan -Path $OutputPath
    Show-AzMonTriageSummary -Plan $plan
    return $plan
}

function Export-AzMonTriageTemplate {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Snapshot, [Parameter(Mandatory)] [string] $Path)

    $entries = @(@($Snapshot['Findings']) | ForEach-Object {
            [ordered]@{
                finding_id  = $_['Id']
                title       = $_['Title']
                severity    = $_['Severity']
                decision    = 'snooze'
                owner       = $null
                target_date = $null
                notes       = $null
            }
        })
    $doc = [ordered]@{ customer_name = $Snapshot['CustomerName']; entries = $entries }

    $dir = Split-Path -Parent $Path
    if ($dir) { New-AzMonOutputDirectory -Path $dir | Out-Null }

    if ($Path -match '\.ya?ml$' -and -not (Get-Module -ListAvailable -Name powershell-yaml)) {
        Write-Warning "'powershell-yaml' module not found - writing JSON instead of YAML."
        $Path = [System.IO.Path]::ChangeExtension($Path, '.json')
    }
    if ($Path -match '\.ya?ml$') {
        Import-Module powershell-yaml -ErrorAction Stop
        (ConvertTo-Yaml $doc) | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
    } else {
        ($doc | ConvertTo-Json -Depth 15) | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
    }
    return $Path
}

function Invoke-AzMonTriage {
    <#
    .SYNOPSIS
        Triage findings interactively (default), from a batch YAML/JSON
        plan (-PlanPath), or emit a pre-populated template (-EmitTemplate).
    .EXAMPLE
        Invoke-AzMonTriage -SnapshotPath ./out/snapshot.json -OutputPath ./out/triage.json
    .EXAMPLE
        Invoke-AzMonTriage -SnapshotPath ./out/snapshot.json -OutputPath ./out/triage.json -EmitTemplate
    .EXAMPLE
        Invoke-AzMonTriage -SnapshotPath ./out/snapshot.json -OutputPath ./out/triage.json -PlanPath ./out/triage-plan.yaml
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SnapshotPath,
        [string] $OutputPath = './out/triage.json',
        [string] $PlanPath,
        [switch] $EmitTemplate,
        [int] $Limit
    )
    $snapshot = Import-AzMonSnapshot -Path $SnapshotPath

    if ($EmitTemplate) {
        $dir = Split-Path -Parent $OutputPath
        if (-not $dir) { $dir = '.' }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
        $ext = if (Get-Module -ListAvailable -Name powershell-yaml) { 'yaml' } else { 'json' }
        $target = Join-Path $dir "$baseName.template.$ext"
        $written = Export-AzMonTriageTemplate -Snapshot $snapshot -Path $target
        Write-Host "[azmon-assess] Template written to $written" -ForegroundColor Green
        return
    }
    if ($PlanPath) {
        return Invoke-AzMonTriageBatch -Snapshot $snapshot -PlanFilePath $PlanPath -OutputPath $OutputPath
    }
    return Invoke-AzMonTriageInteractive -Snapshot $snapshot -OutputPath $OutputPath -Limit $Limit
}
