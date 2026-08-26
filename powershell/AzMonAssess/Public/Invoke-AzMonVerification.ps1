#requires -Version 7.0
# Environment reconciliation: re-checks every finding in a snapshot against
# the customer's current Azure state, and marks findings that no longer
# reproduce as reviewed directly in the existing report.xlsx. Deliberately
# reuses the same Find-AzMonXFinding analyzers used at assessment time
# (rather than a second bespoke set of "is this fixed" checks) so a
# finding is considered resolved exactly when the tool would no longer
# raise it fresh - single source of truth for "what counts as a problem".

$script:AzMonReviewStatus = @{
    NotReviewed  = 'Not Reviewed'
    Implemented  = 'Reviewed - Implemented'
    ManualReview = 'Needs Manual Review'
    Regressed    = 'Regression - Reopened'
}

function Test-AzMonResourceStillExists {
    <#
    .SYNOPSIS
        True if a resource/subscription ID from an original finding can
        still be located in the freshly-collected environment state.
    #>
    [CmdletBinding()]
    param(
        [string] $ResourceId,
        [Parameter(Mandatory)] [hashtable] $Lookup,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]] $SubscriptionId
    )
    if (-not $ResourceId) { return $false }
    $key = $ResourceId.ToLowerInvariant()
    if ($Lookup.ContainsKey($key)) { return $true }
    if ($key -match '^/subscriptions/([^/]+)$') { return $SubscriptionId.Contains($matches[1]) }
    return $false
}

function Get-AzMonFindingCheckKey {
    <#
    .SYNOPSIS
        The signature component identifying WHICH check/rule raised a
        finding - CheckId when present, falling back to Category for
        snapshots saved before CheckId existed. Using bare Category alone
        is NOT enough: several distinct finding types commonly fire for
        the same resource within one category (e.g. an alert rule that is
        both disabled AND missing severity, both category 'alerting') -
        without CheckId, fixing just one makes the OTHER still-open
        finding mask it as "still open" for both.
    .PARAMETER UseCheckId
        Must be the SAME value for both the original and the fresh finding
        set being compared. If the original snapshot predates CheckId, its
        findings can never match a fresh finding's real CheckId (different
        key spaces entirely) - pass $false to force Category-only matching
        consistently on both sides instead of silently comparing
        incompatible keys.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [hashtable] $Finding, [bool] $UseCheckId = $true)
    if ($UseCheckId -and $Finding['CheckId']) { return $Finding['CheckId'] }
    return $Finding['Category']
}

function Find-AzMonVerificationStatusUpdate {
    <#
    .SYNOPSIS
        Core reconciliation: for every (finding, resource) row in the
        original snapshot, decides whether the current environment still
        exhibits the same (CheckId, ResourceId) condition, and returns
        only the rows whose status should actually change.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array] $OriginalFinding,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]] $OpenSignature,
        [Parameter(Mandatory)] [hashtable] $FreshLookup,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]] $FreshSubscriptionId,
        [Parameter(Mandatory)] [hashtable] $PriorStatus,
        [bool] $UseCheckId = $true,
        [System.Collections.Generic.List[hashtable]] $Trace
    )
    $updates = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($f in $OriginalFinding) {
        $checkKey = Get-AzMonFindingCheckKey -Finding $f -UseCheckId $UseCheckId
        $resourceIds = @($f['ResourceIds'])
        foreach ($rid in $resourceIds) {
            if (-not $rid) { continue }
            $ridLower = ([string]$rid).ToLowerInvariant()
            $priorKey = "$($f['Id'])|$ridLower"
            $prior = if ($PriorStatus.ContainsKey($priorKey)) { $PriorStatus[$priorKey] } else { $script:AzMonReviewStatus.NotReviewed }
            $stillOpen = $OpenSignature.Contains("$checkKey|$ridLower")

            $newStatus = $prior
            $reason = $null
            if ($stillOpen) {
                if ($prior -eq $script:AzMonReviewStatus.Implemented) {
                    $newStatus = $script:AzMonReviewStatus.Regressed
                    $reason = 'Previously marked implemented, but the same condition was detected again in the latest check.'
                } elseif ($prior -eq $script:AzMonReviewStatus.ManualReview) {
                    $newStatus = $script:AzMonReviewStatus.NotReviewed
                    $reason = 'Resource is visible again and the condition still reproduces - no longer ambiguous.'
                }
            } elseif (Test-AzMonResourceStillExists -ResourceId $rid -Lookup $FreshLookup -SubscriptionId $FreshSubscriptionId) {
                $newStatus = $script:AzMonReviewStatus.Implemented
                $reason = 'Condition no longer detected for this resource in the latest check; resource confirmed present.'
            } else {
                $newStatus = $script:AzMonReviewStatus.ManualReview
                $reason = 'Condition no longer detected, but the resource could not be confirmed in the latest check (deleted, renamed, or inaccessible) - verify manually.'
            }

            if ($newStatus -ne $prior) {
                $updates.Add(@{
                        FindingId      = $f['Id']
                        ResourceId     = $rid
                        Category       = $f['Category']
                        Title          = $f['Title']
                        PreviousStatus = $prior
                        NewStatus      = $newStatus
                        Reason         = $reason
                    })
            }
            if ($null -ne $Trace) {
                $Trace.Add(@{
                        FindingId      = $f['Id']
                        CheckId        = $f['CheckId']
                        CheckKeyUsed   = $checkKey
                        Category       = $f['Category']
                        Title          = $f['Title']
                        ResourceId     = $rid
                        StillOpen      = $stillOpen
                        PreviousStatus = $prior
                        ComputedStatus = $newStatus
                        WillChange     = ($newStatus -ne $prior)
                        Reason         = $reason
                    })
            }
        }
    }
    return $updates.ToArray()
}

function New-AzMonVerificationRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $FindingId,
        [string] $ResourceId,
        [string] $Category,
        [string] $Title,
        [Parameter(Mandatory)] [string] $PreviousStatus,
        [Parameter(Mandatory)] [string] $NewStatus,
        [string] $Reason,
        [Parameter(Mandatory)] [string] $VerifiedAt,
        [string] $ReportPath
    )
    return @{
        FindingId      = $FindingId
        ResourceId     = $ResourceId
        Category       = $Category
        Title          = $Title
        PreviousStatus = $PreviousStatus
        NewStatus      = $NewStatus
        Reason         = $Reason
        VerifiedAt     = $VerifiedAt
        VerifiedBy     = [System.Environment]::UserName
        ReportPath     = $ReportPath
    }
}

function Invoke-AzMonVerification {
    <#
    .SYNOPSIS
        Re-checks every finding in a snapshot against the customer's
        current Azure environment. Findings whose condition no longer
        reproduces are marked "Reviewed - Implemented" directly in the
        existing report.xlsx's 4.ImpactedResourcesAnalysis sheet, in
        place - every other cell, style, and sheet is left untouched.
        Findings that can't be conclusively re-checked (resource missing
        from the current inventory) are flagged "Needs Manual Review".
        A finding previously marked implemented that reappears is flagged
        "Regression - Reopened". Every status change is appended to
        verification-log.json for audit purposes.
    .PARAMETER CurrentSnapshotPath
        Path to an already-generated, up-to-date snapshot.json to compare
        against (e.g. from re-running `run`/`demo`). Use this OR -Live.
    .PARAMETER Live
        Re-collect and re-analyze directly against Azure right now, scoped
        to the same subscriptions recorded in the original snapshot. The
        caller is responsible for an existing signed-in session (this
        function does not call Connect-AzMonSession itself).
    .EXAMPLE
        Invoke-AzMonVerification -SnapshotPath ./out/snapshot.json -Live
    .EXAMPLE
        Invoke-AzMonVerification -SnapshotPath ./out/snapshot.json -CurrentSnapshotPath ./out-rerun/snapshot.json
    .PARAMETER Detail
        Also write verification-detail.json: one row per (finding,
        resource) actually EVALUATED (not just ones that changed), with
        the CheckId used, whether it was still found open in the fresh
        check, and the resulting decision - use this to see exactly why a
        specific finding did or didn't change status.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SnapshotPath,
        [string] $ReportPath,
        [string] $OutputPath,
        [string] $CurrentSnapshotPath,
        [switch] $Live,
        [string[]] $SubscriptionId,
        [string] $ManagementGroupId,
        [int] $LookbackDays,
        [int] $ThrottleLimit,
        [switch] $Detail
    )

    if (-not $CurrentSnapshotPath -and -not $Live) {
        throw 'Invoke-AzMonVerification requires either -CurrentSnapshotPath (an already-generated fresh snapshot) or -Live (re-collect from Azure now).'
    }
    if (-not (Test-Path -LiteralPath $SnapshotPath)) { throw "-SnapshotPath not found: $SnapshotPath" }
    if ($CurrentSnapshotPath -and -not (Test-Path -LiteralPath $CurrentSnapshotPath)) {
        throw "-CurrentSnapshotPath not found: $CurrentSnapshotPath - check the path (it must already exist; this command does not create it), or use -Live to re-collect from Azure instead."
    }

    $snapshotDir = Split-Path -Parent (Resolve-Path -LiteralPath $SnapshotPath).Path
    if (-not $ReportPath) { $ReportPath = Join-Path $snapshotDir 'report.xlsx' }
    if (-not $OutputPath) { $OutputPath = $snapshotDir }
    if (-not (Test-Path -LiteralPath $ReportPath)) { throw "report.xlsx not found at $ReportPath - generate it first (report/run/demo command), or pass -ReportPath explicitly." }

    $snapshot = Import-AzMonSnapshot -Path $SnapshotPath
    $originalFindings = @($snapshot['Findings'])
    Write-Host "[azmon-assess] Verifying $($originalFindings.Count) finding(s) from $SnapshotPath against the current environment..." -ForegroundColor Cyan

    $useCheckId = [bool]($originalFindings | Where-Object { $_['CheckId'] } | Select-Object -First 1)
    if ($originalFindings.Count -gt 0 -and -not $useCheckId) {
        Write-Warning ("$SnapshotPath predates CheckId support (regenerated before this feature shipped). Falling back to " +
            'Category-only matching for this run, which cannot tell apart multiple finding types on the same resource ' +
            "within one category - re-run 'run'/'demo' to produce a fresh baseline snapshot for full accuracy.")
    }

    if ($CurrentSnapshotPath) {
        $freshSnapshot = Import-AzMonSnapshot -Path $CurrentSnapshotPath
    } else {
        $subsForLive = if ($SubscriptionId) { $SubscriptionId } else { @($snapshot['SubscriptionIds']) }
        $scratchDir = Join-Path $OutputPath '.verification-scratch'
        Write-Host '[azmon-assess] Live mode: re-collecting and re-analyzing current Azure state...' -ForegroundColor DarkCyan
        $freshSnapshot = Invoke-AzMonAssessment -SubscriptionId $subsForLive -ManagementGroupId $ManagementGroupId -CustomerName $snapshot['CustomerName'] `
            -LookbackDays $LookbackDays -ThrottleLimit $ThrottleLimit -OutputPath $scratchDir -Only @('all') -Report @('none') -SkipAiSummary
    }
    $freshFindings = @($freshSnapshot['Findings'])
    Write-Host "[azmon-assess] Current environment yields $($freshFindings.Count) finding(s)." -ForegroundColor DarkCyan

    $openSignature = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($f in $freshFindings) {
        $checkKey = Get-AzMonFindingCheckKey -Finding $f -UseCheckId $useCheckId
        foreach ($rid in @($f['ResourceIds'])) {
            if (-not $rid) { continue }
            [void]$openSignature.Add("$checkKey|$(([string]$rid).ToLowerInvariant())")
        }
    }
    $freshLookup = New-AzMonResourceLookup -Snapshot $freshSnapshot
    $freshSubscriptionId = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($s in @($freshSnapshot['SubscriptionIds'])) { [void]$freshSubscriptionId.Add(([string]$s).ToLowerInvariant()) }

    $priorStatus = Get-AzMonExcelReviewStatus -Path $ReportPath -SheetName '4.ImpactedResourcesAnalysis'
    if ($Detail) {
        $trace = [System.Collections.Generic.List[hashtable]]::new()
    } else {
        $trace = $null
    }
    $updates = @(Find-AzMonVerificationStatusUpdate -OriginalFinding $originalFindings -OpenSignature $openSignature `
            -FreshLookup $freshLookup -FreshSubscriptionId $freshSubscriptionId -PriorStatus $priorStatus -UseCheckId $useCheckId -Trace $trace)
    Write-Host "[azmon-assess] $($updates.Count) row(s) will change status." -ForegroundColor DarkCyan

    $applied = @(Update-AzMonExcelReviewStatus -Path $ReportPath -SheetName '4.ImpactedResourcesAnalysis' -StatusUpdate $updates)

    $outDir = New-AzMonOutputDirectory -Path $OutputPath

    if ($Detail) {
        $detailPath = Join-Path $outDir 'verification-detail.json'
        ($trace.ToArray() | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $detailPath -Encoding utf8NoBOM
        Write-Host "[azmon-assess] Full per-row detail (all $($trace.Count) evaluated rows) written to $detailPath" -ForegroundColor Green
    }

    $updatesByKey = @{}
    foreach ($u in $updates) { $updatesByKey["$($u.FindingId)|$(([string]$u.ResourceId).ToLowerInvariant())"] = $u }

    $logPath = Join-Path $outDir 'verification-log.json'
    $existingLog = @()
    if (Test-Path -LiteralPath $logPath) {
        try { $existingLog = @(Get-Content -LiteralPath $logPath -Raw | ConvertFrom-Json -AsHashtable -Depth 10 | Where-Object { $null -ne $_ }) } catch { $existingLog = @() }
    }
    $verifiedAt = (Get-Date).ToUniversalTime().ToString('o')
    $newRecords = @($applied | ForEach-Object {
            $a = $_
            $u = $updatesByKey["$($a.FindingId)|$(([string]$a.ResourceId).ToLowerInvariant())"]
            New-AzMonVerificationRecord -FindingId $a.FindingId -ResourceId $a.ResourceId -Category $u.Category -Title $u.Title `
                -PreviousStatus $a.PreviousStatus -NewStatus $a.NewStatus -Reason $u.Reason -VerifiedAt $verifiedAt -ReportPath $ReportPath
        })
    $combinedLog = @($existingLog) + $newRecords
    ($combinedLog | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $logPath -Encoding utf8NoBOM
    Write-Host "[azmon-assess] Verification log written to $logPath" -ForegroundColor Green

    $implementedCount = @($applied | Where-Object { $_.NewStatus -eq $script:AzMonReviewStatus.Implemented }).Count
    $manualCount = @($applied | Where-Object { $_.NewStatus -eq $script:AzMonReviewStatus.ManualReview }).Count
    $regressedCount = @($applied | Where-Object { $_.NewStatus -eq $script:AzMonReviewStatus.Regressed }).Count
    Write-Host "[azmon-assess] Summary: $implementedCount implemented, $manualCount need manual review, $regressedCount regressed. $ReportPath updated in place." -ForegroundColor Cyan

    return $applied
}
