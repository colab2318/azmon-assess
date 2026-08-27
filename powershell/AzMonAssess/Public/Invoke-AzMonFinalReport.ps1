#requires -Version 7.0
# Builds a "final" report containing only findings/resources a human has
# manually confirmed remain open, after reviewing the auto-verified
# report.xlsx. Complements Invoke-AzMonVerification (which only ever
# updates status text in place) - this reads the reviewer's own
# "Reviewed - Not Implemented" annotations back out and uses them to
# filter a fresh set of report deliverables, rather than handing over the
# full, unfiltered assessment again.

function Invoke-AzMonFinalReport {
    <#
    .SYNOPSIS
        Regenerates report deliverables containing ONLY the findings/
        resources a reviewer has manually marked "Reviewed - Not
        Implemented" in a copy of report.xlsx (e.g. report_updated.xlsx).
    .DESCRIPTION
        Loads the original assessment snapshot (for full finding detail -
        title, recommendation, evidence, etc.) and cross-references the
        reviewed workbook's "REQUIRED ACTIONS / REVIEW STATUS" column for
        every (finding, resource) row. Any resource NOT marked "Reviewed -
        Not Implemented" is dropped; findings left with zero resources are
        dropped entirely. The resulting filtered snapshot is rendered to
        report-final.<ext> in -OutputPath - the original snapshot.json and
        report.xlsx are never modified.
    .PARAMETER SnapshotPath
        The original assessment snapshot.json (source of full finding
        detail). Not the reviewed workbook.
    .PARAMETER ReviewedReportPath
        Path to the manually-reviewed copy of report.xlsx (e.g.
        report_updated.xlsx) whose REQUIRED ACTIONS / REVIEW STATUS column
        a reviewer has hand-annotated.
    .PARAMETER OutputPath
        Directory to write report-final.<ext> into.
    .PARAMETER Report
        Which format(s) to generate: html, pptx, excel, markdown, all.
        Defaults to html + pptx.
    .EXAMPLE
        Invoke-AzMonFinalReport -SnapshotPath ./out/snapshot.json -ReviewedReportPath ./out/report_updated.xlsx -OutputPath ./out
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SnapshotPath,
        [Parameter(Mandatory)] [string] $ReviewedReportPath,
        [Parameter(Mandatory)] [string] $OutputPath,
        [ValidateSet('all', 'markdown', 'excel', 'pptx', 'html', 'none')]
        [string[]] $Report = @('html', 'pptx')
    )

    if (-not (Test-Path -LiteralPath $SnapshotPath)) { throw "-SnapshotPath not found: $SnapshotPath" }
    if (-not (Test-Path -LiteralPath $ReviewedReportPath)) { throw "-ReviewedReportPath not found: $ReviewedReportPath" }

    $snapshot = Import-AzMonSnapshot -Path $SnapshotPath
    $allFindings = @($snapshot['Findings'])
    Write-Host "[azmon-assess] Filtering $($allFindings.Count) finding(s) from $SnapshotPath against reviewer annotations in $ReviewedReportPath..." -ForegroundColor Cyan

    $statusByKey = Get-AzMonExcelReviewStatus -Path $ReviewedReportPath -SheetName '4.ImpactedResourcesAnalysis'
    $targetStatus = $script:AzMonReviewStatus.NotImplemented
    $confirmedKeys = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($kv in $statusByKey.GetEnumerator()) {
        if ($kv.Value -eq $targetStatus) { [void]$confirmedKeys.Add($kv.Key) }
    }
    Write-Host "[azmon-assess] $($confirmedKeys.Count) row(s) marked '$targetStatus'." -ForegroundColor Cyan
    if ($confirmedKeys.Count -eq 0) {
        Write-Warning "No rows in $ReviewedReportPath are marked '$targetStatus' - the final report will have zero findings. Mark the rows you've confirmed are still open, then re-run."
    }

    $finalFindings = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $allFindings) {
        $resourceIds = @($f['ResourceIds'] | Where-Object { $_ })
        if ($resourceIds.Count -eq 0) {
            # aggregate findings with no resource get one blank-id row in the sheet
            if ($confirmedKeys.Contains("$($f['Id'])|")) { $finalFindings.Add($f) }
            continue
        }
        $keptResourceIds = @($resourceIds | Where-Object { $confirmedKeys.Contains("$($f['Id'])|$(([string]$_).ToLowerInvariant())") })
        if ($keptResourceIds.Count -eq 0) { continue }
        $trimmed = $f.Clone()
        $trimmed['ResourceIds'] = $keptResourceIds
        $finalFindings.Add($trimmed)
    }
    Write-Host "[azmon-assess] $($finalFindings.Count) of $($allFindings.Count) finding(s) retained in the final report." -ForegroundColor Cyan

    $finalSnapshot = $snapshot.Clone()
    $finalSnapshot['Findings'] = $finalFindings.ToArray()
    # the original narrative summary describes ALL findings - carrying it over
    # verbatim would contradict the filtered finding list shown below it.
    $finalSnapshot['AiSummary'] = "This is a filtered final report: $($finalFindings.Count) of $($allFindings.Count) finding(s) a reviewer confirmed remain open (marked '$targetStatus' in $ReviewedReportPath). See the original assessment report for the full narrative executive summary."

    $outDir = New-AzMonOutputDirectory -Path $OutputPath
    $reportAll = $Report -contains 'all'

    if ($reportAll -or $Report -contains 'markdown') {
        try {
            New-AzMonMarkdownReport -Snapshot $finalSnapshot -Path (Join-Path $outDir 'report-final.md') | Out-Null
            Write-Host '[azmon-assess] report-final.md written.' -ForegroundColor Green
        } catch {
            Write-Warning "Final markdown report skipped: $($_.Exception.Message)"
        }
    }
    if ($reportAll -or $Report -contains 'excel') {
        try {
            New-AzMonExcelReport -Snapshot $finalSnapshot -Path (Join-Path $outDir 'report-final.xlsx') | Out-Null
            Write-Host '[azmon-assess] report-final.xlsx written.' -ForegroundColor Green
        } catch {
            Write-Warning "Final Excel report skipped: $($_.Exception.Message)"
        }
    }
    if ($reportAll -or $Report -contains 'pptx') {
        try {
            New-AzMonPptxReport -Snapshot $finalSnapshot -Path (Join-Path $outDir 'report-final.pptx') | Out-Null
            Write-Host '[azmon-assess] report-final.pptx written.' -ForegroundColor Green
        } catch {
            Write-Warning "Final PPTX report skipped: $($_.Exception.Message)"
        }
    }
    if ($reportAll -or $Report -contains 'html') {
        try {
            New-AzMonHtmlReport -Snapshot $finalSnapshot -Path (Join-Path $outDir 'report-final.html') | Out-Null
            Write-Host '[azmon-assess] report-final.html written.' -ForegroundColor Green
        } catch {
            Write-Warning "Final HTML report skipped: $($_.Exception.Message)"
        }
    }

    return $finalSnapshot
}
