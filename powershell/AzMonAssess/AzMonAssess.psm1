#requires -Version 7.0
# Root module — dot-sources every Private/Public script into module scope,
# then exports only the intended public surface.

$here = $PSScriptRoot
$privateFiles = @(Get-ChildItem -Path (Join-Path $here 'Private') -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
$publicFiles = @(Get-ChildItem -Path (Join-Path $here 'Public') -Filter '*.ps1' -File -ErrorAction SilentlyContinue)

foreach ($file in ($privateFiles + $publicFiles)) {
    try {
        . $file.FullName
    } catch {
        throw "AzMonAssess: failed to load $($file.FullName): $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function @(
    'Connect-AzMonSession'
    'Initialize-AzMonPrerequisite'
    'Resolve-AzMonSubscription'
    'Get-AzMonWorkspace'
    'Get-AzMonAppInsight'
    'Get-AzMonAlertRule'
    'Get-AzMonActionGroup'
    'Get-AzMonResource'
    'Get-AzMonDiagnosticSetting'
    'Get-AzMonDataCollectionRule'
    'Find-AzMonConsolidationFinding'
    'Find-AzMonCoverageGapFinding'
    'Find-AzMonAlertQualityFinding'
    'Find-AzMonCostOptimizationFinding'
    'Find-AzMonTracingFinding'
    'Find-AzMonReliabilityFinding'
    'Find-AzMonSecurityFinding'
    'Find-AzMonPerformanceFinding'
    'New-AzMonAiSummary'
    'New-AzMonSnapshotObject'
    'Save-AzMonSnapshot'
    'Import-AzMonSnapshot'
    'New-AzMonDemoSnapshot'
    'Invoke-AzMonAssessment'
    'Invoke-AzMonReportGeneration'
    'New-AzMonMarkdownReport'
    'New-AzMonExcelReport'
    'New-AzMonPptxReport'
    'New-AzMonHtmlReport'
    'Invoke-AzMonTriage'
    'Invoke-AzMonRemediation'
)
