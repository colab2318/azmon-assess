@{
    RootModule           = 'AzMonAssess.psm1'
    ModuleVersion        = '0.1.0'
    GUID                 = 'b6f3a3d2-8c1e-4a7a-9e6b-2f6d6a2d7a10'
    Author               = 'Aminul Chowdhury'
    Description          = 'AI-powered Azure Monitoring & Observability assessment - PowerShell edition. Runs entirely in Azure Cloud Shell or local PowerShell 7+ with no admin rights: reads Log Analytics workspaces, Application Insights, alert rules and diagnostic settings via Azure Resource Graph, scores them against Well-Architected Framework guidance, and produces Excel + PowerPoint deliverables for a customer-facing readout.'
    PowerShellVersion    = '7.0'
    CompatiblePSEditions = @('Core')
    RequiredModules      = @()
    FunctionsToExport    = @(
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
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('Azure', 'Monitoring', 'Observability', 'Assessment', 'CloudShell', 'WAF')
            ProjectUri = ''
        }
    }
}
