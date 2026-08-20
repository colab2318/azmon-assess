#requires -Version 7.0
# Model "constructors" — every record in this tool is a plain (unordered)
# hashtable, matching exactly what ConvertFrom-Json -AsHashtable produces on
# reload, so a live collection run and a reloaded snapshot.json behave
# identically everywhere else in the module. (Deliberately NOT [ordered] —
# OrderedDictionary is not a [hashtable] and would risk conversion issues
# anywhere code is typed [hashtable]/List[hashtable].)

function New-AzMonWorkspace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [Parameter(Mandatory)] [string] $Location,
        [string] $Sku,
        [Nullable[int]] $RetentionDays,
        [Nullable[double]] $DailyQuotaGb,
        [string] $CustomerId,
        [hashtable] $Tags = @{},
        [Nullable[double]] $IngestionGb30d,
        [hashtable] $IngestionByTable = @{},
        [int] $ConnectedSources = 0,
        [string] $PublicNetworkAccessForIngestion,
        [string] $PublicNetworkAccessForQuery,
        [Nullable[bool]] $DisableLocalAuth,
        [Nullable[int]] $CapacityReservationLevel,
        [string] $ClusterResourceId
    )
    @{
        Kind                                = 'Workspace'
        Id                                  = $Id
        Name                                = $Name
        SubscriptionId                      = $SubscriptionId
        ResourceGroup                       = $ResourceGroup
        Location                            = $Location
        Sku                                 = $Sku
        RetentionDays                       = $RetentionDays
        DailyQuotaGb                        = $DailyQuotaGb
        CustomerId                          = $CustomerId
        Tags                                = $Tags
        IngestionGb30d                      = $IngestionGb30d
        IngestionByTable                    = $IngestionByTable
        ConnectedSources                    = $ConnectedSources
        PublicNetworkAccessForIngestion     = $PublicNetworkAccessForIngestion
        PublicNetworkAccessForQuery         = $PublicNetworkAccessForQuery
        DisableLocalAuth                    = $DisableLocalAuth
        CapacityReservationLevel            = $CapacityReservationLevel
        ClusterResourceId                   = $ClusterResourceId
    }
}

function New-AzMonAppInsight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [Parameter(Mandatory)] [string] $Location,
        [string] $Kind,
        [string] $ApplicationType,
        [string] $WorkspaceResourceId,
        [Nullable[double]] $SamplingPercentage,
        [Nullable[int]] $RetentionDays,
        [hashtable] $Tags = @{},
        [string] $PublicNetworkAccessForIngestion,
        [string] $PublicNetworkAccessForQuery,
        [Nullable[bool]] $DisableLocalAuth,
        [Nullable[double]] $DailyCapGb
    )
    @{
        Kind                                = 'AppInsights'
        Id                                  = $Id
        Name                                = $Name
        SubscriptionId                      = $SubscriptionId
        ResourceGroup                       = $ResourceGroup
        Location                            = $Location
        AiKind                              = $Kind
        ApplicationType                     = $ApplicationType
        WorkspaceResourceId                 = $WorkspaceResourceId
        SamplingPercentage                  = $SamplingPercentage
        RetentionDays                       = $RetentionDays
        Tags                                = $Tags
        PublicNetworkAccessForIngestion     = $PublicNetworkAccessForIngestion
        PublicNetworkAccessForQuery         = $PublicNetworkAccessForQuery
        DisableLocalAuth                    = $DisableLocalAuth
        DailyCapGb                          = $DailyCapGb
    }
}

function New-AzMonAlertRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [Parameter(Mandatory)] [ValidateSet('metric', 'log', 'activityLog')] [string] $AlertKind,
        [Parameter(Mandatory)] [bool] $Enabled,
        [Nullable[int]] $Severity,
        [string[]] $Scopes = @(),
        [string[]] $ActionGroupIds = @(),
        [string] $Description,
        [Nullable[int]] $FireCount30d,
        [Nullable[double]] $EvaluationFrequencyMinutes
    )
    @{
        Kind                       = 'AlertRule'
        Id                         = $Id
        Name                       = $Name
        SubscriptionId             = $SubscriptionId
        ResourceGroup              = $ResourceGroup
        AlertKind                  = $AlertKind
        Enabled                    = $Enabled
        Severity                   = $Severity
        Scopes                     = @($Scopes)
        ActionGroupIds             = @($ActionGroupIds)
        Description                = $Description
        FireCount30d               = $FireCount30d
        EvaluationFrequencyMinutes = $EvaluationFrequencyMinutes
    }
}

function New-AzMonActionGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [string] $ShortName,
        [int] $EmailReceivers = 0,
        [int] $SmsReceivers = 0,
        [int] $WebhookReceivers = 0,
        [int] $LogicAppReceivers = 0,
        [int] $ItsmReceivers = 0,
        [int] $UsedByRules = 0
    )
    @{
        Kind              = 'ActionGroup'
        Id                = $Id
        Name              = $Name
        SubscriptionId    = $SubscriptionId
        ResourceGroup     = $ResourceGroup
        ShortName         = $ShortName
        EmailReceivers    = $EmailReceivers
        SmsReceivers      = $SmsReceivers
        WebhookReceivers  = $WebhookReceivers
        LogicAppReceivers = $LogicAppReceivers
        ItsmReceivers     = $ItsmReceivers
        UsedByRules       = $UsedByRules
    }
}

function New-AzMonDiagnosticSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $ResourceId,
        [Parameter(Mandatory)] [string] $Name,
        [string] $WorkspaceId,
        [string] $StorageId,
        [string] $EventHubId,
        [bool] $LogsEnabled = $false,
        [bool] $MetricsEnabled = $false
    )
    @{
        Kind           = 'DiagnosticSetting'
        ResourceId     = $ResourceId
        Name           = $Name
        WorkspaceId    = $WorkspaceId
        StorageId      = $StorageId
        EventHubId     = $EventHubId
        LogsEnabled    = $LogsEnabled
        MetricsEnabled = $MetricsEnabled
    }
}

function New-AzMonResourceRef {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $Type,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [Parameter(Mandatory)] [string] $Location,
        [hashtable] $Tags = @{}
    )
    @{
        Kind           = 'ResourceRef'
        Id             = $Id
        Name           = $Name
        Type           = $Type
        SubscriptionId = $SubscriptionId
        ResourceGroup  = $ResourceGroup
        Location       = $Location
        Tags           = $Tags
    }
}

function New-AzMonDataCollectionRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Id,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $SubscriptionId,
        [Parameter(Mandatory)] [string] $ResourceGroup,
        [Parameter(Mandatory)] [string] $Location,
        [string] $DcrKind,
        [hashtable] $Tags = @{},
        [int] $DataFlowCount = 0,
        [string] $WorkspaceResourceId
    )
    @{
        Kind                = 'DataCollectionRule'
        Id                  = $Id
        Name                = $Name
        SubscriptionId      = $SubscriptionId
        ResourceGroup       = $ResourceGroup
        Location            = $Location
        DcrKind             = $DcrKind
        Tags                = $Tags
        DataFlowCount       = $DataFlowCount
        WorkspaceResourceId = $WorkspaceResourceId
    }
}

function New-AzMonFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Category,
        [Parameter(Mandatory)] [ValidateSet('critical', 'high', 'medium', 'low', 'info')] [string] $Severity,
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] [string] $Detail,
        [string[]] $ResourceIds = @(),
        [Nullable[double]] $EstimatedMonthlySavingsUsd,
        [string] $Recommendation,
        [hashtable] $Evidence = @{}
    )
    @{
        Kind                       = 'Finding'
        Id                         = [guid]::NewGuid().ToString()
        Category                   = $Category
        Severity                   = $Severity
        Title                      = $Title
        Detail                     = $Detail
        ResourceIds                = @($ResourceIds)
        EstimatedMonthlySavingsUsd = $EstimatedMonthlySavingsUsd
        Recommendation             = $Recommendation
        Evidence                   = $Evidence
    }
}

function New-AzMonSnapshotObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $SubscriptionId,
        [Parameter(Mandatory)] [string] $CustomerName,
        [datetime] $GeneratedAt = (Get-Date).ToUniversalTime(),
        [array] $Workspaces = @(),
        [array] $AppInsights = @(),
        [array] $AlertRules = @(),
        [array] $ActionGroups = @(),
        [array] $DiagnosticSettings = @(),
        [array] $Resources = @(),
        [array] $DataCollectionRules = @(),
        [array] $Findings = @(),
        [string] $AiSummary
    )
    @{
        Kind                = 'Snapshot'
        GeneratedAt         = $GeneratedAt.ToString('o')
        SubscriptionIds     = @($SubscriptionId)
        CustomerName        = $CustomerName
        Workspaces          = @($Workspaces)
        AppInsights         = @($AppInsights)
        AlertRules          = @($AlertRules)
        ActionGroups        = @($ActionGroups)
        DiagnosticSettings  = @($DiagnosticSettings)
        Resources           = @($Resources)
        DataCollectionRules = @($DataCollectionRules)
        Findings            = @($Findings)
        AiSummary           = $AiSummary
    }
}
