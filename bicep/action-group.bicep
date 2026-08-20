// Standard action group for platform alerting.
// Deploy per-environment: prod-pager (PagerDuty), non-prod-email, security-siem.

@description('Short name of the action group (max 12 chars).')
param shortName string

@description('Full display name.')
param name string

@description('Location for the action group (metadata only).')
param location string = 'global'

@description('Email addresses to notify.')
param emails array = []

@description('Optional PagerDuty webhook URL.')
param pagerDutyWebhookUrl string = ''

@description('Optional Microsoft Teams webhook URL.')
param teamsWebhookUrl string = ''

resource ag 'Microsoft.Insights/actionGroups@2023-01-01' = {
  name: name
  location: location
  properties: {
    groupShortName: shortName
    enabled: true
    emailReceivers: [for (e, i) in emails: {
      name: 'email-${i}'
      emailAddress: e
      useCommonAlertSchema: true
    }]
    webhookReceivers: concat(
      empty(pagerDutyWebhookUrl) ? [] : [{
        name: 'pagerduty'
        serviceUri: pagerDutyWebhookUrl
        useCommonAlertSchema: true
      }],
      empty(teamsWebhookUrl) ? [] : [{
        name: 'teams'
        serviceUri: teamsWebhookUrl
        useCommonAlertSchema: true
      }]
    )
  }
}

output actionGroupId string = ag.id
