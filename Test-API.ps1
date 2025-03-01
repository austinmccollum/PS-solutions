
# requires Powershell module MSAL.PS https://www.powershellgallery.com/packages/MSAL.PS/4.37.0.0
<#
.SYNOPSIS

This command uses a certificate thumbprint to access the private key of a self-signed certificate and use it to get OAuth 2.0 token for client_credentials grant type (unattended automation).
Then, based on the options you provide, perform a REST API call to the proper endpoint.
    .DESCRIPTION
        Perform a REST API call to an endpoint you provide and output the results in JSON format.
    .PARAMETER WorkspaceName
        Enter the Log Analytics workspace name, this is a required parameter
    .PARAMETER WorkspaceId
        Enter the Log Analytics workspace ID, this is a required parameter
    .PARAMETER ResourceGroupName
        Enter the Resource Group Name where the LA workspace resides, this is a required parameter
    .PARAMETER AppId
        Enter the Application ID of the App Registration you created, this is a required parameter
    .PARAMETER TenantName
        Enter the Tenant Name where the App Registration resides, this is a required parameter
    .PARAMETER CertThumbprint
        Enter the thumbprint of the certificate uploaded to your app registration, this is a required parameter
    .PARAMETER SubscriptionId
        Enter the Subscription ID where the resources reside, this is a required parameter
    .NOTES
        AUTHOR: Austin McCollum
        ALIAS: AustinMc 
        LASTEDIT: 2.24.2025
    .EXAMPLE
        Test-Api -API UploadApi -WorkspaceName "workspacename" -ResourceGroupName "rgname" -AppId "00001111-aaaa-2222-bbbb-3333cccc4444" -TenantName "contoso.onmicrosoft.com" -Path "C:\Users\user\Documents\stixobjects.json"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("UploadApi", "ListDcrApi", "GetDcrApi", "ListAlertRulesApi", "GetAlertRuleApi", "GetAlertRuleTemplateApi")]
    [string]$Api = "uploadApi",

    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$CertThumbprint,

    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $false)]
    [string]$WorkspaceName,

    [Parameter(Mandatory = $false)]
    [string]$WorkspaceId,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$ItemId,

    [Parameter(Mandatory = $false)]
    [string]$ApiVersionOverride,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$Path
)

# Testing APIs
#  Thanks to Nicola Suter for a great example!
#  https://tech.nicolonsky.ch/explaining-microsoft-graph-access-token-acquisition/

# This script focuses on REST APIs that use the management.azure.com scope
$Scope = "https://management.azure.com/.default"

# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

# Connection details for getting initial token with self-signed certificate from local store
$connectionDetails = @{
    'TenantId'          = $TenantName
    'ClientId'          = $AppId
    'ClientCertificate' = Get-Item -Path "Cert:\CurrentUser\My\$CertThumbprint"
    scope               = $Scope
}

# Request the token
#  Using Powershell module MSAL.PS https://www.powershellgallery.com/packages/MSAL.PS/4.37.0.0
#  Get-MsalToken is automatically using OAuth 2.0 token endpoint https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token
#  and sets auth flow to grant_type = 'client_credentials'
$token = Get-MsalToken @connectionDetails

# Create header
#  Again relying on MSAL.PS which has method CreateAuthorizationHeader() getting us the bearer token
$Header = @{
    'Authorization' = $token.CreateAuthorizationHeader()
}

# uploadAPI
# STIX object API endpoint
# Samples in the article I wrote here
#  https://learn.microsoft.com/en-us/azure/sentinel/stix-objects-api?branch=main#sample-indicator-request-body
if ($Api -eq "uploadApi") {
    if (-not $ApiVersionOverride) { $apiVersion = "2023-12-01-preview" } 
    else { $apiVersion = $ApiVersionOverride }
    $Uri = "https://api.ti.sentinel.azure.com/workspaces/$workspaceId/threat-intelligence-stix-objects:upload?api-version=$apiVersion"
    $stixobjects = get-content -path $Path
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Body $stixobjects -Method POST -ContentType "application/json"
}

# ListDcrApi
# REST API call to list DCRs by resource group
# Requires app registration API permissions - Log Analytics API
# https://learn.microsoft.com/en-us/rest/api/monitor/data-collection-rules/list-by-resource-group?view=rest-monitor-2023-03-11&tabs=HTTP
if ($Api -eq "ListDcrApi") {
    if (-not $apiversionoverride) { $apiVersion = "2019-11-01-preview" }
    else { $apiVersion = $ApiVersionOverride }
    $Uri = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Insights/dataCollectionRules?api-version=$apiVersion"
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method GET
}

# GetDcrApi
# REST API call to get a specific DCR
# Requires app registration API permissions - Log Analytics API
# https://learn.microsoft.com/en-us/rest/api/monitor/data-collection-rules/get?view=rest-monitor-2023-03-11&tabs=HTTP
if ($Api -eq "GetDcrApi") {
    if (-not $apiversionoverride) { $apiVersion = "2023-03-11" }
    else { $apiVersion = $ApiVersionOverride }
    $DcrName = $ItemId
    $Uri = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Insights/dataCollectionRules/$($DcrName)?api-version=$apiVersion"
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method GET
}

# ListAlertRulesApi
# REST API call to list Alert rules
# https://learn.microsoft.com/en-us/rest/api/securityinsights/alert-rules/list?view=rest-securityinsights-2024-09-01&tabs=HTTP
if ($Api -eq "ListAlertRulesApi") {
    if (-not $apiversionoverride) { $apiVersion = "2024-09-01" }
    else { $apiVersion = $ApiVersionOverride }
    $Uri = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$workspaceName/providers/Microsoft.SecurityInsights/alertRules?api-version=$apiVersion"
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method GET
}

# GetAlertRuleApi
# REST API call to get an Alert rule
# https://learn.microsoft.com/en-us/rest/api/securityinsights/alert-rules/get?view=rest-securityinsights-2025-01-01-preview&tabs=HTTP
if ($api -eq "GetAlertRuleApi") {
    if (-not $apiversionoverride) { $apiVersion = "2024-09-01" }
    else { $apiVersion = $ApiVersionOverride }
    $alertRuleId = $ItemId
    $Uri = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$workspaceName/providers/Microsoft.SecurityInsights/alertRules/$($alertRuleId)?api-version=$apiVersion"
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method GET
}

# GetAlertRuleTemplateApi
# REST API call to get Alert Rule Template
# https://learn.microsoft.com/en-us/rest/api/securityinsights/alert-rule-templates/get?view=rest-securityinsights-2025-01-01-preview&tabs=HTTP
if ($Api -eq "GetAlertRuleTemplateApi") {
    if (-not $apiversionoverride) { $apiVersion = "2025-01-01-preview" }
    else { $apiVersion = $ApiVersionOverride }
    $alertRuleTemplateId = $ItemId
    $Uri = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$workspaceName/providers/Microsoft.SecurityInsights/alertRuleTemplates/$($alertRuleTemplateId)?api-version=$apiVersion"
    $results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method GET
}

$results | ConvertTo-Json -Depth 8
$results