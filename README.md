# PowerShell solutions

This repository is a collection of automation and scripting bits to test Azure and Microsoft 365 technologies.

Wherever possible, the scripts require least privilege permissions and will never include secrets. APIs will connect using local self-signed certificates with limited lifetimes.

## REST API use with self-signed certificates

Self-signed certificates are useful for secure API testing when the following conditions are met:
- limited validity period
- private key is not exportable
- the private key is located in the profile of a compliant device from which the testing is performed

### Create the Entra App registration

Activate one of these roles:
- Cloud Application Administrator
- Application Developer
- Application Administrator

1. Create an [App registration in Entra admin portal](https://entra.microsoft.com/?feature.msaljs=true#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade/quickStartType~/null/isMSAApp~/false) and give it a good name, like "API testing".

1. Keep the default options > **Register**

### Create self-signed certificate

Use this script to create a secure 30 day self-signed certificate:
[New-SelfSignedApiCert.ps1](New-SelfSignedApiCert.ps1)

### Upload public cert to App registration

A secure self-signed certificate restricts access to the private key (comparable to the secret) with your logon identity (windows user profile) and prevents the certificate from being moved and prevents the private key from being exported.

The public portion of the certificate is going to be the `.cer` file. If you used the sample script, this will be the **SelfSignedApiTestingCert** *yyyy-MM-dd HH:mm* **.cer** exported to your desktop.

1. Navigate to your App registration > Certificates & secrets.
1. Select the `.cer` public certificate and add a description

### Set permissions

Depending on the API, there are 2 different permissions that may be required.

#### Permission type 1 - API Permissions

For example, in order for the app to access certain REST APIs, you need to add API permissions. Here are 3 examples:

- Example 1: [Data Collection Rules - REST API (Azure Monitor)](https://learn.microsoft.com/rest/api/monitor/data-collection-rules?view=rest-monitor-2023-03-11)

  This API requires the Log Analytics API. In this case, the options are straightforward as there's only 1 option. Choose Application permissions for unattended scripts. Then give admin consent.

- Example 2: Microsoft Graph 

  This REST API has extensive and very granular API permissions. It is very difficult to know what you need if you haven't been here before. Graph Explorer is a great place to start to understand the API permissions you need and test.
  [Graph Explorer | Try Microsoft Graph APIs](https://developer.microsoft.com/graph/graph-explorer)

  >Graph explorer shows you what API permissions are required and whether admin consent is needed. Once the Graph Explorer test is successful, mirror the API permission configuration for your App registration.
  >Graph Explorer demonstrates the API permissions in the delegated model with user permissions. For the *client_credentials* grant type where we're using the self-signed certificate, your app needs the application permissions.

- Example 3: Threat Intelligence STIX object [upload API](https://learn.microsoft.com/azure/sentinel/stix-objects-api)

  This API and many other REST APIs like the Microsoft Sentinel REST APIs don't require API permissions, but do require Azure RBAC permissions.

#### Permission Type 2 – Azure RBAC

Some APIs require giving your application Azure RBAC permissions at a certain scope. For example, the Microsoft Sentinel upload API requires the app registration to be granted the Microsoft Sentinel contributor role at the workspace level.

## Securely test APIs with a PowerShell script

Once your App registration, self-signed certificate and permissions are configured, you're ready to securely test APIs! 
[Test-API.ps1](Test-API.ps1)