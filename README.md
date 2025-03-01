# Powershell solutions

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

The public portion of the certificate is going to be the `.cer` file. If you used the sample script, this will be the **SelfSignedApiTestingCert***yyyy-MM-dd HH:mm***.cer** exported to your desktop.

1. Navigate to your App registration > Certificates & secrets.
1. Select the `.cer` public certificate and add a description

:::image type="content" source="resources/public-certificate-upload.png" alt-text="Screenshot showing upload of .cer file in certificate store of the app registration." lightbox="resources/public-certificate-upload.png":::
