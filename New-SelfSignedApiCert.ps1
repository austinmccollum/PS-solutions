# Make private key nonexportable - locks down access to your machine and login
# A secure tenant Azure policy may restrict credential lifetime policy. In this example, the secure environment sets validity to 1 month, so we create cert validity of 1 month

# using date/time to make cert more unique and stand out in the list of certs
#  public service announcement - don't include ":" in the filename as that will cause the export-certificate command to fail.
$today = (Get-Date).ToString("yyyy-MM-dd HHmm.ss")

# even though lifetime policy is max 1 month, we can't use AddDays(30) because in February, it creates a validity period over 1 month
# so we use AddMonths(1) to ensure we're always under the 1 month limit
$newCert = New-SelfSignedCertificate -CertStoreLocation "cert:\CurrentUser\My" -Subject "My API testing cert $today" -NotAfter (Get-Date).AddMonths(1) -KeySpec KeyExchange -KeyExportPolicy NonExportable -KeyUsage KeyEncipherment -KeyProtection None

# This is the public key portion which we'll upload to the app registration
Export-Certificate -Cert $newCert -FilePath "~\desktop\SelfSignedApiTestingCert$($today).cer" -Force

# Here's the thumbprint which you'll need to get your auth token to call your API
$newCert.Thumbprint