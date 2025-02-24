# Make private key nonexportable - locks down access to your machine and login
# Secure tenant restricts credential lifetime policy with max 30 days, so we create cert validity of 30 days
New-SelfSignedCertificate -CertStoreLocation "cert:\CurrentUser\My" -Subject "My API testing cert" -NotAfter (Get-Date).AddDays(30) -KeySpec KeyExchange -KeyExportPolicy NonExportable -KeyUsage KeyEncipherment -KeyProtection None

$certs = Get-ChildItem -Path cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=My API testing cert" }

# This is the public key portion which we'll upload to the app registration
Export-Certificate -Cert $certs.PSPath -FilePath ~\desktop\SelfSignedApiTestingCert.cer

# Here's the thumbprint which you'll need to get your auth token to call your API
$certs.Thumbprint