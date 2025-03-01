# Make private key nonexportable - locks down access to your machine and login
# Secure tenant restricts credential lifetime policy to 1 month, so we create cert validity of 1 month

# using date to make cert more unique and stand out in the list of certs
$today = (Get-Date).ToString("yyyy-MM-dd HH:mm")

# even though lifetime policy is max 1 month, we can't use AddDays(30) because in February, it creates a validity period over 1 month
# so we use AddMonths(1) to ensure we're always under the 1 month limit
New-SelfSignedCertificate -CertStoreLocation "cert:\CurrentUser\My" -Subject "My API testing cert $today" -NotAfter (Get-Date).AddMonths(1) -KeySpec KeyExchange -KeyExportPolicy NonExportable -KeyUsage KeyEncipherment -KeyProtection None

$certs = Get-ChildItem -Path cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=My API testing cert $today" }

# This is the public key portion which we'll upload to the app registration
Export-Certificate -Cert $certs.PSPath -FilePath "~\desktop\SelfSignedApiTestingCert$($today).cer"

# Here's the thumbprint which you'll need to get your auth token to call your API
$certs.Thumbprint