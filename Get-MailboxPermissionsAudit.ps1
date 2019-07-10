<#==============================================================================
                    Exchange Mailbox Permission Dump Script
================================================================================
Programmed By:  Joshua Loos (jloos@microsoft.com)
Programmed Date:  09/20/2016
Last Modified:    07/10/2019 [austinmc@microsoft.com]
 ------------------------------------------------------------------------
DISCLAIMER: Use this powershell script at your own risk and willingness.
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 
 
.SYNOPSIS
  This Powershell script is intended to dump permissions for users in a list or
  queried from AD
   
.DESCRIPTION
   Pulls Send-As, Send on Behalf, and folder level access permissions for a given
   set of users
      
   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
   RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.NOTES 
   Version 3.0 November 29, 2016
   Revsions:
       1.0 Initial script created by customer.
       2.0 Updated functionality to pull additional data and formulate reports.
       3.0 Fixed some minor bugs, added functionality updated to pull additional
           data, added comments.
       4.0 Customer-specific changes, minor bug fixes
       5.0 Cleaned up script with new functions and more comments
       6.0 Improved error handling, purged customer-specific logic
	   7.0 Added checks for full mailbox permissions and quick output of unique relationships
       8.0 Rehauled script - trimmed down to just delegate, full mailbox and sendas with no folder level checks

.PARAMETER OU
   Optional.  Specify a OU from which to query users.  Do not use this parameter
   when using ImportList or AnalyzeBatches
.PARAMETER ImportList
   Optional.  Specify the location of a CSV file that contains a list of primary
   SMTP addresses under the column, "PrimarySMTPAddress".  Do not use this parameter
   when specifying an OU

.EXAMPLE
  '.\Get-MailboxPermissionAudit.ps1' -ImportList "EarlyAdopters.csv"
.EXAMPLE
  '.\Get-MailboxPermissionAudit.ps1' -OU "OU=Contractors,CN=Users,DC=contoso,DC=com"
#>


#region Parameters
Param(
    [string]$OU,
    [string]$ImportList
)
#endregion


#region Script Variables
# These variables are ok to change
$CustomerName = "Contoso by Microsoft"
$Domain = "contoso.local" # Domain only affects the default OU location if no OU is specified
#endregion


#region Script Setup
# Do not change these variables
[array]$Script:UserList = @()
[System.Collections.ArrayList]$Script:Permissions = New-Object System.Collections.ArrayList($null)
#[array]$Script:Permissions = @()
[array]$Script:MigrationBatches = @()

# Default path if no OU is specified
$SplitDomain = $Domain.split(".")
If (-not $OU -and -not $AnalyzeBatches) { $OU = "CN=Users,DC=$($SplitDomain[0]),DC=$($SplitDomain[1])" }

# The following few lines are used later to pull Send As permissions
$DSE = [ADSI]"LDAP://Rootdse"
$EXT = [ADSI]("LDAP://CN=Extended-Rights," + $DSE.ConfigurationNamingContext)
$DN = [ADSI]"LDAP://$($DSE.DefaultNamingContext)"
$DSLookFor = New-Object System.DirectoryServices.DirectorySearcher($DN)
$Right = $EXT.psbase.Children | ? { $_.DisplayName -eq "Send As" }

# Determining the output name of the report
$RunTime = $(Get-Date -Format "yyyyMMddTHHmmss")
$OutputLocation = "Permissions-$Runtime.csv"

# Build error report filename
$ErrorLogPath = "Permissions-ErrorLog-$Runtime.txt"

# Import commands to query AD
If (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}
#endregion


#regions Functions
Function Draw-Banner {
    Write-Host "`n############################################################"
    Write-Host "#                " -NoNewline
    Write-Host "Exchange Permissions Dump" -ForegroundColor Cyan -NoNewline
    Write-Host "                 #"
    Write-Host "#     " -NoNewLine
    Write-Host "Designed for:" -ForegroundColor Yellow -NoNewline
    Write-Host " $CustomerName                   #"
    Write-Host "#     " -NoNewLine
    Write-Host "Author:" -ForegroundColor Yellow -NoNewline
    Write-Host "       Joshua Loos -- jloos@microsoft.com     #"
    Write-Host "############################################################`n"
}

Function Draw-RunSettings {
    If ($OU) { Write-Host " Run Mode:`t`tPermission Dump" }
    Write-Host " Domain:`t`t$Domain"
    If (-not $ImportList) { Write-Host " OU:`t`t`t$OU" } Else { Write-Host " OU:`t`t`tFalse" }
    If ($ImportList) { Write-Host " ImportList:`t`t$ImportList"} Else { Write-Host " ImportList:`t`tFalse" }
}

Function Write-Action ($Message) {
    $TimeStamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$TimeStamp] $Message"
}

Function Write-ErrorLog ($ErrorMessage) {
    $TimeStamp = Get-Date -Format "HH:mm:ss"
    $Output = "[$TimeStamp] [ERROR] $ErrorMessage"
    $Output | Out-File $ErrorLogPath -Encoding ascii -Append
}

Function Get-UserList {
    If ($ImportList) { # Use a list to import users
        $ImportList = Import-Csv $ImportList
        Foreach ($Row in $ImportList) {
            $ImportUser = $($Row.PrimarySMTPAddress)
            try { $Recipients = Get-Recipient $ImportUser.Trim() -ErrorAction Stop | Select SamAccountName,RecipientTypeDetails,PrimarySMTPAddress }
            catch { Write-ErrorLog "Unable to import user $ImportUser, Get-Recipient command failed.  Error Message: $($_.ToString())" }
            Foreach ($Recipient in $Recipients) {
                If ($Recipient.RecipientTypeDetails -match "Mailbox" -and $Recipient.RecipientTypeDetails -notlike "*Remote*") {
                    try {
                        $UserObject = Get-ADUser $($Recipient.SamAccountName) -Properties msExchDelegateListLink,PublicDelegates,PublicDelegatesBL,GivenName,Surname,DistinguishedName,UserPrincipalName,mail,name,msExchRecipientTypeDetails | Select msExchDelegateListLink,PublicDelegates,PublicDelegatesBL,GivenName,Surname,DistinguishedName,UserPrincipalName,mail,name,msExchRecipientTypeDetails
                        If ($UserObject.count -gt 1) {
                            Write-ErrorLog "Unable to import user.  Duplicate users found for $ImportUser (SAM: $($Recipient.SamAccountName))."
                        } Else {
                            $Script:UserList += $UserObject
                        }
                    }
                    catch {Write-ErrorLog "Unable to import user $ImportUser, Get-ADUser command failed.  Error Message: $($_.ToString())"}
                } Else {
                    Write-ErrorLog "Unable to import user $ImportUser.  User is not a User-, Shared-, or ResourceMailbox ($($Recipient.RecipientTypeDetails))."
                }
            }
        }
    } 
    Else { # Used when the OU parameter is specified, imports all users from that particular OU
        try {
            If ([ADSI]::Exists("LDAP://$OU")) {
                $Script:UserList = Get-ADUser -SearchScope Subtree -SearchBase $OU -ResultSetSize $null -Filter {(objectCategory -eq "person") -and (objectClass -eq "user") -and (msExchRecipientTypeDetails -like "*")} -Properties msExchDelegateListLink,PublicDelegates,PublicDelegatesBL,GivenName,Surname,DistinguishedName,UserPrincipalName,mail,name,msExchRecipientTypeDetails | Select msExchDelegateListLink,PublicDelegates,PublicDelegatesBL,GivenName,Surname,DistinguishedName,UserPrincipalName,mail,name,msExchRecipientTypeDetails
            } Else {
                Write-ErrorLog "Unable to get users, the OU that was specified does not exist."
            }
        } catch [System.Management.Automation.RuntimeException] {
            Write-ErrorLog "Unable to get users.  If you specified an OU, verify the path, otherwise, verify that the domain variable is accurate.  Error Message: $($_.ToString())"
        } catch {
            Write-ErrorLog "Unable to get users, Get-ADUser command failed.  Error Message: $($_.ToString())"
        }

    }
}

Function Get-Permissions { # Iterate through each user and grab Send-As permissions, Full Mailbox Access and Delegate Access
    $i=0
    ForEach ($User in $Script:UserList) {
        $i++
        Write-Progress -Activity "Grabbing permissions.." -Status "Processing user $($User.mail) ($i of $($Script:UserList.count))" -PercentComplete $($i/$($Script:UserList.count)*100)

        $HasDelegate = $false

        # User data from Get-ADUser object
        $EmailAddress = $User.mail
        $FirstName = $User.givenname
        $LastName = $User.surname
        $DisplayName= $User.name
        $UPN = $User.UserPrincipalName
        $DistinguishedName= $User.DistinguishedName
        $RecipientType= Switch ($User.msExchRecipientTypeDetails) {
            1 {"User Mailbox"}
            4 {"Shared Mailbox"}
            16 {"Room Mailbox"}
            32 {"Equipment Mailbox"}
            64 {"Mail Contact"}
            128 {"Mail User"}
            256 {"Distribution Group"}
            1024 {"Security Group"}
            2048 {"Dynamic Group"}
            2147483648 {"Remote Mailbox"}
            Default {$user.msExchRecipientTypeDetails}
        }
        
        # Send-as permissions
        $UserDN = [ADSI]("LDAP://$($User.DistinguishedName)")
        $SAPermissions = New-Object -TypeName System.Collections.ArrayList
        # Do not include inherited permissions. Only explicit permissions are migrated https://technet.microsoft.com/en-us/library/jj200581(v=exchg.150).aspx
        $UserDN.psbase.ObjectSecurity.Access | ? { ($_.ObjectType -eq [GUID]$Right.RightsGuid.Value) -and ($_.IsInherited -eq $false) } | Select -ExpandProperty IdentityReference | %{
            If($_ -notlike "NT AUTHORITY\SELF" -and $_ -notlike "*S-1-5-21*") { [void]$SAPermissions.Add($_) }
        }
        If ($SAPermissions) {
            $HasDelegate = $true
            ForEach ($Perm in $SAPermissions) {
                $DelegateName = $Perm.ToString().Replace("NT User:","")
                $SendAsRow = [ordered]@{
                    'First Name' = $FirstName
                    'Last Name' = $LastName
                    'Primary SMTP' = $EmailAddress
                    'Display Name' = $DisplayName
                    UPN = $UPN
                    'DistringuishedName' = $DistinguishedName
                    'Recipient Type' = $RecipientType
                    'Delegate Name' = $DelegateName
                    'Delegate Right' = "SendAs"
                }
                $null = $Script:Permissions.Add((New-Object PSobject -property $SendAsRow))
            }
        }
        # End of send-as permissions

	    # Full Mailbox Permissions added by austinmc
        #  this attribute lists all the other mailboxes your mailbox has FullAccess to, unless AutoMapping was set to $false when assigning the permission
	    If ($User.msExchDelegateListLink) { 
            $HasDelegate = $true
		    ForEach ($DelegateLink in $User.msExchDelegateListLink) {
                $DelegateName = $DelegateLink
                $FullMbxRow = [ordered]@{
                    'First Name' = $FirstName
                    'Last Name' = $LastName
                    'Primary SMTP' = $EmailAddress
                    'Display Name' = $DisplayName
                    UPN = $UPN
                    'DistringuishedName' = $DistinguishedName
                    'Recipient Type' = $RecipientType
                    'Delegate Name' = $DelegateName
                    'Delegate Right' = "Full Mailbox"
                }
                $null = $Script:Permissions.Add((New-Object PSobject -property $FullMbxRow))    
            
            }
        }		

        <# If no data has been found up until this point, this user is reported as having no delegate dependents
        If ($HasDelegate -eq $false) {
            $DelegateName = @{"Delegate" = "None"}
            $DelegateRight = @{"Rights" = "None"}
            $NoneRow = @{
                'First Name' = $FirstName
                'Last Name' = $LastName
                'Primary SMTP' = $EmailAddress
                'Display Name' = $DisplayName
                UPN = $UPN
                'DistringuishedName' = $DistinguishedName
                'Recipient Type' = $RecipientType
                'Delegate Name' = $DelegateName
                'Delegate Right' = $DelegateRight
            }
            $null = $Script:Permissions.Add((New-Object PSobject -property $SendAsRow))
        } #>
    }
}

Function Export-Permissions # Export the permissions to a CSV file.  If AnalyzeBatches was used, the output is a subset of those permissions whose relationship is not satisfied by the batch suggestions.
{

    $Script:Permissions | Export-CSV -path ".\$OutputLocation" -NoTypeInformation

}
#endregion


#region Main
Measure-Command {
    cls
    Draw-Banner
    Draw-RunSettings

    Set-ADServerSettings -ViewEntireForest:$true

    Write-Action "Getting list of users..."
    Get-UserList
    Write-Action "Found $(($Script:UserList | Measure).Count) recipients."

    Write-Action "Getting permissions..."
    Get-Permissions
    $PermCount = $($Script:Permissions | Measure).Count
    Write-Action "Found $PermCount delegation type permissions on those recipients."

    If ($PermCount -gt 0) {
        Write-Action "Exporting permissions to file..."
        Export-Permissions
    }

    Write-Action "Script complete!"
}
#endregion