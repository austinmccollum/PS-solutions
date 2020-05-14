<#
.Notes
    Name: Get-MailboxAudit.ps1
    Author: Austin McCollum [austinmc@microsoft.com]
    Version 1.0: 4/18/2019 finally upgraded sample script to fully featured
    Version 1.1: 10/30/2019 revised to remove AD cached info
    Version 1.2: 2/19/2020 get-mailboxstatistics per user too slow.  calling mailboxstatistics per server instead.
.Description
A comprehensive audit of mailboxes to identify mailboxes that haven't been logged into for a long time.
.Parameter SearchBase
Default is for Fabrikam users, so you'll either need to change the default in the script, or specify a searchbase with the syntax like "OU=Users,DC=fabrikam,DC=com"
.Parameter Days
The number of days a mailbox has had no logon recorded to trigger additional information gathering for the report
.Parameter Resume
By default, we try to resume based on the existence of a temp file.  Setting this parameter to $false allows a fresh start
.Example
Get-MailboxAudit.ps1 -SearchBase "OU=VIPs,OU=Departments,DC=fabrikam,DC=com" -Days 180 -Resume $false -InformationAction 'continue'
#>

[cmdletbinding()]
Param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase="OU=Users,DC=fabrikam,DC=com",

    [Parameter(Mandatory=$false)]
    [boolean]$Resume=$true,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0,365)]
    [int16]$days = 180

)

function Get-ADMailboxUsers($searchbase) {
    # retrieves all mailbox users in Active Directory, flushes users and selected properties to file

    $ADMailboxUsers = Get-ADUser -searchbase $searchbase -Filter {homeMDB -like "*" -and displayname -notlike "HealthMailbox*" -and name -notlike "SystemMailbox{*"} -properties distinguishedname,msDS-parentdistname,title,description,displayname,msExchDelegateListLink,publicdelegates,publicdelegatesBL,extensionattribute1,extensionattribute2,extensionattribute3,extensionattribute4,extensionattribute5,lastlogondate,created,modified,homeMDB,mailnickname,msexchwhenmailboxcreated,passwordlastset,legacyExchangeDN

    $ADMailboxUsers | export-csv -Path $script:ADUserOutput -NoTypeInformation -Force
    $core = get-content -Path $script:ADUserOutput
    Set-Content -Path $script:ADUserOutput -Value $searchbase,$core
    return $ADMailboxUsers
}

function Get-ServerMailboxStatistics() {
    # retrieve all mailbox statistics per server to reduce overall runtime of script

    $ExchangeServers = Get-ExchangeServer | Where-Object{$_.serverrole -like "Mailbox"}
    $script:MBXStatsAll = foreach($ExchangeServer in $ExchangeServers){
        Get-MailboxStatistics -server $ExchangeServer -noADLookup | Select-Object TotalItemSize,TotalDeletedItemSize,mailboxguid,lastlogontime,legacyDn,DisplayName
        Write-Information "Success - retrieved Mailbox Statistics for all mailboxes from server $ExchangeServer"
    }
    Write-Information "$($script:MBXStatsAll.count) total mailboxes"
    $script:MBXStatsAll | export-csv -Path $script:MBXStatsOutput -NoTypeInformation -Force
    return $script:MBXStatsAll
}

function Stop-MailboxAuditStatistics() {
    [cmdletbinding()]
    Param(
        [Parameter( Mandatory=$false)]	
        [bool]$Flush,
    
        [Parameter( Mandatory=$true)]
        $aduser,

        [Parameter( Mandatory=$true)]
        [int]$index
    )
    
    # Now to put all that info into a spreadsheet. 
    $mbxAdCombo | export-csv -path $script:mailboxAuditOutput -notypeinformation -Append
    if ($Flush)
    {
        Write-Information "Flushing..."
        $mbxAdCombo.Clear()
        $null = New-Item -Path $script:ResumeIndexOutput -ItemType File -Force
        Set-Content -Path $script:ResumeIndexOutput -Value $index
        Add-Content -Path $script:ResumeIndexOutput -Value $script:mailboxAuditOutput
    }
    else {
        Write-Verbose "Stopping collection at index of $index and writing to $($script:mailboxAuditOutput)"
        Write-Verbose "Stopping with user $($aduser.displayname)"
        $null = New-Item -Path $script:ResumeIndexOutput -ItemType File -Force
        Set-Content -Path $script:ResumeIndexOutput -Value $index
        Add-Content -Path $script:ResumeIndexOutput -Value $script:mailboxAuditOutput 
        $mbxAdCombo = $null
    }
}

function Exit-MailboxAuditStatistics($auto_errors) {
    foreach ($err in $auto_errors) 
    {  
        $logdata = $null 
        $logdata = $err 
        if ($logdata) 
        { 
            out-file -filepath $script:errfilename -Inputobject $logData -Append 
        } 
    }

    $endDate = Get-Date
    $elapsedTime = $endDate - $startDate
    Write-Verbose "Report started at $($script:startDate)."
    Write-Verbose "Report ended at $($endDate)."
    Write-Verbose "Total Elapsed Time: $($elapsedTime)"
}

function Add-MailboxAuditStatistics ($aduser) {
        [string]$inboxRules=$null
        [string]$SendAs=$null
        
        $mailnickname=$aduser.mailnickname 
        Write-Verbose "aduser legdn $($aduser.legacyExchangeDN)"
        $mailbox = $script:MBXStatsAll.where({$_.legacydn -eq ($aduser.legacyExchangeDN)})

        [datetime]$mbxlogoncompare = $mailbox.lastlogontime
        Write-Verbose "Looking at $($mailbox.DisplayName)"

        if($null -eq $mailbox.lastlogontime)
        {
            $MbxLastLogon="Never"
        }
        else{$MbxLastLogon=$mailbox.lastlogontime}
        write-Verbose "Last Mailbox Logon Time for $($mailbox.legacyexchangeDN) is... $MBXLastLogon"
        if ($mbxlogoncompare -le $agedDate -or ($MbxLastLogon -eq "Never"))
        {
            $MailboxProps = get-mailbox $aduser.distinguishedname | Select-Object RecipientTypeDetails,ProhibitSendQuota
            $ProhibitSendQuota=$MailboxProps.ProhibitSendQuota.ToString()
            Write-Verbose "Quota check is $ProhibitSendQuota"
            $RecipientTypeDetails=($MailboxProps.RecipientTypeDetails | out-string).Trim()
            Write-Verbose "Mailbox type is $RecipientTypeDetails"
            $inboxRules = Get-InboxRule -Mailbox $aduser.displayname | Select-Object name,enabled
            [string]$inboxRulesformatted = $inboxRules -split '-------'
            [string]$SendAs = (Get-ADPermission -Identity $aduser.distinguishedname | 
            `Where-Object {$_.isinherited -eq $false -and $_.extendedrights -like "Send-As" -and $_.User.RawIdentity -ne "NT AUTHORITY\SELF"} | 
            `Select-Object user).user.RawIdentity
        }
        [string]$publicdelegates = $aduser.publicdelegates.value
        [string]$publicdelegatesBL = $aduser.publicdelegatesBL.value
        [string]$fullMbxAccess = $aduser.msExchDelegateListLink.value
        
        # ordered helps maintain the arraylist's value pairs in this explicit orders
        #  order isn't respected when script is run from ISE
        $line = [ordered]@{
            DisplayName=$aduser.displayname
            Title=$aduser.Title
            'Mailbox Type'= $RecipientTypeDetails
            'Mailbox last logon'= $MbxLastLogon
            'Mailbox Created'=$aduser.msexchwhenmailboxcreated
            'AD account created'=$aduser.created
            'AD password last set'=$aduser.passwordlastset
            'AD last logon date'=$aduser.lastlogondate
            'AD account last modified'=$aduser.modified
            mailnickname=$mailnickname
            database=(($aduser.homeMDB -split ',CN=')[0] -split 'CN=')[1]
            TotalItemSize=$mailbox.TotalItemSize
            TotalDeletedItemSize=$mailbox.TotalDeletedItemSize
            ProhibitSendquota = $ProhibitSendQuota
            description=$aduser.description
            Delegates=$publicdelegates
            'Delegate for'=$publicdelegatesBL
            'Full Mailbox Access'=$fullMbxAccess
            'Send As'=$SendAs
            extensionattribute1=$aduser.extensionattribute1
            extensionattribute2=$aduser.extensionattribute2
            extensionattribute3=$aduser.extensionattribute3
            extensionattribute4=$aduser.extensionattribute4
            extensionattribute5=$aduser.extensionattribute5
            'Inbox Rules'= $inboxRulesformatted
            OU=$aduser.'msDS-parentdistname'
    
        }
    return $line
}

# Main script
# Start new timer
$script:startDate = Get-Date

# Clear error variable
$error.Clear()

# Setup time date variables 
$timestamp = Get-Date -Format o | ForEach-Object {$_ -replace ":", "."}
# $daystamp = Get-Date -Format 'MM-dd-yyyy'

# Setup output variables script wide
$script:outputFolder= "$($ENV:HOMEPATH)\desktop\"
$script:mailboxAuditOutput = $script:outputFolder + "MBX audit" + $timestamp + ".csv"
$script:ADUserOutput = $script:outputFolder + "AD users with mailboxes.csv"
$script:MBXStatsOutput = $script:outputFolder + "MBX stats for all mailboxes.csv"
$script:ResumeIndexOutput = $script:outputFolder + "ResumeMailboxAudit.log"
$script:errfilename = $script:outputFolder + "Errorlog_" + $timestamp + ".txt" 

[datetime]$agedDate = (get-date).adddays(-($days))
[int]$ResumeIndex = 0
Set-ADServerSettings -ViewEntireForest $true
#[string]$ADSearchBase = "OU=VIPs,OU=Departments,DC=fabrikam,DC=com"

[int]$mbxcount = 0
[int]$i=1  # variable for progress bar
[int]$index = 0  # variable for iterating through primary mailbox user array 
[int]$flushLimit = 100 # When to flush
[int]$flushCount = 0 # counting to flush limit

write-progress -id 1 -activity "Getting all on prem mailboxes from Active Directory" -PercentComplete (1)

if (Test-Path -Path $script:ADUserOutput)
{
    if ((get-content -Path $script:ADUserOutput | Select-Object -First 1) -eq $SearchBase)
    {
        Write-Information "Using previously saved AD user list from today, $($script:ADUserOutput)"
        $ADUsers = get-content -Path $script:ADUserOutput | Select-Object -skip 1 | Out-String | ConvertFrom-Csv
        
        if ($Resume)
        {
        if (Test-Path -Path $script:ResumeIndexOutput)
        {
            $ResumeContent = Get-Content $script:ResumeIndexOutput
            $ResumeIndex = $ResumeContent[0]
            [string]$ResumeMbxOutput = $ResumeContent[1]
            $script:mailboxAuditOutput = $ResumeMbxOutput

            Write-Information "Resuming report at index of $ResumeIndex with output file $ResumeMbxOutput"
            Write-Information "Index $ResumeIndex correlates to user $($adUsers[$ResumeIndex].displayname)" 
        }
        }
        else {
            Write-Information "... but starting fresh mailbox audit"}
    }
    else
    {
        Write-Information "Performing new AD lookups as SearchBase has changed"
        $adusers=Get-ADMailboxUsers($SearchBase)
    }

}
else 
{
    $adusers=Get-ADMailboxUsers($SearchBase)
}

if (Test-Path -Path $script:MBXStatsOutput)
{
    Write-Information "Using previously saved Mailbox statistics from earlier, $($script:MBXStatsOutput)"
    $script:MBXstatsAll = get-content -Path $script:MBXStatsOutput | out-string | ConvertFrom-Csv
}
else {
    $script:MBXstatsAll = Get-ServerMailboxStatistics
}

$i = $ResumeIndex
$mbxcount = ($adusers | Measure-Object).count

write-progress -id 1 -activity "Getting all Audit info for $mbxcount on prem mailboxes" -PercentComplete (10)
Write-Host "Press the Ctrl-C key to stop and save progress so far..."

# Ctrl-C can't be constrained or treated as input from ISE
[System.Console]::TreatControlCAsInput = $True
[System.Collections.ArrayList]$mbxAdCombo = New-Object System.Collections.ArrayList($null)

$Index = $ResumeIndex
While ($Index -lt $mbxcount) 
{
    if ($Host.UI.RawUI.KeyAvailable -and ($Keypress = $Host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp")))
    {
        if([Int]$Keypress.Character -eq 3)
        {
            Write-Warning "CTRL-C was used - Shutting down any running jobs before exiting the script."
            Stop-MailboxAuditStatistics -aduser $adusers[$index] -index $Index -Verbose
            Break
        }
    }
   
    $percentage=(($i/$mbxcount)*90) + 10
    write-progress -id 1 -activity "Processing $mbxcount on prem mailboxes" -PercentComplete ($percentage) -Status "Currently getting stats for mbx # $i ... $($adusers[$index].DisplayName)"
    $stats = Add-MailboxAuditStatistics($ADUsers[$index])

    $null = $mbxAdCombo.Add((New-Object PSobject -property $stats))
    if ($flushCount -ge $flushLimit)
    {
        Write-Verbose "Flushing at index $index and appending to $($script:mailboxAuditOutput)"
        Stop-MailboxAuditStatistics -Flush $true -aduser $ADUsers[$index] -index $Index -Verbose
        $flushCount=0
    }
 
    $i++
    $index++
    $flushCount++
}

# If we get through the entire list of AD users, we need to write output since last flush and cleanup resume file
if ($index -eq $mbxcount) 
{
    $mbxAdCombo | export-csv -path $script:mailboxAuditOutput -notypeinformation -Append

    if (Test-Path -Path $script:ResumeIndexOutput)
    {
        Remove-Item $script:ResumeIndexOutput -Force
    }
}

Exit-MailboxAuditStatistics($error) -Verbose