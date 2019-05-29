<#
.Notes
    Name: Get-MailboxAudit.ps1
    Author: Austin McCollum [austinmc@microsoft.com]
    Version 1.0: 4/18/2019 finally upgraded sample script to fully featured
.Description
A comprehensive audit of mailboxes to identify mailboxes that haven't been logged into for a long time.
.Parameter Resume
Look for saved cache of ad results and compare results already collected
.Example

#>
[cmdletbinding()]
Param(
    [Parameter( Mandatory=$false)]	
    [string]$Resume
)

function Get-ADMailboxUsers {
    # retrieves all mailbox users in Active Directory, flushes users and selected properties to file 
    $ADMailboxUsers = Get-ADUser -Filter {homeMDB -like "*" -and displayname -notlike "HealthMailbox*" -and name -notlike "SystemMailbox{*"} -properties distinguishedname,
        `msDS-parentdistname,title,description,displayname,msExchDelegateListLink,publicdelegates,publicdelegatesBL,extensionattribute1,extensionattribute2,extensionattribute3,
        `extensionattribute4,extensionattribute5,lastlogondate,created,modified,homeMDB,mailnickname,msexchwhenmailboxcreated,passwordlastset
    $ADMailboxUsers | export-csv -Path $script:ADUserOutput -NoTypeInformation    
    return $ADMailboxUsers
}

function Stop-MailboxAuditStatistics() {
    [cmdletbinding()]
    Param(
        [Parameter( Mandatory=$false)]	
        [bool]$Flush,
    
        [Parameter( Mandatory=$true)]
        [string]$aduser
    )
    
    
    # Now to put all that info into a spreadsheet. 
    $mbxAdCombo | export-csv -path $script:mailboxAuditOutput -notypeinformation -Append
    if ($Flush)
    {
       
        [System.Collections.ArrayList]$mbxAdCombo = New-Object System.Collections.ArrayList($null)
    }
    else {
        Write-Verbose "Stopping collection $i and writing to $($script:mailboxAuditOutput)"
    }
}

function Set-MailboxAuditStatistics($aduser) {
    
    $mbxAdCombo | Export-Csv -Path $script:mailboxAuditOutput -NoTypeInformation -Append
    
}

function Add-MailboxAuditStatistics ($aduser) {
        [string]$RecipientTypeDetails=$null
        [string]$inboxRules=$null
        [string]$SendAs=$null
    
        $mailnickname=$aduser.mailnickname    
        $mailbox = Get-MailboxStatistics -identity $mailnickname | Select-Object TotalItemSize,TotalDeletedItemSize,database,mailboxguid,lastlogontime,ProhibitSendquota
    
        if($null -eq $mailbox.lastlogontime)
        {
            $MbxLastLogon="Never"
        }
        else{$MbxLastLogon=$mailbox.lastlogontime}
    
        if ($mailbox.lastlogontime -gt $agedDate)
        {
            $RecipientTypeDetails = (get-recipient $aduser.displayname | Select-Object RecipientTypeDetails).RecipientTypeDetails
            $inboxRules = Get-InboxRule -Mailbox $aduser.displayname | Select-Object name,enabled
            [string]$inboxRulesformatted = $inboxRules -split '-------'
            [string]$SendAs = (Get-ADPermission -Identity $aduser.distinguishedname | Where-Object {$_.isinherited -eq $false -and $_.extendedrights -like "Send-As" -and $_.User.RawIdentity -ne "NT AUTHORITY\SELF"} | select user).user.RawIdentity
        }
        [string]$publicdelegates = $aduser.publicdelegates.value
        [string]$publicdelegatesBL = $aduser.publicdelegatesBL.value
        [string]$fullMbxAccess = $aduser.msExchDelegateListLink.value
        
        
        $line = @{
            DisplayName=$aduser.displayname
            'AD account created'=$aduser.created
            'AD password last set'=$aduser.passwordlastset
            'AD last logon date'=$aduser.lastlogondate
            'AD account last modified'=$aduser.modified
              
            mailnickname=$mailnickname
            'Mailbox Created'=$aduser.msexchwhenmailboxcreated
            database=(($aduser.homeMDB -split ',CN=')[0] -split 'CN=')[1]
            Title=$aduser.Title
    
            TotalItemSize=$mailbox.TotalItemSize
            TotalDeletedItemSize=$mailbox.TotalDeletedItemSize
            
            'Mailbox last logon'= $MbxLastLogon
            ProhibitSendquota = $mailbox.ProhibitSendquota
            
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
    
            'Mailbox Type'= $RecipientTypeDetails
            'Inbox Rules'= $inboxRulesformatted
    
            OU=$aduser.'msDS-parentdistname'
    
        }
        return $line
    }

# Main script
$startDate = Get-Date
$error=$null

# Setup some variables for flushing data at regular intervals
$timestamp = Get-Date -Format o | ForEach-Object {$_ -replace ":", "."}
$daystamp = Get-Date -Format 'MM-dd-yyyy'
$outputFolder= "$($ENV:HOMEPATH)\desktop\"
$script:mailboxAuditOutput = $outputfolder + "MBX audit" + $timestamp + ".csv"
$script:mailboxAuditOutputTest = $outputfolder + "MBX audit Test" + $timestamp + ".csv"
$script:ADUserOutput = $outputfolder + "AD users with mailboxes" + $daystamp + ".csv"


#$agedADaccount = (get-date).adddays(-365)
$agedDate = (get-date).adddays(-365)
#$everyoneDate= (get-date)
Set-ADServerSettings -ViewEntireForest $true

write-progress -id 1 -activity "Getting all on prem mailboxes from Active Directory" -PercentComplete (1)

if (Test-Path -Path $script:ADUserOutput)
{
    Write-Information "Using previously saved AD user list from today, $($script:ADUserOutput)"
    $ADUsers = Import-Csv -Path $script:ADUserOutput
}
else 
{
    $adusers=Get-ADMailboxUsers
}

[int]$mbxcount = ($adusers | Measure-Object).count
[int]$i=1
[int]$flushLimit = 100
[int]$flushCount = 0

write-progress -id 1 -activity "Getting all Audit info for $mbxcount on prem mailboxes" -PercentComplete (10)
Write-Host "Press the Ctrl-C key to stop and save progress so far..."
[Console]::TreatControlCAsInput = $True
[System.Collections.ArrayList]$mbxAdCombo = New-Object System.Collections.ArrayList($null)

foreach ($mbxuser in $adusers) 
{
    
    if ($Host.UI.RawUI.KeyAvailable -and ($Keypress = $Host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp")))
    {
        if([Int]$Keypress.Character -eq 3)
        {
            Write-Warning "CTRL-C was used - Shutting down any running jobs before exiting the script."
            Stop-MailboxAuditStatistics -aduser $mbxuser -Verbose
        }
    }
   
    $percentage=(($i/$mbxcount)*90) + 10
    write-progress -id 1 -activity "Processing $mbxcount on prem mailboxes" -PercentComplete ($percentage) -Status "Currently getting stats for mbx # $i ... $($mbxuser.DisplayName)"
    $stats = Add-MailboxAuditStatistics($mbxuser)

    $null = $mbxAdCombo.Add((New-Object PSobject -property $stats))
    if ($flushCount -ge $flushLimit)
    {
        Write-Verbose "Flushing collection $i and appending to $($script:mailboxAuditOutput)"
        Stop-MailboxAuditStatistics -Flush $true -aduser $mbxuser -Verbose
        $flushCount=0
    }
 
    $i++
    $flushCount++
}

$errfilename = $outputfolder + "Errorlog_" + $timestamp + ".txt" 

foreach ($err in $error) 
{  
    $logdata = $null 
    $logdata = $err 
    if ($logdata) 
    { 
        out-file -filepath $errfilename -Inputobject $logData -Append 
    } 
}

$endDate = Get-Date
$elapsedTime = $endDate - $startDate
Write-Host "Report started at $($startDate)."
Write-Host "Report ended at $($endDate)."
Write-Host "Total Elapsed Time: $($elapsedTime)"