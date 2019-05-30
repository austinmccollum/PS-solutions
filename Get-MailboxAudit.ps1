<#
.Notes
    Name: Get-MailboxAudit.ps1
    Author: Austin McCollum [austinmc@microsoft.com]
    Version 1.0: 4/18/2019 finally upgraded sample script to fully featured
.Description
A comprehensive audit of mailboxes to identify mailboxes that haven't been logged into for a long time.
.Example

#>
function Get-ADMailboxUsers {
    # retrieves all mailbox users in Active Directory, flushes users and selected properties to file 
    $ADMailboxUsers = Get-ADUser -searchbase "OU=VIPs,OU=Departments,DC=fabrikam,DC=com" -Filter {homeMDB -like "*" -and displayname -notlike "HealthMailbox*" 
        `-and name -notlike "SystemMailbox{*"} -properties distinguishedname,msDS-parentdistname,title,description,displayname,msExchDelegateListLink,
        `publicdelegates,publicdelegatesBL,extensionattribute1,extensionattribute2,extensionattribute3,extensionattribute4,extensionattribute5,lastlogondate,
        `created,modified,homeMDB,mailnickname,msexchwhenmailboxcreated,passwordlastset
    $ADMailboxUsers | export-csv -Path $script:ADUserOutput -NoTypeInformation    
    return $ADMailboxUsers
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

function Exit-MailboxAuditStatistics($error) {
    foreach ($err in $error) 
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
# Start new timer
$script:startDate = Get-Date

# Clear error variable
$error.Clear()

# Setup time date variables 
$timestamp = Get-Date -Format o | ForEach-Object {$_ -replace ":", "."}
$daystamp = Get-Date -Format 'MM-dd-yyyy'

# Setup output variables script wide
$script:outputFolder= "$($ENV:HOMEPATH)\desktop\"
$script:mailboxAuditOutput = $script:outputFolder + "MBX audit" + $timestamp + ".csv"
$script:ADUserOutput = $script:outputFolder + "AD users with mailboxes.csv"
$script:ResumeIndexOutput = $script:outputFolder + "ResumeMailboxAudit.log"
$script:errfilename = $script:outputFolder + "Errorlog_" + $timestamp + ".txt" 

$agedDate = (get-date).adddays(-365)
[int]$ResumeIndex = 0
Set-ADServerSettings -ViewEntireForest $true

[int]$mbxcount = 0
[int]$i=1  # variable for progress bar
[int]$index = 0  # variable for iterating through primary mailbox user array 
[int]$flushLimit = 100 # flush count to flush at
[int]$flushCount = 0 # current flush count

write-progress -id 1 -activity "Getting all on prem mailboxes from Active Directory" -PercentComplete (1)

if (Test-Path -Path $script:ADUserOutput)
{
    Write-Information "Using previously saved AD user list from today, $($script:ADUserOutput)"
    $ADUsers = Import-Csv -Path $script:ADUserOutput
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
else 
{
    $adusers=Get-ADMailboxUsers
}

$i = $ResumeIndex
$mbxcount = ($adusers | Measure-Object).count

write-progress -id 1 -activity "Getting all Audit info for $mbxcount on prem mailboxes" -PercentComplete (10)
Write-Host "Press the Ctrl-C key to stop and save progress so far..."
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