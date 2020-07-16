###################################################################################### 
# Get-EASlogs_scheduled.ps1  
#  based on script created by Jim Martin and Matt Stehle 
# 
# This script can help in data collection where it may take some time to reproduce the  
#  issue or the logs are overwritten too quickly. You can then grab all of the logs and  
#  open them with the Mailbox Log Parser tool for quick analysis. 
# 
# MailboxLogParser is tool that is able to consolidate multiple Activesync Mailbox Logs 
#  collected by this script and present in a flattened view for analysis. 
#  https://github.com/edwin-huber/MailboxLogParser/releases/tag/v2.2
# 
# This script is modified to work best with a hard coded list 
#  of mailbox users to collect logs using task scheduler. 
#  http://social.technet.microsoft.com/wiki/contents/articles/23150.how-to-use-task-scheduler-for-exchange-scripts.aspx 
#  https://social.technet.microsoft.com/wiki/contents/articles/32656.office-365-how-to-schedule-a-script-using-task-scheduler.aspx 
#
# Script must be run with Exchange powershell commands loaded. 
# 
# 0. Clear the error log so that sending errors to file relate only to this run of the script 
#   
$error.clear() 
 
# 1. Folder to save the mailbox logs to ** this requires manual input ** 
#  ex. $outputFolder='C:\temp\logs\' 
$outputFolder = 'c:\EASmbxlogs\' 
 
# 2. SMTP addrss of mailbox to retrieve logs from ** this requires manual input ** 
#  ex. $targetMailboxes ="user1@fabrikam.info","user2@fabrikam.com" 
$targetMailboxes = "VIP1@contoso.com", "VIP2@fabrikam.com" 
 
# 3. Check Exchange version for cmdlet 
if ((Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup\")) { 
    if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup\").MsiProductMajor -eq 15) { 
        $version = "15" 
    } 
} 
if ((Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup\")) { 
    if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup\").MsiProductMajor -eq 14) { 
        $version = "14" 
    } 
} 
 
# 4. Turn on mailbox logging for each target mailbox - no changes will be made if logging is already on 
foreach ($targetMailbox in $targetMailboxes) { 
    # Turn logging on 
    Set-CasMailbox $targetMailbox -ActiveSyncDebugLogging:$true 
} 
foreach ($targetMailbox in $targetMailboxes) { 
    Write-Host "Getting all devices for mailbox:" $targetMailbox 
 
    # ...get all devices syncing with mailbox... 
    if ($version -eq "15") { $devices = Get-MobileDeviceStatistics -Mailbox $targetMailbox } 
    else { $devices = Get-ActiveSyncDeviceStatistics -Mailbox $targetMailbox } 
     
    # ..and for each device... 
    foreach ($device in $devices) { 
        Write-Host "Downloading logs for device: " $device.DeviceFriendlyName $device.DeviceID 
 
        # ...create an output file... 
        $fileName = $outputFolder + "MailboxLog_" + $device.DeviceFriendlyName + "_" + $device.DeviceID + "_" + (Get-Date).Ticks + ".txt" 
 
        # ...and write the mailbox log to the output file... 
        $logData = $null 
        if ($version -eq "15") { $logData = (Get-MobileDeviceStatistics $device.Identity -GetMailboxLog -ErrorAction SilentlyContinue | select -ExpandProperty MailboxLogReport) } 
        if ($version -eq "14") { $logData = (Get-ActiveSyncDeviceStatistics $device.Identity -GetMailboxLog:$true -ErrorAction SilentlyContinue | select -ExpandProperty MailboxLogReport) } 
             
        if ($logData.Length -gt 0) { 
            Write-Host "Saving logs to: " $fileName 
            Out-File -FilePath $fileName -InputObject $logData 
        } 
    } 
}      
 
# 5. Write errors to a log 
$errfilename = $outputfolder + "Errorlog_" + (Get-Date).Ticks + ".txt" 
 
foreach ($err in $error) {  
    $logdata = $null 
    $logdata = $err 
    if ($logdata) { 
        out-file -filepath $errfilename -Inputobject $logData -Append 
    } 
}