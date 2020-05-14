# Need credentials for many of the invoking commands - here's a quick way to get creds into a variable temporarily
$adminpass=read-host -Prompt "enter password" -AsSecureString
$adminuser="Fabrikam\fabadmin"
$admincreds=new-object System.Management.Automation.PSCredential -ArgumentList $adminuser,$adminpass 

#setup firewall rules for robocopy
$computernames = @("ex2019","ex2019-2","ex2019-3")

Invoke-Command -ComputerName $computernames -ScriptBlock {
New-NetFirewallRule -Name "RoboCopy_tcp_in_TEMP" -DisplayName "Allow Robocopy TCP-In" -Description "Need to copy files for Exchange Setup - disable after Exchange install successful" -Enabled True -Protocol TCP -RemotePort 445 -Action Allow -RemoteAddress "10.0.0.5-10.0.0.7" -Direction Inbound; New-NetFirewallRule -Name "RoboCopy_tcp_out_TEMP" -DisplayName "Allow Robocopy TCP-Out" -Description "Need to copy files for Exchange Setup - disable after Exchange install successful" -Enabled True -Protocol TCP -RemotePort 445 -Action Allow -RemoteAddress "10.0.0.5-10.0.0.7" -Direction Outbound}

Invoke-Command -ComputerName $computernames -ScriptBlock {Get-NetFirewallRule -name "Robocopy_tcp_in_TEMP" | Get-NetFirewallApplicationFilter | set-NetFirewallApplicationFilter -program "%SystemRoot%\System32\Robocopy.exe";Get-NetFirewallRule -name "Robocopy_tcp_out_TEMP" | Get-NetFirewallApplicationFilter | set-NetFirewallApplicationFilter -program "%SystemRoot%\System32\Robocopy.exe"}

Invoke-Command -ComputerName $computernames -ScriptBlock {mkdir c:\Automation}

#create share on jump server with everyone Read permissions at c:\automation
# include scripts, UCMA, input files, etc.

Invoke-Command -ComputerName $computernames -ScriptBlock {Robocopy.exe /S \\10.0.0.5\Automation\ c:\Automation}


Invoke-Command -ThrottleLimit 32 -ComputerName $computernames-scriptblock {param([securestring]$admincreds) New-PSDrive -Name X -PSProvider FileSystem -Root \\Ex1\g$\js -Credential $admincreds; Copy-Item -Path x:\ -Destination "g:\" -Recurse} -ArgumentList $admincreds
#1..2 | foreach {"EX$_"} | Get-ADComputer | % {Invoke-Command -ComputerName $_.name {get-disk | ft}}
$computernames=(1..2 | ForEach-Object {"EX$($_)"} | Get-ADComputer).name
$computernames=1..2 | ForEach-Object {"EX$($_)"}

#Invoke-Command -ThrottleLimit 32 -ComputerName $computernames -ScriptBlock {param($admincreds) New-Item -Path f:\ -ItemType directory -Name "Automation"} -argumentlist $admincreds

## copy Exchange setup files to all the Exchange servers before Exchange is installed
## can also be used to copy UC managed API prereq http://go.microsoft.com/fwlink/p/?linkId=258269
Invoke-Command -ThrottleLimit 32 -ComputerName EX2 -scriptblock {param([SecureString] $admincreds) New-PSDrive -Name X -PSProvider FileSystem -Root \\EX1\f$\Automation -Credential $admincreds; Copy-Item -Path x:\ -Destination "f:\" -Recurse} -ArgumentList $admincreds

Invoke-Command -ThrottleLimit 32 -ComputerName $computernames -scriptblock {param([securestring]$admincreds) New-PSDrive -Name X -PSProvider FileSystem -Root \\ex1\g$\js -Credential $admincreds; Copy-Item -Path x:\ -Destination "g:\" -Recurse} -ArgumentList $admincreds

## running diskpart.ps1 on all the machines to set / reset mountpoints and volumes -- requires DBmap.csv
Invoke-Command -ThrottleLimit 32 -ComputerName $computernames -scriptblock {param([securestring]$admincreds) f:\automation\diskpart.ps1 -serverfile f:\automation\labdag_servers.csv -whatif:$false} -ArgumentList $admincreds

Invoke-Command -ComputerName EX2 -ScriptBlock { 
    param([securestring]$admincreds) Install-package -name "F:\Automation\Jetstress.msi" -Force
} -ArgumentList $admincreds

#http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/30/learn-how-to-configure-powershell-memory.aspx
# in order for Jetstress to run through invoke-command, need to increase memory thresholds,
 
Set-Item WSMan:\localhost\Shell\MaxMemoryPerShellMB 20480
Set-Item WSMan:\localhost\Plugin\Microsoft.PowerShell\Quotas\MaxMemoryPerShellMB 20480
restart-service winrm

#After the ese files have been copied to the Jetstress directory, you need a first run of Jetstress to register the ese perf objects.

Invoke-Command -ComputerName EX2 { & 'C:\Program Files\Exchange Jetstress\jetstresscmd.exe' /?}

#Here’s what the expected output looks like:
<#
[PS] C:\Windows\system32>Invoke-Command -ComputerName $computerlist { & 'C:\Program Files\Exchange Jetstress\jetstresscmd.exe' /?}
6/3/2019 7:41:47 PM -- Microsoft Exchange Jetstress 2013 Core Engine (version: 15.00.0995.000) detected.
6/3/2019 7:41:47 PM -- Windows Server 2012 R2 Standard  (6.2.9200.0) detected.
6/3/2019 7:41:47 PM -- Microsoft Exchange Server Database Storage Engine (version: 15.00.0995.021) was detected.
6/3/2019 7:41:47 PM -- Microsoft Exchange Server Database Storage Engine Performance Library (version: 15.00.0995.021) was detect
ed.
6/3/2019 7:41:48 PM -- The MSExchange Database or MSExchange Database ==> Instances performance counter category isn't
registered.
    + CategoryInfo          : NotSpecified: (6/3/2019 7:41:4...n't registered.:String) [], RemoteException
    + FullyQualifiedErrorId : NativeCommandError
    + PSComputerName        : ex15-01

6/3/2019 7:41:48 PM -- The Database Storage Engine Performance Library was registered successfully.
6/3/2019 7:43:10 PM -- Database Storage Engine performance objects and counters were successfully loaded.
6/3/2019 7:43:10 PM -- Advanced Database Storage Engine Performance counters were successfully enabled.
This process will restart itself within 5 seconds.
6/3/2019 7:43:10 PM -- This process must be restarted for the changes to take effect.
[PS] C:\Windows\system32>
#>


# "C:\Program Files\Exchange Jetstress\JetstressCmd.exe"  /c "C:\Program Files\Exchange Jetstress\JetstressConfigInitialize.xml" /timeout 0H0M0S /new /threads 0

Invoke-Command -ComputerName $computernames  { & 'C:\Program Files\Exchange Jetstress\jetstresscmd.exe' /c "F:\Automation\$($env:COMPUTERNAME).xml" /timeout 0H0M0S /new /threads 0}
Invoke-Command -ComputerName EX2  { & 'C:\Program Files\Exchange Jetstress\jetstresscmd.exe' /c "F:\Automation\EX2.xml" /timeout 0H5M0S /open /threads 2}

$computernames_contoso_DAG1 = 11..18 | ForEach-Object {"Ex$($_)SITE1"} 
$computernames_contoso_DAG1 += 19..26 | ForEach-Object {"Ex$($_)SITE2"} 

# Cleaning up after Jetstress
#  uninstalling Jetstress.  Uninstall-package works great for providername msi, but if you attempt providername "programs" it will silently fail.
Invoke-Command -ComputerName EX2 -ScriptBlock { 
    param([securestring]$admincreds) Uninstall-package -name "Microsoft Exchange Jetstress 2013" -Force
} -ArgumentList $admincreds

#  run diskpart again to remove all Jetstress DBs

#validate diskpart
$computernames=(1..2 | ForEach-Object {"EX$($_)"} | Get-ADComputer).name
Invoke-Command -ComputerName $computernames {
    $VolVerify=get-volume -FileSystemLabel ExVol3
    if($volverify.path){
        get-partition | Where-Object {$_.accesspaths -like $volverify.path}
    }
} `
    | Format-Table disknumber,accesspaths,size,pscomputername

#    disknumber accesspaths                                                                                                   size PSComputerName
#    ---------- -----------                                                                                                   ---- --------------
#             3 {F:\ExchangeVols\ExVol3\, F:\ExchangeDBs\DAG1-DB1\, \\?\Volume{0b097e7e-c8d5-42d2-a826-d8a786de3e1a}\} 53550776320 EX2
# --> missing 3 {F:\ExchangeVols\ExVol3\, \\?\Volume{8cc5f485-5525-4eb6-bcac-5485144e3858}\}                           53550776320 EX1

# Determin if HyperThreading is on
$ComputerName = $env:COMPUTERNAME
$LogicalCPU = 0
$PhysicalCPU = 0
$Core = 0

# Get the Processor information from the WMI object
$Proc = [object[]]$(get-WMIObject Win32_Processor -ComputerName $ComputerName)
if ($($Proc | measure-object -Property NumberOfLogicalProcessors -sum).Sum -gt $($Proc | measure-object -Property NumberOfCores -sum).Sum)
{write-host "HypterThreading is enabled" -ForegroundColor Yellow}
Else {write-host "HyperThreading is not enabled" -ForegroundColor Green}


$accesspaths = get-partition
$accessRemoval =$accesspaths | ?{$_.accesspaths -like 'F:\Exchange*'}
$accessRemoval | ft disknumber,accesspaths

1..8 | %{"DAG1-DB0$_"} | %{New-Item -Type directory -Path "F:\ExchangeDatabases\$($_)" -WhatIf}


[Array]$DiskPart = Import-CSV -Path F:\Automation\labdagservers.csv
$Machine = get-wmiobject "Win32_ComputerSystem"
$Machine.Name
Foreach ($Server in $Diskpart) {
    If ($Machine.name -eq $Server.ServerName) {
        $DBperVolume = [int]$Server.DBperVolume
        $DiskStart = [int]$Server.StartDrive
        $ExchangeDBs = $Server.DatabasesRootFolder
        [Array]$DbMap = $Server.DbMap -split ","
        $LastDrive = $DiskStart - 1 + $DbMap.Count/$DBperVolume
    }
}

$DBmounts = get-partition | ?{$_.accesspaths -like 'F:\ExchangeVols*'}
$DBcount=$DBmap.Count
$count=0
Foreach ($DBmount in $DbMounts) {
    if($count -gt $DBcount){write-host "something wrong here!"}
    $CurrentDisk = $DBmount.disknumber
    $CurrentPartition = $DBmount.partitionnumber
    for ($CurrentMount = 1; $CurrentMount -eq $DBperVolume)
    {
        [string]$DB=$DBmap[$count]
        $DBPath = "$ExchangeDBs\$DB"
        Add-PartitionAccessPath -DiskNumber $currentdisk -PartitionNumber $currentpartition -AccessPath "$DBPath"-Passthru |Set-Partition -NoDefaultDriveLetter:$True
        $count++
    }
}