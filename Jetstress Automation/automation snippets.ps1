# Need credentials for many of the invoking commands - here's a quick way to get creds into a variable temporarily
$adminpass=read-host -Prompt "enter password" -AsSecureString
$adminuser="Fabrikam\fabadmin"
[securestring]$admincreds=new-object System.Management.Automation.PSCredential -ArgumentList $adminuser,$adminpass -

Invoke-Command -ThrottleLimit 32 -ComputerName $computernames-scriptblock {param([securestring]$admincreds) New-PSDrive -Name X -PSProvider FileSystem -Root \\Ex1\g$\js -Credential $admincreds; Copy-Item -Path x:\ -Destination "g:\" -Recurse} -ArgumentList $admincreds
#1..2 | foreach {"EX$_"} | Get-ADComputer | % {Invoke-Command -ComputerName $_.name {get-disk | ft}}
$computernames=(1..2 | ForEach-Object {"EX$($_)"} | Get-ADComputer).name

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

Invoke-Command -ComputerName EX2  { & 'C:\Program Files\Exchange Jetstress\jetstresscmd.exe' /c "F:\Automation\EX2.xml" /timeout 0H0M0S /new /threads 0}

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