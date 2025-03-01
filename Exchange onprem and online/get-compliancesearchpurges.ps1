# Collect the target email address
$addressOrSite = Read-Host "Enter an email address"

# Authenticate with Exchange Online and the Security & Complaince Center (Exchange Online Protection - EOP)
if (!$credentials) {
    $credentials = Get-Credential
}

if ($addressOrSite.IndexOf("@") -ige 0) {
    # List the folder Ids for the target mailbox
    $emailAddress = $addressOrSite

    # Authenticate with Exchange Online
    if (!$ExoSession) {
        $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
        Import-PSSession $ExoSession -AllowClobber -DisableNameChecking
    }

    $folderQueries = @()
    $folderStatistics = Get-MailboxFolderStatistics $emailAddress
    foreach ($folderStatistic in $folderStatistics) {
        $folderId = $folderStatistic.FolderId;
        $folderPath = $folderStatistic.FolderPath;

        $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler = $encoding.GetBytes("0123456789ABCDEF");
        $folderIdBytes = [Convert]::FromBase64String($folderId);
        $indexIdBytes = New-Object byte[] 48;
        $indexIdIdx = 0;
        $folderIdBytes | select -skip 23 -First 24 | % { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
        $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

        $folderStat = New-Object PSObject
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery

        $folderQueries += $folderStat
    }
    Write-Host "-----Exchange Folders-----"
    $folderQueries | ft
}