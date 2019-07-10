$templatePath = 'F:\Automation\JetstressConfig.xml'
$dbmaps=Import-Csv 'f:\automation\LABDAG_Servers.csv'
$jsconfig= New-Object -Typename XML
$jsconfig.Load($templatePath)


foreach($dbmap in $dbmaps){

	$dbs=$dbmap.dbmap.split(",") 
    $dbcount = $dbs.Count
	write-host "Server $($dbmap.servername)"
	write-host "DB list $dbs"
    [int]$i=0

	foreach($db in $dbs){

            $dbPath = $dbmap.DatabasesRootFolder + $db + "\" + $db + ".db"
            $logPath = $dbmap.DatabasesRootFolder + $db + "\" + $db + ".log"
            write-host "DB $db"
            $oldDbPath=$jsconfig.SelectNodes("//Path")[$i]
            $oldDbPath.InnerText = $dbPath

            $oldLogPath=$jsconfig.SelectNodes("//LogPath")[$i]
            $oldLogPath.InnerText = $logPath

            $i++
		}
	$jsconfig.save("f:\automation\"+$dbmap.servername+".xml")
}


