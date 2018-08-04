#############################################################################
#																			#
# name: wsus.ps1															#
# script version 0.9														#
# 																			#	
# author: eduard-albert.prinz@univie.ac.at									#
#																			#
# comment: 	check wuau self update status and correct gpo values, 			#
#  			download and install from the windows server update service 	#
#																			#
# 06.11.2012 :: prototype													#
# 23.01.2013 :: alpha														#
# 28.01.2013 :: beta														#
# 29.01.2013 :: test in combination with matrix42 workplace v15				#
# 30.01.2013 :: rc1															#
# 01.02.2013 :: rc2															#
#																			#
#############################################################################


$erroractionpreference = "stop"
# ---------------------------------------------------------------------------
#to customize
# ---------------------------------------------------------------------------
$program = $myinvocation.mycommand.name											#if you only want the filename (not the full path)
$nl = [environment]::newline													#gets the newline string defined for this environment
$versionstring = ""																#global wuau versionstring
[int]$waitselfupdate = 60 														#wait for selfupdate to finish
[int]$checkpolicytries = 2														#how often check correct policy values
[int]$recursiveloopstodo = 10													#how often loops recursive script
[int]$recursivecounter = 0														#counter for recursive loop
[int]$runrecursiveselfupdate = 10												#how often loops recursive self updates
$runrecursiveloops = $false 													#exit recursive run of script
$forcepolicyupdate = $true 														#forcing a policy update
$regwususname = "wuserver"														#registry name for the wsus server
$usewuserver = "usewuserver"
$selfupdatestring = "service restarted after self update" 						#selfupdatestring in windowsupdatelog
$wuauclientpath = "$env:windir\system32\wuauclt.exe"							#path to windows update autoupdate client for version check
$windowsupdatelog = "$env:windir\windowsupdate.log"								#path to ms windowsupdate logs
$regwususkey = "hklm:\software\policies\microsoft\windows\windowsupdate"		#registry path for windowsupdate
$lastruntime = ""																#each script execution on writes a date to the registry
$patchday = ""																	#the wsus administrator sets the date for the execution of this script
[int]$taskscheduler = 0															#flaq executed by the taskscheduler
[int]$retriescounter = 5														#how often loops the resultcheck
[int]$debug = 0																	#flaq debug

$logerrortype = "error"
$logselfupdatetype = "selfupdate"
$logwuautype = "wuau"
$logdir = "$env:windir\sysadmin\log-wsus\"
if (!$(test-path $logdir)) {new-item $logdir -itemtype directory | out-null}
$selfupdatelog = [string]::join('',($logdir, "wuauselfupdate.log"))
$wuaulog = [string]::join('',($logdir, "wuau.log"))
$errorlog = [string]::join('',($logdir, "wuauerror.log"))
if (!$(test-path $selfupdatelog)) {new-item -type file $selfupdatelog -force}
if (!$(test-path $wuaulog)) {new-item -type file $wuaulog -force}
if (!$(test-path $errorlog)) {new-item -type file $errorlog -force}

function showarguments {
  write-host @"
	$nl
	path to logs: $logdir	
	task ... executed by the taskscheduler
	debug ... with shell output
	help ... show help and arguments
	$nl
"@
  exit(0)
}

$args | %{ 
	switch ($_) {
		help { showarguments }
		task {[int]$taskscheduler = 1}
		debug {[int]$debug = 1}
		default {[string]$matrixpackage = $_}
	}
}

function writelog($level, $msg, $log, $divider)
{
	switch ($level) {
		selfupdate {$filename = $selfupdatelog}
		wuau {$filename = $wuaulog} 
		default {$filename = $errorlog}
	}	
	if($divider){ $logrecord = "------------------------------------------------------------$nl $((get-date).tostring())" }
	$logrecord += $msg
	if($debug){ write-host $logrecord }
	if($log){ add-content -path $filename -encoding utf8 -value $logrecord }
}

function disablewsus{
	set-itemproperty -path $regwususkey -name ElevateNonAdmins -value 1 -type dword
	set-itemproperty -path $regwususkey"\au" -name AUOptions -value 2 -type dword
	set-itemproperty -path $regwususkey"\au" -name UseWUServer -value 0	-type dword
	set-itemproperty -path $regwususkey -name DisableWindowsUpdateAccess -value 0 -type dword	
	
	$logmsg =  " change regvalues"
	writelog $logwuautype $logmsg 1 1
	set-itemproperty -path "registry::hklm\system\currentcontrolset\services\wuauserv" -name "delayedautostart" -value 1 -type dword
	restart-service wuauserv
	$startype = get-wmiobject win32_service | where {$_.name -eq "wuauserv"} | select -expand startmode
	$logmsg = " wuauserv starttype: $startype"
	writelog $logwuautype $logmsg 1 1
	if ($startype -eq "manual") {
		set-service wuauserv -startuptype automatic 
		$logmsg =  " wuauserv starttype changed to: $(get-wmiobject win32_service | where {$_.name -eq "wuauserv"} | select -expand startmode)"
		writelog $logwuautype $logmsg 1 1
	}
	$logmsg = " now check directly from microsoft landscape for updates and quitting script!"			
	writelog $logerrortype $logmsg 1 1
}

function checkwsusstatus{
	try{if($env:notebook -eq 1){$nb=1;}else{$nb=0;}}
	catch{$nb=0}
	if($nb){
		$logmsg =  " it is a notebook - check wsus status"
		writelog $logwuautype $logmsg 1 1
		try{ $wsussrv = (get-itemproperty -path $regwususkey).$regwususname }
		catch{
			$logmsg = " could not resolve path to $regwususkey \ $regwususname on this system, quitting script!"			
			writelog $logerrortype $logmsg 1 1
			setlastruntime
			setlastrunfailed(1) 
			exit(1)
		}			
		$on = $true
        $check = [system.net.httpwebrequest]::create($wsussrv)
        try{ $response = $check.getresponse() }
        catch{ $on = $false } 		
        if ($on){
			$response.close()
			$logmsg =  " wsus statuscode 200 - script running!"
			writelog $logwuautype $logmsg 1 1
		}
		else{
			$logmsg = " could not reach $wsussrv, disablewsus settings!"			
			writelog $logerrortype $logmsg 1 1
			disablewsus
			exit(1)
		}		
	}
	else{
		$logmsg =  " it is not a notebook go on as usual"
		writelog $logwuautype $logmsg 1 1
	}
}

function checkregconditions($regvalue)
{	
	try{ $value = (get-itemproperty -path $regwususkey).$regvalue }
	catch{ 
		$logmsg = " missing value on $regwususkey \ $regvalue on this system - function checkregconditions!"			
		writelog $logerrortype $logmsg 1 1
		#setlastrunfailed(1) 
	}
	
	if (($value -eq $null) -or ($value.length -eq 0)){
		$logmsg =  " couldn't find $regwususkey $regvalue on this system or $regvalue is not set - function checkregconditions!"			
		writelog $logerrortype $logmsg 1 1
		return $value
		#setlastrunfailed(1) 
	}
	else{ return $value	}	
}

function getfileversion($filepath)
{
	return $fileversion = [system.diagnostics.fileversioninfo]::getversioninfo($filepath).productversion
}

function changefileext($filename, $newext)
{
	return $newfilename = [system.io.path]::changeextension($filename, $newext)
}

function checkversion
{
	$versionstring = checkregconditions("versionstring")	
	$updateinfo = getfileversion $wuauclientpath 

	if($updateinfo.compareto($versionstring) -eq 0)
	{
		$logmsg = "  wua version $updateinfo matches with the referenced version $versionstring on $wsussrv, windows update agent is up to date."
		writelog $logselfupdatetype $logmsg 1 1
		return $true
	}
	elseif ($updateinfo.compareto($versionstring) -eq 1)
	{
		$logmsg = "  installed wua version $updateinfo is newer than the referenced version $versionstring on $wsussrv "
		writelog $logselfupdatetype $logmsg 1 1
		return $true
	}	
	else
	{
		$logmsg = "  wua version $updateinfo does not match with version $versionstring on $wsussrv, wua selfupdate required."
		writelog $logselfupdatetype $logmsg 1 1
		return $false
	}	
}

function setlastrunfailed($lastrunfailed)
{	
	new-itemproperty -path $regwususkey -name "lastrunfailed" –force
	set-itemproperty -path $regwususkey -name "lastrunfailed" -value $lastrunfailed
}

function setlastruntime
{	
	new-itemproperty -path $regwususkey -name "lastruntime" –force
	$d=new-object system.globalization.cultureinfo("de-at");
	$f=$d.datetimeformat.shortdatepattern;
	$lastruntime=get-date -format $f;
	set-itemproperty -path $regwususkey -name "lastruntime" -value $lastruntime	
}

function checklastrunfailed
{
	$lastrunfailed = checkregconditions("lastrunfailed")
	if($lastrunfailed -eq 1){ return $true }
	else{ return $false }
}

function finishedselfupdate
{
	$versionstring = checkregconditions("versionstring")
	$logmsg = "  selfupdate to version $versionstring finished successfully"
	writelog $logselfupdatetype $logmsg  1 1
	stop-service wuauserv -force 
	$logmsg = "  stopped wuauserv after selfupdate"
	writelog $logselfupdatetype $logmsg 1 1  
	start-sleep -s $waitselfupdate
}

#script begin
#############################################################################

$logmsg = " script begin"
writelog $logwuautype $logmsg 1 1
$updateinfo = getfileversion $wuauclientpath 
$logmsg =  " current wua version $updateinfo "
writelog $logselfupdatetype $logmsg 1 1
checkwsusstatus
$startype = get-wmiobject win32_service | where {$_.name -eq "wuauserv"} | select -expand startmode
$logmsg = " wuauserv starttype: $startype"
writelog $logwuautype $logmsg 1 1
if ($startype -ne "manual") {
	set-service wuauserv -startuptype manual
	$logmsg =  " wuauserv starttype changed to: $(get-wmiobject win32_service | where {$_.name -eq "wuauserv"} | select -expand startmode)"
	writelog $logwuautype $logmsg 1 1
}

if(($matrixpackage.length -ne 0) -and ($matrixpackage -ne "task"))
{
	$logmsg = " executed by $matrixpackage $nl"
	writelog $logwuautype $logmsg 1 1
}

if (!(test-path $regwususkey))
{
	$logmsg = " couldn't resolve path to $regwususkey on this system, quitting script!"			
	writelog $logerrortype $logmsg 1 1
	setlastruntime
	setlastrunfailed(1) 
	exit(1)
}	

if($taskscheduler)
{
	$lastruntime  = checkregconditions("lastruntime")
	$patchday  = checkregconditions("patchday")
	$logmsg = " script has been executed by the taskscheduler"
	writelog $logwuautype $logmsg 1 1	
	if($lastruntime.length -gt 0)
	{
		$patchday = [datetime]::ParseExact($patchday, "dd.MM.yyyy", $null)	
		$lastruntime = [datetime]::ParseExact($lastruntime, "dd.MM.yyyy", $null)
		
		if($patchday -eq $lastruntime) 
		{
			if(checklastrunfailed)
			{
				$logmsg = " patchday: $($patchday.toshortdatestring()) and lastruntime: $($lastruntime.toshortdatestring()) same-day but with error, script is running!"
				writelog $logwuautype $logmsg 1 1
			}else{
				$logmsg = " patchday: $($patchday.toshortdatestring()) and lastruntime: $($lastruntime.toshortdatestring()) same-day and no error, quitting script!"
				writelog $logwuautype $logmsg 1 1
				exit(1)
			}
		}
		elseif($lastruntime -lt $patchday) 
		{
			$logmsg = " lastruntime: $($lastruntime.toshortdatestring()) before the patchday: $($patchday.toshortdatestring()), script is running!"
			writelog $logwuautype $logmsg 1 1		
		}
		elseif($lastruntime -gt $patchday) 
		{			
			if(checklastrunfailed)
			{
				$logmsg = " lastruntime: $($lastruntime.toshortdatestring()) greater than patchday: $($patchday.toshortdatestring()) but with error, script is running!"
				writelog $logwuautype $logmsg 1 1
			}else{
				$logmsg = " lastruntime: $($lastruntime.toshortdatestring()) greater than patchday: $($patchday.toshortdatestring()) and no error, quitting script!"
				writelog $logwuautype $logmsg 1 1
				exit(1)
			}
		}
	}
	elseif($lastruntime.length -eq 0){
		$logmsg = " lastruntime is not set, however yet executed script!"
		writelog $logwuautype $logmsg 1 1	
	}
}
$policycheckcounter = 0	
do{
	try{ $wsussrv = (get-itemproperty -path $regwususkey).$regwususname }
	catch{
		$logmsg = " couldn't resolve path to $regwususkey \ $regwususname on this system, quitting script!"			
		writelog $logerrortype $logmsg 1 1
		setlastruntime
		setlastrunfailed(1) 
		exit(1)
	}	
	if ($wsussrv -ne $null){ break }
	else{
		$logmsg = "------------------------------------------------------------"		
		writelog $logerrortype $logmsg 1 1
		$logmsg = "  missing or wrong registry key in $regwususkey\wuserver: $polcheck"
		writelog $logerrortype $logmsg 1 1
		if ($forcepolicyupdate)
		{
			$logmsg = " forcing a policy update!!!"
			writelog $logerrortype $logmsg 1 1
			start-process "gpupdate.exe" -argumentlist "/force" -wait -windowstyle hidden		   
			start-sleep -s 5
			$policycheckcounter++
			if ($policycheckcounter -ge $checkpolicytries) 
			{
				$logmsg = " quitting script after second unsuccessfull policy enforcement! no updates have been done!"
				writelog $logerrortype $logmsg 1 1
				setlastruntime
				setlastrunfailed(1) 
				exit(1)
			}
		}
		else
		{	
			setlastruntime
			setlastrunfailed(0) 
		}            
	}	
}
until ($policycheckcounter -ge $checkpolicytries)

if (!(test-path "$wuauclientpath"))
{
	$logmsg = " couldn't resolve path to $wuauclientpath on this system, quitting script!"			
	writelog $logerrortype $logmsg 1 1
	setlastruntime
	setlastrunfailed(1) 
	exit(1)
}
stop-service wuauserv -force
do{
if((get-itemproperty -path $regwususkey"\au").$usewuserver){ 
    $httpcount = 0
    do{
        $online = $true
        $httpcheck = [system.net.httpwebrequest]::create($wsussrv)
        try{ $response = $httpcheck.getresponse() }
        catch{ $online = $false } 		
        if ($online){$response.close();break}
        $httpcount++
        start-sleep -s 5
    }	
    until ($httpcount -ge 4) 
	if (!$online){
        $logmsg = " could not connect to $($wsussrv.substring(7)) after $httpcount retries, a connection problem might exist!"		
		writelog $logerrortype $logmsg 1 1
		setlastruntime
        setlastrunfailed(1) 
		exit(1)
    }
	if(!(checkversion)){
		$logmsg = "  stopped wuauserv before selfupdate"
		writelog $logselfupdatetype $logmsg 1 1		
		$testupdatelog = get-content($windowsupdatelog) | select-string $selfupdatestring -quiet 
		if ($testupdatelog)
		{		
			if (test-path "$(split-path $windowsupdatelog -parent)\$(changefileext(split-path $windowsupdatelog -leaf) 'bak')"){remove-item "$(split-path $windowsupdatelog -parent)\$(changefileext(split-path $windowsupdatelog -leaf) 'bak')"}
			rename-item $windowsupdatelog $(changefileext (split-path $windowsupdatelog -leaf) "bak")
		}		
		if($recursivecounter -eq 0){
			$logmsg = "  executing: wuauclt.exe detectnow"
			writelog $logselfupdatetype $logmsg 1 1
			start-process $wuauclientpath -argumentlist "/detectnow" -wait -windowstyle hidden
			$logmsg = "  the local wuau version is checked as long as they match up with the server version!"
			writelog $logselfupdatetype $logmsg 1 1	
		}
		$count = 0		
		do{	
			$count++
			$logmsg = " $count of $runrecursiveselfupdate retry check local wuau version and compare client version with server version"
			writelog $logselfupdatetype $logmsg 1 1
			$testupdatelog = get-content($windowsupdatelog) | select-string $selfupdatestring -quiet 
			start-sleep -s $waitselfupdate				
		}
		until((!($testupdatelog -eq $null) -and (checkversion)) -or ($count -gt $runrecursiveselfupdate))		
		$testupdatelog = get-content($windowsupdatelog) | select-string $selfupdatestring -quiet 
		if (($testupdatelog -eq $null) -and !(checkversion))
		{
			$logmsg = "  final selfupdate check failed, exited script!"
			writelog $logselfupdatetype $logmsg 1 1  
			stop-service wuauserv -force 
			$logmsg = "  stopped wuauserv"
			writelog $logselfupdatetype $logmsg 1 1    
			setlastruntime
			setlastrunfailed(1)
			exit(1)
		}
		else{ finishedselfupdate }
	}
}
	$logmsg = " starting windows update ($($program)) $nl"
	writelog $logwuautype $logmsg 1 1
	$updatesession = new-object -comobject "microsoft.update.session" 
	$updatesearcher = $updatesession.createupdatesearcher()
	$logmsg = " searching for updates...$nl" 
	writelog $logwuautype $logmsg 1 1
	$resultcounter = 0		
	do{		
		$searchresult = $updatesearcher.search("isinstalled = 0 and type = 'software'")
		if ($searchresult -is [object]){break}
			$resultcounter + 1
			$logmsg = "  result loop: $resultcounter"
			writelog $logerrortype $logmsg 1 1
			start-sleep -s 5
	}
	until ($resultcounter -gt $retriescounter)			
	$logmsg = " list of applicable items on the machine: $nl" 
	writelog $logwuautype $logmsg 0 1
	for ($i = 0; $i -le $($searchresult.updates.count)-1; $i++)
	{
		$update = $searchresult.updates.item($i)
		$logmsg = "  $($i+1) > $($update.title)" 
		writelog $logwuautype $logmsg 0 1
	}

	$newupdates = $false 
	if ($($searchresult.updates.count) -gt 0)
	{   
		$newupdates = $true  
		$logmsg = " creating collection of updates to download: $nl"
		writelog $logwuautype $logmsg 0 1
		$updatestodownload = new-object -comobject "microsoft.update.updatecoll"
		for ($i = 0; $i -le $($searchresult.updates.count)-1; $i++)
		{
			$update = $searchresult.updates.item($i)
			$logmsg =  "  $($i+1) > adding: $($update.title)"
			writelog $logwuautype $logmsg 0 0
			$updatestodownload.add($update) | out-null
		}
		$logmsg = " wsus: downloading updates..."
		writelog $logwuautype $logmsg 1 1							
		$downloader = $updatesession.createupdatedownloader()
		$downloader.updates = $updatestodownload
		$downloader.download() | out-null
		$logmsg =  " list of downloaded updates: $nl"
		writelog $logwuautype $logmsg 0 1
		for ($i = 0; $i -le $($searchresult.updates.count)-1; $i++)
		{
			$update = $searchresult.updates.item($i)
			if ($update.isdownloaded -eq $true)
			{
				$logmsg =  "  $($i+1) > $($update.title)"
				writelog $logwuautype $logmsg 0 0
			} 
		}
		$updatestoinstall = new-object -comobject "microsoft.update.updatecoll"
		$logmsg =  " creating collection of downloaded updates to install: $nl"
		writelog $logwuautype $logmsg 0 1
		for ($i = 0; $i -le $($searchresult.updates.count)-1; $i++)
		{
			$update = $searchresult.updates.item($i)
			if ($update.isdownloaded)
			{   $logmsg =  "  $($i+1) > adding: $($update.title)"
				writelog $logwuautype $logmsg 0 0
				$updatestoinstall.add($update) | out-null
			}
		}

		if ($($updatestoinstall.count) -gt 0)
		{
			$logmsg = "  wsus: installing updates... $nl"
			writelog $logwuautype $logmsg 1 1						
			$installer = $updatesession.createupdateinstaller()
			$installer.updates = $updatestoinstall
			$installationresult = $installer.install()			
			$logmsg = "  listing of updates installed and individual installation results: $nl"
			writelog $logwuautype $logmsg 0 1
			$check = 0
			$logmsg = " updatestoinstall: $($updatestoinstall.count) $nl"
			writelog $logwuautype $logmsg 1 1
			for ($i = 0; $i -le $($updatestoinstall.count)-1; $i++)
			{
				$resultcode = $installationresult.getupdateresult($i).resultcode
				if($resultcode -eq 4 -and $check -eq 0){ setlastrunfailed(1); $check = 1 }
				$logmsg = "  $($i+1) >  $($updatestoinstall.item($i).title): $($resultcode)"
				writelog $logwuautype $logmsg 1 0
			}			
			$logmsg = " program finished installing $i patches: reboot required: $($installationresult.rebootrequired) ! $nl"
			writelog $logwuautype $logmsg 1	1		   
			if ($($installationresult.rebootrequired) -eq "true"){ writelog $logwuautype " need reboot yes $nl" 1 1	}
			setlastruntime
			setlastrunfailed(0) 		
		}	
		else
		{
			$logmsg = "  no updates have been downloaded or are ready for install. $nl"
			writelog $logwuautype $logmsg 1	1
			setlastruntime
			setlastrunfailed(0) 					
		}
	}
	else
	{
		$logmsg = "  wsus: there are no applicable updates for this system. $nl"
		writelog $logwuautype $logmsg 1 1
		setlastruntime
		setlastrunfailed(0) 		
	}
	$recursivecounter++
	if (!$runrecursiveloops){break} 
}
until (!($newupdates) -or ($recursivecounter -ge $recursiveloopstodo))
stop-service wuauserv -force