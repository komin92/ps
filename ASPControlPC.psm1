#############################################################################
#																			#
# name: aspControlPC.psm1													#
# script version 1.01														#
# 																			#
# author: eduard-albert.prinz												#
#																			#
# comment: 	asp powershell module to control one or more client(s)			#
#			This program is written in the hope that it will be useful, 	#
#			but WITHOUT ANY WARRANTY ^^										#
#																			#
# 16.01.2013 :: pre-alpha													#
# 07.02.2013 :: alpha 														#
# 22.02.2013 :: beta for groups asp											#
# 22.02.2013 :: release shell for groups asp								#
# 25.07.2013 :: prerelease gui for groups asp								#
# 12.08.2013 :: release gui for groups sd, ps                               #
# 03.02.2015 :: prerelease for specific deploymentmanager (experiential)    #
# 04.08.2015 :: release for all deploymentmanager 							#							
#		                                                                    #						
#############################################################################

$erroractionpreference = 'silentlycontinue'

# ---------------------------------------------------------------------------
# to customize
# ---------------------------------------------------------------------------
$global:logdir = "$env:utils"
$global:file = "aspControlPC.log"
$global:logfile = [string]::join('',($global:logdir, $global:file))
$clientfile = "$env:path2pclist"
$readmefile = "\\$env:server\pub_utils$\readme_aspcontrolpc.txt"
$nl = [environment]::newline
$global:lvresult = New-Object System.Windows.Forms.DataGridView 
$global:customcsv=0
$global:deploymanager=0;
$xmlfile = "aspcontrolpc.options.xml"
$icon = "gui.ico"
$script:parentfolder = split-path (get-variable myinvocation -scope 1 -valueonly).mycommand.definition
$xmlfile = join-path $parentfolder $xmlfile
$icon = join-path $parentfolder $icon
[xml]$script:xml = get-content $xmlfile
$global:xml=$xml
$regaspcmkey = "hkcu:\software\policies\asp"
$global:dlanguage = "en"
$domain = "dc1.ad.domain"
# ---------------------------------------------------------------------------
# to customize
# ---------------------------------------------------------------------------

register-engineevent powershell.exiting –action { handleexit }	

# ---------------------------------------------------------------------------
# check deployment manager pc´s
# ---------------------------------------------------------------------------
function checkdeploymentmanager{
	$global:deploymanager=1;
	dontquit;
	if ((get-module) -notcontains "psterminalservices"){ import-module psterminalservices }
	$session=get-tscurrentsession
	$session|%{$username=$_.username}	
	#$username=$env:username
	$check = $true
	$loader="."
	$status="please wait.."
		while($check){
		dontquit;
		$status+=$loader
			write-progress -activity "check user credentials" -status $status	
			if ((get-module) -notcontains "activedirectory"){ import-module activedirectory }
			try{
				$groups = get-adgroup -server $domain:3268 -filter 'Name -like"*Deploymanager"' | select-object name | ? {$_.name -notlike 'DL_DeployManager'}
				$result=@();$groups|%{if(get-adgroupmember -server $domain -identity $_.name -recursive  | ? {$_.name -eq $username}){$result+=$_.name}}				
				if($result.length -gt 0){
					dontquit;										
					$b=@();$c=@();[string]$regex = " ";
					$result|%{$b+=$_ -split("_");$c+=$b[2];$b=$null;}
					$regex="(";
					#$regex+=$result
					$count=$null;
					$c|%{$regex+=$_;if($c.count-1 -ne $count -and $c.count -gt 1){$regex+="|";}$count++;}
					$regex+=")";
					aspgetclientlist($regex);					
				}
			}catch{
				$errormessage = $_.exception.message;write-host "an error has occurred - $errormessage";
			}
			finally{$check = $false;write-progress -activity "check user credentials" -status "successfully" -completed;dontquit;}			
			dontquit;
		}
}
# ---------------------------------------------------------------------------
# check deployment manager pc´s
# ---------------------------------------------------------------------------
function test-registrykeyvalue{
     [cmdletbinding()]
    param(
        [parameter(mandatory=$true)]
        [string]
        $path,
        [parameter(mandatory=$true)]
        [string]
        $name
    )

    if( -not (test-path -path $path -pathtype container) ){return $false}
    $properties = get-itemproperty -path $path 
    if( -not $properties ){return $false}
    $member = get-member -inputobject $properties -name $name
    if( $member ){return $true}
    else{return $false}
}
function dontquit{
	[console]::treatcontrolcasinput = $true
	
		if ([console]::keyavailable) {
			$key = [system.console]::readkey($true)
			if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "c")){write-warning "this function is blocked !";}
		}	
}
function resize-datagrid{
	[reflection.assembly]::loadwithpartialname("system.windows.forms") | out-null
	#$global:lvresult.autoresizecolumns([system.windows.forms.datagridviewautosizecolumnsmode.allcells]::allcells)
	#$datagridviewautosizecolumnsmode.allcells
	$global:lvresult.autoresizecolumnheadersheight();
}
function refreshgui{if($formmain){ resize-datagrid;$formmain.refresh()}}

#$servicestatus = (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").status
  #if ($servicestatus -eq "stopped") { (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").startservice() }

function checkquit{
	[console]::treatcontrolcasinput = $true
	
		if ([console]::keyavailable) {
			$key = [system.console]::readkey($true)
			if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "c")){write-warning "script was aborted !";break;}
		}	
}

function aspwriteevent{
	[cmdletbinding()]
	param (
			$eventmessage = '',
			$eventtype = 'information',
			$eventid = '9999'
	)			
	if(-not([system.diagnostics.eventlog]::sourceexists('asp'))){[system.diagnostics.eventlog]::createeventsource('asp','application')}
	$eventlog = new-object system.diagnostics.eventlog('application');$eventlog.source = 'asp';$eventlog.writeentry($eventmessage,$eventtype,$eventid)
}

function aspshowevent{get-eventlog -log application  | where {$_.eventID -eq 9999}}

function aspwritetext{	
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			$text = @(""),
			[consolecolor[]]$color = ("white")	
	)
	if($text.count -le 1){ write-host $text -f $color -nonewline }
	elseif($color.count -le 1){
		for ($i = 0; $i -lt $text.count; $i++) {write-host $text[$i] -f $color -nonewline}
	}
	else{
		for ($i = 0; $i -lt $text.count; $i++) {write-host $text[$i] -f $color[$i] -nonewline}
	}
}

function setcolumn{
[cmdletbinding()]
		param([parameter(mandatory=$true,valuefromremainingarguments=$true)]
			[int]$columncount = 0,
			[string[]]$columns = $null
		)
	if($global:gui){
		$global:lvresult.columncount=$columncount;	
		for($i=0;$i -lt $columncount;$i++){
			$global:lvresult.columns[$i].name = $columns[$i];
		}
		refreshgui;		
	}
}

function aspgethelp{
	if(!$(test-path $readmefile)){ $readmefile = "\\$env:server\utils$\readme_aspcontrolpc.txt" }
	try { test-path $readmefile }
	catch {$readmefile = "\\$env:empmasterserver\readme$\readme_aspcontrolpc.txt"}
	try { invoke-item $readmefile }
	catch { aspwritetext "could not access $readmefile" red }
}

function aspwritelog{
	[cmdletbinding()]
	param (
			[string]$msg = "",
			[validaterange(0,1)] 
			[int]$debugview = 0			
	)
	if(!$(test-path $global:logdir)) {new-item $global:logdir -itemtype directory | out-null}	
	if(!$(test-path $global:logfile)) {new-item -type file $global:logfile -force}
	$logrecord = "------------------------------------------------------------$nl $((get-date).tostring())" 
	$logrecord += " [$env:username]: $msg";if($debugview){ aspwritetext $logrecord };aspwriteevent " [$env:username]: $msg" 'information' '9999';add-content -path $global:logfile -encoding utf8 -value $logrecord
}

function aspshowlog{
	[int]$now = 0
	[int]$count = 0
	$args | %{
		switch ($_) { -now { [int]$now = 1 } }
	}
	if(!$(test-path $global:logdir)) {new-item $global:logdir -itemtype directory | out-null}	
	if(!$(test-path $global:logfile)) {new-item -type file $global:logfile -force}	
	[string]$date = get-date -uformat "%m%/%d%/%Y";if($now){ $logs = get-content $global:logfile | select-string -pattern $date; $count = $logs.count; $logs; aspwritetext "$nl logs: $count $nl" cyan }
	else{ $logs = get-content $global:logfile; $count = $logs.count;  $logs | %{ if(!($_ -match "---")){ aspwritetext "$nl $_ " }else{ $count-- } }; write-host $nl; aspwritetext " logs: $count " cyan; write-host $nl ; aspwritetext " logsfile:  $global:logfile " cyan; write-host $nl }
}

function asptestconnectivity{	
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = "",
			[validaterange(0,1)] 
			[int]$debugview = 0,
			[validaterange(0,1)] 
			[int]$log = 0,
			[validaterange(0,1)] 
			[int]$show = 1,
			[int]$loops = 0,
			[int]$returnvalue = 1
	)
	if($global:gui -and ($global:lvresult.columncount -eq 0)){
		setcolumn -columncount 2 -columns "computername","result";
	}	
	if($loops -eq 0){$loops = 1}	
	for($i=0;$i -lt $loops;$i++){
		checkquit; if($global:gui){refreshgui;}
		if($debugview){ aspwritetext " try connecting to computer: $computername $nl" yellow  }	
		  try{
			$ping = new-object system.net.networkinformation.ping
			$pingreturns = $ping.send($computername, 1000) 
			if($pingreturns.status -eq "success"){			
				$logmsg = " $computername up"
				if($show -eq 1){ aspwritetext " $computername"," up $nl"cyan, green;if($global:gui){$global:lvresult.rows.add("$computername", "up");refreshgui;}}
				if($log){ aspwritelog $logmsg }; return "yes"	#dirty hack				
			}
			else{
				$logmsg = " problem connecting to computer: $computername."
				aspwritetext " $computername ","down $nl" cyan, red;if($global:gui){$global:lvresult.rows.add("$computername", "down" );refreshgui;}
				if($log){ aspwritelog $logmsg }; return "no" 	#dirty hack	
			}
			if($loops -gt 1){delay(1)}
			}catch{ 
				$errormessage = $_.exception.message				
				$logmsg = " ping failed for $computername the error message was $errormessage $nl"
				if($log){ aspwritelog $logmsg }
				aspwritetext $logmsg red
				if($global:gui){$global:lvresult.rows.add("$computername",$logmsg);refreshgui;};if($log){ aspwritelog $logmsg };  return "no" 	#dirty hack	
			}
				checkquit; if($global:gui){refreshgui;}
		}
}

function aspsendmsg{
	[cmdletbinding()]
	param(
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = "",
			[string]$message="please log off.",	
			[validaterange(0,1)] 
			[int]$debugview = 0,
			[validaterange(0,1)] 
			[int]$log = 0,
			[int]$wait="",		
			[string]$session="*",
			[int]$seconds=55,
			[validaterange(0,1)] 
			[int]$show = 1				
	)
	checkquit; if($global:gui){refreshgui;}	
	if($debugview){aspwritetext " $nl sending the following message with a $seconds second delay: $message $nl" yellow;}	
	$command = "msg.exe $session /time:$($seconds)";
	if($computername){ $command += " /server:$($computername)"; }
	if($debugview){ $command += " /v"; }
	if($wait){ $command += " /w"; }
	$command += " $($message)";
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","result";
		invoke-expression $command; 
		if($lastexitcode -eq 0){
			if($show -eq 1){aspwritetext " $computername message '$message' sent $nl"  green;if($global:gui){$global:lvresult.rows.add("$computername", "message '$message' sent")}}
		}
		else{
			$logmsg =  " there was no message sent to $computername."
			if($show -eq 1){aspwritetext " $logmsg $nl " red;if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)}};if($log){ aspwritelog $logmsg }			
		}	
	}
	else{		
		$logmsg =  " there was no message sent to $computername."
		if($show -eq 1){aspwritetext " $logmsg $nl " red;if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)}};if($log){ aspwritelog $logmsg }			
	}	
	checkquit; if($global:gui){refreshgui;}
}



function aspgetlastlogon{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){		
		setcolumn -columncount 2 -columns "computername","username";
		$computer = $computername
		$adminpath = test-path \\$computer\admin$
		if($adminpath -eq "true")
		{
			$key = "software\microsoft\windows\currentversion\authentication\logonui"
			$type = [microsoft.win32.registryhive]::localmachine
			try{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer,[microsoft.win32.registryview]::registry64) }
			catch{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer) }
			$logon = $regkey.opensubkey($key)
			$username = $logon.getvalue("lastloggedonuser")
			if($username -ne ""){
				aspwritetext " $computer :" cyan;aspwritetext " $username $nl"  green  
				if($global:gui){$global:lvresult.rows.add("$computername", " $username")}
				$regkey.close();$logon.close()
			}
		}
		else{$logmsg = " can not access $computer $nl";if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetinstalledsw{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","software";		
		checkquit; if($global:gui){refreshgui;}	
	try{
		aspwritetext " $computername : $nl" cyan;
		$global:swlist =$null;		
			$global:swlist = invoke-command -scriptblock {
				param (
					[string] $hostname,
					[string] $filter
				)
				[int]$count = 0;
				$global:swlist = get-childitem "hklm:\software\empirum\asp team\*\*" -erroraction silentlycontinue | where-object {$_ -match $filter}
				if (!$global:swlist) {write-host " no installed sw found!" -f red;exit}
				else{
					write-host " total software count: $($global:swlist.count)" -f green;			
						foreach ($sw in $global:swlist)
						{
							$count++;
							write-host ("{0:d2}" -f $count)" $($sw -replace '^.*\\.*\\(.*\\.*\d.*)$', '$1')" -f yellow;		
						}
						return $global:swlist;
					}
				
			} -computername $computername -args $computername,$filter
			if ($global:swlist.count -gt 0){
				if($global:gui){			
					foreach ($sw in $global:swlist){$global:lvresult.rows.add("$computername", " $($sw -replace '^.*\\.*\\(.*\\.*\d.*)$', '$1')" );}				
				}
			}else{
				if($global:gui){$global:lvresult.rows.add("$computername"," no installed sw found!");}
			}
		}catch{$logmsg = " can not access $computer $nl";if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetwsusfailcount{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","wsusfailcount";		
		checkquit; if($global:gui){refreshgui;}
		$computer = $computername
		$adminpath = test-path \\$computer\admin$
		if($adminpath -eq "true")
		{
			$key = "software\asp-zid\inv"
			$type = [microsoft.win32.registryhive]::localmachine
			try{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer,[microsoft.win32.registryview]::registry64) }
			catch{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer) }
			$logon = $regkey.opensubkey($key)
			$wsusfailcount = $logon.getvalue("wsusfailcount")
			if($wsusfailcount -ne ""){
				aspwritetext " $computer :" cyan;aspwritetext " $wsusfailcount $nl"  green  
				if($global:gui){$global:lvresult.rows.add("$computername", " $wsusfailcount")}
				$regkey.close();$logon.close()
			}
		}
		else{$logmsg = " can not access $computer $nl";if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspshowprofiles{
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""
		)
	$drive = ""
	checkquit; if($global:gui){refreshgui;}		
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){	
		setcolumn -columncount 6 -columns "computername","accountname","lastaccesstime","size MB","count file","count dir";
		#$p =  get-item env:\userprofile
		#invoke-command -computername $computername {[environment]::getenvironmentvariable(“temp”,”machine”)}
		try{ gwmi win32_userprofile -computername $computername -erroraction $erroractionpreference | % { if($_.localpath -notmatch "service" -and $_.localpath -notmatch "system32"){ $path = $_.localpath } } }
		catch {aspwritetext " can not access $computername" red;if($global:gui){$global:lvresult.rows.add("$computername", " can not access ");}}
		if($path -ne "" -and $path.length -gt 1){ $drive = $path.split(":") }
		if($drive.count -gt 0){ $drive = $drive[0] }	
		$userpath = "$computername\"+$drive+"$\users";$userprofile = test-path \\$userpath;$profiles = "";$i=0
		try{ if($userprofile -eq "true"){ $profiles = get-item "\\$userpath\*" } }
		catch {aspwritetext " can not access $computername $nl" red;if($global:gui){$global:lvresult.rows.add("$computername", " can not access ");}}
		$count=0
		if($profiles -ne "" -and $profiles.length -gt 0 -and $profiles.count -gt 0){
			foreach ($profile in $profiles) 
			{					
				$accountname = (get-item $profile).pschildname 
				$lastaccesstime = (get-item $profile).lastaccesstime
				 if($accountname -notmatch "public" -and $accountname -notmatch "temp" -and $accountname -notmatch "guest"){ 
				  $robocopyargs = @("/l","/s","/njh","/bytes","/fp","/nc","/ndl","/ts","/xj","/r:0","/w:0")       
					[string] $summary = robocopy "\\$userpath\$accountname" null $robocopyargs | select-object -last 8
					[regex] $headerregex    = '\s+total\s+copied\s+skipped\s+mismatch\s+failed\s+extras'
					[regex] $dirlineregex   = 'dirs\s:\s+(?<dircount>\d+)(?:\s+\d+){3}\s+(?<dirfailed>\d+)\s+\d+'
					[regex] $filelineregex  = 'files\s:\s+(?<filecount>\d+)(?:\s+\d+){3}\s+(?<filefailed>\d+)\s+\d+'
					[regex] $byteslineregex = 'bytes\s:\s+(?<bytecount>\d+)(?:\s+\d+){3}\s+(?<bytefailed>\d+)\s+\d+'
					[regex] $timelineregex  = 'times\s:\s+(?<timeelapsed>\d+).*'
					[regex] $endedlineregex = 'ended\s:\s+(?<endedtime>.+)'
					if ($summary -match "$headerregex\s+$dirlineregex\s+$filelineregex\s+$byteslineregex\s+$timelineregex\s+$endedlineregex") {
					 $size = ([math]::round(([int64] $matches['bytecount'] / 1mb), 4)); $filecount = [int64] $matches['filecount'];$dircount=[int64] $matches['dircount'];
					 }
					if($i -eq 0){
						aspwritetext " $computername : $nl" cyan
						aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
						aspwritetext " profile`t`t`t`tlast access date`tsize MB`t`t`tfilecount`t`t`tdircount $nl" yellow
						aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
					}				
					aspwritetext  " $accountname`t`t`t`t$($lastaccesstime.tostring('dd.MM.yyyy HH:mm'))`t$size`t`t`t$filecount`t`t`t$dircount $nl" green
					aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue					
					if($global:gui){$global:lvresult.rows.add("$computername", "$accountname","  $($lastaccesstime.tostring('dd.MM.yyyy HH:mm'))","  $size ","  $filecount","  $dircount")}
				}
				else{
					if($i -eq 0){$logmsg = " no user profiles exists on $computername";if($global:gui){$global:lvresult.rows.add("$computername", " no user profiles exists")};aspwritetext " $nl $logmsg $nl $nl" red}
				}	
			$i++
			}
		}else{$logmsg = " no profiles exists on $computername";if($global:gui){$global:lvresult.rows.add("$computername", " no user profiles exists");aspwritetext "$nl $logmsg $nl $nl " red}}		
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetlogon{
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""
		)
	checkquit; if($global:gui){refreshgui;}		
	$error=0;$global:userloggedin=0;
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","username";		
		<#
		$users = ""
		try{
			$users = gwmi win32_loggedonuser -computername $computername -erroraction $erroractionpreference | `
			where-object { (($_.antecedent -match "t") -and ($_.antecedent -notmatch '.*name="x.+x"') -and ($_.antecedent -notmatch '.*\$')) } | `
			%{ $_.antecedent -replace '.*name="(.*)"','$1' } | select-object -unique
		}catch{ aspwritetext  " $computername : can not access $nl" red 
				$global:resultlist += " $computername : can not access $nl"
				
		}#>	
		$user2 = @();$i=0		
		try
            {
                $info = @(gwmi -class win32_process -computername $computername -filter "name='explorer.exe'" -erroraction $erroractionpreference)				
                if ($info){$info | % {$user2 += $_.getowner().user;#$sid +=$_.getownersid().sid
				}}
            }
		catch{ $error=1;  }			
		#if($users.count -gt 0){ $users|%{ if($_ -notmatch "system" -and $_ -notmatch "netzwerk" -and $_ -notmatch "anonymous" -and $_ -notmatch "lokaler"){$user2 = $_} } }
		if($user2.length -gt 0){		
				aspwritetext " $computername :" cyan;aspwritetext " $user2 $nl" green;if($global:gui){$global:lvresult.rows.add("$computername", " $user2")};$i++;
		}else{			
			$computer = "$computername";$sessionname=@();$username=@();$sessionid=@();
			try{
				$sessions = query session /server:$computer
				if($sessions.count -gt 0){
					1..($sessions.count -1) | % {					
						try{ $state = $sessions[$_].substring(48,8).trim() }
						catch{ $error=1; }
					   if($state -eq 'active' -or  $state -eq 'aktiv'){	$username += $sessions[$_].substring(19,20).trim() }   
					}
				}
				if($username.count -gt 0){			
					foreach($user in $username){
						aspwritetext " $computer :" cyan;aspwritetext " $user $nl" green;  #"`t`t"($sessionname[$i])"`t`t"($sessionid[$i])
						if($global:gui){$global:lvresult.rows.add("$computername", " $user")}
					}
				}				
			}
			catch{ $error=1; }			
		}		
		if(($username.count -eq 0) -and ($user2.length -eq 0)){
			$global:userloggedin=0;
			aspwritetext " $computer :" cyan;aspwritetext " no user $nl" red;if($global:gui){$global:lvresult.rows.add("$computername", " no user ")};
		}
		else{
			$global:userloggedin=1;
		}
		if($error){ aspwritetext  " $computer : can not access $nl" red;if($global:gui){$global:lvresult.rows.add("$computername", " can not access ");}		}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetcpuload{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)
		checkquit; 
		try{
			if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
				setcolumn -columncount 5 -columns "computername","name","cpupercent","cpu","description";						
					aspwritetext " $computername : $nl" cyan;
					$global:list =$null;		
						$global:list = invoke-command -scriptblock {				
							$cpupercent = @{
							  name = 'cpupercent'
							  expression = {
								$totalsec = (new-timespan -start $_.starttime).totalseconds
								[math]::round( ($_.cpu * 100 / $totalsec), 2)
								}
							}
							$list=get-process  | select-object -property name, $cpupercent, cpu, description | sort-object -property cpupercent -descending 
							if($list.count -gt 0){write-host "Name`t`t`tCPUPercent`t`tCPU`t`t`tDescription" -f green;}
							$list|%{$c="{0:n2}" -f $($_.cpu);write-host "$($_.name)`t`t`t $($_.cpupercent)`t`t $c`t`t`t $($_.description)" -f yellow;}
							return $list
					} -computername $computername  				
				if ($global:list.count -gt 0){
					if($global:gui){			
						$global:list|%{$c="{0:n2}" -f $($_.cpu);$global:lvresult.rows.add("$computername", "$($_.name)"," $($_.cpupercent)","$c"," $($_.description)" );}				
					}
				}else{
					if($global:gui){$global:lvresult.rows.add("$computername"," no cpu load found!");}
				}
			}
		}catch{
			$errormessage = $_.exception.message				
			$logmsg = " $cpuload failed for $computername the error message was $errormessage $nl"
			if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;
		}	
	checkquit; if($global:gui){refreshgui;}			
}

function aspgetfwstatus{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)
		checkquit; 
		try{
			if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
				setcolumn -columncount 3 -columns "computername","name","fwstatus";										
				aspwritetext " $computername : $nl" cyan;
				$hash = @{};
				$hash = invoke-command -scriptblock {
					$hash = $null
					$hash = @{}						
					$firewallkey = "hklm:\software\policies\microsoft\windowsfirewall";$type="gpo";	
					if( -not (test-path -path $firewallkey) )
					{
						$firewallkey = "hklm:\system\currentcontrolset\services\sharedaccess\parameters\firewallpolicy";$type="nogpo";						
					}
					write-host "`nfirewall-status(0=off,1=on), configtype $type" -f yellow;
					$firewall = get-childitem $firewallkey -recurse -erroraction silentlycontinue | where {$_.property -eq "enablefirewall";}
					 if (! $firewall)
						 {
							write-host "firewall is not configured or activated!" -f red;
							exit;
						 }
					foreach ($fwstatus in $firewall)
					{
						write-host "$($fwstatus.pschildname): $($fwstatus.getvalue('enablefirewall'))" -f green;
						$hash.add($($fwstatus.pschildname),$($fwstatus.getvalue('enablefirewall')))
					}					
					return $hash
				} -computername $computername			
				if ($hash){
					if($global:gui){
						$hash.getenumerator()|%{ 
							$global:lvresult.rows.add("$computername", "$($_.key)","$($_.value)");
						}
					}
				}else{
					if($global:gui){$global:lvresult.rows.add("$computername","firewall is not configured or activated!");}
				}
			}
		}catch{
			$errormessage = $_.exception.message;				
			$logmsg = " $firewall status failed for $computername the error message was $errormessage $nl";
			if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;
		}	
	checkquit; if($global:gui){refreshgui;}			
}

function aspgetmonitor{
	[cmdletbinding()]
		param (
				[parameter(mandatory=$true,valuefrompipeline=$true)] 
				[string]$computername = ""
			)
		checkquit; 
		try{
			if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
				setcolumn -columncount 4 -columns "computername","modelname","screenwidth","screenheight";							
				aspwritetext " $computername : $nl" cyan;			
				$monitors = gwmi -namespace root\wmi -class wmimonitorid -computername $computername -erroraction $erroractionpreference;
				$monitor2 = gwmi -class win32_desktopmonitor -computername $computername -erroraction $erroractionpreference;
				foreach ($singlemon in $monitors)
				{
					$modelname = ($singlemon.userfriendlyname | % {[char]$_}) -join "";
					$serialnum = ($singlemon.serialnumberid | % {[char]$_}) -join "";
					aspwritetext " $modelname $nl" green
					$monitor2| % {$widh=$_.screenwidth;$height=$_.screenheight;aspwritetext " screenwidth: $widh screenheight: $height $nl" yellow}
					if($global:gui){$global:lvresult.rows.add("$computername",$modelname,$widh,$height);}
				}
				if(!$monitors){
					$logmsg = "no monitor informations available! $nl";
					if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;
				}
			}			
		}catch{
			$errormessage = $_.exception.message;				
			$logmsg = " $monitor informations failed for $computername the error message was $errormessage $nl";
			if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg red;
		}	
	checkquit; if($global:gui){refreshgui;}			
}


function delay($seconds)
{
	aspwritetext " $nl start in " green
	while ($seconds -ge 1){
		aspwritetext "$seconds seconds..." green;start-sleep 1;#clear-host
		$seconds --;	
	}
	aspwritetext "go... $nl $nl";
}	

function aspviewwsusupdates{		
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[string]$logselection=""
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		if($logselection -eq "" -or $logselection -eq 0){
			aspwritetext "$nl which wsus log? $nl" yellow
			aspwritetext "$nl[1] asp wsus [2] asp wsus error [3] asp wsus selfupdate [4] ms wsus [5] wsus report (default is '1'): $nl"  yellow 
			$answer = read-host
		}else{
			$answer = $logselection
			}		
		try{
			$system = gwmi win32_logicaldisk -computername $computername -filter "deviceid='c:'"  -erroraction $erroractionpreference 
			$system | %{ $drive = $_.deviceid.chars(0) }
			}
		catch{
			$logmsg = " could not access $drive for $computername $nl";aspwritetext $logmsg  red;if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
		}		
		if(($answer.length -eq 0) -or ($answer -eq "1")){ 
			$path = "\\$computername\$drive`$\windows\asp-zid\log-wsus\wuau.log"
		}elseif($answer -eq "2"){
			$path = "\\$computername\$drive`$\windows\asp-zid\log-wsus\wuauerror.log"
		}elseif($answer -eq "3"){
			$path = "\\$computername\$drive`$\windows\asp-zid\log-wsus\wuauselfupdate.log"
		}elseif($answer -eq "4"){
			$path = "\\$computername\$drive`$\windows\windowsupdate.log"
		}elseif($answer -eq "5"){
			$path = "\\$computername\$drive`$\windows\softwaredistribution\reportingevents.log"
		}		
		try{ if (test-path $path){invoke-item $path} }
		catch{
			$logmsg = " could not access $path for $computername $nl";aspwritetext $logmsg  red;if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
		}
	}
	else{
		$logmsg = " $computername down $nl";aspwritetext $logmsg red;if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspvieweventvwr{		
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){eventvwr $computername}
	checkquit; if($global:gui){refreshgui;}	
}

function aspviewservices{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 1 -columns "computername";										
			services.msc /computer:$computername
			aspgetservices -computername $computername
		checkquit; if($global:gui){refreshgui;}	
	}
}

function aspgetprocess{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)	
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","process";										
		try{ 
			$proc = get-process -computername $computername  -erroraction $erroractionpreference |  select -expandproperty name 
			aspwritetext " $computername : $nl" cyan;
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			$proc|%{aspwritetext "$_ $nl" green}	
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue		
			$proc|%{if($global:gui){$global:lvresult.rows.add("$computername","$_")}}
		}
		catch{$logmsg = " error: could not retrieve process for $computername $nl";if($global:gui){$global:lvresult.rows.add("$computername", "error: could not retrieve process ")};aspwritetext $logmsg  red;}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspviewuser{		
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){lusrmgr.msc /computer:$computername}
	checkquit; if($global:gui){refreshgui;}	
}

function aspcdrive{				
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
	#$viewcdrive = "\\$computername\c$";explorer.exe $viewcdrive;
	$remotetempfile="\\$computername\c$\Users\ankae7"
	#Get-ChildItem -Path $remotetempfile -Recurse | Remove-Item -force -recurse
	#Remove-Item $remotetempfile -Force 
	rm -Force -Recurse -Confirm:$false $remotetempfile
	}	
	checkquit; if($global:gui){refreshgui;}
}

function readtextfile{
param( [string]$file )
 if ( test-path $file ) {
	$msg = ( ( get-content $file ) -join "`n" )	
}else {
	$msg = "cannot read '$mcafeelogs'" 
	}
	return $msg
}	

function aspvsedrive{				
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	setcolumn -columncount 2 -columns "computername","result";										
	try{	
		if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
			$mcafeelogs = "\\$computername\c$\windows\asp-zid\log-vse\ondemandscanlog.txt";
			$mcafeelogs2 = "\\$computername\c$\windows\asp-zid\log-vse\onaccessscanlog.txt";				
				$gui = new-object system.windows.forms.form
				$transparent = [system.drawing.color]::fromargb(200,225,225,225)
				$gui.Text = "McAfee Logs on $computername"
				$gui.formborderstyle = 'fixedsingle'
				$gui.startposition = 'centerscreen'
				$gui.minimumsize = new-object system.drawing.size(890,890)
				$gui.maximumsize = new-object system.drawing.size(890,890)
				$gui.icon = $icon
				$info = new-object system.windows.forms.label
				$info.text = readtextfile $mcafeelogs
				$info.text += readtextfile $mcafeelogs2
				$info.backcolor = "transparent"
				$info.autosize = $true
				$canvas = new-object system.windows.forms.panel
				$canvas.size = new-object system.drawing.size(865,800)
				$canvas.location  = new-object system.drawing.size(1,10)
				$canvas.backcolor = $transparent
				$canvas.autoscroll = $true
				$canvas.controls.add($info)
				$gui.controls.add($canvas)
				$cancelbutton = new-object system.windows.forms.button
				$cancelbutton.location = '810, 840'
				$cancelbutton.size = new-object system.drawing.size(75,23)
				$cancelbutton.text = "close"
				$cancelbutton.add_click({$gui.close()})
				$gui.controls.add($cancelbutton)
				$gui.add_shown({$gui.activate()})
				$gui.showdialog()			
		}	
	}catch{$logmsg = " error: could not open log-vse on $computername $nl";if($global:gui){$global:lvresult.rows.add("$computername", $logmsg)};aspwritetext $logmsg  red;}
	checkquit; if($global:gui){refreshgui;}
}
	
function aspviewlocaladmins{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)	
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 4 -columns "computername","name","domain","group";										
		try{
			$oscode = gwmi win32_operatingsystem -computername  $computername -erroraction $erroractionpreference | foreach {$_.oslanguage}
			  switch($oscode)
				{
					1033 {$group = [ADSI]("WinNT://$computername/ADMINISTRATORS");}				
					default {$group = [ADSI]("WinNT://$computername/ADMINISTRATOREN");}
				}	
			$group.members()|%{$groupcount++;}				
			if($groupcount -gt 0){
				$group.members() | %{ $adspath = $_.gettype().invokemember("adspath", 'getproperty', $null, $_, $null); $prop = $adspath.split('/',[stringsplitoptions]::removeemptyentries); $name = $prop[-1]; $domain = $prop[-2]; $class = $_.gettype().invokemember("class", 'getproperty', $null, $_, $null)
				aspwritetext "$nl $computername : $nl" cyan
				aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
				aspwritetext " name`t`t`t`t`t`tdomain`t`t`t`tgroup $nl" yellow
				aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
				if($domain -eq $computername){ $domain = "local" }
				aspwritetext " $name`t`t`t`t`t$domain`t`t`t`t$class $nl" green;if($global:gui){$global:lvresult.rows.add("$computername", " $name "," $domain "," $class")} ;
				}
			}else{$logmsg = " no administrators group exists on $computername $nl";if($global:gui){$global:lvresult.rows.add("$computername", "no administrators group exists")};aspwritetext $logmsg  red;}			
		}
		catch{$logmsg = " error: could not retrieve local administrators for $computername $nl";if($global:gui){$global:lvresult.rows.add("$computername", "error: could not retrieve local administrators")};aspwritetext $logmsg  red;}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspviewfreespace{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)	
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 6 -columns "computername","name","drive","totalsize","freespace","freespace percent";										
		try{
			aspwritetext "$nl $computername : $nl" cyan;
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " name`tdrive`ttotalsize`tfreespace`tfreespace percent $nl" yellow
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			$diskinfo= gwmi -computername $computername win32_logicaldisk -erroraction $erroractionpreference | where-object{$_.drivetype -eq 3} | select-object volumename, name, size, freespace, percentfree 
			$diskinfo | % { 
			$name = $_.volumename;$drive = $_.name;$size = ($_.size/1gb).tostring("0.0",$culturede) + " GB ";$freespace = ($_.freespace/1gb).tostring("0.0",$culturede) + " GB ";$percentfree = ($_.freespace/$_.size*100).tostring("0.0",$culturede) + " % ";aspwritetext " $name `t $drive `t $size `t $freespace `t $percentfree " green;
			if($global:gui){$global:lvresult.rows.add("$computername",  " $name "," $drive"," $size "," $freespace "," $percentfree")};
			aspwritetext $nl;}				
		}catch{$logmsg = " error: could not retrieve space informations for $computername $nl";aspwritetext $logmsg  red;if($global:gui){$global:lvresult.rows.add("$computername", "error: could not retrieve space informations");}}
	}
	checkquit; if($global:gui){refreshgui;}	
}

function asprdp{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		try{ mstsc.exe /v:$computername }
		catch{$logmsg = " error: could not retrieve space informations for $computername $nl";aspwritetext $logmsg  red;}
	}	
	checkquit; if($global:gui){refreshgui;}
}

function aspcompmgmt{		
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){compmgmt.msc /computer:$computername}
	checkquit; if($global:gui){refreshgui;}	
}

function aspgetipaddress{		
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	setcolumn -columncount 2 -columns "computername","ipaddress";										
	if($global:ipaddress -eq ""){
		$ip = $global:iplist[$computername]
	}else{$ip=$global:ipaddress}
	if($ip -eq ""){
		try { $addresslist = @(([net.dns]::gethostentry($computername)).addresslist) }
		catch {aspwritetext "cannot determine the ip address on $computername $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","cannot determine the ip address ");}}
		if ($addresslist.count -gt 0){ $addresslist | % { if ($_.addressfamily -eq "internetwork"){ $ip = $_.ipaddresstostring; }}}
	}		
	aspwritetext " $computername ","ip address: ","$ip $nl" cyan,white,green; 
	if($global:gui){$global:lvresult.rows.add("$computername","$ip")}	
	checkquit; if($global:gui){refreshgui;}
}

function aspgetmacaddress{    
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	setcolumn -columncount 2 -columns "computername","mac address";
	if($global:mac -eq ""){
		$mac =  $global:maclist[$computername]
	}else{$mac=$global:mac}
	if($mac -ne ""){
		aspwritetext " $computername ","mac address: ","$mac $nl" cyan,white,green; 
		if($global:gui){$global:lvresult.rows.add("$computername","$mac")}		
	}
	else{
		if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
			try {
				$networkinfos = gwmi -class "win32_networkadapterconfiguration" -computername $computername -erroraction $erroractionpreference|?{$_.ipenabled -match "true"} 
				$networkinfos|%{ 
					$nwservername = $_.dnshostname;$nwipaddr = $_.ipaddress;$nwsubnet = $_.ipsubnet;$nwgateway = $_.defaultipgateway;$nwdns = $_.dnsserversearchorder;
					$desc = $_.description; $mac = $_.macaddress; if($global:mac -ne ""){ aspwritetext " $computername ","mac address: ","$mac ","desc: ","$desc $nl" cyan,white,green,white,cyan; $global:lvresult.rows.add("$computername","$mac");}
				}
			}
			catch {	aspwritetext " cannot determine the macaddress on $computername $nl" red; if($global:gui){$global:lvresult.rows.add("$computername", " cannot determine the macaddress ")}}  
		}
	}
	checkquit; if($global:gui){refreshgui;}
} 

function aspgetlastboot{   	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}
	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 5 -columns "computername","lastboot","days","hours","minutes";
		try {
			$os = gwmi -class "win32_operatingsystem" -computername $computername -erroraction $erroractionpreference 	
			$lastboot = $os.converttodatetime($os.lastbootuptime);$lastboot = $lastboot.tostring('dd.MM.yyyy HH:mm');$uptime = ((get-date) - ($os.converttodatetime($os.lastbootuptime)));$d = $uptime.days;$h = $uptime.hours;$m = $uptime.minutes;
			aspwritetext "$nl $computername : $nl" cyan
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " lastboot`tuptime days hours minutes $nl" yellow
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue							
			aspwritetext " $lastboot `t $d $h $m $nl" green	
			if($global:gui){$global:lvresult.rows.add("$computername", "$lastboot",$d,$h,$m)}
		}
		catch {	aspwritetext " cannot determine the lastboot on $computername $nl" red; if($global:gui){$global:lvresult.rows.add("$computername", " cannot determine the lastboot ")}}  
	}
	checkquit; if($global:gui){refreshgui;};resize-datagrid		
}

function aspgetuserlang{   	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 3 -columns "computername","user","mui language";		
		try{ gwmi win32_userprofile -computername $computername  -erroraction $erroractionpreference | % { if($_.localpath -notmatch "service" -and $_.localpath -notmatch "system32"){ $path = $_.localpath } } }
		catch { 
			aspwritetext " can not access $computername" red 
			if($global:gui){$global:lvresult.rows.add(" can not access $computername")}
		}
		if($path -ne "" -and $path.length -gt 1){ $drive = $path.split(":") }
		if($drive.count -gt 0){ $drive = $drive[0] }	
		$userpath = "$computername\"+$drive+"$\users"		
		$userprofile = test-path \\$userpath
		try{ if($userprofile -eq "true"){			
				$userfolders = get-childitem "\\$userpath" | ?{$_.psiscontainer -and $_.name -ne 'public' -and $_.name -ne 'temp' -and $_.name -ne 'guest'}
				$userfolders | %{
					$userdat = join-path $_.fullname "ntuser.dat";$key = "control panel\desktop\muicached";$type = [microsoft.win32.registryhive]::currentuser;
					try{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computername,[microsoft.win32.registryview]::registry64) }
					catch{ $regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computername) }
					$subkey = $regkey.opensubkey($key);$lang = $subkey.getvalue("machinepreferreduilanguages");start-sleep -seconds 1;aspwritetext "$nl $computername : $nl" cyan;if($global:gui){$global:lvresult.rows.add("$computername"," $_ "," $lang" )}
					aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
					aspwritetext " user`tmui language $nl" yellow
					aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue	
					aspwritetext " $_ `t $lang $nl " green
					$lang = $null;$regkey.close();$subkey.close();
				}
			}			
		}
		catch { 
			aspwritetext " no user profiles on $computername $nl" red 
			if($global:gui){$global:lvresult.rows.add(" no user profiles on $computername ") }
		}
	}	
	checkquit; if($global:gui){refreshgui;}	
}

function aspgetenv{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 5 -columns "computername","user","variable value","system variable","user name";		
		try{ 
			$envs = gwmi -class "win32_environment" -namespace "root\cimv2" -computername $computername  -erroraction $erroractionpreference 
			aspwritetext "$nl $computername : $nl" cyan;
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " name`tvariable value`tsystem variable`tuser name: $nl" yellow
			$envs | % {					
					aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue	
					aspwritetext " ", $_.name,"`t",$_.variablevalue,"`t",$_.systemvariable,"`t",$_.username, $nl  green	
					if($global:gui){$global:lvresult.rows.add("$computername"," $($_.name)","$($_.variablevalue)","$($_.systemvariable)","$($_.username)")}
			}			
		}
		catch { aspwritetext " could not access environment vairables on $computername $nl" red; if($global:gui){$global:lvresult.rows.add(" $computername"," could not access environment vairables ") } }
	}
	checkquit; if($global:gui){refreshgui;}
}

function asptestport{ 
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[int]$port
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){		
		try{			 
			$failed = $null;$build = "" | select port, protocol, open, message;$tcpobject = new-object system.net.sockets.tcpclient;
			$connect = $tcpobject.beginconnect($computername,$port,$null,$null);$wait = $connect.asyncwaithandle.waitone(1000,$false);
			$build.port = $port;$build.protocol = "tcp";
			if(!$wait){  
				$tcpobject.close();$build.open = "false";$build.message = "connection to port timed out";
			}else{  
				$error.clear()  
				$tcpobject.endconnect($connect) | out-null 
				if($error[0]){ 
					[string]$string = ($error[0].exception).message;$message = (($string.split(":")[1]).replace('"',"")).trimstart();$failed = $true;
				} 
				$tcpobject.close() 
				if($failed){
					$build.open = "false";$build.message = "$message";
				}else{ 
					$build.open = "true";$build.message = "";
				}  
			}			       
		} 
		catch { aspwritetext " could not access port $port on $computername $nl" red } 
		if($build -ne $null){ 
			aspwritetext "$nl $computername : $nl" cyan
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " port`tprotocol`topen`tmessage $nl" yellow
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue	
			aspwritetext " $($build.port)`t$($build.protocol)`t$($build.open)`t$($build.message) $nl " green
		}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetmember{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){				 
		$computer = "$computername"			
		$adminpath = test-path \\$computer\admin$
		if ($adminpath -eq "true"){
			$type = [microsoft.win32.registryhive]::localmachine
			try{$regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer,[microsoft.win32.registryview]::registry64)}
			catch{$regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer)}
			$subkey = $regkey.opensubkey("system\\currentcontrolset\\services\\tcpip\\parameters") 
			$domainname = $subkey.getvalue("domain");aspwritetext "$computer : $domainname $nl" green;$regkey.close();$subkey.close();
		}else { aspwritetext " could not access $computername $nl" red }
	checkquit; if($global:gui){refreshgui;}			
	}
}

function showtip{         
	[cmdletbinding()]            
	param(            
	 [parameter(mandatory=$true)]            
	 [string]$title,            
	 [validateset("info","warning","error")]             
	 [string]$messagetype = "info",            
	 [parameter(mandatory=$true)]            
	 [string]$message,            
	 [string]$duration=10000            
	)            
	checkquit; if($global:gui){refreshgui;}  
	[system.reflection.assembly]::loadwithpartialname('system.windows.forms') | out-null            
	$balloon = new-object system.windows.forms.notifyicon;$path = get-process -id $pid | select-object -expandproperty path;
	$balloon.icon = $icon ; $balloon.balloontipicon = $messagetype; $balloon.balloontiptext = $message; $balloon.balloontiptitle = $title; $balloon.visible = $true; $balloon.showballoontip($duration);
	checkquit; if($global:gui){refreshgui;}	
}

function aspgetproductkey {    
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}
	$map = "bcdfghjkmpqrtvwxy2346789" 
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){   
		try { $os = gwmi -computername $computername win32_operatingsystem -erroraction $erroractionpreference } 
		catch {
				$os = new-object psobject -property @{
				caption = $_.exception.message
				version = $_.exception.message
			}
		}
		try {		  
			$remotereg = [microsoft.win32.registrykey]::openremotebasekey([microsoft.win32.registryhive]::localmachine,$computername)
			if ($os.osarchitecture -eq '64-bit') { $value = $remotereg.opensubkey("software\microsoft\windows nt\currentversion").getvalue('digitalproductid4')[0x34..0x42] } 
			else { $value = $remotereg.opensubkey("software\microsoft\windows nt\currentversion").getvalue('digitalproductid')[0x34..0x42] }
			$pk = "" 		   
			for ($i = 24; $i -ge 0; $i--) { 
			  $k = 0 
			  for ($j = 14; $j -ge 0; $j--) { 
				$k = ($k * 256) -bxor $value[$j] 
				$value[$j] = [math]::floor([double]($k/24)) 
				$k %=  24 
			  } 
			  $pk = $map[$k] + $pk 
			  if (($i % 5) -eq 0 -and $i -ne 0) { 
				$pk = "-" + $pk 
			  } 
			}
		} catch { $pk = $_.exception.message }
			aspwritetext "$nl $computername : $nl" cyan
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " osdescription`tosversion`tproductkey $nl" yellow
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue	
			aspwritetext " $($os.caption)`t$($os.version)`t$($pk) $nl " green 
			aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue	
	}else { aspwritetext " could not access $computername $nl" red } 	
	checkquit; if($global:gui){refreshgui;}	
}
        
function aspgettasks{ 
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}		
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 12 -columns "computername","name","path","state","enabled","lastruntime","lasttaskresult","numberofmissedruns","nextruntime","autor","userid","desc";		
		try { $schedule = new-object -com("schedule.service") } 
		catch {
			aspwritetext " schedule.service com object not found, this script requires this object $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","schedule.service com object not found, this script requires this object")}
		}
		try {$schedule.connect($computername) }
		catch {
			aspwritetext " schedule.service com object not found, this script requires this object $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","schedule.service com object not found, this script requires this object")}
		}
		$tasks = $schedule.getfolder("\").gettasks(1)
		$results = @()
		aspwritetext " $computername : $nl" cyan
		aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue	
		$tasks | % {
			$autor = ([regex]::split($_.xml,'<Author>|</Author>'))[1]
			$userid = ([regex]::split($_.xml,'<UserId>|</UserId>'))[1]
			$desc = ([regex]::split($_.xml,'<Description>|</Description>'))[1]
			aspwritetext "`tname:","`t$($_.name) $nl" yellow, green
			aspwritetext "`tpath:","`t$($_.path) $nl" yellow, green
			aspwritetext "`tstate:","`t$($_.state) $nl" yellow, green
			aspwritetext "`tenabled:","`t$($_.enabled) $nl" yellow, green 
			aspwritetext "`tlastruntime:","`t$($_.lastruntime) $nl" yellow, green 
			aspwritetext "`tlasttaskresult:","`t$($_.lasttaskresult) $nl" yellow, green 
			aspwritetext "`tnumberofmissedruns:","`t$($_.numberofmissedruns) $nl" yellow, green 
			aspwritetext "`tnextruntime:","`t$($_.nextruntime) $nl" yellow, green 
			aspwritetext "`tauthor:","`t$($autor) $nl" yellow, green 
			aspwritetext "`tuserid:","`t$($userid) $nl" yellow, green 
			aspwritetext "`tdescription:","`t$($desc) $nl" yellow, green 
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			if($global:gui){$global:lvresult.rows.add("$computername","$($_.name)","$($_.path)","$($_.state)","$($_.enabled)","$($_.lastruntime)","$($_.lasttaskresult)","$($_.numberofmissedruns)","$($_.nextruntime)","$($autor)","$($userid)","$($desc)")}
		}
	}else { aspwritetext " could not access $computername $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","could not access")}
		} 
	checkquit; if($global:gui){refreshgui;}
} 

function aspgetservices{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 3 -columns "computername","service","status";		
		try { 
		#gwmi win32_service -computername $computername  -erroraction $erroractionpreference | select-object name 
			get-service -computername $computername | %{if ($_.status -eq "stopped"){aspwritetext "$($_.name) : $($_.status) $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","$($_.name)","$($_.status)")}} elseif ($_.status -eq "running") {aspwritetext "$($_.name) : $($_.status) $nl" green;if($global:gui){$global:lvresult.rows.add("$computername","$($_.name)","$($_.status)") }}}
		}
		catch {
				aspwritetext " win32_service object not found, this script requires this object $nl" red 
				
			}	
	}else { aspwritetext " could not access $computername $nl" red } 
	checkquit; if($global:gui){refreshgui;}	
}   
 
function aspsetservice{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[string]$service = "",	
		[string]$action = "",
		[validaterange(0,1)] 
		[int]$log = 0	
	)
	checkquit; if($global:gui){refreshgui;} 	
	$res = ""	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","result";		
   		if($service -eq "" -and $action -eq "" -and $msg.length -gt 0 -and $msg -ne "please log off."){
			$arg = @()
			$arg = $msg.split(";")
			$service = $arg[0]
			$action = $arg[1]
		}
		#try{		
		aspwritetext " $computername : $nl" cyan			
		aspwritetext "  try $action $service $nl" yellow
		$global:status = "  try $action $service $nl"
		switch ($action) {
			start { $res = (gwmi win32_service -computername $computername -filter "name='$service'" -erroraction $erroractionpreference ).startservice()}
			stop { $res = (gwmi win32_service -computername $computername -filter "name='$service'" -erroraction $erroractionpreference ).stopservice()}
			disabled { $res = (gwmi win32_service -computername $computername -filter "name='$service'" -erroraction $erroractionpreference ).changestartmode("disabled")}
			automatic { $res = (gwmi win32_service -computername $computername -filter "name='$service'" -erroraction $erroractionpreference ).changestartmode("automatic")}			
		}
			
		if($res.returnvalue -eq "0"){ 
			$logmsg = " $action $service successfully $nl"
			aspwritetext $logmsg green; 
			if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}			
			if($log){ aspwritelog $logmsg }
		}
		else { 
			$logmsg = " an error has occurred $nl"
			aspwritetext $logmsg red; 
			if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
			if($log){ aspwritelog $logmsg}
		}
	<#}catch {
			$logmsg = " win32_service object not found, this script requires this object $nl"
			aspwritetext $logmsg red 
			
			if($log){ aspwritelog $logmsg}
		}#>	
	}else { 
			$logmsg = " could not access $computername $nl"
			aspwritetext $logmsg red; 
			if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
			if($log){ aspwritelog $logmsg}
		} 
	checkquit; if($global:gui){refreshgui;}		
}

function asprestarteris{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[validaterange(0,1)] 
		[int]$log = 0
	)
	checkquit; if($global:gui){refreshgui;} 	
	if(asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0){	
		#try { 
	aspsetservice -computername $computername -service "eris" -action "stop" -log $log	
	delay(5)
	aspsetservice -computername $computername -service "eris" -action "start" -log $log
	<#}catch {
			$logmsg = " win32_service object not found, this script requires this object $nl"
			aspwritetext $logmsg red 
			if($log){ aspwritelog $logmsg }
	}#>	
	}else{ 
		$logmsg = " could not access $computername $nl"
		aspwritetext $logmsg red 
		if($log){ aspwritelog $logmsg	}
	}
	checkquit; if($global:gui){refreshgui;}	
}

function aspgetinstalldate{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;} 	
	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","install date";		
		try{ 
			$date = (gwmi -class win32_operatingsystem -computername $computername -erroraction $erroractionpreference ).installdate 
			aspwritetext " $computername : $nl" cyan
			aspwritetext "`tinstall date:","`t$($date.substring(6,2)).$($date.substring(4,2)).$($date.substring(0,4)) $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername", " $($date.substring(6,2)).$($date.substring(4,2)).$($date.substring(0,4))")	}	
		}	
		catch{
				aspwritetext " the rpc server is unavailable $nl" red
				if($global:gui){$global:lvresult.rows.add("$computername", "  the rpc server is unavailable ")}				
			}		
    }else{ aspwritetext " could not access $computername $nl" red; if($global:gui){$global:lvresult.rows.add("$computername", "  could not access ") } }
	checkquit; if($global:gui){refreshgui;}
}

function aspgetshares{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	)
	checkquit; if($global:gui){refreshgui;}	
	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 3 -columns "computername","share","username";				
		try{ $shares = gwmi -class win32_share -computername $computername -erroraction $erroractionpreference  | select -expandproperty name }
		catch {
				aspwritetext " the rpc server is unavailable $nl" red 
				if($global:gui){$global:lvresult.rows.add("$computername",  " the rpc server is unavailable")}
			}		
		aspwritetext " $computername : $nl" cyan
		aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
		foreach($share in $shares){ 
			$rule = $null 
			aspwritetext "`t$share" green 
			$sec = gwmi -class win32_logicalsharesecuritysetting -filter "name='$share'" -computername $computername -erroraction $erroractionpreference 
			try { 
				$sd = $sec.getsecuritydescriptor().descriptor   
				$sd.dacl | % {  
					$username = $_.trustee.name     
					if($_.trustee.domain -ne $null) {$username = "$($_.trustee.domain)\$username"}   
					if($_.trustee.name -eq $null) {$username = $_.trustee.sidstring}     
					[array]$rule += new-object security.accesscontrol.filesystemaccessrule($username, $($_.accessmask), $($_.acetype)) 
				}          
			}
			catch{aspwritetext "`tunable to obtain permissions for $share $nl" red;if($global:gui){$global:lvresult.rows.add("$computername", "$share"," unable to obtain permissions ")} }
			$rule
			if($rule.count -gt 0){ if($global:gui){$global:lvresult.rows.add("$computername", "$share"," $rule") }}
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
		} 		
	}
	else{ aspwritetext " could not access $computername $nl" red;if($global:gui){$global:lvresult.rows.add("$computername", "  could not access ");} }
	checkquit; if($global:gui){refreshgui;}
}

function checkfor45dotversion($releaseKey){
	if ($releaseKey -ge 393273) {
		  return "4.6 RC or later";
	   }
	   if (($releaseKey -ge 379893)) {
			return "4.5.2 or later";
		}
		if (($releaseKey -ge 378675)) {
			return "4.5.1 or later";
		}
		if (($releaseKey -ge 378389)) {
			return "4.5 or later";
		}
	return "No 4.5 or later version detected";
}

function aspgetnetVersion{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
			setcolumn -columncount 2 -columns "computername",".net version";				
		try {
				aspwritetext " $computername : $nl" cyan
				$w32reg = [microsoft.win32.registrykey]::openremotebasekey('localmachine',$computername)
				$keypath = 'software\\microsoft\\net framework setup\\ndp\\v4\\full\\'
				$net = $w32reg.opensubkey($keypath)
				$netrelease = $net.getvalue('release')
				$bla=checkfor45dotversion($netrelease)
				aspwritetext "  .NET Version: $bla $nl" yellow
				if($global:gui){$global:lvresult.rows.add("$computername",$bla);}
			}
			catch{aspwritetext "` rpc server is unavailable $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","  RPC server is unavailable ");}} 
	}
}


function aspseterislogon{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){ 
		try {
		aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
			aspgeterislogstatus -computername $computername 
			aspwritetext " $computername : $nl" cyan;
			$w32reg = "hklm:\system\currentcontrolset\services\eris";
			$logon = "c:\windows\system32\empirum\eris.exe /log";
			$value = "imagepath";
			invoke-command -scriptblock {
			param (
			   [string] $reg,
			   [string] $logon,
			   [string] $value
			   )
			   if (test-path $reg){
				write-host " try to set eris log on " -f yellow;
				set-itemproperty -path $reg -name $value -value $logon;				
			   }
			}-computername $computername -argumentlist $w32reg,$logon,$value;
			asprestarteris -computername $computername;			
			aspgeterislogstatus -computername $computername;
		}catch{
			$errormessage = $_.exception.message;
			$log = "` failed for $computername the error message was $errormessage $nl";
			aspwritetext $log red;if($global:gui){$global:lvresult.rows.add("$computername",$log);}
		}
		aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
	}			
}

function aspseterislogoff{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){ 
		try {
		aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue;
			aspgeterislogstatus -computername $computername 
			aspwritetext " $computername : $nl" cyan;
			$w32reg = "hklm:\system\currentcontrolset\services\eris";
			$logoff = "c:\windows\system32\empirum\eris.exe";
			$value = "imagepath";
			invoke-command -scriptblock {
			param (
			   [string] $reg,
			   [string] $logoff,
			   [string] $value
			   )
			   if (test-path $reg){
				write-host " try to set eris log off " -f yellow;
				set-itemproperty -path $reg -name $value -value $logoff;				
			   }
			}-computername $computername -argumentlist $w32reg,$logoff,$value;
			asprestarteris -computername $computername;			
			aspgeterislogstatus -computername $computername;			
		}catch{
			$errormessage = $_.exception.message;
			$log = "` failed for $computername the error message was $errormessage $nl";
			aspwritetext $log red;if($global:gui){$global:lvresult.rows.add("$computername",$log);}
		}
		aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue;
	}			
}


function aspgetnetstat{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 9 -columns "computername","pid","processname","protocol","localaddress","localport","remoteaddress","remoteport""state";				
		try {
			aspwritetext " $computername : $nl" cyan;
			#[outputtype('system.management.automation.psobject')]
			[system.string]$processname='*';
			[system.string]$address='*';
			$port='*';
			[system.string]$protocol='*';
			[system.string]$state='*';
			$showhostnames=$false; 
			$showprocessnames = $true;
			[system.string]$tempfile = "c:\temp\net.dat";
			[string]$addressfamily = '*';
			$properties = 'computername','protocol','localaddress','localport','remoteaddress','remoteport','state','processname','pid';
			$dnscache = @{};
			if($showprocessnames){
				try {$processes = get-process -computername $computername -erroraction $erroractionpreference | select name, id;}
				catch {
					aspwritetext "could not run get-process -computername $computername.  verify permissions and connectivity.  defaulting to no showprocessnames";
					$showprocessnames = $false;
				}
			}
			if($computername -ne $env:computername){
				[string]$cmd = "cmd /c c:\windows\system32\netstat.exe -ano >> $tempfile";
				$remotetempfile = "\\{0}\{1}`${2}" -f "$computername", (split-path $tempfile -qualifier).trimend(":"), (split-path $tempfile -noqualifier);
				try{$null = invoke-wmimethod -class win32_process -name create -argumentlist "cmd /c del $tempfile" -computername $computername -erroraction $erroractionpreference;}
				catch{aspwritetext "could not invoke create win32_process on $computername to delete $tempfile";}
				try{$processid = (invoke-wmimethod -class win32_process -name create -argumentlist $cmd -computername $computername -erroraction $erroractionpreference).processid}
					catch{throw $_;break;}
						while($(try{get-process -id $processid -computername $computername -erroraction $erroractionpreference}
								catch{$false;})
						)
						{start-sleep -seconds 2;}
						if(test-path $remotetempfile){
							try{$results = get-content $remotetempfile | select-string -pattern '\s+(tcp|udp)';}
							catch{throw "could not get content from $remotetempfile for results";break;}
							remove-item $remotetempfile -force;
						}
						else{throw "'$tempfile' on $computername converted to '$remotetempfile'.  this path is not accessible from your system.";break;}
				}
				else{$results = netstat -ano | select-string -pattern '\s+(tcp|udp)';}
					$totalcount = $results.count;$count = 0;
					$arr=@();
					if($global:gui){
						aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
						aspwritetext " pid`t`t`t`tprocessname`tprotocol`t`t`tlocaladdress`t`t`tlocalport`t`t`tremoteaddress`t`t`tremoteport`t`t`tstate $nl" yellow
						aspwritetext " -------------------------------------------------------------------------------------------------- $nl" blue
					}
				foreach($result in $results){
    	            $item = $result.line.split(' ',[system.stringsplitoptions]::removeemptyentries);
    	            if($item[1] -notmatch '^\[::'){
    	                    if (($la = $item[1] -as [ipaddress]).addressfamily -eq 'internetworkv6'){
    	                        $localaddress = $la.ipaddresstostring;
    	                        $localport = $item[1].split('\]:')[-1];
    	                    }
    	                    else {
    	                        $localaddress = $item[1].split(':')[0];
    	                        $localport = $item[1].split(':')[-1];
    	                    }
    	                    if (($ra = $item[2] -as [ipaddress]).addressfamily -eq 'internetworkv6'){
    	                        $remoteaddress = $ra.ipaddresstostring;
    	                        $remoteport = $item[2].split('\]:')[-1];
    	                    }
    	                    else {
    	                        $remoteaddress = $item[2].split(':')[0];
    	                        $remoteport = $item[2].split(':')[-1];
    	                    }
                            if($addressfamily -ne "*")
                            {
                                if($addressfamily -eq 'ipv4' -and $localaddress -match ':' -and $remoteaddress -match ':|\*' ){aspwritetext "filtered by addressfamily:`n$result";continue;}
                                elseif($addressfamily -eq 'ipv6' -and $localaddress -notmatch ':' -and ( $remoteaddress -notmatch ':' -or $remoteaddress -match '*' )){aspwritetext "filtered by addressfamily:`n$result";continue;}
                            }
    	    		        $procid = $item[-1];$proto = $item[0];$status = if($item[0] -eq 'tcp'){$item[3];}else{$null;}
		    		        if($remoteport -notlike $port -and $localport -notlike $port){ if(-not $global:gui){aspwritetext "remote $remoteport local $localport port $port";aspwritetext "filtered by port:`n$result";continue;}}
		    		        if($remoteaddress -notlike $address -and $localaddress -notlike $address){ if(-not $global:gui){aspwritetext "filtered by address:`n$result";continue;}}
    	    			    if($status -notlike $state){ if(-not $global:gui){aspwritetext "filtered by state:`n$result";continue;}}
    	    			    if($proto -notlike $protocol){ if(-not $global:gui){aspwritetext "filtered by protocol:`n$result";continue;}}
                            write-progress -activity "resolving host and process names" -status "resolving process id $procid with remote address $remoteaddress and local address $localaddress" -percentcomplete (( $count / $totalcount ) * 100);
                            if($showprocessnames -or $psboundparameters.containskey -eq 'processname'){if($procname = $processes|?{$_.id -eq $procid} | select -expandproperty name ){;}
                                else{$procname = "unknown";}
                            }
                            else{$procname = "na";}
		    		        if($procname -notlike $processname){ if(-not $global:gui){aspwritetext "filtered by processname:`n$result";continue;}}
                            if($showhostnames){
                                $tmpaddress = $null;
                                try{
                                    if($remoteaddress -eq "127.0.0.1" -or $remoteaddress -eq "0.0.0.0"){$remoteaddress = $computername;}
                                    elseif($remoteaddress -match "\w"){
                                            if ($dnscache.containskey( $remoteaddress)){$remoteaddress = $dnscache[$remoteaddress]; if(-not $global:gui){aspwritetext "using cached remote '$remoteaddress'";}}
                                            else{
												$tmpaddress = $remoteaddress;
												$remoteaddress = [system.net.dns]::gethostbyaddress("$remoteaddress").hostname;
												$dnscache.add($tmpaddress, $remoteaddress);
												 if(-not $global:gui){aspwritetext "using non cached remote '$remoteaddress`t$tmpaddress";}
                                            }
                                    }
                                }
                                catch{continue;}
                                try{
                                    if($localaddress -eq "127.0.0.1" -or $localaddress -eq "0.0.0.0"){$localaddress = $computername;}
                                    elseif($localaddress -match "\w"){
										if($dnscache.containskey($localaddress)){$localaddress = $dnscache[$localaddress]; if(-not $global:gui){aspwritetext "using cached local '$localaddress'";}}
										else{
											$tmpaddress = $localaddress;
											$localaddress = [system.net.dns]::gethostbyaddress("$localaddress").hostname;
											$dnscache.add($localaddress, $tmpaddress);
											 if(-not $global:gui){aspwritetext "using non cached local '$localaddress'`t'$tmpaddress'";}
										}
                                    }
                                }
                                catch{continue;}
                            }
							$netresult = new-object -typename psobject -property @{
								computername = $computername;
								pid = $procid;
								processname = $procname;
								protocol = $proto;
								localaddress = $localaddress;
								localport = $localport;
								remoteaddress =$remoteaddress;
								remoteport = $remoteport;
								state = $status;
							} | select-object -property $properties;
							 if(-not $global:gui){$netresult|ft -autosize;}
							 else{aspwritetext " $procid`t`t`t`t$procname`t$proto`t`t`t$localaddress`t`t`t$localport`t`t`t$remoteaddress`t`t`t$remoteport`t`t`t$status $nl" green;}
							$arr+=$netresult;
						$count++;
                    }
                }write-progress -completed -activity "completed" -status "completed";
				#for($i=0;$i -lt $count;$i++){
						#$arr[$i]|ft -autosize;
						$arr|%{if($global:gui){$global:lvresult.rows.add("$computername",$_.pid,$_.processname,$_.protocol,$_.localaddress,$_.localport,$_.remoteaddress,$_.remoteport,$_.state);}}
					#}
        }		
		catch{
			$errormessage = $_.exception.message;
			$log = "` failed for $computername the error message was $errormessage $nl";
			aspwritetext $log red;if($global:gui){$global:lvresult.rows.add("$computername",$log);}
		} 
	 }
}

function aspgetrebootpending{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 3 -columns "computername","location","value";				
		aspwritetext " $computername : $nl" cyan
		try {
			$comppendren,$pendfilerename,$pending,$sccm = $false,$false,$false,$false
			$cbsrebootpend = $null
			$wmi_os = gwmi -class win32_operatingsystem -property buildnumber, csname -computername $computername -erroraction $erroractionpreference
			$hklm = [uint32] "0x80000002"
			$wmi_reg = [wmiclass] "\\$computername\root\default:stdregprov"
			if ([int32]$wmi_os.buildnumber -ge 6001) {
				$regsubkeyscbs = $wmi_reg.enumkey($hklm,"software\microsoft\windows\currentversion\component based servicing\")
				$cbsrebootpend = $regsubkeyscbs.snames -contains "rebootpending"		
			}
			$regwuaurebootreq = $wmi_reg.enumkey($hklm,"software\microsoft\windows\currentversion\windowsupdate\auto update\")
			$wuaurebootreq = $regwuaurebootreq.snames -contains "rebootrequired"
			$regsubkeysm = $wmi_reg.getmultistringvalue($hklm,"system\currentcontrolset\control\session manager\","pendingfilerenameoperations")
			$regvaluepfro = $regsubkeysm.svalue
			$netlogon = $wmi_reg.enumkey($hklm,"system\currentcontrolset\services\netlogon").snames
			$penddomjoin = ($netlogon -contains 'joindomain') -or ($netlogon -contains 'avoidspnset')
			$actcompnm = $wmi_reg.getstringvalue($hklm,"system\currentcontrolset\control\computername\activecomputername\","computername")            
			$compnm = $wmi_reg.getstringvalue($hklm,"system\currentcontrolset\control\computername\computername\","computername")
			if (($actcompnm -ne $compnm) -or $penddomjoin) {
				$comppendren = $true
			}
			if ($regvaluepfro) {
				$pendfilerename = $true
			}
			$ccmclientsdk = $null
			$ccmsplat = @{
				namespace='root\ccm\clientsdk'
				class='ccm_clientutilities'
				name='determineifrebootpending'
				computername=$computername
				erroraction='stop'
			}
			try {
				$ccmclientsdk = invoke-wmimethod @ccmsplat
			} catch [system.unauthorizedaccessexception] {
				$ccmstatus = get-service -name ccmexec -computername $computername -erroraction $erroractionpreference
				if ($ccmstatus.status -ne 'running') {
					$log = " error: ccmexec service is not running"
					aspwritetext " $log $nl" red
					if($global:gui){$global:lvresult.rows.add("$computername",$log);}
					$ccmclientsdk = $null
				}
			} catch {$ccmclientsdk = $null}
			if ($ccmclientsdk) {
				if ($ccmclientsdk.returnvalue -ne 0) {
					$log = " error: determineifrebootpending returned error code $($ccmclientsdk.returnvalue) "
					aspwritetext " $log $nl" red						
					if($global:gui){$global:lvresult.rows.add("$computername",$log);}						
				}
				if ($ccmclientsdk.ishardrebootpending -or $ccmclientsdk.rebootpending) {
					$sccm = $true
				}
			}					
			else{$sccm = $null}
			aspwritetext "  component based servicing:"," $cbsrebootpend $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername","component based servicing:",$cbsrebootpend);}	
			aspwritetext "  windowsupdate:"," $wuaurebootreq $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername","windowsupdate:",$wuaurebootreq);}				
			aspwritetext "  pending computer after rename:"," $comppendren $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername","pending computer after rename:",$comppendren);}	
			aspwritetext "  pending file rename operations:"," $pendfilerename $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername","pending file rename operations:",$pendfilerename);}	
			$regvaluepfro | %{$val=$_;aspwritetext "  pending file rename operation files:"," $val $nl" yellow, green;if($global:gui){$global:lvresult.rows.add("$computername","pending file rename operation files:",$val);}}
			$value=($comppendren -or $cbsrebootpend -or $wuaurebootreq -or $sccm -or $pendfilerename)
			aspwritetext "  rebootpending:"," $value $nl" yellow, green
			if($global:gui){$global:lvresult.rows.add("$computername","rebootpending:",$value);}				
		} catch{aspwritetext "` rpc server is unavailable $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","  RPC server is unavailable ");}} 
	}
}

function aspgetprinter{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){ 
		setcolumn -columncount 2 -columns "computername","printer";				
		try {
				aspwritetext " $computername : $nl" cyan
				$objwmi = gwmi -computername $computername -class 'win32_printer' -erroraction $erroractionpreference
				foreach( $strwmiprinter in $objwmi )
				{				  
					$status = $strwmiprinter.printerstatus
					$result = ''
					switch( $status )
					{
					   1{ $result = 'other' }
					   2{ $result = 'unknown' }
					   3{ $result = 'running/full power' }
					   4{ $result = 'warning' }
					   5{ $result = 'in test' }
					   6{ $result = 'not applicable' }
					   7{ $result = 'power off' }
					   8{ $result = 'off line' }
					   9{ $result = 'off duty' }
					  10{ $result = 'degraded' }
					  11{ $result = 'not installed' }
					  12{ $result = 'install error' }
					  13{ $result = 'power save - unknown' }
					  14{ $result = 'power save - low power mode' }
					  15{ $result = 'power save - standby' }
					  16{ $result = 'power cylce' }
					  17{ $result = 'power save - warning' }					  
					  default{ $result = 'unknown' }
					}
					if( $result -ne $null )
					{
					  $res = " printer: "+$($strwmiprinter.name)+" - printerstatus: " + $result				  
					}
					else
					{
					  $res = " error: no printer found on " + $computername 
					}
					aspwritetext " $res $nl" yellow
					if($global:gui){$global:lvresult.rows.add("$computername",$res);}
				} 				
			}
			catch{aspwritetext "` rpc server is unavailable $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","  RPC server is unavailable ");}} 
	}
}


function aspgetPSVersion{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","powershell version";				
	aspwritetext " $computername : $nl" cyan		
		try {				
				$ver=invoke-command -scriptblock{
					$ver = $psversiontable.psversion
					write-host "  PS-Version: "$ver -f yellow
					return $ver
				}-computername $computername
				if($global:gui){$global:lvresult.rows.add("$computername",$ver);}
			}
			catch{aspwritetext "` rpc server is unavailable $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","  RPC server is unavailable ");}} 
	}
}


function aspgetInternetExVersion{
[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","ie version";				
		aspwritetext " $computername : $nl" cyan		
		try {				
				#gwmi win32_product -filter "name like '%internet%'" -computername $computername -erroraction $erroractionpreference			
				$name = @{name="name";expression= {split-path -leaf $_.filename}};	$path = @{name="path";expression= {split-path $_.filename}}
				dir -recurse -path "\\$computername\C$\program files\internet explorer" | % { if ($_.name -match "(.*exe)$") {$version=$_.versioninfo;}} 				
				$v = $version | select fileversion, $name, $path | ? {$_.name -eq "iexplore.exe"};
				aspwritetext "  Internet Explorer Version: $($v.fileversion) $nl" yellow
				if($global:gui){$global:lvresult.rows.add("$computername",$($v.fileversion))};
			}
			catch{aspwritetext "` rpc server is unavailable $nl" red;if($global:gui){$global:lvresult.rows.add("$computername","  RPC server is unavailable ");}} 
	}
}

function aspstopprocess{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){ 
		try{ 
			if($msg.length -gt 0 -and $msg -ne "please log off."){		
				$processes = gwmi -class win32_process -computername $computername -filter "name='$msg.exe'" -erroraction $erroractionpreference 
				if($processes){
					foreach ($process in $processes) {
						$returnval = $process.terminate();$processid = $process.handle;
						if($returnval.returnvalue -eq 0) {aspwritetext "`nthe process $processname `($processid`) terminated successfully";}
						else {aspwritetext "`nthe process $processname `($processid`) termination has some problems";}
					}
				}else {aspwritetext "`n no processes found with the name $msg";}	
			}
		}
		catch {aspwritetext " the rpc server is unavailable $nl" red;}	
	}else { aspwritetext " could not access $computername $nl" red }
	checkquit; if($global:gui){refreshgui;}	
}

function aspdelsetupinf{	
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[string]$drive = ""
	) 
	checkquit; if($global:gui){refreshgui;}	
	if($msg.length -eq 0 -or $msg -eq "please log off."){ aspwritetext "$nl no drive parameter! $nl" red; return }
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){ 
	$check = $true
	try{ 
		while($check){
			write-progress -activity "search setup.inf(s) on $computername" -status "please wait..."
			$setups = gci \\$computername\$drive$\empirumagent\packages\ -rec | where {($_.name -like "setup.inf")}
			$check = $false
		}
		write-progress "search done" "done" -completed		
		if($setups.count -gt 0){
			if($force -eq 0){
				aspwritetext "$nl $($setups.count) were found - do you really want to delete all? $nl"  yellow 
				aspwritetext "$nl [y] yes  [n] no (default is 'y'): "  yellow 
				$answer = read-host
				if($answer.length -eq 0){ $answer = "y" }
				if($answer.tolower() -eq "y"){ $setups | % { remove-item $_.fullname -force }; aspwritetext "$nl delete setup.inf(s) successfully $nl" green  }
				else { aspwritetext "$nl process aborted $nl" red  }
			}else{	
					$count = 0
					try{ $setups | % { remove-item $_.fullname -force; $count++ }; aspwritetext "$nl delete $count setup.inf(s) successfully $nl" green }
					catch {
						aspwritetext "$nl $computername an error has occurred $nl" red 						
					}	
				}
		}else{ aspwritetext "$nl 0 setup.inf(s) on $computername were found! $nl" red }			
	}
	catch {
			aspwritetext "$nl an error has occurred $nl" red 
			
		}	
	}else { aspwritetext " could not access $computername $nl" red }
	checkquit; if($global:gui){refreshgui;}	
}

function getprog($start=$false){
	if($start){ aspwritetext "." yellow; start-sleep 1 }
}

function aspwake{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[validaterange(0,1)] 
		[int]$log = 0
	) 
	setcolumn -columncount 2 -columns "computername","result";				
	if($force -eq 0){
		aspwritetext "$nl do you really want to wake $computername up? $nl"  yellow 
		aspwritetext "$nl [y] yes  [n] no (default is 'y'): "  yellow 
		$answer = read-host
		if($answer.length -eq 0){ $answer = "y" }
		if($answer.tolower() -eq "y"){  }
		else { aspwritetext "$nl process aborted $nl" red; break }
	}	
	#try { $addresslist = @(([net.dns]::gethostentry($computername)).addresslist) }
	#catch{}
	delay(1);
	#if ($addresslist.count -gt 0){ $addresslist | % { if ($_.addressfamily -eq "internetwork"){ $global:ipaddress = $_.ipaddresstostring } } }
	$ip = $global:iplist[$computername]
	$mac = $global:maclist[$computername]
	if ($mac -eq $null){ 
		$logmsg = "$computername not found. $nl"
		aspwritetext $logmsg red; if($global:gui){$global:lvresult.rows.add("$computername","not found") }
		if($log){ aspwritelog $logmsg}
	}
	else
	{
		$mac  -match "(..)(..)(..)(..)(..)(..)" | out-null
		$mac = [byte[]]($matches[1..6] |% {[int]"0x$_"})
		$udpclient = new-object system.net.sockets.udpclient				
		#$udpclient.connect(([system.net.ipaddress]::broadcast),7)
		$packet = [byte[]](,0xff * 102)
		6..101 |% { $packet[$_] = $mac[($_%6)]}
		<#$udpclient.send($packet, $packet.length)
		$udpclient.connect(([system.net.ipaddress]::broadcast),9)
		$packet = [byte[]](,0xff * 102)
		6..101 |% { $packet[$_] = $mac[($_%6)]}
		$udpclient.send($packet, $packet.length)
		$udpclient.connect(([system.net.ipaddress]::broadcast),4000)
		$packet = [byte[]](,0xff * 102)
		6..101 |% { $packet[$_] = $mac[($_%6)]}
		$udpclient.send($packet, $packet.length)#>		
		$address = [system.net.ipaddress]::parse($ip)
		$endpoint = new-object system.net.ipendpoint $address, "9"
		for($i = 0;$i -lt 5;$i++){
			$udpclient.send($packet, $packet.length, $endpoint) | out-null
		}
		$logmsg	="`t$computername is booting up...$nl";	
		aspwritetext $logmsg green
		if($global:gui){$global:lvresult.rows.add("$computername","is booting up...")}
		if($log){ aspwritelog $logmsg }
	}
	<#$i=0
	while(-not (test-connection -buffersize 32 -count 1 -computername $computername -quiet)){
		asptestconnectivity -computername $computername  -show 1 
		$i++
		if($i -eq 10){break;}
	}
	asptestconnectivity -computername $computername  -show 1 #>
	checkquit; if($global:gui){refreshgui;}
}

function aspgetadinfos{	
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	setcolumn -columncount 7 -columns "computername","last log on","when changed","when created","distinguished name","canonical name","organisation unit";				
	$ou = "";	
	if ((get-module) -notcontains "activedirectory") { import-module activedirectory }
	$computer = "'$computername'"
	$adinfos = get-adcomputer -filter "name -eq $computer" -prop *
	$lastlogondate = $adinfos.lastlogondate
	$whenchanged = $adinfos.whenchanged
	$whencreated = $adinfos.whencreated
	$canonicalname = $adinfos.canonicalname
	$dsn = $adinfos.distinguishedname
	if(($dsn -ne $null) -and ($dsn -ne "")){
		$cn = $dsn.substring(0,$dsn.indexof(",")+1)
		$cnlength = $dsn.length - $dsn.substring(0,$dsn.indexof(",")+1).length
		$newdsn =  $dsn.substring($cn.length,$cnlength)
		$ou =  $newdsn.substring(3, $newdsn.indexof(",")-3)
	}
	aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
	aspwritetext " $computername : $nl" cyan
	aspwritetext "`tlast log on:","`t$lastlogondate $nl" yellow, green
	aspwritetext "`twhen changed:","`t$whenchanged $nl" yellow, green
	aspwritetext "`twhen created:","`t$whencreated $nl" yellow, green
	aspwritetext "`tdistinguished name:","`t$dsn $nl" yellow, green
	aspwritetext "`tcanonical name:","`t$canonicalname $nl" yellow, green
	aspwritetext "`torganisation unit:","`t$ou $nl" yellow, green 
	aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue	
	if($global:gui){$global:lvresult.rows.add("$computername","$lastlogondate","$whenchanged","$whencreated","$dsn","$canonicalname","$ou")}	
	checkquit; if($global:gui){refreshgui;}
}

function aspgetupdatecount{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","installed updates";				
		try{
			$servicestatus = (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").status
			if ($servicestatus -eq "stopped") { (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").startservice() }		
			$patches=get-hotfix -computername $computername | select hotfixid, description, installedon | sort-object installedon
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " $computername : $nl" cyan
			aspwritetext "`installed updates:","`t$($patches.count) $nl" yellow, green 
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			if($global:gui){$global:lvresult.rows.add("$computername","$($patches.count)")}
			checkquit; if($global:gui){refreshgui;}
		}catch {
			$msg="$nl an error has occurred $nl"
			#aspwritetext $msg red 
			#$global:lvresult.rows.add("$computername","$msg")
		}	
	}
	checkquit; if($global:gui){refreshgui;};
}

function aspgetupdatelist{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 4 -columns "computername","hotfixid","description","installedon";				
		try{
			$servicestatus = (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").status
			if ($servicestatus -eq "stopped") { (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").startservice() }
			$patches=get-hotfix -computername $computername | select hotfixid, description, installedon | sort-object installedon -descending
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " $computername : $nl" cyan
			$patches | %{
				if($global:gui){$global:lvresult.rows.add("$computername","$($_.hotfixid) ","$($_.description) ","$($_.installedon)")}
				aspwritetext "`hotfixid:`t","$($_.hotfixid) $nl" yellow, green			
				aspwritetext "`description:`t","$($_.description) $nl" yellow, green			
				aspwritetext "`installedon:`t","$($_.installedon) $nl" yellow, green
			}
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
		}catch {
			$msg="$nl an error has occurred $nl"
			aspwritetext $msg red 
			if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}		
		}	
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspsearchhotfix{

[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = "",
			[string]$hotfixid
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 4 -columns "computername","hotfixid","description","installedon";				
		try{
			$servicestatus = (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").status
			if ($servicestatus -eq "stopped") { (gwmi -computername $computername -class win32_service -filter "name='remoteregistry'").startservice() }
			$patches=get-hotfix -computername $computername | select hotfixid, description, installedon | sort-object installedon -descending
			$result=@();
			$msg= "search $hotfixid on $computername$nl" 
			aspwritetext $msg yellow 	
			$patches|%{if($_.hotfixid -match $hotfixid){$result+=$_}}
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			aspwritetext " $computername : $nl" cyan		
			$result | %{
				if($global:gui){$global:lvresult.rows.add("$computername","$($_.hotfixid) ","$($_.description) ","$($_.installedon)")}
				aspwritetext "`hotfixid:`t","$($_.hotfixid) $nl" yellow, green
				aspwritetext "`description:`t","$($_.description) $nl" yellow, green
				aspwritetext "`installedon:`t","$($_.installedon) $nl" yellow, green
			}
			aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
		}catch {
			$msg="$nl an error has occurred $nl"
			aspwritetext $msg red 
			if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}			
		}	
	}
	checkquit; if($global:gui){refreshgui;};
}

function aspgetprograms{
	[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){			
	try{ $remoteregistry = gwmi -computername $computername -class win32_service -erroraction 'stop' | where-object { $_.name -eq "remoteregistry" } }
	catch{aspwritetext "$nl an error has occurred $nl" red }	
	if ($remoteregistry.state -eq "stopped" -and $computername -ne $env:computername) 
	{                        
		$return = $remoteregistry.changestartmode("automatic") 
		if($return.returnvalue -eq 0) {
			$remoteregistry.startservice()
			start-sleep -seconds 5                     
		}
	}
				
	aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
	aspwritetext " $computername : $nl" cyan		
	$uninstallkey="software\\microsoft\\windows\\currentversion\\uninstall" 
	$reg=[microsoft.win32.registrykey]::openremotebasekey('localmachine',$computername) 
	$regkey=$reg.opensubkey($uninstallkey) 	
	if ($regkey -ne $null)
	{	
		$subkeys=$regkey.getsubkeynames() 		 
		foreach($key in $subkeys)
		{
			$thiskey=$uninstallkey+"\\"+$key 
			$thissubkey=$reg.opensubkey($thiskey) 
			if ($($thissubkey.getvalue("displayname")) -ne $null){
				aspwritetext "`tname:","`t$($thissubkey.getvalue("displayname")) $nl" yellow, green
				aspwritetext "`tversion:","`t$($thissubkey.getvalue("displayversion")) $nl" yellow, green
				aspwritetext "`tinstalldate:","`t$($thissubkey.getvalue("installdate")) $nl" yellow, green
				aspwritetext "`tpublisher:","`t$($thissubkey.getvalue("publisher")) $nl" yellow, green
				aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
			}
		} 		
	
		# 32bit app installed on a 64 bit    
		$uninstallkey="software\\wow6432node\\microsoft\\windows\\currentversion\\uninstall"
		$reg=[microsoft.win32.registrykey]::openremotebasekey('localmachine',$computername)			
		$regkey=$reg.opensubkey($uninstallkey)		
		if ($regkey -ne $null)
		{		 
			$subkeys=$regkey.getsubkeynames()		 
			foreach($key in $subkeys){		 
				$thiskey=$uninstallkey+"\\"+$key  
				$thissubkey=$reg.opensubkey($thiskey)  
				if ($($thissubkey.getvalue("displayname")) -ne $null){
				aspwritetext "`tname:","`t$($thissubkey.getvalue("displayname")) $nl" yellow, green
				aspwritetext "`tversion:","`t$($thissubkey.getvalue("displayversion")) $nl" yellow, green
				aspwritetext "`tinstalllocation:","`t$($thissubkey.getvalue("installlocation")) $nl" yellow, green
				aspwritetext "`tinstalldate:","`t$($thissubkey.getvalue("installdate")) $nl" yellow, green
				aspwritetext "`tpublisher:","`t$($thissubkey.getvalue("publisher")) $nl" yellow, green
				aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
				}
			}
		} 
		aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue		
	}
}
	checkquit; if($global:gui){refreshgui;}	
}


function aspgetopensessions{ 
   [cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){	   
        $serversessions = @();$serversessions2 = @()  
        $server = [ADSI]"WinNT://$($computername)/LanmanServer" 
        try{$sessions = $server.psbase.invoke("sessions") }catch{}  
        foreach ($session in $sessions) 
        { 
            try 
            { 
                $usersession = new-object -typename psobject -property @{ 
                    user = $session.gettype().invokemember("user","getproperty",$null,$session,$null) 
                    computer = $session.gettype().invokemember("computer","getproperty",$null,$session,$null) 
                    connecttime = $session.gettype().invokemember("connecttime","getproperty",$null,$session,$null) 
                    ideltime = $session.gettype().invokemember("idletime","getproperty",$null,$session,$null) 					
                    } 
                } 
            catch 
            { 
				aspwritetext "$nl an error has occurred $nl" red 
				
            } 
            $serversessions += $usersession 
        }
			$eventid = 1149
			$logname = 'microsoft-windows-terminalservices-remoteconnectionmanager/operational' 
		try 
		{ 
			$events = get-winevent -logname $logname -computername $computername -erroraction $erroractionpreference |where-object {$_.id -eq $eventid} 
			
				foreach ($event in $events) 
				{ 
					$loginattempt = new-object -typename psobject -property @{ 
						user = $event.properties[0].value 
						domain = $event.properties[1].value 
						sourcenetworkaddress = [net.ipaddress]$event.properties[2].value 
						timecreated = $event.timecreated 
						} 
					$serversessions2 += $loginattempt 
					} 
				} 						 
		catch { aspwritetext "$nl an error has occurred $nl" red } 				
        aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue
		aspwritetext " $computername : $nl" cyan		
        $serversessions 
		$serversessions2
        aspwritetext "`t-------------------------------------------------------------------------------------------------- $nl" blue 
    } 
	checkquit; if($global:gui){refreshgui;}
}

function aspcopyx{ 
   [cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		#$sourcefile="c:\wsus.ps1"
		if($msg.length -gt 0){ $sourcefile=$msg}
		$dest="\\$($computername)\c$\Windows\asp-zid\"
		if((test-Path $dest))
		{
			aspwritetext "`ntry to copy $sourcefile to $dest `n" yellow
			try{copy-item -path $sourcefile -dest $dest -force}catch{
			aspwritetext "$nl an error has occurred $nl" red 
			}
			aspwritetext "`ncopy $sourcefile to $dest successfully `n" green
		}
	}
}

function aspgetmappedDrive{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 3 -columns "computername","drive","share";				
		aspwritetext "$nl $computername : " cyan; 
		try{ $mdriver=gwmi -computername $computername -class win32_mappedlogicaldisk -erroraction 'stop' }
		catch{
			aspwritetext "$nl an error has occurred $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","$nl an error has occurred $nl")}
		}
		$i=0;
		if($mdriver.count -eq 0){aspwritetext "no drives are mapped $nl";if($global:gui){$global:lvresult.rows.add("$computername","no drives are mapped")}}
		else{
			$mdriver|%{if($i -eq 0){aspwritetext "$nl";};$i++;aspwritetext "`t$($_.deviceid) $($_.providername) $nl" green;if($global:gui){$global:lvresult.rows.add("$computername","$($_.deviceid)"," $($_.providername)")}}
		}
        aspwritetext "-------------------------------------------------------------------------------------------------- $nl" blue 			
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspgetMemory{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 4 -columns "computername","free physical memory MB","used physical memory MB","total MB";				
		try{ 
			$objmem = gwmi -computername $computername -class win32_operatingsystem -erroraction 'stop' 
			if( $objmem -eq $null )
			  {
				$res = "uncertain: unable to connect. please make sure that powershell and wmi are both installed on the monitered system. also check your credentials"
				aspwritetext $res red
				#$global:lvresult.rows.add("$computername","	$res")			
			  }
			  else{			 
				$freemb = [math]::round( ( $objmem.freephysicalmemory / 1024 ), 0 )
				$totalmb = [math]::round( ( $objmem.totalvisiblememorysize / 1024 ), 0 )
				$usedmb = $totalmb - $freemb
				aspwritetext "$nl $computername : $nl" cyan; 				
				aspwritetext "`tfree physical memory:","`t$($freemb) mb $nl" yellow, green
				aspwritetext "`tused physical memory:","`t$($usedmb) mb $nl" yellow, green
				aspwritetext "`ttotal:","`t$($totalmb) mb $nl" yellow, green	
				if($global:gui){$global:lvresult.rows.add("$computername","$freemb "," $usedmb"," $totalmb")}
				aspwritetext "-------------------------------------------------------------------------------------------------- $nl" blue 	
			}		
		}
		catch{
			aspwritetext "$nl an error has occurred $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","$nl an error has occurred $nl")}
		}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspUninstall-Hotfix{
	[cmdletbinding()]
	param(
		$computername = $env:computername,
		[string] $hotfixid
	) 
	setcolumn -columncount 2 -columns "computername","result";				
	$hotfixes = gwmi -computername $computername -class win32_quickfixengineering -erroraction 'stop' | select hotfixid 
	if($hotfixes -match $hotfixid) {
		$hotfixid = $hotfixid.replace("kb","")
		$msg= "found the hotfix kb" + $hotfixid
		aspwritetext $msg yellow
		$msg= "uninstalling the hotfix"
		aspwritetext $msg yellow
		if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}
		$uninstallstring = "cmd.exe /c wusa.exe /uninstall /kb:$hotfixid /quiet /norestart"
		([wmiclass]"\\$computername\root\cimv2:win32_process").create($uninstallstring) | out-null  
		while (@(get-process wusa -computername $computername -erroraction $erroractionpreference).count -ne 0) {
			start-sleep 3
			$msg= "waiting for update removal to finish ..."
			aspwritetext $msg
			if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}
		}
		$msg= "completed the uninstallation of $hotfixid"
		aspwritetext $msg
		if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}
	}
	else {
		$msg= "given hotfix($hotfixid) not found"
		aspwritetext $msg
		if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}
		return
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspdelCred{
	setcolumn -columncount 2 -columns "computername","result";				
	$aspcm="c:\windows\temp\$($env:username).dat"
	if (test-path $aspcm){
		try{remove-item $aspcm -force;$msg="$nl delete credentials completed successfully! $nl";aspwritetext $msg;if($global:gui){$global:lvresult.rows.add("$env:computername","$msg	")}}
		catch{
			$msg="$nl an error has occurred $nl";
			aspwritetext $msg red;
			if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}	
		}
	}else{
		$msg="$nl file does not exist! $nl";
		aspwritetext $msg red;
		if($global:gui){$global:lvresult.rows.add("$computername","$msg	")}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspstartcommand{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = "",
			[validaterange(0,1)] 
			[int]$debugview = 0,
			[validaterange(0,1)] 
			[int]$log = 0,
			[string]$cmd = ""
	)
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		if($cmd -eq ""){ if($global:restart -or $global:restartloggedoff){ $cmd = "restart" } elseif($global:stop -or $global:stoploggedoff) { $cmd = "stop" } }
		aspwritetext " start command $cmd $computername $nl" darkcyan
		if($debugview){ write-host "$nl $cmd computer: $computername $nl"  -f yellow }	
		setcolumn -columncount 2 -columns "computername","result";
		try { 
			$global:logmsg = " try $cmd : $computername $nl"
			if($debugview){ aspwritetext $global:logmsg yellow }									
			if($cmd -eq "restart"){ restart-computer -comp $computername -force }
			elseif($cmd -eq "stop"){ stop-computer -comp $computername -force }					
			$logmsg = " $cmd computer: $computername successfully."
			if($log){ aspwritelog $logmsg }
			aspwritetext " $computername $cmd success $nl" green	
			if($global:gui){ if($global:gui){$global:lvresult.rows.add("$computername"," $cmd computer successfully.")}}
		}
		catch{ 
			$errormessage = $_.exception.message				
			$logmsg = " $cmd failed for $computername the error message was $errormessage $nl"
			if($log){ aspwritelog $logmsg }
			aspwritetext $logmsg red
			if($global:gui){ if($global:gui){$global:lvresult.rows.add("$computername","$cmd failed for $computername the error message was $errormessage ") }}
		}		
	}
	checkquit; if($global:gui){refreshgui;}	
}

function aspgetInstallApps{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","result";				
		try{
			aspwritetext "$nl $computername : " cyan; 
			$apps = gwmi -computername $computername win32_product -erroraction 'stop' | select-object name, version, installdate
			$apps
			if($global:gui){$global:lvresult.rows.add("$computername","$apps")}
			aspwritetext "-------------------------------------------------------------------------------------------------- $nl" blue 	
		}
		catch{
			aspwritetext "$nl an error has occurred $nl" red 
			if($global:gui){$global:lvresult.rows.add("$computername","$nl an error has occurred $nl")}
		}
	}
	checkquit; if($global:gui){refreshgui;}
}

function aspremoveRegistry{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)] 
			[string]$computername = ""			
	)
	checkquit; if($global:gui){refreshgui;}	
	if((asptestconnectivity -computername $computername -debugview $debugview -log $log -show 0 ) -eq "yes"){
		setcolumn -columncount 2 -columns "computername","result";						
		aspwritetext "$nl $computername : " cyan;
		<#invoke-command -scriptblock {		
			param (
			[Parameter(Position=0)]
			[string] $computername
			)
            $count = 0
			$swlist = get-childitem "hklm:\software\agent\software\*\*" -erroraction silentlycontinue #| where-object {$_ -match $filter}
			if (!$swlist) {$logmsg="no installed sw found!";write-host $logmsg;if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)};exit}
             $logmsg="max number of failed installations reached for:";write-host $logmsg -foregroundcolor yellow
			 if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)}
			$swmatch = "hklm:\software\agent\software\asp team"
			$swmatch|%{ aspwritetext " $sw $nl" green;if($global:gui){$global:lvresult.rows.add("$computername",$sw);$count += 1 }	}		
			if($pscmdlet.shouldprocess($swmatch,"delete failed counter registry key for all sw-packages on clients"))
            {               
				$logmsg="deleted $swmatch";aspwritetext "deleted $swmatch" green;if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)}
            }
			if ($count -eq 0){$logmsg="no matching sw found!";aspwritetext $logmsg red;if($global:gui){$global:lvresult.rows.add("$computername",$logmsg)}}
		} -computername $computername #>			
			$type = [microsoft.win32.registryhive]::localmachine;			
			try{
				$reg = [microsoft.win32.registrykey]::openremotebasekey($type, $computername)	
				$regkey= $reg.opensubkey("software\agent\software\asp team")
				if($regkey){					
					$subkeynames = $regkey.getsubkeynames()
					$regkey= $reg.opensubkey("software\agent\software",$true)					
					$regkey.deletesubkeytree("asp team")
					$regkey.close()						
					$array = $subkeynames -split "\s+"					
					$logmsg += " removed "					
					$array|%{if($global:gui){$global:lvresult.rows.add("$computername","removed: $_ successfully.");$logmsg +=" $_ "}}					
					$logmsg +=" successfully."			
					aspwritetext " $logmsg $nl" green										
				}else{
					$logmsg = " $registry asp team not exists."			
					aspwritetext " $logmsg $nl" red	
					if($global:gui){$global:lvresult.rows.add("$computername","$logmsg")}
				}
				aspwritetext "-------------------------------------------------------------------------------------------------- $nl" blue 	
			}catch{
				$errormessage = $_.exception.message				
				$logmsg = " remove failed for $computername the error message was $errormessage $nl"
				if($log){ aspwritelog $logmsg }
				aspwritetext "$logmsg $nl" red
			}		
	}
	checkquit; if($global:gui){refreshgui;}
}	

function aspgetclientlist([string]$regex="nothing"){
	
	#if($global:nclist.count -eq 0){
		if(!(test-path $clientfile)){			
			aspwritetext "$nl $clientfile not found - please insert the file-path: " yellow
			$path = read-host
			if($path.length -eq 0) { break }
			if(!(test-path $path)){ aspwritetext "$nl wrong again, try it another time.$nl" red ; break }
			else{ $clientfile = $path }
			aspwritetext $nl		
		}
		#}
	if($global:newclientfile.length -gt 0){ $clientfile = $global:newclientfile }
	if($global:clientlist -eq $null -or $global:customcsv -eq 1){#}
		#elseif(($global:clientlist -ne $null) -and (!($global:clientlist.contains("edlovesmoneyandbigtits:)")))){
		write-progress -id 9999 -activity "initialize and loading data" -status "please wait a few seconds";
		$global:clientlist = @();$global:maclist = @{};$global:iplist = @{};
		$data = get-content –path $clientfile  | where-object { $_.trim() -ne '' }
		write-progress -id 9999 -activity "loading $($data.count) computer(s)" -status "please wait a few seconds";
		$count = $null;
			$data | % {
				#$dot+=".";
				#write-progress -id 9999 -activity "initialize and loading data" -status "please wait a few seconds$dot";
				$fields = $_.split(',');if(($regex -ne "") -and ($fields[0] -match $regex) -xor ($regex -eq "nothing")){							
					$count++;$global:clientlist += ,@($count,$fields[0],$fields[1],$fields[2]);
					$global:maclist[$fields[0]] =  $fields[1]
					$global:iplist[$fields[0]] =$fields[2]
				}				
			}
	#$global:clientlist+="edlovesmoneyandbigtits:)";
	$data=$null;
	write-progress -id 9999 -completed -activity "completed" -status "completed";
	aspwritetext "$nl client(s) total: $($global:clientlist.count) $nl" green;aspwritetext  "$nl";				
	}		
}	

function aspjobs{
	param(
		[string]$regex = " "
	)
	$starttime = get-date;
	[int]$func = 0
	[int]$global:restart = 0
	[int]$global:stop = 0
	[int]$help = 0
	[int]$send = 0
	[int]$test = 0
	[int]$log = 0
	[int]$sec = 0	
	[int]$lastlogon = 0
	[int]$logon = 0
	[int]$profiles = 0
	[int]$wsusupdates = 0
	[int]$services = 0
	[int]$cdrive = 0
	[int]$space = 0
	[int]$rdp = 0
	[int]$force = 0
	[int]$comp = 0
	[int]$ip = 0
	[int]$mac = 0
	[int]$lastboot = 0
	[int]$userlang = 0
	[int]$env = 0
	[int]$port = 0
	[int]$member = 0
	[int]$pkey = 0
	[int]$tasks = 0
	[int]$getservices = 0
	[int]$setservice = 0
	[int]$installdate = 0
	[int]$eris = 0
	[int]$proc = 0
	[int]$shares = 0
	[int]$csv = 0
	[int]$setup = 0
	[int]$adinfos = 0
	[int]$vsedrive = 0
	[int]$ie = 0
	[int]$updcount=0
	[int]$updlist=0
	[int]$programs=0
	[int]$opensessions=0
	[int]$copys=0
	[int]$mdrive=0
	[int]$memo=0
	[int]$uninstall=0
	[string]$clients = ""
	[string]$global:client = ""
	[int]$debugview = 0	
	[string]$msg = "please log off."
	[int]$installapps=0
	[int]$remreg=0
	[int]$dontshowlist=0
	[int]$fill=0
	[int]$noshell = 0
	[int]$noregex = 0
	[int]$instsw = 0
	[int]$cpu = 0
	[int]$moni = 0
	[int]$psv = 0
	[int]$nsv = 0
	[int]$printer = 0
	[int]$rp = 0
	[int]$netstat = 0
	[int]$erislog = 0
	[int]$erislogon = 0
	[int]$erislogoff = 0
	[int]$global:restartloggedoff = 0
	[int]$global:stoploggedoff  = 0
	[int]$global:userloggedin = 0
	$jobs = @()
	$global:clientlistcount = 0	

	function showarguments {
  aspwritetext @"
	$nl
optional parameters:
		
-test ...  sends ICMP echo request packets 
-restart ... reboot the operating system
-force ... force reboot or shutdown
-stop ... shutdown the operating system
-send ... send a message 
-debug ... with shell output, what´s going on
-log ... log every working process
-logon ... show currently logged in user 
-lastlogon ... show the last registered user 
-profiles ... show all existing profiles
-wsusupdates ... show asp, ms wsusupdates, wsus report
-services ... start remote services management
-users ... start remote user management
-events ... start remote event management
-cdrive ... open c drive
-ladmins ... view local admins group and user
-space ... view total and freespace disk
-rdp ... start remote desktop
-comp ... start remote computer management
-ip ... show ip address
-mac ... show mac address
-lastboot ... show lastboot and uptime
-userlang ... show existing user profile mui language
-env ... show all existing environment variables
-port ... test port via tcp
-member ... show domain membership
-pkey ... show the product key and os information
-tasks ... show the informations about scheduled tasks
-getservices ... show all services
-setservice [servicename action(stop,start,enable,disable)] ... stop,start,enable or disable service
-installdate ... get the system installation date
-shares ... get all shares and the share permissions for each share
-stopproc ... stop process [processname]
-proc ... get process
-csv ... import an individuell source witch client(s) for current session
-setup ... del all setup.inf´s -setup c -force
-wake ... sends a specified number of magic packets to a mac address in order to wake up the machine
-adinfos ... get computer infos via ldap 
-ie ... get internet explorer version
-updcount ... get installed ms update count
-updlist ... get installed ms update list
-programs ... get all installed programs
-getsessions ... get all open sessions
-mdrive ... get all mapped drives 
-memo ... get memory usage
-uninstall ... uninstall hotfix -uninstall kb2918614
-supdate ... search hotfix 
-installapps ... get all install applications
-remreg ... remove subkeytree asp team 
-wsusfailcount ... show wsusfailcount
-cpu ... show cpu load sort descending by cpupercent
-fw ... show firewall status
-moni ... show monitor informations
-psv ... show powershell version
-nsv ... show .net version
-printer ... show printer and printer status
-rp ... show reboot pending
-netstat ... show networkig activity
-erislog ... show eris log status (on/off)
-erislogon ... set eris log on
-erislogoff ... set eris log off
-restartlo ... reboot the operating system only if user logged off
-stoplo ... shutdown the operating system only if user logged off
-dontshowlist ... do not show extra computer search result list


"@
#for more detailed information - just execute:	
#$nl
#aspwritetext "aspgethelp $nl $nl" yellow 
}		
	$args | %{
		switch ($_) {
			-help { $help = 1; showarguments }
			-restart { [int]$global:restart = 1;$func =1 }
			-stop { [int]$global:stop = 1;$func =1 }
			-send { [int]$send = 1;$func =1 }
			-test { [int]$test = 1;$func =1 }
			-debug { [int]$debugview = 1 } 
			-log { [int]$log = 1 } 
			-lastlogon { [int]$lastlogon = 1;$func =1 }
			-logon { [int]$logon = 1;$func =1 }
			-profiles { [int]$profiles = 1;$func =1 }
			-wsusupdates { [int]$wsusupdates = 1;$func =1 }
			-services { [int]$services = 1;$func =1 }
			-users { [int]$users = 1;$func =1 }
			-events { [int]$events = 1;$func =1 }
			-cdrive { [int]$cdrive = 1;$func =1 }
			-vsedrive { [int]$vsedrive = 1;$func =1 }
			-ladmins { [int]$ladmins = 1;$func =1 }
			-space { [int]$space = 1;$func =1 }
			-rdp { [int]$rdp = 1;$func =1 }
			-comp { [int]$comp = 1;$func =1 }
			-ip { [int]$ip = 1;$func =1 }
			-mac { [int]$mac = 1;$func =1 }
			-lastboot { [int]$lastboot = 1;$func =1 }
			-userlang { [int]$userlang = 1;$func =1 }
			-env { [int]$env = 1;$func =1 }
			-port { [int]$port = 1;$func =1 }
			-member { [int]$member = 1;$func =1 }
			-pkey { [int]$pkey = 1;$func =1 }
			-tasks { [int]$tasks = 1;$func =1 }
			-getservices { [int]$getservices = 1;$func =1 }
			-setservice { [int]$setservice = 1;$func =1 }
			-eris { [int]$eris = 1;$func =1 }
			-shares { [int]$shares = 1;$func =1 }
			-installdate { [int]$installdate = 1;$func =1 }
			-stopproc { [int]$stopproc = 1;$func =1 }
			-proc { [int]$proc = 1;$func =1 }
			-csv { [int]$csv = 1;$func =1 }
			-setup { [int]$setup = 1;$func =1 }
			-adinfos { [int]$adinfos = 1;$func =1 }
			-wake { [int]$wake = 1;$func =1 }
			-force { [int]$force = 1;$func =1 }
			-ie { [int]$ie=1;$func =1 }
			-updcount { [int]$updcount=1;$func =1 }
			-updlist { [int]$updlist=1;$func =1 }
			-programs { [int]$programs=1;$func =1 }
			-getsessions { [int]$getsessions=1;$func =1 }
			-copyx { [int]$copyx=1;$func =1 }
			-mdrive { [int]$mdrive=1;$func =1 }
			-memo { [int]$memo=1;$func =1 }
			-uninstall { [int]$uninstall=1;$func =1 }
			-supdate { [int]$supdate=1;$func =1 }
			-installapps { [int]$installapps=1;$func =1 }
			-remreg { [int]$remreg=1;$func =1 }
			-dontshowlist { [int]$dontshowlist=1;$func =1 }
			-gui { [int]$gui=1;$func =1 }
			-fill { [int]$fill=1}
			-noshell { [int]$noshell=1}
			-noregex { [int]$noregex=1}
			-wsusfailcount { [int]$wsusfailcount = 1;$func =1 }
			-instsw { [int]$instsw = 1;$func =1 }
			-cpu { [int]$cpu = 1;$func =1 }
			-fw { [int]$fw = 1;$func =1 }
			-moni { [int]$moni = 1;$func =1 }
			-psv { [int]$psv = 1;$func =1 }
			-nsv { [int]$nsv = 1;$func =1 }
			-printer { [int]$printer = 1;$func =1 }
			-rp { [int]$rp = 1;$func =1 }
			-netstat { [int]$netstat = 1;$func =1 }
			-erislog { [int]$erislog = 1;$func =1 }
			-erislogon { [int]$erislogon = 1;$func =1 }
			-erislogoff { [int]$erislogoff = 1;$func =1 }
			-restartlo { [int]$global:restartloggedoff = 1;$func =1 }
			-stoplo { [int]$global:stoploggedoff = 1;$func =1 }
			default { if($_ -match "^[\d\.]+$"){ $sec = $_ }elseif($_ -ne ""){ [string]$msg = $_ } }
		}
	}
	#set-location $env:homedrive -passthru;	
	if($csv -eq 1){
		if($global:gui){$path=$global:inputfile;}
		else{	
			aspwritetext "$nl please insert the csv-file-path: " yellow
			$path = read-host		
			while(($path.length -eq 0) -or !(test-path $path)){
				aspwritetext "$nl wrong again, try it another time.$nl" red
				aspwritetext "$nl please insert the csv-file-path: " yellow
				$path = read-host			
			}
		}		
		if($path.length -gt 0){ 
			$global:newclientfile = $path
			$global:customcsv=1;
			aspwritetext "$nl set csv-file: $global:newclientfile successfully " green
			aspwritetext $nl
			aspgetclientlist;
		}			
	}	
	if($help -eq 0){
		function checkanswer([string]$control="control"){
			aspwritetext "$nl try to $control computer: $($global:resultset.count) pc(s) $nl"  yellow 
			aspwritetext "$nl [y] yes  [n] no (default is 'y'): "  yellow 
			$answer = read-host
			if($answer.length -eq 0){ $answer = "y" }
			if($answer.tolower() -eq "y"){ return $true }
			else { return $false }	
		}
		#$global:joblist = @()
		#$jobs = get-job
		#if($jobs.count -gt 0){ $jobs | % { if(($_.state -eq "completed") -or ($_.state -eq "stopped") -or ($_.state -eq "failed")){remove-job $_.id} } }
		
		#if($global:deploymanager -eq 0){
			#aspgetclientlist;	
			if($global:customcsv -eq 1){ aspwritetext "$nl ATTENTION this is the individual csv-file session! $nl" magenta;}
			if($global:clientlist -eq $null -or $global:clientlist.count -lt 0){aspgetclientlist;}
			aspwritetext "$nl client(s) total: $($global:clientlist.count) $nl";aspwritetext  "$nl";			
			$global:resultset = @();
			$count = $null;
			#write-progress -id 9999 -activity "searching computer with regex: $regex" -status "search...";
			#foreach ( $i in ( 0 .. $global:clientlist.count -1 ) ) { 		
			#for($i = 0;$i -lt $global:clientlist.count -1;$i++){	
				$global:clientlist|%{
				#for($j = 0;$j -lt $($global:clientlist[$i]).count;$j++){ 
				<#for($i=0; $i -lt $global:clientlist.count; $i++){#>
					#$global:client = $global:clientlist[$i][0]					
					try{
						[string]$client=$_[1];[string]$clientmac=$_[2];[string]$clientip=$_[3];
						if( $client -eq 0 -and $clientmac -eq 0 -and $clientip -eq 0){break;}					
						if(($noregex -eq 0 -and ($client -match $regex -or $clientip -match $regex)) -xor ($noregex -eq 1 -and ($client -eq $regex -or $clientip -eq $regex))){	
						#if( $global:clientlist[$i][0] -match $regex -or $global:clientlist[$i][2] -match $regex){ 
							#$global:clientlist += $global:client
							$count++;
							$global:resultset += ,@($count, $client, $clientmac,$clientip);
							if($fill){$global:lvmain.rows.add($($client))};
							if($dontshowlist -eq 0){ 
								#write-progress -id 9999 -activity "searching computer with regex: $regex" -status "match: $($global:resultset.count) $($client)" -percentcomplete (( $count / $global:clientlist.count ) * 100);
								aspwritetext "`t $($client) $nl " cyan;
							}
						}
					}catch{
						$errormessage = $_.exception.message
						aspwritetext "$nl incorrect regex syntax: $regex - please try again. - $errormessage $nl"  red
						break
					}    	
				#}			
				}
			#write-progress -id 9999 -completed -activity "completed" -status "completed";				
		#}	
		if($global:customcsv -eq 1){ aspwritetext "$nl ATTENTION this is the individual csv-file session! $nl" magenta;}
		#$global:clientlist=$global:clientlist.getenumerator() | sort-object value
		$global:clientlistcount = $global:resultset.count
		$global:status = "$nl matching client(s):  $global:clientlistcount $nl "
		aspwritetext $global:status;		
		function startcommand{			
			if($global:restart){ $cmd = "restart" }	else { $cmd = "stop" }
			aspwritetext " start command $cmd $global:client $nl" darkcyan
			$global:joblist += start-job -name $global:client -scriptblock{
			param($computername, $debugview, $log, $cmd, $logdir, $logfile)
			$nl = [environment]::newline
			function aspwriteevent{
				[cmdletbinding()]
				param (
						$eventmessage = '',
						$eventtype = 'information',
						$eventid = '9999'
				)			
				if(-not([system.diagnostics.eventlog]::sourceexists('asp'))){[system.diagnostics.eventlog]::createeventsource('asp','application')}
				$eventlog = new-object system.diagnostics.eventlog('application')
				$eventlog.source = 'asp'
				$eventlog.writeentry($eventmessage,$eventtype,$eventid)
			}
			function aspwritelog{
				[cmdletbinding()]
				param (
						[string]$msg = ""		
				)
				if(!$(test-path $logdir)) {new-item $logdir -itemtype directory | out-null}	
				if(!$(test-path $logfile)) {new-item -type file $logfile -force}
				$logrecord = "------------------------------------------------------------$nl  $((get-date).toshortdatestring()) " 
				$logrecord += $msg				
				aspwriteevent $msg 'information' '9999' 
				add-content -path $logfile -encoding utf8 -value $logrecord
			}
			if($debugview){ write-host "$nl $cmd computer: $computername $nl"  -f yellow }
			$ping = new-object system.net.networkinformation.ping
				try{ $pingreturns = $ping.send($computername, 1000) }
				catch{ 	
					$logmsg = "Ping request could host $computername not found."
					aspwritelog $logmsg
					write-host " $logmsg" -f red
					$global:resultlist += "Ping request could host $computername not found."
				}					  
				if($pingreturns.status -ne "success")
				{				
					$logmsg = " problem connecting to computer: $computername."
					aspwritelog $logmsg
					write-host " $computername down" -f red	
					$global:resultlist += " problem connecting to computer: $computername."					
				}
				else{
					try { 
						$global:logmsg = " try $cmd : $computername $nl"
						if($debugview){ write-host $global:logmsg -f yellow }									
						if($cmd -eq "restart"){ restart-computer -comp $computername -force }
						elseif($cmd -eq "stop"){ stop-computer -comp $computername -force }					
						$logmsg = " $cmd computer: $computername successfully."
						aspwritelog $logmsg
						write-host " $computername $cmd success"  -f green	
						$global:resultlist += " $cmd computer: $computername successfully."	
					}
					catch{ 
						$errormessage = $_.exception.message				
						$logmsg = " $cmd failed for $computername the error message was $errormessage "
						aspwritelog $logmsg
						write-host " $computername $cmd failed."  -f red
						$global:resultlist += " $cmd failed for $computername the error message was $errormessage "						
					}		
				}		
			} -argumentlist $global:client, $debugview, $log, $cmd, $global:logdir, $global:logfile
		}	
		function getjobdescription{		
			param (	[string]$job = "job" )
			$action = "command"
			#if($job -eq "shutdown" -or $job -eq "reboot"){ $action = "job(s)" }			
			$global:jobdescription = "$action $job to one or more client(s)"
			#aspwritetext "$nl start $global:jobdescription : $nl $nl" magenta
		}						
		function nightlatch($cmd){
			if($global:resultset.count -gt 0){
				if(checkanswer($cmd)){
					aspwritetext $nl
					$global:logmsg = " $((get-date).toshortdatestring()), $cmd computer for '$regex', $($global:resultset.count) pc(s): $nl$nl" 
					aspwritetext $global:logmsg cyan
					aspwritelog $global:logmsg
				}else{
					$global:logmsg = " process $cmd was aborted. " 
					aspwritetext $global:logmsg  red
					aspwritelog $global:logmsg
					break
				}
			}
		}	
		aspwritetext "$nl-------------------------------------------------------------------------------------------------------------$nl";		
	if($args -gt 0 -or $func -eq 1){
		$i=0;
		 #for($i=0; $i -lt $global:clientlistcount; $i++){
		 $global:resultset|%{
			$global:client=$_[1];[string]$global:mac=$_[2];[string]$global:ipaddress=$_[3];
			checkquit; if($global:gui){refreshgui;}
			if($send){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.send.text)}; aspsendmsg -message $msg -computername $global:client -debugview $debugview -log $log 
			}	
			if($test){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.test.text)}; 
				if($global:gui){
					asptestconnectivity -computername $global:client -debugview $debugview -log $log
					checkquit;if($global:gui){refreshgui;}	
					}
				else{
					if($sec -eq 0){$sec=1}	
					for($j=0;$j -lt $sec;$j++){
						if($debugview){ aspwritetext " try connecting to computer: $global:client $nl" yellow  }	
						  try{
							$ping = new-object system.net.networkinformation.ping
							$pingreturns = $ping.send($global:client, 1000) 
							if($pingreturns.status -eq "success"){						  
								$logmsg = " $global:client up"
								aspwritetext " $global:client"," up $nl"cyan, green;
								if($log){ aspwritelog $logmsg }; 				
							}
							else{
								$logmsg = " problem connecting to computer: $global:client."
								aspwritetext " $global:client ","down $nl" cyan, red;
								if($log){ aspwritelog $logmsg }; 
								continue;
							}
							if($sec -gt 1){delay(1)}
							}catch{ 
								$errormessage = $_.exception.message				
								$logmsg = " ping failed for $global:client the error message was $errormessage $nl"
								if($log){ aspwritelog $logmsg }
								aspwritetext $logmsg red
								continue;	
							}
						checkquit;if($global:gui){refreshgui;}	 
					}					
				}			
			}		
			if($global:stop){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.stop.text)};if($force -eq 0){nightlatch('stop') };
				showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start shutdown to computer: $global:client" -duration 5000 ; 
				if($sec -gt 0){ delay($sec) }; aspstartcommand -computername $global:client -debugview $debugview -log $log #startcommand
			}
			if($global:stoploggedoff){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.stop.text)};if($force -eq 0){nightlatch('stop') };
				aspgetlogon -computername $global:client;
				if($global:userloggedin){
				aspwritetext "$nl user logged in - skip shutdown to computer: $global:client $nl$nl" red;
				if($global:gui){$global:lvresult.rows.add("$global:client","user logged in - skip shutdown");}
				}elseif($global:userloggedin -eq 0){showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start shutdown to computer: $global:client" -duration 5000; 
				if($sec -gt 0){ delay($sec) }; aspstartcommand -computername $global:client -debugview $debugview -log $log} #startcommand
			}
			if($global:restart){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.restart.text)}; if($force -eq 0){ nightlatch('restart') };showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start reboot to computer: $global:client" -duration 5000 ; 
				if($sec -gt 0){ delay($sec) }; aspstartcommand -computername $global:client -debugview $debugview -log $log #startcommand
			}
			if($global:restartloggedoff){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.restart.text)}; if($force -eq 0){ nightlatch('restart') };				
				aspgetlogon -computername $global:client;
				if($global:userloggedin){
				aspwritetext "$nl user logged in - skip reboot to computer: $global:client $nl$nl" red;
				if($global:gui){$global:lvresult.rows.add("$global:client","user logged in - skip restart");}
				}elseif($global:userloggedin -eq 0){showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start reboot to computer: $global:client" -duration 5000; 
				if($sec -gt 0){ delay($sec) }; aspstartcommand -computername $global:client -debugview $debugview -log $log} #startcommand
			}
			if($erislogoff){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.erislogoff.text)}; aspseterislogoff -computername $global:client
			}
			if($erislogon){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.erislogon.text)}; aspseterislogon -computername $global:client
			}
			if($erislog){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.erislog.text)}; aspgeterislogstatus -computername $global:client
			}
			if($netstat){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.netstat.text)}; aspgetnetstat -computername $global:client
			}
			if($rp){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.rp.text)}; aspgetrebootpending -computername $global:client
			}
			if($printer){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.printer.text)}; aspgetprinter -computername $global:client
			}
			if($nsv){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.nsv.text)}; aspgetnetversion -computername $global:client
			}
			if($psv){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.psv.text)}; aspgetpsversion -computername $global:client
			}
			if($moni){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.moni.text)}; aspgetmonitor -computername $global:client
			}
			if($fw){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.fw.text)}; aspgetfwstatus -computername $global:client
			}
			if($cpu){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.cpu.text)}; aspgetcpuload -computername $global:client
			}
			if($instsw){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.instsw.text)}; aspgetinstalledsw -computername $global:client
			}
			if($logon){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.logon.text)}; aspgetlogon -computername $global:client
			}
			if($lastlogon){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.lastlogon.text)}; aspgetlastlogon -computername $global:client
			}
			if($wsusfailcount){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.wsusfailcount.text)}; aspgetwsusfailcount -computername $global:client
			}
			if($profiles){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.profiles.text)}; aspshowprofiles -computername $global:client
			}
			if($wsusupdates){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.wsusupdates.text)}; aspviewwsusupdates -computername $global:client -logselection $sec
			}
			if($services){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.services.text)}; aspviewservices -computername $global:client
			}
			if($events){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.events.text)}; aspvieweventvwr -computername $global:client 
			}
			if($users){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.users.text)}; aspviewuser -computername $global:client
			}
			if($ladmins){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.ladmins.text)}; aspviewlocaladmins -computername $global:client
			}
			if($cdrive){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.cdrive.text)}; aspcdrive -computername $global:client
			}			
			if($vsedrive){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.cdrive.text)}; aspvsedrive -computername $global:client
			}
			if($space){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.space.text)}; aspviewfreespace -computername $global:client
			}
			if($rdp){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.rdp.text)}; asprdp -computername $global:client
			}
			if($comp){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.comp.text)}; aspcompmgmt -computername $global:client
			}
			if($ip){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.ip.text)}; aspgetipaddress -computername $global:client
			}
			if($mac){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.mac.text)}; aspgetmacaddress -computername $global:client
			}
			if($lastboot){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.lastboot.text)}; aspgetlastboot -computername $global:client
			}
			if($userlang){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.userlang.text)}; aspgetuserlang -computername $global:client
			}
			if($env){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.env.text)}; aspgetenv -computername $global:client
			}
			if($port){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.port.text)}; asptestport -computername $global:client -port $sec
			}
			if($member){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.member.text)}; aspgetmember -computername $global:client
			}
			if($pkey){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.pkey.text)}; aspgetproductkey -computername $global:client
			}
			if($tasks){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.tasks.text)}; aspgettasks -computername $global:client
			}
			if($getservices){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.getservices.text)}; aspgetservices -computername $global:client
			}
			if($setservice){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.setservice.text)}; aspsetservice -computername $global:client
			}
			if($eris){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.eris.text)}; asprestarteris -computername $global:client -log $log
			}
			if($installdate){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.installdate.text)}; aspgetinstalldate -computername $global:client
			}
			if($shares){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.shares.text)}; aspgetshares -computername $global:client
			}
			if($stopproc){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.stopproc.text)}; aspstopprocess -computername $global:client
			}
			if($proc){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.proc.text)}; aspgetprocess -computername $global:client
			}
			if($setup){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.setup.text)};aspdelsetupinf -computername $global:client -drive $msg
			}
			if($wake){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.wake.text)}; if($sec -gt 0){ delay($sec) }; aspwake -computername $global:client -log $log
			}
			if($adinfos){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.adinfos.text)}; aspgetadinfos -computername $global:client
			}
			if($ie){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.ie.text)}; aspgetInternetExVersion -computername $global:client
			}
			if($updcount){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.updcount.text)}; aspgetupdatecount -computername $global:client
			}
			if($updlist){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.updlist.text)}; aspgetupdatelist -computername $global:client
			}
			if($programs){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.programs.text)}; aspgetprograms -computername $global:client
			}
			if($getsessions){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.getsessions.text)}; aspgetopensessions -computername $global:client
			}
			if($copyx){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.copyx.text)}; aspcopyx -computername $global:client
			}
			if($mdrive){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.mdrive.text)}; aspgetmappedDrive -computername $global:client
			}
			if($memo){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.memo.text)}; aspgetMemory -computername $global:client
			}
			if($uninstall){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.uninstall.text)}; aspUninstall-Hotfix -computername $global:client -hotfixid $msg
			}
			if($supdate){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.searchhotfix.text)}; aspsearchhotfix -computername $global:client -hotfixid $sec
			}
			if($installapps){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.installapps.text)}; aspgetInstallApps -computername $global:client 
			}
			if($remreg){
				if($i -eq 0){getjobdescription($global:xml.options.jobdescriptions.remreg.text)}; aspremoveRegistry -computername $global:client 
			}
			$global:client = "";$global:mac = "";$global:ipaddress = "";$i++;
		}
	}
		#$global:clientlistcount = $global:resultset.count
		 if($global:clientlistcount -gt 20){
			aspwritetext "$nl client(s) total: $($global:clientlist.count) $nl"
			aspwritetext " matching client(s): $global:clientlistcount $nl"
		}		
		<#$count=0
		if($global:joblist.count -gt 0){
			$complete = $false
			$jobsinprogress = ""
			while (-not $complete) {	
				$jobsinprogress = $global:joblist | where-object { $jobname = $_.name; $state = $_.state; $_.state -match 'running' }
				if(-not $jobsinprogress) { $complete = $true; aspwritetext "$nl $nl" }
				else{ if($count -eq 0) { aspwritetext "$nl job(s) running $nl" yellow ;$count++ } getprog($true) }#aspwritetext "." yellow; start-sleep 2 }
			}	
		}#>
		<#only experimentation
		if($match -match "\[([^\[]*)\]"){
			$praestring = $match.substring(0,$match.indexof("["))
			$rest = ($match.length - ($match.indexof("]")+1))
			$poststring = $match.substring(($match.indexof("]")+1),$rest)
			$rest = ($match.indexof("]")+1) -  ($match.indexof("[")+1)
			$between = $match.substring(($match.indexof("[")+1),$rest-1)
			$match = "^$praestring([$between])$poststring$"
		}#>
		<#$livecred = get-credential	 
		$j=  invoke-command -computername localhost -credential $livecred -filepath $env:emp_utils\controlclients.ps1 -argumentlist $client,restart,task -asjob
		if($global:clientlistcount -eq 1){ restartcomputer -computername $client }
		else{ controlcommand "restart" }
		$scriptblock = (. $env:emp_utils\restart.ps1 -computername $client -debugview $debugview -log $log -sec $sec)
		$jobs += start-job -name $client -scriptblock{$scriptblock}	
		#write-host "started job: $client"
		foreach ($job in $jobs) {
			while ($job.state -eq 'running') {
				$progress=$job.childjobs[0].progress 
				$desc = $progress | %{$_.statusdescription} 
				write-progress activity "$desc"; start-sleep -sec 1
			}
		}#if($count -eq 0){clear-host}
			#$count++
			#write-host " $nl $jobdescription running $count seconds $nl" -foregroundcolor yellow 
			#start-sleep -seconds 1
			#clear-host
			#get-job | select name, state | ft *			
			#countdown(3)			
			#start-sleep -seconds 5				get-job | where {$_.childjobs[0].id -eq  $job.childjobs[0].id} | select -expand name -outvariable jobinfo | out-null
			#Write-Output "finished on" -outvariable +jobinfo | out-null
			#$jobinfo+= ($job.childjobs[0].psendtime)			
			#(get-content $dumpfile ) | where {$_ -ne ""} | set-content $dumpfile
			#select-string -pattern "\w" -path $dumpfile | foreach {$_.line} | out-file -filepath $dumpfile
			#$content= get-content $dumpfile  | where {$_ -ne ""} | select-object -first 1
			#$content+= get-content $dumpfile  | where {$_ -ne ""} | select-object -last 2
			#if($count -eq 0) { write-host "$nl $jobdescription have completed. $nl" -foregroundcolor green }
			$global:jobcount++
			#write-host "$count. " -nonewline			
			#write-host ($jobinfo)  -nonewline -foregroundcolor cyan			
			#write-host $nl			
			#write-host (receive-job $job)
			<#$joboutput = [string](receive-job -job $job)
			$a = @("exeption", "error", "problem")
			$joboutput.length
			$check = ($joboutput | Select-String -patter $a)
			if($check.Length -gt 0){write-host " problem. $nl" -foregroundcolor red}
			else{write-host " success. $nl" -foregroundcolor green}
		>
		clear-host#>
		<#$global:jobcount = 0
		$error = 0
		$results = @()
		#if(($log) -and ($global:restart)) { start-transcript $global:logfile -append -noclobber | out-null }
		$global:joblist | % {
			try{
				if($_.state -eq 'failed') {
					aspwritetext ($_.childjobs[0].jobstateinfo.reason.message) red
					$error = 1
				} 
				else {
					#get-job | select -expand name -outvariable jobinfo | out-null			
					$global:jobcount++
					#wait-job -name $jobinfo
					$results = receive-job $_ | out-null				
					remove-job -name $_.name | out-null
				}
			}catch{	aspwritetext " could´t receive job - try get-job for more details $nl" red }	
		 }		
		$global:joblist | receive-job 
		if(($log) -and ($global:restart)) { stop-transcript | out-null }
		get-job | % { wait-job -name $_.name; $results += receive-job $_.id}
		if($log){ $logfile = [string]::join('',($logdir, $logfile)); out-file -inputobject $results -filepath $logfile -force }	
		$results | % { if($log){ aspwritelog $_ } }
		if(($error -eq 0) -and ($global:jobcount -gt 0)){ aspwritetext "$nl $global:jobdescription have completed. $nl"  cyan }
		elseif($error -eq 1){ aspwritetext "$nl $global:jobdescription failed. $nl"  cyan }
		elseif($global:jobcount -eq 0){ aspwritetext "$nl no job(s) running. $nl"  yellow }#>
		aspwritetext "$nl-------------------------------------------------------------------------------------------------------------$nl";
		aspwritetext $nl;
		$stoptime = get-date;
		$timerunning = ($stoptime - $starttime).totalseconds;
		if($timerunning -gt 60){$timerunning = ($timerunning / 60); $minsec = "min.";}
		else{$minsec = "sec.";}
		$run = "{0:n2}" -f ($timerunning);
		aspwritetext "$nl script after $run $minsec done $nl $nl" cyan;
		aspwritetext $nl;
	}
}

function aspgetgui{
	[cmdletbinding()]
	param (
		[parameter(mandatory=$false,valuefrompipeline=$true)] 
		[string]$computername = "",
		[int]$select=0
	) 	
	$global:gui=1;	
	function get-translation{
		$btnchlang.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index = 'btnchlang']"| % { $_.Node.InnerText}
		if($global:dlanguage -eq "en"){$btnchlang.text += " DE";}else{$btnchlang.text += " EN"}
		$btnimportcsv.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index = 'btnimportcsv']"| % { $_.Node.InnerText}
		$btninstsw.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index = 'btninstsw']"| % { $_.Node.InnerText}
		$btnrestart.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index = 'btnrestart']"| % { $_.Node.InnerText}
		$btnrestartloggedoff.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index = 'btnrestartloggedoff']"| % { $_.Node.InnerText}
		$btnmsg.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnmsg']"| % { $_.Node.InnerText}
		$btncdrive.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btncdrive']"| % { $_.Node.InnerText}
		$btnvsedrive.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnvsedrive']"| % { $_.Node.InnerText}
		$btnstop.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnstop']"| % { $_.Node.InnerText}
		$btnstoploggedoff.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnstoploggedoff']"| % { $_.Node.InnerText}
		$btnrdp.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnrdp']"| % { $_.Node.InnerText}
		$btneris.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btneris']"| % { $_.Node.InnerText}
		$btnservices.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnservices']"| % { $_.Node.InnerText}
		$btnexit.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnexit']"| % { $_.Node.InnerText}
		$btnprocesses.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnprocesses']"| % { $_.Node.InnerText}
		$btnprofiles.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnprofiles']"| % { $_.Node.InnerText}
		$btnlastlogon.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnlastlogon']"| % { $_.Node.InnerText}
		$btnlastboot.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnlastboot']"| % { $_.Node.InnerText}
		$btninstalldate.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btninstalldate']"| % { $_.Node.InnerText}
		$btnlocaladmins.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnlocaladmins']"| % { $_.Node.InnerText}
		$btnlogon.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnlogon']"| % { $_.Node.InnerText}
		$btncpu.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btncpu']"| % { $_.Node.InnerText}
		$btnfw.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnfw']"| % { $_.Node.InnerText}
		$btnmonitor.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnmonitor']"| % { $_.Node.InnerText}
		$btnieversion.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnieversion']"| % { $_.Node.InnerText}
		$btnpsversion.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnpsversion']"| % { $_.Node.InnerText}
		$btnnetversion.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnnetversion']"| % { $_.Node.InnerText}
		$btnprinter.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnprinter']"| % { $_.Node.InnerText}
		$btnrebootpending.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnrebootpending']"| % { $_.Node.InnerText}
		$btnnetstat.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnnetstat']"| % { $_.Node.InnerText}
		$btnselectall.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnselectall']"| % { $_.Node.InnerText}
		$btnunselectall.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnunselectall']"| % { $_.Node.InnerText}
		$btnsearch.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnsearch']"| % { $_.Node.InnerText}
		$btnusers.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnusers']"| % { $_.Node.InnerText}
		$btnevents.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnevents']"| % { $_.Node.InnerText}
		$btnip.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnip']"| % { $_.Node.InnerText}
		$btnmac.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnmac']"| % { $_.Node.InnerText}
		$btntestcon.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btntestcon']"| % { $_.Node.InnerText}
		$btnspace.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnspace']"| % { $_.Node.InnerText}
		$btntasks.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btntasks']"| % { $_.Node.InnerText}
		$btnshares.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnshares']"| % { $_.Node.InnerText}
		$btnuserlang.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnuserlang']"| % { $_.Node.InnerText}
		$btnwakeup.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnwakeup']"| % { $_.Node.InnerText}
		$btnremreg.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnremreg']"| % { $_.Node.InnerText}
		$btnadinfos.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnadinfos']"| % { $_.Node.InnerText}
		$btnenv.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnenv']"| % { $_.Node.InnerText}
		$btnmdrive.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnmdrive']"| % { $_.Node.InnerText}
		$btnmemory.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnmemory']"| % { $_.Node.InnerText}
		$btndelcred.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btndelcred']"| % { $_.Node.InnerText}
		$btnupdcount.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnupdcount']"| % { $_.Node.InnerText}
		$btndelupdate.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btndelupdate']"| % { $_.Node.InnerText}
		$btnsupdate.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnsupdate']"| % { $_.Node.InnerText}
		$btnupdlist.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnupdlist']"| % { $_.Node.InnerText}
		$btnupdatelogs.text = $global:xml | Select-Xml "//languageKey[@index = '$global:dlanguage']//label[@index ='btnupdatelogs']"| % { $_.Node.InnerText}
		if($global:gui){refreshgui;};
	}
	
	function generateform{		
		[reflection.assembly]::loadwithpartialname("system.windows.forms") | out-null
		[reflection.assembly]::loadwithpartialname("microsoft.visualbasic") | out-null
		[system.windows.forms.application]::enablevisualstyles()
		$formmain = new-object system.windows.forms.form
		$groupadministration = new-object system.windows.forms.groupbox
		$btnimportcsv = new-object system.windows.forms.button
		$btninstsw = new-object system.windows.forms.button
		$btnrestart = new-object system.windows.forms.button
		$btnrestartloggedoff = new-object system.windows.forms.button
		$btnchlang = new-object system.windows.forms.button
		$btnmsg = new-object system.windows.forms.button
		$btncdrive = new-object system.windows.forms.button
		$btnvsedrive = new-object system.windows.forms.button
		$btnstop = new-object system.windows.forms.button
		$btnstoploggedoff = new-object system.windows.forms.button
		$btnrdp = new-object system.windows.forms.button
		$btneris = new-object system.windows.forms.button
		$groupinfo = new-object system.windows.forms.groupbox
		$groupinfo2 = new-object system.windows.forms.groupbox
		$groupadds = new-object system.windows.forms.groupbox
		$groupupdates = new-object system.windows.forms.groupbox
		$btnservices = new-object system.windows.forms.button
		$btnexit = new-object system.windows.forms.button
		$btnprocesses = new-object system.windows.forms.button
		$btnprofiles = new-object system.windows.forms.button
		$btnlastlogon = new-object system.windows.forms.button
		$btnlastboot = new-object system.windows.forms.button
		$btninstalldate = new-object system.windows.forms.button
		$btnlocaladmins = new-object system.windows.forms.button
		$btnlogon = new-object system.windows.forms.button
		$btncpu = new-object system.windows.forms.button
		$btnfw = new-object system.windows.forms.button
		$btnmonitor = new-object system.windows.forms.button
		$btnieversion = new-object system.windows.forms.button
		$btnpsversion = new-object system.windows.forms.button
		$btnnetversion = new-object system.windows.forms.button
		$btnprinter = new-object system.windows.forms.button
		$btnrebootpending = new-object system.windows.forms.button
		$btnnetstat = new-object system.windows.forms.button
		$btnselectall = new-object system.windows.forms.button
		$btnunselectall = new-object system.windows.forms.button
		$global:lvmain = new-object system.windows.forms.datagridview				
		$global:lvmain.autosize = $true
		$global:lvmain.allowsorting = $true
		$global:lvmain.readonly = $true
		$global:lvmain.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
		$global:lvmain.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
		$global:lvmain.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Silver
		$global:lvmain.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
		$global:lvmain.ColumnHeadersDefaultCellSTyle.ForeColor = [System.Drawing.Color]::HighlightText
		$global:lvmain.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::lightsteelblue
		$global:lvmain.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::Tan
		$global:lvmain.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
		#$global:lvmain.captiontext = "result"
		$global:lvmain.headerfont = new-object system.drawing.font("verdana",8.25,1,3,0)
		$global:lvmain.headerforecolor = [system.drawing.color]::fromargb(255,0,0,0)
		$global:lvmain.font = new-object system.drawing.font("verdana",8.25,[system.drawing.fontstyle]::bold)
		$global:lvmain.backcolor = [system.drawing.color]::fromargb(255,0,160,250)
		$global:lvmain.alternatingbackcolor = [system.drawing.color]::fromargb(255,133,194,255)
		$global:lvmain.name = "$global:lvmain"
		$global:lvmain.rowheadersvisible=$false
		$global:lvmain.allowusertoaddrows = $false;
		$global:lvmain.allowusertodeleterows = $false;
		$global:lvmain.readonly = $true;
		$global:lvmain._cellcontentclick($global:lvmain_CellContentClick);
		#$global:lvmain = new-object system.windows.forms.listview		
		$global:lvresult = new-object system.windows.forms.datagridview
		#$global:lvresult.location = new-object system.drawing.size(0,0)
		#$global:lvresult.size = new-object system.drawing.size(592,400)
		$global:lvresult.AllowUserToAddRows = $False
		$global:lvresult.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
		$global:lvresult.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
		$global:lvresult.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Silver
		$global:lvresult.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
		$global:lvresult.ColumnHeadersDefaultCellSTyle.ForeColor = [System.Drawing.Color]::HighlightText
		$global:lvresult.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::lightsteelblue
		$global:lvresult.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::Tan
		$global:lvresult.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
		$global:lvresult.autosize = $true
		$global:lvresult.allowsorting = $true
		$global:lvresult.readonly = $true
		#$global:lvresult.captiontext = "result"
		$global:lvresult.headerfont = new-object system.drawing.font("verdana",8.25,1,3,0)
		$global:lvresult.headerforecolor = [system.drawing.color]::fromargb(255,0,0,0)
		$global:lvresult.font = new-object system.drawing.font("verdana",8.25,[system.drawing.fontstyle]::bold)
		$global:lvresult.backcolor = [system.drawing.color]::fromargb(255,0,160,250)
		$global:lvresult.alternatingbackcolor = [system.drawing.color]::fromargb(255,133,194,255)
		$global:lvresult.name = "lvresult"
		$global:lvresult.rowheadersvisible=$false
		$global:lvresult.allowusertoaddrows = $false;
		$global:lvresult.allowusertodeleterows = $false;
		$global:lvresult.readonly = $true;		
		$btnsearch = new-object system.windows.forms.button
		$btnusers = new-object system.windows.forms.button
		$btnevents = new-object system.windows.forms.button
		$btnip = new-object system.windows.forms.button
		$btnmac = new-object system.windows.forms.button
		$btntestcon = new-object system.windows.forms.button
		$btnspace = new-object system.windows.forms.button
		$btntasks = new-object system.windows.forms.button
		$btnshares = new-object system.windows.forms.button
		$btnuserlang = new-object system.windows.forms.button
		$btnwakeup = new-object system.windows.forms.button
		$btnremreg = new-object system.windows.forms.button
		$btnadinfos = new-object system.windows.forms.button
		$btnenv = new-object system.windows.forms.button
		$btnmdrive = new-object system.windows.forms.button
		$btnmemory = new-object system.windows.forms.button
		$btndelcred = new-object system.windows.forms.button
		$btnupdcount = new-object system.windows.forms.button
		$btndelupdate = new-object system.windows.forms.button
		$btnsupdate = new-object system.windows.forms.button
		$btnupdlist = new-object system.windows.forms.button
		$btnupdatelogs = new-object system.windows.forms.button
		$txtcomputer = new-object system.windows.forms.textbox
		$sb = new-object system.windows.forms.statusbar
		$picturebox = new-object windows.forms.picturebox
		$lblaction = new-object system.windows.forms.label
		$lblselected = new-object system.windows.forms.label
		$lblclient = new-object system.windows.forms.label		
		$lblstatus = new-object system.windows.forms.statusbarpanel
		$lblzid = new-object system.windows.forms.statusbarpanel
		$initialformwindowstate = new-object system.windows.forms.formwindowstate
		$contextmenu = new-object system.windows.forms.contextmenustrip
		$cmsselect = new-object system.windows.forms.toolstripmenuitem
		$cmsunselect = new-object system.windows.forms.toolstripmenuitem
		$tabctrl = new-object system.windows.forms.tabcontrol
		$tabpageadmin = new-object system.windows.forms.tabpage
		$tabpageinfo = new-object system.windows.forms.tabpage
		$tabpageinfo2 = new-object system.windows.forms.tabpage
		$tabpageadds = new-object system.windows.forms.tabpage
		$tabpageupdates = new-object system.windows.forms.tabpage
		
		$formmain_load={
			$formmain.windowstate = $initialformwindowstate;
			$txtcomputer.focus()			
			$vbmsg = new-object -comobject wscript.shell
			set-formtitle			
			$global:lvmain.autoresizecolumns();            
			$global:lvmain.autosizecolumnsmode = datagridviewautosizecolumnsmode.allcells;	
			$global:lvresult.autoresizecolumns();            
			$global:lvresult.autosizecolumnsmode = datagridviewautosizecolumnsmode.allcells;				
			$global:selectedcount = 0
			if($computername -ne "" -and $computername.length -gt 0){ $txtcomputer.text = $computername; &$btnsearch_click
				$global:selectedcount = 0
				#$global:lvmain.items | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::yellowgreen; $global:selected += $global:lvmain.items[$_.index].text; $global:selectedcount++ }
				 if($select -eq 1){set-selectedall;}
				 for($i=0;$i -lt $lmain.rows.count;$i++){
					 if ($global:lvmain.rows[$i].cells[0].selected)
					{
						$global:selected += $global:lvmain.rows[$i].cells[0].value
						$global:selectedcount++
					}
				}
				$lblselected.text = "current selected client(s) $global:selectedcount"
			}
			
			<#if($global:deploymanager){
				$regaspcmname="credmanvault"
				$aspcm="c:\windows\temp\$($env:username).dat"
				#write-host $aspcm
				#try{ $credmanvault = (get-itemproperty -path $regaspcmkey).$regaspcmname }
				if(test-path $aspcm){
					$cm=get-content $aspcm
					if($cm -eq 1){aspwritetext "$nl adm Credentials already exists $nl" cyan}
				}
				else{
					$formmain.refresh()
					$vbmsg = new-object -comobject wscript.shell; $vbmsg.popup($global:xml.options.messages.notifyinginformation.text ,0,"attention !"); 
					control /name Microsoft.CredentialManager
				}
				#new-item -path $regaspcmkey -name $regaspcmname –force
				#set-itemproperty -path $regaspcmkey -name $regaspcmname -value 1 -type dword
				new-item c:\windows\temp\$env:username.dat -type file -force -value "1"
			}#>			
		}
		$onapplicationexit={stop-process -id $pid}
	
	function checkselectedcomputer{
		if($txtcomputer.text -ne "" -and $txtcomputer.text.length -gt 0){ $picturebox.visible = $true;$lblaction.text = "loading computer(s)..."; $cmsselect.visible = $true;  return $true; 
		}elseif($txtcomputer.text -eq ""){ $picturebox.visible = $false;$lblstatus.text = "computername is empty..."; $lblaction.text = "computername is empty..."; $lblselected.text = " ";$cmsselect.visible = $false; return $false; }
	}		
		$txtcomputer_keypress=[system.windows.forms.keypresseventhandler]{ if ($_.keychar -eq [system.windows.forms.keys]::enter){ &$btnsearch_click}}		
		$formmain.add_keydown({
		if ($_.keycode -eq "Enter") { &$btnsearch_click} 
		if($_.keycode -eq "ControlKey") { $global:key = $true } 
		if($global:key -eq $true -and $_.keycode -eq "C" ){ checkquit;}#$vbmsg = new-object -comobject wscript.shell; $vbmsg.popup("script was aborted !" ,0,"script was aborted !"); }
		})						
		$form_statecorrection_load=
		{
			$formmain.windowstate = $initialformwindowstate
		}
	function handleexit{
	if ((get-module) -notcontains "psterminalservices"){ import-module psterminalservices }
				get-tssession -computername localhost -username $env:username | stop-tssession -force		
				$formmain.close()	
	}		
		$btnexit_click={
			handleexit		
		}		
		$cmsselect_click={
			set-selected
		}		
		$cmsunselect_click={
			set-unselected
		}		
		$btnselectall_click={
			set-selectedall
		}		
		$btnunselectall_click={		
			set-unselectedall
		}		
		$btnsearch_click={
			initialize-listviewmain
			#initialize-listviewresult
			$lblstatus.text = "retrieving computers..."		
			#$global:lvmain.items.add		
			if($txtcomputer.text -ne "" -and $txtcomputer.text.length -gt 0){ $picturebox.visible = $true;$lblaction.text = "loading computer(s)..."; 
			checkquit; if($global:gui){refreshgui;}
			$global:lvmain.columncount=1
			$global:lvmain.Columns[0].Name ="computername"
			$global:lvmain.autoresizecolumns();            
			$global:lvmain.autosizecolumnsmode = datagridviewautosizecolumnsmode.allcells;	
			$global:lvresult.autoresizecolumns();            
			$global:lvresult.autosizecolumnsmode = datagridviewautosizecolumnsmode.allcells;		
			aspjobs $txtcomputer.text -fill; 
			set-unselectedall			
			if($global:clientlistcount -eq 0){ $lblaction.text = " no computer(s) found... ";initialize-listviewresult;}
			#for($i=0; $i -lt $global:clientlistcount; $i++){
				#$item = new-object system.windows.forms.listviewitem($global:clientlist[$global:resultset[$i]][0])
				#$global:lvmain.items.add($item)
			#}		
			$lblstatus.text = $global:status
			#get-selected
			}elseif($txtcomputer.text -eq ""){ $lblstatus.text = "computername is empty..."; $lblaction.text = "computername is empty..."; $lblselected.text = " ";$cmsselect.visible = $false; return }
			$lblaction.text = " "
			$picturebox.visible = $false 
			checkquit; if($global:gui){refreshgui;}
		}		
		#$global:lvmain.add_Click( {$cmsunselect.visible = $true})			
		
		function get-openfile{ 
			[system.reflection.assembly]::loadwithpartialname("system.windows.forms") | out-null
			$openfiledialog = new-object system.windows.forms.openfiledialog
			$openfiledialog.initialdirectory = "%systemdrive%"
			$openfiledialog.filter = "text files (*.csv)|*.csv"
			$openfiledialog.showdialog() | out-null
			$openfiledialog.filename
			$openfiledialog.showhelp = $true
		}
		$btninstsw_click={	
			if(checkselectedcomputer){		
				set-sbpstate("show installed software")
				fill-listviewresult("instsw")	
			}
		}
		
		$btnimportcsv_click={		
			set-sbpstate("import csv")
			$global:inputfile = get-openfile
			fill-listviewresult("importcsv")		
		}
		$btnupdatelogs_click={
			if(checkselectedcomputer){
				set-sbpstate("show update logs")			
				fill-listviewresult("wsusupdates")
			}
		}		
		$btnupdlist_click={
			if(checkselectedcomputer){
				set-sbpstate("show update list")			
				fill-listviewresult("updlist")
			}
		}		
		$btnsupdate_click={
			if(checkselectedcomputer){
				set-sbpstate("search update")			
				fill-listviewresult("supdate")
			}
		}		
		$btndelupdate_click={
			if(checkselectedcomputer){
				set-sbpstate("uninstall update")			
				fill-listviewresult("delupdate")
			}
		}		
		$btnupdcount_click={
			if(checkselectedcomputer){
				set-sbpstate("show update count")			
				fill-listviewresult("updcount")
			}
		}				
		$btnuserlang_click={
			if(checkselectedcomputer){
				set-sbpstate("existing user language")			
				fill-listviewresult("userlang")
			}
		}		
		$btntasks_click={
			if(checkselectedcomputer){
				set-sbpstate("existing scheduled tasks")			
				fill-listviewresult("tasks")
			}
		}		
		$btnshares_click={
			if(checkselectedcomputer){
				set-sbpstate("existing shares")			
				fill-listviewresult("shares")
			}
		}		
		$btnspace_click={
			if(checkselectedcomputer){
				set-sbpstate("partition space")			
				fill-listviewresult("space")
			}
		}		
		$btnusers_click={
			if(checkselectedcomputer){
				set-sbpstate("open user management")			
				fill-listviewresult("users")
			}
		}
		$btnevents_click={
			if(checkselectedcomputer){
				set-sbpstate("open event management")			
				fill-listviewresult("events")
			}
		}		
		$btnlogon_click={
			if(checkselectedcomputer){
				set-sbpstate("current logon")			
				fill-listviewresult("getlogon")
			}
		}
		$btnnetstat_click={
			if(checkselectedcomputer){
				set-sbpstate("show network activity")			
				fill-listviewresult("netstat")
			}
		}	
		$btnrebootpending_click={
			if(checkselectedcomputer){
				set-sbpstate("show reboot pending files")			
				fill-listviewresult("rp")
			}
		}	
		$btnprinter_click={
			if(checkselectedcomputer){
				set-sbpstate("show printer and printer status")			
				fill-listviewresult("printer")
			}
		}	
		$btnnetversion_click={
			if(checkselectedcomputer){
				set-sbpstate("show .net version")			
				fill-listviewresult("netversion")
			}
		}		
		$btnpsversion_click={
			if(checkselectedcomputer){
				set-sbpstate("show powershell version")			
				fill-listviewresult("psversion")
			}
		}		
		$btnieversion_click={
			if(checkselectedcomputer){
				set-sbpstate("show internet explorer version")			
				fill-listviewresult("ieversion")
			}
		}		
		$btnmonitor_click={
			if(checkselectedcomputer){
				set-sbpstate("show monitor informations")			
				fill-listviewresult("monitor")
			}
		}		
		$btnfw_click={
			if(checkselectedcomputer){
				set-sbpstate("show firewall status")			
				fill-listviewresult("fwstatus")
			}
		}				
		$btncpu_click={
			if(checkselectedcomputer){
				set-sbpstate("show cpu load")			
				fill-listviewresult("cpuload")
			}
		}			
		$btnlastlogon_click={
			if(checkselectedcomputer){
				set-sbpstate("last logon")			
				fill-listviewresult("getlastlogon")
			}
		}
		$btnlocaladmins_click={
			if(checkselectedcomputer){
				set-sbpstate("local admins")			
				fill-listviewresult("viewlocaladmins")
			}
		}		
		$btnprofiles_click={
			if(checkselectedcomputer){
				set-sbpstate("profiles")			
				fill-listviewresult("showprofiles")
			}
		}		
		$btnprocesses_click={
			if(checkselectedcomputer){		
				set-sbpstate("processes")			
				fill-listviewresult("getprocess")
			}			
		}		
		$btnservices_click={
			if(checkselectedcomputer){
				set-sbpstate("services")			
				fill-listviewresult("viewservices")
			}				
		}		
		$btntestcon_click={
			if(checkselectedcomputer){
				set-sbpstate("ping request")			
				fill-listviewresult("test")
			}
		}		
		$btnip_click={
			if(checkselectedcomputer){
				set-sbpstate("ip address")			
				fill-listviewresult("ip")
			}			
		}		
		$btnmac_click={
			if(checkselectedcomputer){
				set-sbpstate("mac address")			
				fill-listviewresult("mac")
			}			
		}		
		$btnlastboot_click={
			if(checkselectedcomputer){
				set-sbpstate("lastboot")			
				fill-listviewresult("lastboot")
			}			
		}		
		$btninstalldate_click={
			if(checkselectedcomputer){
				set-sbpstate("installdate")			
				fill-listviewresult("installdate")
			}			
		}		
		$btnrestart_click={
			if(checkselectedcomputer){
				set-sbpstate("restart")			
				fill-listviewresult("restart")
			}
		}		
		$btnstop_click={
			if(checkselectedcomputer){
				set-sbpstate("stop")			
				fill-listviewresult("stop")
			}
		}
		$btnrestartloggedoff_click={
			if(checkselectedcomputer){
				set-sbpstate("restart")			
				fill-listviewresult("restartlo")
			}
		}		
		$btnstoploggedoff_click={
			if(checkselectedcomputer){
				set-sbpstate("stop")			
				fill-listviewresult("stoplo")
			}
		}				
		$btnrdp_click={
			if(checkselectedcomputer){
				set-sbpstate("rdp")			
				fill-listviewresult("rdp")
			}
		}		
		$btnmsg_click={
			if(checkselectedcomputer){
				set-sbpstate("send message")			
				fill-listviewresult("sendmsg")
			}			
		}		
		$btncdrive_click={
			if(checkselectedcomputer){
				set-sbpstate("c drive")			
				fill-listviewresult("cdrive")
			}			
		}	
		$btnvsedrive_click={
			if(checkselectedcomputer){
				set-sbpstate("vse drive")			
				fill-listviewresult("vsedrive")
			}			
		}				
		$btneris_click={
			if(checkselectedcomputer){
				set-sbpstate("status restart eris")			
				fill-listviewresult("eris")
			}
		}		
		$btnwakeup_click={
			if(checkselectedcomputer){
				set-sbpstate("wake up client")			
				fill-listviewresult("wake")
			}
		}
		$btnremreg_click={
			if(checkselectedcomputer){
				set-sbpstate("remove registry asp team")			
				fill-listviewresult("remreg")
			}
		}		
		$btnadinfos_click={
			if(checkselectedcomputer){
				set-sbpstate("ad computer infos")			
				fill-listviewresult("adinfos")
			}			
		}		
		$btnenv_click={
			if(checkselectedcomputer){
				set-sbpstate("environments")			
				fill-listviewresult("env")
			}		
		}
		$btnchlang_click={
			if($global:dlanguage -eq "en"){			
				$global:dlanguage = "de"			
			}else{
				$global:dlanguage = "en"
			}			
				get-translation
		}					
		$btnmdrive_click={
			if(checkselectedcomputer){
				set-sbpstate("mapped drive")			
				fill-listviewresult("mdrive")
			}			
		}	
		$btndelcred_click={
			set-sbpstate("delete credentials")
			$params = @('-noprofile', '-noexit', 'c:\windows\system32\windowspowershell\v1.0\modules\aspcontrolpc\gui.ps1','-check')
			$file="c:\windows\system32\windowspowershell\v1.0\powershell.exe"
			try{start-process -filepath $file -workingdirectory "c:\windows\system32\windowspowershell\v1.0\modules\aspcontrolpc" -argumentlist $params -passthru -credential $cred;handleexit}
			catch{
				$errormessage = $_.exception.message	
				write-host "an error has occurred : $errormessage" -f red
				start-sleep -s 5
				handleexit
			}
		}
		$btnmemory_click={
			if(checkselectedcomputer){
				set-sbpstate("memory usage")			
				fill-listviewresult("memory")
			}
		}		
			
		function set-sbpstate($state="data"){
			$picturebox.visible = $true
			$lblstatus.text = "retrieving $state ..."
		}				
		function initialize-listviewmain{
			#for($i=0; $i -lt $global:lvmain.items.count; $i++){
			#	$global:lvmain.items[$i].remove()	
			#}
			#$global:lvmain.items.clear()
			$global:lvmain.rows.clear();			
			$global:status = ""	
			refreshgui;
		}		
		function initialize-listviewresult{
			#$global:lvresult.items.clear()
			#get-selected
			$global:lvresult.rows.clear();
			refreshgui;
		}		
		function fill-listviewresult{			
			param($aspfunction)
			#write-progress -id 9988 -activity "initialize and loading data" -status "please wait a few seconds";
			initialize-listviewresult
			if($txtcomputer.text -eq ""){ $lblstatus.text = "computername is empty..."; $lblaction.text = "computername is empty...";  }			
			#$global:lvresult.items.add
			$lblaction.text = "loading data - please wait, it may take 2 minutes..."
			#$formmain.refresh()		
			$global:selected = @()
			if($global:selected.count -lt $global:lvmain.rows.count){
				$starttime2 = get-date;
				aspwritetext "$nl-------------------------------------------------------------------------------------------------------------$nl"
				aspwritetext $nl	
			}
			$global:lvresult.columncount=0	
			$global:lvresult.Columns.clear();
			$global:lvresult.rows.clear();
			if($global:gui){refreshgui;}
			get-selected
			switch($aspfunction){
				"getlastlogon" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetlastlogon $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -lastlogon -dontshowlist} }
				"monitor" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetmonitor $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -moni -dontshowlist} }
				"printer" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetprinter $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -printer -dontshowlist} }
				"rp" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetrebootpending $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -rp -dontshowlist} }
				"netstat" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetnetstat $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -netstat -dontshowlist} }
				"ieversion" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetinternetexversion $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -ie -dontshowlist} }
				"psversion" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetpsversion $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -psv -dontshowlist} }
				"netversion" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetnetversion $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -nsv -dontshowlist} }
				"cpuload" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetcpuload $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -cpu -dontshowlist} }
				"fwstatus" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetfwstatus $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -fw -dontshowlist} }
				"instsw" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetinstalledsw $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -instsw -dontshowlist} }
				"importcsv" {  aspjobs $txtcomputer.text -csv -dontshowlist; $lblaction.text = "import csv-file successfully"; }
				"getlogon" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetlogon $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -logon -dontshowlist} }
				"viewlocaladmins" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspviewlocaladmins $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -ladmins -dontshowlist} }
				"showprofiles" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspshowprofiles $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -profiles -dontshowlist} }
				"getprocess" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetprocess $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -proc -dontshowlist} }
				"viewservices" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspviewservices $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -services -dontshowlist} }
				"restart" { if(check-answer($aspfunction)){$global:restart=1; if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start $aspfunction to computer: $_" -duration 5000; aspstartcommand -computername $_ -debugview $debugview -log $log}}elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -restart -force -dontshowlist}}} 					
				"stop" { if(check-answer($aspfunction)){$global:stop=1; if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start $aspfunction to computer: $_" -duration 5000;aspstartcommand -computername $_ -debugview $debugview -log $log}}elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -stop -force -dontshowlist}}} 					
				"restartlo" {if(check-answer("restart if user logged off")){$global:restartloggedoff=1;if($global:selected.count -lt $global:lvmain.rows.count){$global:selected | % { if($global:restartloggedoff){aspgetlogon -computername $_;if($global:userloggedin){aspwritetext "$nl user logged in - skip reboot to computer: $_ $nl$nl" red;$global:lvresult.rows.add("$_","user logged in - skip reboot");}elseif($global:userloggedin -eq 0){showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start command reboot to computer: $_" -duration 5000;if($sec -gt 0){ delay($sec) };aspstartcommand -computername $_ -debugview $debugview -log $log}}}}elseif($global:lvmain.rows.count -eq $global:selectedcount){aspjobs $txtcomputer.text -restartlo -force -dontshowlist}}}
				"stoplo" {if(check-answer("shutdown if user logged off")){$global:stoploggedoff=1; if($global:selected.count -lt $global:lvmain.rows.count){$global:selected | % {if($global:stoploggedoff){aspgetlogon -computername $_;if($global:userloggedin){aspwritetext "$nl user logged in - skip shutdown to computer: $_ $nl$nl" red;$global:lvresult.rows.add("$_","user logged in - skip shutdown");}elseif($global:userloggedin -eq 0){showtip -title “asp_control_pc” -messagetype warning -message "$env:username you have start command shutdown to computer: $_" -duration 5000;if($sec -gt 0){ delay($sec) }; aspstartcommand -computername $_ -debugview $debugview -log $log}}}}elseif($global:lvmain.rows.count -eq $global:selectedcount){aspjobs $txtcomputer.text -stoplo -force -dontshowlist}}}
				"rdp" { if($global:selected.count -lt $global:lvmain.rows.count){$global:selected | % { asprdp $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -rdp -dontshowlist} }
				"sendmsg" { $message = [microsoft.visualbasic.interaction]::inputbox("enter a message to send", "send message ", ""); if($message.length -gt 0){if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspsendmsg $_ -message $message} } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -send $message -dontshowlist}} }
				"cdrive" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspcdrive $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -cdrive -dontshowlist} }
				"vsedrive" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspvsedrive $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -vsedrive -dontshowlist} }
				"test" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { asptestconnectivity $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -test -dontshowlist -gui} }
				"eris" { if(check-answer("restart eris")){ if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { asprestarteris $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -eris -dontshowlist} } }					
				"wake" { if(check-answer("wake up")){ if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspwake $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -wake -force -dontshowlist} } }					
				"ip" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetipaddress $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount){ aspjobs $txtcomputer.text -ip -dontshowlist} }
				"env" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetenv $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount){ aspjobs $txtcomputer.text -env -dontshowlist} }
				"adinfos" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetadinfos $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount){ aspjobs $txtcomputer.text -adinfos -dontshowlist} }
				"mac" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetmacaddress $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -mac -dontshowlist} }
				"lastboot" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetlastboot $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -lastboot -dontshowlist} }
				"installdate" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetinstalldate $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -installdate -dontshowlist} }
				"users" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspviewuser $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -users -dontshowlist} }
				"events" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspvieweventvwr $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -events -dontshowlist} }
				"tasks" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgettasks $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -tasks -dontshowlist} }
				"shares" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetshares $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -shares -dontshowlist} }
				"space" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspviewfreespace $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -space -dontshowlist} }
				"userlang" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetuserlang $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -userlang -dontshowlist} }
				"mdrive" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetmappedDrive $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -mdrive -dontshowlist} }
				"memory" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetMemory $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -memo -dontshowlist} }
				"delcred" { aspdelCred; }
				"updcount" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetupdatecount $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -updcount -dontshowlist} }
				"delupdate" { $message = [microsoft.visualbasic.interaction]::inputbox("please enter a kb number: [only number]", "uninstall kb number ", ""); if($message.length -gt 0){if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspUninstall-Hotfix $_ -hotfixid $message} } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -uninstall $message -dontshowlist} }}
				"supdate" { $message = [microsoft.visualbasic.interaction]::inputbox("please enter a kb number: [only number]", "search kb number", ""); if($message.length -gt 0){if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspsearchhotfix $_ -hotfixid $message} } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -supdate $message -dontshowlist}} }
				"updlist" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspgetupdatelist $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -updlist -dontshowlist} }
				"wsusupdates" { $message = [microsoft.visualbasic.interaction]::inputbox("[1] asp wsus [2] asp wsus error [3] asp wsus selfupdate [4] ms wsus [5] wsus report :", "which wsus log?", ""); if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspviewwsusupdates $_ -logselection $message} } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -wsusupdates $message -dontshowlist} }
				"remreg" { if($global:selected.count -lt $global:lvmain.rows.count){ $global:selected | % { aspremoveRegistry $_ } } elseif($global:lvmain.rows.count -eq $global:selectedcount) { aspjobs $txtcomputer.text -remreg -dontshowlist} }
			
			}			
			$lblstatus.text = "$global:status selected client(s):  $($global:selected.count)  ";
			$lblaction.text = "  "
			$picturebox.visible = $false
			if($global:gui){refreshgui;}
			<#for($i=0; $i -lt $global:resultlist.count; $i++){
				$item = new-object system.windows.forms.listviewitem("item")
				$item.text = $global:resultlist[$i]
				#$global:lvresult.items.add($item)					
			}#>
			$global:restart=0;
			$global:restartloggedoff=0;
			$global:stop=0;
			$global:stoploggedoff=0;
			#$global:lvresult.rows.removeat($global:lvresult.rows.count - 1);
			#$global:lvresult.rows.remove($global:lvresult.rows[$global:lvresult.rows.count - 1]);
			#$global:resultlist = @()
			#write-progress -id 9988 -activity "completed" -status "completed" -completed;
			if($global:selected.count -lt $global:lvmain.rows.count){
				aspwritetext "$nl-------------------------------------------------------------------------------------------------------------$nl"
				aspwritetext $nl		
				$stoptime2 = get-date 
				$timerunning2 = ($stoptime2 - $starttime2).totalseconds
				if($timerunning2 -gt 60){ $timerunning2 = ($timerunning2 / 60); $minsec2 = "min." }
				else { $minsec2 = "sec." }
				$run2 = "{0:n2}" -f ($timerunning2)
				aspwritetext "$nl script after $run2 $minsec2 done $nl $nl" cyan
				aspwritetext $nl
			}
		}		
		function set-selectedall{
			$global:selectedcount = $global:lvmain.rows.count
			#if($global:lvmain.items.count -gt 0){ $global:lvmain.items | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::yellowgreen; $global:selected += $global:lvmain.items[$_.index].text; $global:selectedcount++ } } 
			#if($count -eq 0){ $count = $global:clientlist.count}
			$global:lvmain.SelectAll();
			$lblselected.text = "current selected client(s) $$global:lvmain.rows.count" 
		}
		function $global:lvmain_CellContentClick={
		#event argument: $_ = [system.windows.forms.datagridviewcelleventargs]
		write-host $_.rowindex 
		write-host $_.columnindex 
		write-host $global:lvmain.rows[$_.rowindex].cells[0].value 
		write-host $global:lvmain.rows[$_.rowindex].cells[$_.columnindex].value 
		}		
		function set-selected{
			$global:selectedcount = 0
			if($global:lvmain.selecteditems.count -gt 0){ $global:lvmain.selecteditems | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::yellowgreen }}
			if($global:lvmain.items.count -gt 0){ $global:lvmain.items | % { if($global:lvmain.items[$_.index].backcolor.name -eq "yellowgreen"){ $global:selectedcount++ } } }
			#if($count -eq 0){ $count = $global:clientlist.count}
			$lblselected.text = "current selected client(s) $global:selectedcount" 
		}		
		function get-selected{
			$global:selectedcount = 0
			$global:mac = ""
			$global:ipaddress = ""		
			#$global:lvmain.items | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::white }
			#if($global:lvmain.items.count -gt 0){ $global:lvmain.items | % { if($global:lvmain.items[$_.index].backcolor.name -eq "yellowgreen"){ $global:selected += $global:lvmain.items[$_.index].text; $global:selectedcount++ } } }
			#if($count -eq 0){ $count = $global:clientlist.count}
			 for($i=0;$i -lt $global:lvmain.rows.count;$i++){
				if($global:lvmain.rows[$i].cells[0].selected){
					$global:selected += $global:lvmain.rows[$i].cells[0].value
					$global:selectedcount++	
				}
			}
			$lblselected.text = "current selected client(s) $global:selectedcount" 
		}		
		function set-unselectedall{
			$global:selectedcount = 0
			#if($global:lvmain.items.count -gt 0){ $global:lvmain.items | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::white } }
			$global:lvmain.clearselection();
			$lblselected.text = "current selected client(s) 0" 
		}		
		function set-unselected{			
			if($global:lvmain.selecteditems.count -gt 0){$global:lvmain.selecteditems | % { $global:lvmain.items[$_.index].backcolor = [system.drawing.color]::white} }
			get-selected
		}		
		function check-answer($command=""){
			#if($global:lvmain.selecteditems.count -gt 0){ $global:lvmain.selecteditems | % { $global:selected += $global:lvmain.items[$_.index].text } }
			$vbmsg = new-object -comobject wscript.shell
			if($global:selected.count -eq 1){ $answer = $vbmsg.popup("try to $command " + $global:selected + " ?",0,"$command " + $global:selected + " ?",4) }
			else { $answer = $vbmsg.popup("try to $command for all " + $global:selectedcount + " computer(s) ?",0,"$command for all " + $global:selectedcount + " computer(s) ?",4) }
			switch ($answer){ 6 { return $true }; 7 { return $false } }		
		}		
		function add-column{
			param([string]$column)
			$global:lvmain.columns.add($column)
		}				
		function set-formtitle{
			$formmain.text = $env:username + '@' + [system.environment]::machinename + ' - ' + "ASP PC Control"
		}		
		<#$path = ${env:psmodulepath}.split(';') | where-object { $_ -match ${env:systemroot}.replace('\','\\') } | select-object -first 1 | foreach-object { return $_ }
		$file = $path+"asp_pc_control\gui.ico"	
		if(!(test-path $path+"asp_pc_control\gui.ico")){$path+="\";$file = $path+"asp_pc_control\gui.ico"}				
		#$path = ${env:psmodulepath}.split(';') | where-object { $_ -match ${env:systemroot}.replace('\','\\') } | select-object -first 1 | foreach-object { return $_ }#>
		$file = "\\$env:server\utils$\PS-Modules\asp_PC_Control\loader.gif"	
		#write-host $file;	
		if(!(test-path $file)){$file = (get-item $file) }
		$img = [system.drawing.image]::fromfile($file)
		$picturebox.location = '720, 10'
		$picturebox.width = $img.size.width
		$picturebox.height = $img.size.height		 
		$picturebox.image = $img
		$picturebox.visible = $false
		$formmain.controls.add($global:lvmain)
		$formmain.controls.add($global:lvresult)
		$formmain.controls.add($btnsearch)
		$formmain.controls.add($btnexit)
		$formmain.controls.add($txtcomputer)
		$formmain.controls.add($sb)
		$formmain.controls.add($lblaction)
		$formmain.controls.add($lblselected)
		$formmain.controls.add($lblclient)
		$formmain.controls.add($picturebox)
		$formmain.controls.add($btnselectall)
		$formmain.controls.add($btnunselectall)
		$formmain.controls.add($tabctrl)
		$formmain.clientsize = '1024, 950'
		$formmain.name = "formmain"
		$dim = ( [system.windows.forms.screen]::allscreens | ? { $_.primary} ).workingarea   
		$formmain.startposition = "manual"
		$system_drawing_point = new-object system.drawing.point
		$system_drawing_point.x = $dim.width*0.32
		$system_drawing_point.y = $dim.height*0.04
		$formmain.location = $system_drawing_point
		$formmain.text = "asp_control_pc"
		$formmain.sizegripstyle = "hide"
		$formmain.topmost = true;
		$formmain.add_load($formmain_load)
		$formmain.add_Closing($onapplicationexit)
		$formmain.icon = [system.drawing.icon]::extractassociatedicon($file)
		$formmain.formborderstyle = [system.windows.forms.formborderstyle]::fixedsingle
		$formmain.keypreview = $true
		$formmain.maximizebox = $false
		$formmain.icon = $icon		
		$fontbold = new-object system.drawing.font("arial",12,[drawing.fontstyle]'bold' )
		$tabctrl.controls.add($tabpageinfo)
		$tabctrl.controls.add($tabpageinfo2)
		$tabctrl.controls.add($tabpageupdates)
		$tabctrl.controls.add($tabpageadmin)
		$tabctrl.controls.add($tabpageadds)
		$tabctrl.location = '820, 155'
		$tabctrl.name = "tabctrl"
		$tabctrl.size = '200, 710'
		$tabctrl.selectedindex = 0
		#$tabctrl.font = $fontbold		
		$tabpageadds.controls.add($groupadds)
		$tabpageadds.name = "tabpageadds"
		$tabpageadds.tabindex = 5
		$tabpageadds.text = "active directory"
		$tabpageadds.usevisualstylebackcolor = $true		
		$tabpageupdates.controls.add($groupupdates)
		#$tabpageupdates.location = '13, 155'
		$tabpageupdates.name = "tabpageupdates"
		#$tabpageupdates.size = '195, 860'
		$tabpageupdates.tabindex = 3
		$tabpageupdates.text = "updates"
		$tabpageupdates.usevisualstylebackcolor = $true		
		$tabpageinfo.controls.add($groupinfo)
		$btnchlang.location = '840, 890'
		$btnchlang.name = "btnchlang"
		$btnchlang.size = '166, 30'
		$btnchlang.tabindex = 1	
		$btnchlang.backcolor=[System.Drawing.Color]::LightGreen		
		$btnchlang.usevisualstylebackcolor = $true
		$btnchlang.add_click($btnchlang_click)
		$formmain.controls.add($btnchlang)		
		#$tabpageinfo.location = '13, 155'
		$tabpageinfo.name = "tabpageinfo"
		#$tabpageinfo.size = '195, 860'
		$tabpageinfo.tabindex = 0
		$tabpageinfo.text = "information"
		$tabpageinfo.usevisualstylebackcolor = $true
		$tabpageinfo2.name = "tabpageinfo2"
		#$tabpageinfo2.size = '195, 860'
		$tabpageinfo2.tabindex = 1
		$tabpageinfo2.text = "information2"
		$tabpageinfo2.usevisualstylebackcolor = $true
		$tabpageinfo2.controls.add($groupinfo2)		
		$tabpageadmin.controls.add($groupadministration)
		#$tabpageadmin.location = '13, 155'
		$tabpageadmin.name = "tabpageadmin"
		#$tabpageadmin.size = '195, 860'
		$tabpageadmin.tabindex = 4
		$tabpageadmin.text = "administration"
		$tabpageadmin.usevisualstylebackcolor = $true
		$fontbold = new-object system.drawing.font("arial",10,[drawing.fontstyle]'bold' )		
		$lblclient.location = '820, 10'
		$lblclient.name = "lblclient"
		$lblclient.size = '170, 18'
		$lblclient.text = "computername:"
		$lblclient.forecolor = [system.drawing.color]::black
		$lblclient.font = $fontbold
		$fontbold = new-object system.drawing.font("arial",12,[drawing.fontstyle]'bold' )
		$lblaction.location = '250, 16'
		$lblaction.name = "lblaction"
		$lblaction.size = '400, 22'
		$lblaction.text = ""
		$lblaction.flatstyle = [system.windows.forms.flatstyle]::system
		$lblaction.forecolor = [system.drawing.systemcolors]::desktop
		$lblaction.font = $fontbold
		$lblaction.textalign = [system.drawing.contentalignment]::middleleft

		$fontregular = new-object system.drawing.font("arial",11,[drawing.fontstyle]'regular' )	
		$lblselected.location = '5, 16'
		$lblselected.name = "lblselected"
		$lblselected.size = '350, 22'
		$lblselected.text = ""
		$lblselected.flatstyle = [system.windows.forms.flatstyle]::system
		$lblselected.forecolor = [system.drawing.systemcolors]::highlight
		$lblselected.font = $fontregular
		$lblselected.textalign = [system.drawing.contentalignment]::middleleft	
		$lblselected.visible=$false
		
		if($global:xml.options.functions.restartcomputer.enabled -eq $true){$groupadministration.controls.add($btnrestart)}
		if($global:xml.options.functions.restartcomputer.enabled -eq $true){$groupadministration.controls.add($btnrestartloggedoff)}
		if($global:xml.options.functions.sendmessage.enabled -eq $true){$groupadministration.controls.add($btnmsg)}
		if($global:xml.options.functions.opencdrive.enabled -eq $true){$groupadministration.controls.add($btncdrive)}
		if($global:xml.options.functions.openvsedrive.enabled -eq $true){$groupadministration.controls.add($btnvsedrive)}
		if($global:xml.options.functions.importcsv.enabled -eq $true){$groupadministration.controls.add($btnimportcsv)}
		if($global:xml.options.functions.stopcomputer.enabled -eq $true){$groupadministration.controls.add($btnstop)}
		if($global:xml.options.functions.stopcomputer.enabled -eq $true){$groupadministration.controls.add($btnstoploggedoff)}
		if($global:xml.options.functions.remotedesktop.enabled -eq $true){$groupadministration.controls.add($btnrdp)}
		if($global:xml.options.functions.restarteris.enabled -eq $true){$groupadministration.controls.add($btneris)}
		if($global:xml.options.functions.wakeup.enabled -eq $true){$groupadministration.controls.add($btnwakeup)}
		if($global:xml.options.functions.remreg.enabled -eq $true){$groupadministration.controls.add($btnremreg)}
		
		
		#$groupadministration.location = '820, 710'
		$groupadministration.name = "grouptools"
		$groupadministration.size = '190, 680'
		$groupadministration.tabindex = 8
		$groupadministration.tabstop = $false
		#$groupadministration.text = "administration"
		
		$btnrestart.location = '13, 19'
		$btnrestart.name = "btnrestart"
		$btnrestart.size = '166, 30'
		$btnrestart.tabindex = 15
		$btnrestart.usevisualstylebackcolor = $true
		$btnrestart.add_click($btnrestart_click)
		
		$btnrestartloggedoff.location = '13, 50'
		$btnrestartloggedoff.name = "btnrestartloggedoff"
		$btnrestartloggedoff.size = '166, 30'
		$btnrestartloggedoff.tabindex = 15
		$btnrestartloggedoff.usevisualstylebackcolor = $true
		$btnrestartloggedoff.add_click($btnrestartloggedoff_click)
		
		$btnmsg.location = '13, 359'
		$btnmsg.name = "btnmsg"
		$btnmsg.size = '166, 30'
		$btnmsg.tabindex = 18
		$btnmsg.usevisualstylebackcolor = $true
		$btnmsg.add_click($btnmsg_click)
		
		$btncdrive.location = '13, 143'
		$btncdrive.name = "btncdrive"
		$btncdrive.size = '166, 30'
		$btncdrive.tabindex = 19
		$btncdrive.usevisualstylebackcolor = $true
		$btncdrive.add_click($btncdrive_click)
		
		$btnvsedrive.location = '13, 266'
		$btnvsedrive.name = "btnvsedrive"
		$btnvsedrive.size = '166, 30'
		$btnvsedrive.tabindex = 21
		$btnvsedrive.usevisualstylebackcolor = $true
		$btnvsedrive.add_click($btnvsedrive_click)
		
		$btnimportcsv.location = '13, 297'
		$btnimportcsv.name = "btnimportcsv"
		$btnimportcsv.size = '166, 30'
		$btnimportcsv.tabindex = 22
		$btnimportcsv.usevisualstylebackcolor = $true
		$btnimportcsv.add_click($btnimportcsv_click)
				
		$btneris.location = '13, 174'
		$btneris.name = "btneris"
		$btneris.size = '166, 30'
		$btneris.tabindex = 20
		$btneris.usevisualstylebackcolor = $true
		$btneris.add_click($btneris_click)
		
		$btnstop.location = '13, 81'
		$btnstop.name = "btnstop"
		$btnstop.size = '166, 30'
		$btnstop.tabindex = 16
		$btnstop.usevisualstylebackcolor = $true
		$btnstop.add_click($btnstop_click)
		
		$btnstoploggedoff.location = '13, 112'
		$btnstoploggedoff.name = "btnstoploggedoff"
		$btnstoploggedoff.size = '166, 30'
		$btnstoploggedoff.tabindex = 16
		$btnstoploggedoff.usevisualstylebackcolor = $true
		$btnstoploggedoff.add_click($btnstoploggedoff_click)
		
		$btnrdp.location = '13, 328'
		$btnrdp.name = "btnrdp"
		$btnrdp.size = '166, 30'
		$btnrdp.tabindex = 17
		$btnrdp.usevisualstylebackcolor = $true
		$btnrdp.add_click($btnrdp_click)
		
		$btnwakeup.location = '13, 205'
		$btnwakeup.name = "btnwakeup"
		$btnwakeup.size = '166, 30'
		$btnwakeup.tabindex = 18
		$btnwakeup.usevisualstylebackcolor = $true
		$btnwakeup.add_click($btnwakeup_click)
		
		$btnremreg.location = '13, 236'
		$btnremreg.name = "btnremreg"
		$btnremreg.size = '166, 30'
		$btnremreg.tabindex = 19
		$btnremreg.usevisualstylebackcolor = $true
		$btnremreg.add_click($btnremreg_click)
		
		if($global:xml.options.functions.adcomputerinfos.enabled -eq $true){$groupadds.controls.add($btnadinfos)}
		
		$groupadds.name = "groupadds"
		$groupadds.size = '190, 680'
		$groupadds.tabindex = 7
		$groupadds.tabstop = $false	

		$groupupdates.name = "groupupdates"
		$groupupdates.size = '190, 680'
		$groupupdates.tabindex = 7
		$groupupdates.tabstop = $false	
		
		$btnupdatelogs.location = '13, 143'
		$btnupdatelogs.name = "btnupdatelogs"
		$btnupdatelogs.size = '166, 30'
		$btnupdatelogs.tabindex = 6
		$btnupdatelogs.usevisualstylebackcolor = $true
		$btnupdatelogs.add_click($btnupdatelogs_click)

		$btnupdcount.location = '13, 112'
		$btnupdcount.name = "btnupdcount"
		$btnupdcount.size = '166, 30'
		$btnupdcount.tabindex = 5
		$btnupdcount.usevisualstylebackcolor = $true
		$btnupdcount.add_click($btnupdcount_click)
		
		$btndelupdate.location = '13, 50'
		$btndelupdate.name = "btndelupdate"
		$btndelupdate.size = '166, 30'
		$btndelupdate.tabindex = 3
		$btndelupdate.usevisualstylebackcolor = $true
		$btndelupdate.add_click($btndelupdate_click)
		
		$btnsupdate.location = '13, 81'
		$btnsupdate.name = "btnsupdate"
		$btnsupdate.size = '166, 30'
		$btnsupdate.tabindex = 4
		$btnsupdate.usevisualstylebackcolor = $true
		$btnsupdate.add_click($btnsupdate_click)
		
		$btnupdlist.location = '13, 19'
		$btnupdlist.name = "btnupdlist"
		$btnupdlist.size = '166, 30'
		$btnupdlist.tabindex = 2
		$btnupdlist.usevisualstylebackcolor = $true
		$btnupdlist.add_click($btnupdlist_click)		
		
		$btnadinfos.location = '13, 19'
		$btnadinfos.name = "btnadinfos"
		$btnadinfos.size = '166, 30'
		$btnadinfos.tabindex = 1
		$btnadinfos.usevisualstylebackcolor = $true
		$btnadinfos.add_click($btnadinfos_click)
		
		if($global:xml.options.functions.updatelist.enabled -eq $true){$groupupdates.controls.add($btnupdlist)}
		if($global:xml.options.functions.searchupdate.enabled -eq $true){$groupupdates.controls.add($btnsupdate)}
		if($global:xml.options.functions.deleteupdate.enabled -eq $true){$groupupdates.controls.add($btndelupdate)}
		if($global:xml.options.functions.updatecount.enabled -eq $true){$groupupdates.controls.add($btnupdcount)}
		if($global:xml.options.functions.updatelogs.enabled -eq $true){$groupupdates.controls.add($btnupdatelogs)}
		if($global:xml.options.functions.showservices.enabled -eq $true){$groupinfo.controls.add($btnservices)}
		if($global:xml.options.functions.showprocesses.enabled -eq $true){$groupinfo.controls.add($btnprocesses)}
		if($global:xml.options.functions.showlocalprofiles.enabled -eq $true){$groupinfo.controls.add($btnprofiles)}
		if($global:xml.options.functions.showlastlogon.enabled -eq $true){$groupinfo.controls.add($btnlastlogon)}
		if($global:xml.options.functions.showlocaladmins.enabled -eq $true){$groupinfo.controls.add($btnlocaladmins)}
		if($global:xml.options.functions.showlogon.enabled -eq $true){$groupinfo.controls.add($btnlogon)}
		if($global:xml.options.functions.testconnection.enabled -eq $true){$groupinfo.controls.add($btntestcon)}
		if($global:xml.options.functions.showip.enabled -eq $true){$groupinfo.controls.add($btnip)}
		if($global:xml.options.functions.showmac.enabled -eq $true){$groupinfo.controls.add($btnmac)}
		if($global:xml.options.functions.showlastboot.enabled -eq $true){$groupinfo.controls.add($btnlastboot)}
		if($global:xml.options.functions.showinstalldate.enabled -eq $true){$groupinfo.controls.add($btninstalldate)}
		if($global:xml.options.functions.openusermanagement.enabled -eq $true){$groupinfo.controls.add($btnusers)}
		if($global:xml.options.functions.openeventmanagement.enabled -eq $true){$groupinfo.controls.add($btnevents)}
		if($global:xml.options.functions.showpartitionspace.enabled -eq $true){$groupinfo.controls.add($btnspace)}
		if($global:xml.options.functions.showscheduledtasks.enabled -eq $true){$groupinfo.controls.add($btntasks)}
		if($global:xml.options.functions.showshares.enabled -eq $true){$groupinfo.controls.add($btnshares)}
		if($global:xml.options.functions.showuserlanguage.enabled -eq $true){$groupinfo.controls.add($btnuserlang)}
		if($global:xml.options.functions.showenvironments.enabled -eq $true){$groupinfo.controls.add($btnenv)}
		if($global:xml.options.functions.showmappeddrive.enabled -eq $true){$groupinfo.controls.add($btnmdrive)}
		if($global:xml.options.functions.checkmemoryusage.enabled -eq $true){$groupinfo.controls.add($btnmemory)}		
		if($global:xml.options.functions.instsw.enabled -eq $true){$groupinfo.controls.add($btninstsw)}
		
		if($global:xml.options.functions.cpuload.enabled -eq $true){$groupinfo2.controls.add($btncpu)}
		if($global:xml.options.functions.fwstatus.enabled -eq $true){$groupinfo2.controls.add($btnfw)}
		if($global:xml.options.functions.monitor.enabled -eq $true){$groupinfo2.controls.add($btnmonitor)}
		if($global:xml.options.functions.ieversion.enabled -eq $true){$groupinfo2.controls.add($btnieversion)}
		if($global:xml.options.functions.psversion.enabled -eq $true){$groupinfo2.controls.add($btnpsversion)}
		if($global:xml.options.functions.netversion.enabled -eq $true){$groupinfo2.controls.add($btnnetversion)}
		if($global:xml.options.functions.printer.enabled -eq $true){$groupinfo2.controls.add($btnprinter)}
		if($global:xml.options.functions.rp.enabled -eq $true){$groupinfo2.controls.add($btnrebootpending)}
		if($global:xml.options.functions.netstat.enabled -eq $true){$groupinfo2.controls.add($btnnetstat)}
		
		#$groupinfo.location = '818, 165'
		$groupinfo.name = "groupinfo"
		$groupinfo.size = '200, 680'
		$groupinfo.tabindex = 7
		$groupinfo.tabstop = $false
		#$groupinfo.text = "information"
		
		#$groupinfo.location = '818, 165'
		$groupinfo2.name = "groupinfo2"
		$groupinfo2.size = '200, 680'
		$groupinfo2.tabindex = 8
		$groupinfo2.tabstop = $false
		#$groupinfo2.text = "information"
		
		$btndelcred.location = '840, 859'
		$btndelcred.name = "btndelcred"
		$btndelcred.size = '166, 30'
		$btndelcred.tabindex = 20
		$btndelcred.usevisualstylebackcolor = $true
		$btndelcred.backcolor = [System.Drawing.Color]::hotpink
		$btndelcred.add_click($btndelcred_click)
		if($global:xml.options.functions.deletecredentials.enabled -eq $true){$formmain.controls.add($btndelcred)}
		
		$btninstsw.location = '13, 639'
		$btninstsw.name = "btninstsw"
		$btninstsw.size = '166, 30'
		$btninstsw.tabindex = 23
		$btninstsw.usevisualstylebackcolor = $true
		$btninstsw.add_click($btninstsw_click)
		
		$btnmemory.location = '13, 608'
		$btnmemory.name = "btnmemory"
		$btnmemory.size = '166, 30'
		$btnmemory.tabindex = 20
		$btnmemory.usevisualstylebackcolor = $true
		$btnmemory.add_click($btnmemory_click)
		
		$btnmdrive.location = '13, 577'
		$btnmdrive.name = "btnmdrive"
		$btnmdrive.size = '166, 30'
		$btnmdrive.tabindex = 20
		$btnmdrive.usevisualstylebackcolor = $true
		$btnmdrive.add_click($btnmdrive_click)
		
		$btnenv.location = '13, 546'
		$btnenv.name = "btnenv"
		$btnenv.size = '166, 30'
		$btnenv.tabindex = 19
		$btnenv.usevisualstylebackcolor = $true
		$btnenv.add_click($btnenv_click)
		
		$btnuserlang.location = '13, 515'
		$btnuserlang.name = "btnuserlang"
		$btnuserlang.size = '166, 30'
		$btnuserlang.tabindex = 18
		$btnuserlang.usevisualstylebackcolor = $true
		$btnuserlang.add_click($btnuserlang_click)
		
		$btnshares.location = '13, 484'
		$btnshares.name = "btnshares"
		$btnshares.size = '166, 30'
		$btnshares.tabindex = 17
		$btnshares.usevisualstylebackcolor = $true
		$btnshares.add_click($btnshares_click)
		
		$btntasks.location = '13, 453'
		$btntasks.name = "btntasks"
		$btntasks.size = '166, 30'
		$btntasks.tabindex = 16
		$btntasks.usevisualstylebackcolor = $true
		$btntasks.add_click($btntasks_click)
		
		$btnspace.location = '13, 422'
		$btnspace.name = "btnspace"
		$btnspace.size = '166, 30'
		$btnspace.tabindex = 15
		$btnspace.usevisualstylebackcolor = $true
		$btnspace.add_click($btnspace_click)
		
		$btnevents.location = '13, 391'
		$btnevents.name = "btnevents"
		$btnevents.size = '166, 30'
		$btnevents.tabindex = 14
		$btnevents.usevisualstylebackcolor = $true
		$btnevents.add_click($btnevents_click)
		
		$btnusers.location = '13, 360'
		$btnusers.name = "btnusers"
		$btnusers.size = '166, 30'
		$btnusers.tabindex = 13
		$btnusers.usevisualstylebackcolor = $true
		$btnusers.add_click($btnusers_click)
		
		$btnlastboot.location = '13, 329'
		$btnlastboot.name = "btnlastboot"
		$btnlastboot.size = '166, 30'
		$btnlastboot.tabindex = 12
		$btnlastboot.usevisualstylebackcolor = $true
		$btnlastboot.add_click($btnlastboot_click)
		
		$btninstalldate.location = '13, 298'
		$btninstalldate.name = "btninstalldate"
		$btninstalldate.size = '166, 30'
		$btninstalldate.tabindex = 11
		$btninstalldate.usevisualstylebackcolor = $true
		$btninstalldate.add_click($btninstalldate_click)
		
		$btnmac.location = '13, 267'
		$btnmac.name = "btnmac"
		$btnmac.size = '166, 30'
		$btnmac.tabindex = 10
		$btnmac.usevisualstylebackcolor = $true
		$btnmac.add_click($btnmac_click)
		
		$btnip.location = '13, 236'
		$btnip.name = "btnip"
		$btnip.size = '166, 30'
		$btnip.tabindex = 9
		$btnip.usevisualstylebackcolor = $true
		$btnip.add_click($btnip_click)
		
		$btntestcon.location = '13, 205'
		$btntestcon.name = "btntestcon"
		$btntestcon.size = '166, 30'
		$btntestcon.tabindex = 8
		$btntestcon.usevisualstylebackcolor = $true
		$btntestcon.add_click($btntestcon_click)
		
		$btnservices.location = '13, 174'
		$btnservices.name = "btnservices"
		$btnservices.size = '166, 30'
		$btnservices.tabindex = 7
		$btnservices.usevisualstylebackcolor = $true
		$btnservices.add_click($btnservices_click)
		
		$btnprocesses.location = '13, 143'
		$btnprocesses.name = "btnprocesses"
		$btnprocesses.size = '166, 30'
		$btnprocesses.tabindex = 6
		$btnprocesses.usevisualstylebackcolor = $true
		$btnprocesses.add_click($btnprocesses_click)
		
		$btnprofiles.location = '13, 112'
		$btnprofiles.name = "btnstartupitems"
		$btnprofiles.size = '166, 30'
		$btnprofiles.tabindex = 5
		$btnprofiles.usevisualstylebackcolor = $true
		$btnprofiles.add_click($btnprofiles_click)
		
		$btnlastlogon.location = '13, 50'
		$btnlastlogon.name = "btnlastlogon"
		$btnlastlogon.size = '166, 30'
		$btnlastlogon.tabindex = 3
		$btnlastlogon.usevisualstylebackcolor = $true
		$btnlastlogon.add_click($btnlastlogon_click)
		
		$btnlocaladmins.location = '13, 81'
		$btnlocaladmins.name = "btnlocaladmins"
		$btnlocaladmins.size = '166, 30'
		$btnlocaladmins.tabindex = 4
		$btnlocaladmins.usevisualstylebackcolor = $true
		$btnlocaladmins.add_click($btnlocaladmins_click)
		
		$btnlogon.location = '13, 19'
		$btnlogon.name = "btnlogon"
		$btnlogon.size = '166, 30'
		$btnlogon.tabindex = 2
		$btnlogon.usevisualstylebackcolor = $true
		$btnlogon.add_click($btnlogon_click)
		
		$btncpu.location = '13, 19'
		$btncpu.name = "btncpu"
		$btncpu.size = '166, 30'
		$btncpu.tabindex = 1
		$btncpu.usevisualstylebackcolor = $true
		$btncpu.add_click($btncpu_click)
		
		$btnfw.location = '13, 50'
		$btnfw.name = "btnfw"
		$btnfw.size = '166, 30'
		$btnfw.tabindex = 2
		$btnfw.usevisualstylebackcolor = $true
		$btnfw.add_click($btnfw_click)
		
		$btnmonitor.location = '13, 81'
		$btnmonitor.name = "btnmonitor"
		$btnmonitor.size = '166, 30'
		$btnmonitor.tabindex = 3
		$btnmonitor.usevisualstylebackcolor = $true
		$btnmonitor.add_click($btnmonitor_click)
		
		$btnieversion.location = '13, 112'
		$btnieversion.name = "btnieversion"
		$btnieversion.size = '166, 30'
		$btnieversion.tabindex = 4
		$btnieversion.usevisualstylebackcolor = $true
		$btnieversion.add_click($btnieversion_click)
		
		$btnpsversion.location = '13, 143'
		$btnpsversion.name = "btnpsversion"
		$btnpsversion.size = '166, 30'
		$btnpsversion.tabindex = 5
		$btnpsversion.usevisualstylebackcolor = $true
		$btnpsversion.add_click($btnpsversion_click)
		
		$btnnetversion.location = '13, 174'
		$btnnetversion.name = "btnnetversion"
		$btnnetversion.size = '166, 30'
		$btnnetversion.tabindex = 6
		$btnnetversion.usevisualstylebackcolor = $true
		$btnnetversion.add_click($btnnetversion_click)
		
		$btnprinter.location = '13, 205'
		$btnprinter.name = "btnprinter"
		$btnprinter.size = '166, 30'
		$btnprinter.tabindex = 7
		$btnprinter.usevisualstylebackcolor = $true
		$btnprinter.add_click($btnprinter_click)
		
		$btnrebootpending.location = '13, 236'
		$btnrebootpending.name = "btnrebootpending"
		$btnrebootpending.size = '166, 30'
		$btnrebootpending.tabindex = 8
		$btnrebootpending.usevisualstylebackcolor = $true
		$btnrebootpending.add_click($btnrebootpending_click)
		
		$btnnetstat.location = '13, 267'
		$btnnetstat.name = "btnrebootpending"
		$btnnetstat.size = '166, 30'
		$btnnetstat.tabindex = 9
		$btnnetstat.usevisualstylebackcolor = $true
		$btnnetstat.add_click($btnnetstat_click)
		
		$btnexit.location = '840, 920'
		$btnexit.name = "btnexit"
		$btnexit.size = '166, 30'
		$btnexit.tabindex = 21
		#$btnexit.text = "exit"
		$btnexit.usevisualstylebackcolor = $true
		$btnexit.backcolor = [System.Drawing.Color]::silver
		$btnexit.add_click($btnexit_click)
		
		$cmsselect.name = "cmsselect"
		$cmsselect.size = '187, 22'
		$cmsselect.text = "select computer"
		$cmsselect.visible = $false
		$cmsselect.add_click($cmsselect_click)
		
		$cmsunselect.name = "cmsunselect"
		$cmsunselect.size = '187, 22'
		$cmsunselect.text = "unselect computer"
		$cmsunselect.visible = $false
		$cmsunselect.add_click($cmsunselect_click)
		
		[void]$contextmenu.items.add($cmsselect)
		[void]$contextmenu.items.add($cmsunselect)
		$contextmenu.name = "contextmenu"
		$contextmenu.size = '188, 114'
			
		$fontregular = new-object system.drawing.font("arial",11,[drawing.fontstyle]'regular' )	
		$global:lvmain.font = $fontregular
		$global:lvmain.anchor = 'top, bottom, left, right'
		$global:lvmain.contextmenustrip = $contextmenu
		$global:lvmain.fullrowselect = $true
		$global:lvmain.gridlines = $true
		$global:lvmain.location = '660, 58'
		$global:lvmain.name = "$global:lvmain"
		$global:lvmain.size = '155, 860'
		$global:lvmain.tabindex = 22
		$global:lvmain.usecompatiblestateimagebehavior = $false
		$global:lvmain.view = 'details'
		
		$fontregular = new-object system.drawing.font("arial",11,[drawing.fontstyle]'regular' )	
		$global:lvresult.font = $fontregular
		$global:lvresult.anchor = 'top, bottom, left, right'
		$global:lvresult.fullrowselect = $false
		$global:lvresult.gridlines = $false
		$global:lvresult.location = '10, 58'
		$global:lvresult.name = "lvresult"
		$global:lvresult.size = '640, 860'
		$global:lvresult.tabindex = 23
		$global:lvresult.usecompatiblestateimagebehavior = $false
		$global:lvresult.view = 'details'
		
		$btnsearch.location = '822, 58'
		$btnsearch.name = "btnsearch"
		$btnsearch.size = '200, 25'
		$btnsearch.tabindex = 1
		$btnsearch.usevisualstylebackcolor = $true
		#$btnsearch.backcolor = [System.Drawing.Color]::deepskyblue
		$btnsearch.add_click($btnsearch_click)
		
		$btnselectall.location = '822, 89'
		$btnselectall.name = "btnselectall"
		$btnselectall.size = '200, 25'
		$btnselectall.tabindex = 1
		$btnselectall.usevisualstylebackcolor = $true
		$btnselectall.add_click($btnselectall_click)
		
		$btnunselectall.location = '822, 120'
		$btnunselectall.name = "btnunselectall"
		$btnunselectall.size = '200, 25'
		$btnunselectall.tabindex = 1
		$btnunselectall.usevisualstylebackcolor = $true
		$btnunselectall.add_click($btnunselectall_click)
		
		$fontbold = new-object system.drawing.font("arial",14,[drawing.fontstyle]'bold' )
		$txtcomputer.font = $fontbold
		$txtcomputer.location = '820, 28'
		$txtcomputer.name = "txtcomputer"
		$txtcomputer.size = '200, 20'
		$txtcomputer.tabindex = 0
		$txtcomputer.backcolor = [system.drawing.color]::gold
		
		$sb.anchor = 'bottom, left, right'
		#$sb.dock = 'none'
		$sb.location = '10, 920'
		$sb.name = "sb"
		[void]$sb.panels.add($lblzid)
		[void]$sb.panels.add($lblstatus)
		$sb.showpanels = $true
		$sb.size = '800, 22'
		$sb.text = "ready"
		#$sb.sizingrip = $false		
		
		$lblstatus.alignment = 'center'
		$lblstatus.name = "lblstatus"
		$lblstatus.text = "$($global:clientlist.count) client(s) ready"
		$lblstatus.width = 610		
		
		$lblzid.alignment = 'center'
		$lblzid.name = "lblzid"
		$lblzid.text = "eduard-albert.prinz"
		$lblzid.width = 180
		
		#$initialformwindowstate = $formmain.windowstate	
		#$formmain.add_load($form_statecorrection_load)
		$formmain.add_formclosed($form_cleanup_formclosed)
		$formmain.add_shown({$formmain.activate()})
		get-translation;
		return $formmain.showdialog()
	} 	
	generateform | out-null
	#onapplicationexit
}
export-modulemember -function aspgetclientlist, aspjobs, aspgetgui, aspgethelp, aspshowlog, aspshowevent, checkdeploymentmanager