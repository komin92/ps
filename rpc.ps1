#############################################################################
#																			#
# name: rpc.ps1																#
# script version 0.9														#
# 																			#	
# author: edi																#
#																			#
# comment: 	get infos about rpc											 	#
#																			#
# 20.02.2013 :: release														#
#																			#
#############################################################################

[cmdletbinding()]
param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = ""
    )

import-module activedirectory
$nl = [environment]::newline
function getrpcinfos()
{
if((. $env:emp_utils\testconn.ps1 -computername $computername -debugview $debugview -log 0 -show 0)){
	$computer = "$computername"

	write-host -fore cyan "$nl remotepc infos`t$computer :" $nl

	$sessionname=@()
	$username=@()
	$sessionid=@()
	$sessions = query session /server:$computer
	1..($sessions.count -1) | % {
		$temp = "" | select sessionname, username, id, state
		try{$temp.state = $sessions[$_].substring(48,8).trim()}
		catch{write-host -fore red "`t`tcan not access $computer"
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"}
	   if($temp.state -eq 'active'){
		$sessionname+=$temp.sessionname = $sessions[$_].substring(1,18).trim()
		$username+=$temp.username = $sessions[$_].substring(19,20).trim()
		$sessionid+=$temp.id = $sessions[$_].substring(39,9).trim()		
	   }   
	}        
	  
	$i=0
	if($username.count -gt 0){
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
		write-host -fore yellow "`tlogged on user`t`tsessionname`t`tsessionid"
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
		foreach($user in $username){
			write-host -fore green "`t"$user"`t`t"($sessionname[$i])"`t`t"($sessionid[$i])
			write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
			$i++
		}
	}

	

	$computer = $computername
	$adminpath = test-path \\$computer\admin$
	if ($adminpath -eq "true")
	{
		$key = "software\microsoft\windows\currentversion\authentication\logonui"
		$type = [microsoft.win32.registryhive]::localmachine
		try{$regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer,[microsoft.win32.registryview]::registry64)}
		catch{$regkey = [microsoft.win32.registrykey]::openremotebasekey($type, $computer)}
		$logon = $regkey.opensubkey($key)
		$username = $logon.getvalue("lastloggedonuser")
		if($username -ne ""){
			write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
			write-host -fore yellow "`tlast logged on user"
			write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
			write-host -fore green "`t"$username
			write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
			$regkey.close()
			$logon.close()
		}
	}
	else
	{
		write-host -fore red "`t`tcan not access $computer"
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
	}


	

	$path1 = "$compuer\c$\documents and settings"
	$path2 = "$computer\c$\users"
	$profile1 = test-path \\$path1
	$profile2 = test-path \\$path2

	if ($profile1 -eq "true"){$profiles = get-item "\\$path1\*" }
	elseif ($profile2 -eq "true"){$profiles = get-item "\\$path2\*" }
	#else {write-host -fore red "`t`tcan not access $computer"}
	if($profiles -ne $null){
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
		write-host -fore yellow "`tprofile`t`tlast access date`t"
		write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
		foreach ($profile in $profiles) 
		{	
			$accountname = (get-item $profile).pschildname 
			$lastaccesstime = (get-item $profile).lastaccesstime | 	get-date  –f "mm/dd/yyyy"
			write-host -fore green "`t$accountname`t`t$lastaccesstime`t"
			write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
		}
	}
}
else{
		$logmsg = " $computername down "
		write-host $logmsg  -foregroundcolor red
		#if($global:log){ . $env:emp_utils\writelog $logmsg }	
	}
}

getrpcinfos	
			
	

