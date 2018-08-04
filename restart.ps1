#############################################################################
#																			#
# name: restart.ps1															#
# script version 0.9														#
# 																			#
# author: edi																#
#																			#
# comment: 	restart pc													 	#
#																			#
# 20.02.2013 :: release														#
#																			#
#############################################################################
[CmdletBinding()]
param (
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[validaterange(0,1)] 
		[int]$debugview = 0,
		[validaterange(0,1)] 
		[int]$log = 0,
		[validaterange(0,999)] 
		[int]$sec = 0					
    )	
	

$global:seconds
$global:log = $log

function restart(){

if((. $env:emp_utils\testconn.ps1 -computername $computername -debugview $debugview -log 0 -show 0)){
	
	if($debugview){ write-host "$nl restart computer: $computername $nl" -foregroundcolor yellow }
		
		try { 
				$logmsg = " try restart $computername "
				if($debugview){ write-host $logmsg -foregroundcolor cyan }
				#if($log){ . $env:emp_utils\writelog $logmsg }
				#restart-computer $computername -force 
				$command = "shutdown -r -t $sec -f -m \\"											
				$command += $computername
				
				invoke-expression $command 
				
				if($lastexitcode -eq 0){
					$logmsg = " restart computer: $computername successfully. "
					write-host " $computername restart success"  -foregroundcolor green
					if($global:log){ . $env:emp_utils\writelog $logmsg }	
				}else{
					$logmsg = " restart failed for $computername."
					write-host " $computername restart failed"  -foregroundcolor red
					if($global:log){ . $env:emp_utils\writelog $logmsg }				
				}				
			}
		catch{ 
				$errormessage = $_.exception.message				
				$logmsg = " restart failed for $computername the error message was $errormessage "
				write-host " $computername restart failed"  -foregroundcolor red
				if($global:log){ . $env:emp_utils\writelog $logmsg }		
				continue
			}
}	
else{
		$logmsg = " $computername down "
		write-host $logmsg  -foregroundcolor red
		if($global:log){ . $env:emp_utils\writelog $logmsg }	
	}
}

restart
