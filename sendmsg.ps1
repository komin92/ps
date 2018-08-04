#############################################################################
#																			#
# name: sendmsg.ps1															#
# script version 0.9														#
# 																			#
# author: edi																#
#																			#
# comment: 	send message to pc											 	#
#																			#
# 20.02.2013 :: release														#
#																			#
#############################################################################
[CmdletBinding()]
param(
		[parameter(mandatory=$true,valuefrompipeline=$true)] 
		[string]$computername = "",
		[string]$message="please log off.",		
		[int]$msgverbose ="",
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

$global:log = $log

function sendmsg()
{
	if($debugview){ write-host " $nl sending the following message with a $seconds second delay: $message $nl" -foregroundcolor yellow }
	
	$command = "msg.exe $session /time:$($seconds)"
	if ($computername){$command += " /server:$($computername)"}
	if ($msgverbose){$command += " /v"}
	if ($wait){$command += " /w"}
	$command += " $($message)"

	if((. $env:emp_utils\testconn.ps1 -computername $computername -debugview $debugview -log $global:log -show $show)){
			invoke-expression $command 
			if($lastexitcode -eq 0){
				if($show -eq 1){ write-host " $computername message sent $nl" -foregroundcolor green }
			}
			else{
				$logmsg =  " there was no message sent to $computername."
				if($show -eq 1){ write-host " $computername no message sent" $nl -foregroundcolor red }
				if($global:log){ . $env:emp_utils\writelog $logmsg }			
			}	
		}
		else{
				$logmsg =  " there was no message sent to $computername. "
				if($show -eq 1){ write-host " $computername no message sent" $nl -foregroundcolor red }
				if($global:log){ . $env:emp_utils\writelog $logmsg }					
			}	
}

sendmsg