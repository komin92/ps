#############################################################################
#																			#
# name: ads.ps1																#
# script version 0.9														#
# 																			#	
# author: edi																#
#																			#
# comment: 	get infos from activedirectory								 	#
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
$ou = ""
function getads()
{
	$computer = "'$computername'"
	$lastlogondate = get-adcomputer -filter "name -eq $computer" -prop lastlogondate | select -expand lastlogondate
	$dsn = get-adcomputer -filter "name -eq $computer" -prop distinguishedname | select -expand distinguishedname
	if(($dsn -ne $null) -and ($dsn -ne "")){
		$cn = $dsn.substring(0,$dsn.indexof(",")+1)
		$cnlength = $dsn.length - $dsn.substring(0,$dsn.indexof(",")+1).length
		$newdsn =  $dsn.substring($cn.length,$cnlength)
		$ou =  $newdsn.substring(3, $newdsn.indexof(",")-3)
	}
	
	write-host -fore cyan -nonewline " $computername :" 
	write-host -fore green " $ou`t" $lastlogondate
			
			
	<#write-host -fore cyan "$nl activedirectory infos`t$computer :" $nl
	write-host -fore blue "`t--------------------------------------------------------------------------------------------------"
	write-host -fore yellow "`tlast logon date`t`t`tou"
	write-host -fore blue "`t--------------------------------------------------------------------------------------------------"

	write-host -fore green "`t$lastlogondate`t`t$ou"
	write-host -fore blue "`t--------------------------------------------------------------------------------------------------"#>
}
getads

		
			
	

