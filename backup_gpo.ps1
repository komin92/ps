#############################################################################
#																			#
# name: backup_gpo.ps1														#
# script version 0.1														#
# 																			#	
# author: edi																#
#																			#
# comment: 	asp powershell script to backup gpo´s dir(s) as	guid	 		#
#																			#
# 10.04.2013 :: beta														#
#																			#
#############################################################################

$count = 0
write-host "`n backupfolder-path? `n" -f yellow
$answer = read-host
if($answer.length -gt 0){ 		
	$backupfolder = $answer
	get-gpo -all | % {
	  $name = $_.displayname
	  if(!(test-path $backupfolder)){ $backupfolder = new-item $backupfolder -type directory }
	  backup-gpo -guid $_.id -path $backupfolder | out-null
	  $count++
	  write-progress activity " copying $count gpo(s) to $backupfolder "
	}
	write-host "`n copying" $count "gpo(s) to $backupfolder successfully `n" -f green
}