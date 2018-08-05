#############################################################################
#																			#
# name: approveupdates.ps1													#
# script version 0.9														#
# 																			#	
# author: edi																#
#																			#
# comment: 	approve updates to the windows server update service			#		
#																			#
# 02.11.2012 :: beta														#
#																			#
#############################################################################


[void][reflection.assembly]::loadwithpartialname("microsoft.updateservices.administration")

#log
$date = ( get-date ).tostring('dd_MM_yyyy_HH_mm_ss')
$log = [string]::join('',("c:\", $date, "_approveupdates.txt"))
$file = new-item -type file $log -force
get-date | out-file $file

$nl = [environment]::newline
#$updateserver = read-host "please enter the fqdn of the wsus-server"
$updateserver = "wsussv01.cc.univie.ac.at"

$wsus = [microsoft.updateservices.administration.adminproxy]::getupdateserver($updateserver,$false,"80")
$subscription = $wsus.GetSubscription()
$subscription.StartSynchronization()
$subscription.GetSynchronizationProgress()
write-host $nl"existing groups:"$nl
$groups = $wsus.getcomputertargetgroups()
$count=0
foreach($g in $groups)
{
	write-host $count  $g.name
	$count++
}
write-host $nl
$group = read-host "please select the group"
#$groupname = "edi_comp"
$groupname = $groups[$group].name
write-host  " $nl you select $groupname $nl " -f cyan
$group = $wsus.getcomputertargetgroups() | where {$_.name –eq $groupname}

$updates = $wsus.getupdates()
#$updates = $wsus.searchupdates("windows 7")

#debug
#echo $updates.count

$choicemessage = @"
"please select the updateclassification
	1 service packs
	2 security updates
	3 updates
	4 update rollups
	5 critical updates
	6 feature packs
	7 update rollups
	8 definition updates
	default is all "
"@

$choice = read-host $nl$choicemessage
$classification = "all"
switch ($choice) { 
        1 {$classification = "service packs"} 
        2 {$classification = "security updates"} 
		3 {$classification = "updates"} 
        4 {$classification = "update rollups"} 
        5 {$classification = "critical updates"} 
        6 {$classification = "feature packs"} 
		7 {$classification = "update rollups"} 
		8 {$classification = "definition updates"} 
		default {$classification = "all"}
    }


write-host  " $nl you select $classification $nl " -f cyan
#[microsoft.updateservices.administration.updateapprovalaction] | gm -static -type property | select –expand name	
$choicemessage = @"	
"please select the updateapprovalaction
	1 notapproved
	2 uninstall
	default is install "
"@

$choice = read-host $nl$choicemessage
$approvalaction = "install"
switch ($choice) { 
        1 {$approvalaction = "notapproved"} 
		2 {$approvalaction = "uninstall"}         
		default {$approvalaction = "install"}
    }
write-host  " $nl you select $approvalaction $nl " -f cyan
$count = 0
write-host "$nl try to $approvalaction $classification for $groupname - is that correct? $nl"  -f yellow 
write-host "$nl [y] yes  [n] no (default is 'y'): " -f yellow 
$answer = read-host
if($answer.length -eq 0){ $answer = "y" }
if($answer.tolower() -eq "y"){
		
	foreach($u in $updates)
	{
		#debug
		#write-host $u.title  | where {$u.updateclassificationtitle –ne "security updates"}
		
		if ($classification -eq "all" -xor $u.updateclassificationtitle –eq $classification)
		{		
			$count++
			$updateapprovel = $u.approve($approvalaction,$group)
			"$($count) $($u.title) is $($updateapprovel.state) " | out-file $file -append		
		}
	}
	 
	write-host $nl $count "successfully applied to the action" $approvalaction -f green
	write-host $nl "for more details show log:" $log $nl -f green
}
else { 
	write-host " process was aborted. " -f red
	break
}	