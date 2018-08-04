
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)]
			[string]$source = "",
			[string]$target = ""
	)
if($source -eq $null -or $target -eq $null){
	write-host " please check params!" -f red;
	exit;
}
# ---------------------------------------------------------------------------
#to customize
# ---------------------------------------------------------------------------
$nl = [environment]::newline;
$global:source=$source;
$global:target=$target;
$global:count = 0;
$global:logdir = "$target\log\"
$logfile = "copy_pics"
$gd = get-date;
$filedate = "$($gd.year)"+"-"+"$($gd.month)"+"-"+"$($gd.day)"+"_"+"$($gd.hour)"+"_"+"$($gd.minute)"+"_"+"$($gd.second)";
$file=$logfile +"_"+ $filedate+".log";
$global:logfile = [string]::join('',($global:logdir, $file))


function ASPwritelog{
	[cmdletbinding()]
	param (
			[string]$msg = ""		
	)
	if(!$(test-path $global:logdir)) {new-item $global:logdir -itemtype directory | out-null}	
	if(!$(test-path $global:logfile)) {new-item -type file $global:logfile -force}
	$logrecord = "------------------------------------------------------------$nl $((get-date).tostring())" 
	$logrecord += " $msg";add-content -path $global:logfile -encoding utf8 -value $logrecord
}

function copyfile{
[cmdletbinding()]
	param (
			[parameter(mandatory=$true,valuefrompipeline=$true)]
			[string]$filename = "",
			[string]$sourcefile = "",
			[string]$dest = "",
			[int]$count = 0
		)
	try{		
		if((test-Path $dest))
		{
			if(test-path "$dest\$filename"){
				ASPwritelog "$count file: $filename already exist! "
			}else{
				ASPwritelog "$count try to copy file: $sourcefile to $dest! "
				copy-item -path $sourcefile -dest $dest
				ASPwritelog "$count copy file: $sourcefile to $dest successfully! " 
			}
		}else{
			$logmsg = " destination: $dest does not exist! $nl"
			ASPwritelog $logmsg;
		}
	}catch{
	$errormessage = $_.exception.message				
	$logmsg = " copy failed - error message was: $errormessage $nl"
	ASPwritelog $logmsg;
}
}

try{
	ASPwritelog " start script!"
	$allFiles = (gci $global:source -recurse -force  | where-object {$_.psiscontainer -eq $false })
	$totalcount = $allFiles.count;$count = 0;	
	$allFiles|%{		
		$lw = $_.lastwritetime
		$directoryname = "$($lw.year)"+"-"+"$($lw.month)"+"-"+"$($lw.day)";
		$targetdirectory = "$global:target\$directoryname";
		write-progress -activity "start backup pics" -status "read pic: $count $_ and copy them to $targetdirectory" -percentcomplete (( $count / $totalcount ) * 100);
		if(!(test-path $targetdirectory)){
			ASPwritelog " try to create dir: $targetdirectory!"
			new-item -itemtype directory -force -path $targetdirectory
			ASPwritelog " create dir: $targetdirectory successfully! " 
		}
		ASPwritelog " dir: $targetdirectory already exist! " 		
		copyfile -filename $_ -sourcefile $_.fullname -dest $targetdirectory -count $count;
		$count++;
		$global:count=$count;
	}

}catch{
	$errormessage = $_.exception.message				
	$logmsg = " failed - error message was: $errormessage $nl"
	ASPwritelog $logmsg;
}
write-progress -completed -activity "completed" -status "completed";
write-host "$nl $global:count files were read from '$global:source' completed! $nl" -f yellow
write-host " logfile: $global:logfile" -f yellow;
ASPwritelog " end script!"
