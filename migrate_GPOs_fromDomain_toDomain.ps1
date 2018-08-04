    # -----------------------------------------
    # import needed grouppolicy module
    # -----------------------------------------
    if ((get-module) -notcontains "grouppolicy") {
		import-module grouppolicy
    }     
    # -----------------------------------------
    # define global variables.
    # -----------------------------------------
    $destdomains = @("child01.domain.local","child02.domain.local")
    $sourcedomain = "source.domain.local"
    $sourcepoldef = "\\$sourcedomain\sysvol\$sourcedomain\policies\policydefinitions"
    $backupfolder = "e:\gpobackup"
    $descprefix = "gpo copy 2 other domain - "
    $gpos = $null     
    # -----------------------------------------
    # create backupfolder and/or empty it
    # -----------------------------------------
    if (!(test-path $backupfolder)) {
		new-item -path $backupfolder -type directory
    } else {
		get-childitem $backupfolder -force | remove-item -recurse -force -confirm:$false
    }     
    # -----------------------------------------
    # get all gpo's that need to be copied
    # -----------------------------------------
    write-host "backup selected gpo's" -foregroundcolor green
    $gpos = get-gpo -all -domain $sourcedomain | where {($_.displayname -notlike "default domain policy") -or ($_.displayname -like "<name>")}   
    $gpos | backup-gpo -path $backupfolder     
    # -----------------------------------------
    # start processing the destination domains
    # -----------------------------------------
    foreach ($destdomain in $destdomains) {
		$destpoldef = "\\$destdomain\sysvol\$destdomain\policies"		 
		# -----------------------------------------
		# make sure the policydefinitions are
		# available
		# -----------------------------------------
		write-host "copying to $destdomain" -foregroundcolor green
		copy-item -path $sourcepoldef -destination $destpoldef -recurse -force     
		# -----------------------------------------
		# looping through the gpo's that need to
		# be copied over and create and import it.
		# -----------------------------------------
		write-host "restoring gpo's on $destdomain" -foregroundcolor green
		$gpos | foreach-object {import-gpo -backupgponame $_.displayname -targetname $_.displayname -backuplocation $backupfolder -domain $destdomain -createifneeded}
    }