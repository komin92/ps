[reflection.assembly]::loadwithpartialname("microsoft.updateservices.administration") 
$wsus = [microsoft.updateservices.administration.adminproxy]::getupdateserver(); 
$wsus.getupdates() | where {$_.isdeclined -eq $true} | foreach-object { if($_.title -match 'Hotfix'){$wsus.deleteupdate($_.id.updateid.tostring()); write-host $_.title removed } }