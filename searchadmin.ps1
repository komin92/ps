
# Search AD for comptuer objects
$ObjFilter = "(objectClass=Computer)" #update the filter with Server specfic filter based on your requirement. 
$objSearch = New-Object System.DirectoryServices.DirectorySearcher
$objSearch.PageSize = 15000
$objSearch.Filter = $ObjFilter
$dsn = get-addomain | select -expand distinguishedname
$objsearch.searchroot = "ldap://$dsn"
$AllObj = $objSearch.FindAll()
foreach ($Obj in $AllObj)
	   	{	
			$objItemT = $Obj.Properties
			$CName = $objItemT.name
			"$CName" | Out-File $ComputerInfoFile -encoding ASCII -append
		}
