
# Search AD for comptuer objects
$ComputerInfoFile="c:\\test.txt"
$ObjFilter = "(objectClass=user)" #update the filter with Server specfic filter based on your requirement. 
$objSearch = New-Object System.DirectoryServices.DirectorySearcher
$objSearch.PageSize = 15000
$objSearch.Filter = $ObjFilter
$dsn = get-addomain | select -expand distinguishedname
$objsearch.searchroot = "LDAP://$dsn"
$AllObj = $objSearch.FindAll()
foreach ($Obj in $AllObj)
	   	{	
			$objItemT = $Obj.Properties
			$CName = $objItemT.name
			$objItemT
			write-host "`n"
			$Obj
			write-host "`n"
		}
