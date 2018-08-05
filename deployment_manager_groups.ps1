
#$username=$Env:USERNAME
$username='prinzu6'



$groups=@{}

$gc="GC://" + $([adsi] "LDAP://RootDSE").Get("RootDomainNamingContext")

$filter = "(&(objectCategory=User)(|(cn=" + $username + ")(samaccountname=" + $username + ")(displayName=" + $username + ")(distinguishedName=" + $username + ")))"
$domain = New-Object System.DirectoryServices.DirectoryEntry($gc)
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot = $domain
$searcher.Filter = $filter
$results = $searcher.FindAll()
if($results.count -eq 0){ "User Not Found"; }else{
 foreach ($result in $results){
  $user=$result.GetDirectoryEntry();
  $user.GetInfoEx(@("tokenGroups"),0)
  $tokenGroups=$user.Get("tokenGroups")
  foreach ($token in $tokenGroups){
   $principal = New-Object System.Security.Principal.SecurityIdentifier($token,0)
   $group = $principal.Translate([System.Security.Principal.NTAccount])
   $groups[$group]=1
  }
 }
}
$a = @();$b=@();$c=@();$clist=@();$nclist=@();[string]$regex = " ";

$groups.keys |%{if($_ -like "`*deploy`*"){$a+=$_;} } 
if($a.count -eq 0){break;}
$a|%{$b+=$_ -split("_");$c+=$b[2];$b=$null;}
$regex="(";
$count=$null;
$c|%{$regex+=$_;if($c.Count-1 -ne $count){$regex+="|";}$count++;}
$regex+=")";
$data = get-content –path $env:path2pclist  | where-object { $_.trim() -ne '' }
$data | % { $f = $_.split(','); $clist += $f[0];}
for($i=0; $i -lt $clist.count; $i++){
try{
				if($clist[$i] -match $regex){ 
					$nclist+=$clist[$i];
}
}catch{
				
				break
			}    	
}
$path="\\"+$env:empirumserver+"\pub_utils$\"+$username+".csv";
$nclist| % {$_ -replace """", ""} | set-content $path 
[Environment]::SetEnvironmentVariable("path2pclist", $path, "Machine")