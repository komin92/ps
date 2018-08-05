import-module activedirectory
#$user=$env:username
$user="auvt61"
$groups = Get-ADPrincipalGroupMembership $user

    foreach ($group in $groups)
    {
        $username = $user.samaccountname
        $groupname = $group.name
        $line = "$groupname"
        write-host $line
    }