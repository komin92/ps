#############################################################################
#																			#
# name: sad.ps1																#
# script version 0.9														#
# 																			#	
# author: edi																#
#																			#
# comment: 	search ad								 	#
#																			#
# 20.02.2013 :: release														#
#																			#
#############################################################################

import-module activedirectory

function getclients
{
	$n = 0
	$f = "(objectclass=computer)"
	$srch = new-object system.directoryservices.directorysearcher

	$srch.filter = $f
	#$dsn = get-addomain | select -expand distinguishedname
	$srch.searchroot = "LDAP://dc=asp,dc=t,dc=univie,dc=ac,dc=at"
	#$srch.searchroot = "LDAP://$dsn"
	$clients = $srch.findall()
	foreach ($client in $clients)
	{ 
		$prop = $client.properties
		$cname = $prop.name
		"$cname"
		$n++
	}
	write-host "....$n computer objects found"

	
}

function getusers
{
	
	$f = "(objectcategory=user)"
	$srch = new-object system.directoryservices.directorysearcher
	#$dsn = get-addomain | select -expand distinguishedname
	$srch.searchroot = "LDAP://dc=t,dc=univie,dc=ac,dc=at"
	
	$srch.filter = $f
	$users = $srch.findall()
	foreach ($user in $users)
		{	$prop = $user.properties
			$prop.samaccountname
		}
}


write-host "`t`tpress 1 - ad users" -foregroundcolor cyan
write-host "`t`tpress 2 - ad clients" -foregroundcolor cyan
$whatsup = read-host

switch ($whatsup) 
    { 
        1 {getusers}
		2 {getclients}
		default {write-host "only [1,2]"}
    }
