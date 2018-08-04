# ps
powershell scripts / modules for sysadmin | devops


#############################################################################
#																			#
# name: aspControlPC.psm1													#
# script version 1.01														#
# 																			#
# author: eduard-albert.prinz												#
#																			#
# comment: 	asp powershell module to control one or more client(s)			#
#			This program is written in the hope that it will be useful, 	#
#			but WITHOUT ANY WARRANTY ^^										#
#																			#
# 16.01.2013 :: pre-alpha													#
# 07.02.2013 :: alpha 														#
# 22.02.2013 :: beta for groups asp											#
# 22.02.2013 :: release shell for groups asp								#
# 25.07.2013 :: prerelease gui for groups asp								#
# 12.08.2013 :: release gui for groups sd, ps                               #
# 03.02.2015 :: prerelease for specific deploymentmanager (experiential)    #
# 04.08.2015 :: release for all deploymentmanager 							#							
#		                                                                    #						
#############################################################################

functions
-----------------------------------------------------------------------------

ASPJOBS : to controlling one or more clients (start sequentiell commands or backgroundjobs)
ASPGETHELP : to get help
ASPGETGUI : to get graphical user interface to control aspjobs via mouse clicks
ASPSHOWLOG : to show log(s) on management station
ASPSHOWEVENT : to show event(s) on management station

-----------------------------------------------------------------------------

ASPJOBS [-help]

optional parameters:
		
-test ...  sends ICMP echo request packets 
-restart ... reboot the operating system
-force ... force reboot or shutdown
-stop ... shutdown the operating system
-send ... send a message 
-debug ... with shell output, what´s going on
-log ... log every working process
-logon ... show currently logged in user 
-lastlogon ... show the last registered user 
-profiles ... show all existing profiles
-wsusupdates ... show asp, ms wsusupdates, wsus report
-services ... start remote services management
-users ... start remote user management
-events ... start remote event management
-cdrive ... open c drive
-ladmins ... view local admins group and user
-space ... view total and freespace disk
-rdp ... start remote desktop
-comp ... start remote computer management
-ip ... show ip address
-mac ... show mac address
-lastboot ... show lastboot and uptime
-userlang ... show existing user profile mui language
-env ... show all existing environment variables
-port ... test port via tcp
-member ... show domain membership
-pkey ... show the product key and os information
-tasks ... show the informations about scheduled tasks
-getservices ... show all services
-setservice [servicename action(stop,start,enable,disable)] ... stop,start,enable or disable service
-eris ... restart empirum remote installation service
-installdate ... get the system installation date
-shares ... get all shares and the share permissions for each share
-stopproc ... stop process [processname]
-proc ... get process
-csv ... import an individuell source witch client(s) for current session
-setup ... del all setup.inf´s -setup c -force
-wake ... sends a specified number of magic packets to a mac address in order to wake up the machine
-adinfos ... get computer infos via ldap 
-ie ... get internet explorer version
-updcount ... get installed ms update count
-updlist ... get installed ms update list
-programs ... get all installed programs
-getsessions ... get all open sessions
-mdrive ... get all mapped drives 
-memo ... get memory usage
-uninstall ... uninstall hotfix -uninstall kb2918614
-supdate ... search hotfix 
-installapps ... get all install applications
-remreg ... remove subkeytree asp team 
-wsusfailcount ... show wsusfailcount
-cpu ... show cpu load sort descending by cpupercent
-fw ... show firewall status
-moni ... show monitor informations
-psv ... show powershell version
-nsv ... show .net version
-printer ... show printer and printer status
-rp ... show reboot pending
-netstat ... show networkig activity
-erislog ... show eris log status (on/off)
-erislogon ... set eris log on
-erislogoff ... set eris log off
-restartlo ... reboot the operating system only if user logged off
-stoplo ... shutdown the operating system only if user logged off
-dontshowlist ... do not show extra computer search result list





ASPJobs [regexpattern] [-test] [-send] [-restart] [-stop] [-force] [-logon] [-lastlogon] [-profiles] [-debug] [-log] 
		[-wsusupdates] [-services] [-users] [-events] [-cdrive] [-ladmins] [-space] [-rdp] [-comp] [-ip] [-mac] 
		[-lastboot] [-userlang] [-env] [-port] [-member] [-pkey] [-tasks] [-getservices] [-setservices] [-eris] [-installdate] 
		[-shares] [-stopproc] [-proc] [-csv] [-setup] [-wake] [-adinfos] [-ie] [0-999 seconds]

default values:

    restart delay:		0 seconds
    send message: 		"please log off"
	
	
ASPJobs examples: 

	ASPJobs d131 -stopproc "wuaserv"
	
	ASPJobs -csv
		please insert the csv-file-path:
		E:\asp\pub_utils\edi.csv
		
	ASPJobs v131*24 -setup c -force //[-setup] {drive} without confirming [-force]	
	
	ASPJobs 55 -setservice "eris;stop"
	
	ASPJobs nigpcr -setservice "eris;start"
	
	ASPJobs d131130 -debug -restart
	ASPJobs auv -send "achtung der pc wird in 60 sekunden heruntergefahren." -restart 60 -log
	ASPJobs 99 -restart 5 -debug
	ASPJobs auv -restart 15 -force -send "poweroff in 15sec." -log
	ASPJobs v131 -log -send
	ASPJobs . -debug -send "nachricht von asp"
	ASPJobs . -logon
	ASPJobs d13 -lastlogon
	ASPJobs v131130 -profiles
	  
	ASPJobs d131130 -restart
	ASPJobs auv -restart -force
	ASPJobs auv -restart 15 -force
	ASPJobs d131130 -restart 10
	ASPJobs d131130 -restart 10 -debug
	ASPJobs d131130 -restart 10 -log
	  
	ASPJobs d131130 -stop
	ASPJobs d131130 -stop 10
	ASPJobs d131130 -stop 10 -debug
	ASPJobs d131130 -stop 10 -log
	  
	ASPJobs d131130 -test [loops]
	ASPJobs d131130 -test 3
	ASPJobs d131130 -test 8
	ASPJobs d131130 -test -debug
	ASPJobs d131130 -test -log
	ASPJobs d131130 -test 4 -log
	  
	ASPJobs d131130 -send
	ASPJobs d131130 -send “hallo computer!”
	ASPJobs d131130 -send -debug
	ASPJobs d131130 -send -log
	
	ASPJobs d131 -setservice "wuaserv;stop"
	ASPJobs d131 -setservice "eris;start"
	ASPJobs d131 -setservice "wuaserv;disabled"
	ASPJobs d131 -setservice "wuaserv;automatic"
	
	ASPJobs d131 -setup "q" -force
	
	ASPJobs d131 -stopproc "powershell.exe"

-----------------------------------------------------------------------------

ASPGETGUI

-----------------------------------------------------------------------------

ASPGETHELP

-----------------------------------------------------------------------------

ASPShowEvent

-----------------------------------------------------------------------------

ASPShowLog [-now]

default: show all log´s

optional parameters:

-now ... show all log from today

ASPShowLog examples:

	ASPShowLog -now

-----------------------------------------------------------------------------
