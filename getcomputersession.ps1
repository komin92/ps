Function Get-ComputerSessions {
<#
.SYNOPSIS
    Retrieves tall user sessions from local or remote server/s
.DESCRIPTION
    Retrieves tall user sessions from local or remote server/s
.PARAMETER computer
    Name of computer/s to run session query against.
.NOTES
    Name: Get-ComputerSessions
    Author: Boe Prox
    DateCreated: 01Nov2010

.LINK
    https://boeprox.wordpress.org
.EXAMPLE
Get-ComputerSessions -computer "server1"

Description
-----------
This command will query all current user sessions on 'server1'.

#>
[cmdletbinding(
	DefaultParameterSetName = 'session',
	ConfirmImpact = 'low'
)]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline = $True)]
            [string[]]$computer
            )
Begin {
    $report = @()
    }
Process {
    ForEach($c in $computer) {
        # Parse 'query session' and store in $sessions:
        $sessions = query session /server:$c
            1..($sessions.count -1) | % {
                $temp = "" | Select Computer,SessionName, Username, Id, State, Type, Device
                $temp.State = $sessions[$_].Substring(48,8).Trim()
               if($temp.state -eq 'active'){
                $temp.Computer = $c
                $temp.SessionName = $sessions[$_].Substring(1,18).Trim()
                $temp.Username = $sessions[$_].Substring(19,20).Trim()
                $temp.Id = $sessions[$_].Substring(39,9).Trim()
                $temp.State = $sessions[$_].Substring(48,8).Trim()
                $temp.Type = $sessions[$_].Substring(56,12).Trim()
                $temp.Device = $sessions[$_].Substring(68).Trim()
                $report += $temp
               }
               
            }
        }
    }
End {
    $report
    }
}
