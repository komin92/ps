
param(
	[int]$count = 0
)

$AllClients = Get-ADComputer -Filter {Enabled -eq $False}
foreach ($client in $AllClients)
{
	if ($AllClients)
	{
		$count += 1
		write-Host "$($client.Name) is disabled: $($client.enabled)"	
	} else {write-Host "No disabled computers found!"}	
}
#write-Host "Found $count disabled clients."
#Get-ADComputer -Filter {Enabled -eq $False} #| Remove-ADComputer