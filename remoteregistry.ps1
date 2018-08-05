param(
    [parameter(mandatory = $true)]
    $computername,
    [parameter(mandatory = $true)]
    $path
)

set-strictmode -version latest

if($path -match "^hklm:\\(.*)")
{
    $basekey = [microsoft.win32.registrykey]::openremotebasekey(
        "localmachine", $computername)
}
elseif($path -match "^hkcu:\\(.*)")
{
    $basekey = [microsoft.win32.registrykey]::openremotebasekey(
        "currentuser", $computername)
}
else
{
    write-error ("please specify a fully-qualified registry path " +
        "(i.e.: hklm:\software) of the registry key to open.")
    return
}

$key = $basekey.opensubkey($matches[1])
foreach($subkeyname in $key.getsubkeynames())
{
    $subkey = $key.opensubkey($subkeyname)
    $returnobject = [psobject] $subkey
    $returnobject | add-member noteproperty pschildname $subkeyname
    $name = $returnobject | add-member noteproperty property $subkey.getvaluenames()
    $name | % { $subkey.getvalue($_).tostring() }
    $returnobject
    $subkey.close()
}

$key.close()
$basekey.close()