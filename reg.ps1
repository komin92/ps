param($keyPath)
#$keyPath="Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.a\PersistentHandler"
$keyPath.replace("Microsoft.PowerShell.Core\Registry::" , "")

$find = 'Matrix42'
$replace = 'bla'

(Get-Item $keyPath).Property |
  % {
    $value = (Get-ItemProperty $keyPath $_).$_
    write-host "find: " $value
    if ($value -match $find) {

      Set-ItemProperty $keyPath $_ ($value -replace $find, $replace)
    }
  }