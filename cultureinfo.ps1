$cultures = "en-US","en-GB","fr-CA","fr-FR","ms-MY","zh-HK","de-at","de-de"

foreach ($c in $cultures)

{

 $culture = New-Object system.globalization.cultureinfo($c)

 $date = get-date -format ($culture.DateTimeFormat.ShortDatePattern)

 New-Object psobject -Property @{"name"=$culture.displayname; "date"=$date}

}