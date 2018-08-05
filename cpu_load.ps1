$cpupercent = @{
  name = 'cpupercent'
  expression = {
    $totalsec = (new-timespan -start $_.starttime).totalseconds
    [math]::round( ($_.cpu * 100 / $totalsec), 2)
  }
}

get-process  | 
 select-object -property name, cpu, $cpupercent, description |
 sort-object -property cpupercent -descending |
 select-object -first 4