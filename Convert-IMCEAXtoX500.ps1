[CmdletBinding()]
Param(
    [Parameter (Mandatory=$True,Position=1)]
    [string] $imceaex
)

$x500 = $imceaex.Replace("IMCEAEX-","").Replace("_","/").Replace("+20"," ").Replace("+28","(").Replace("+29",")").Replace("+40","@").Replace("+2E",".").Replace("+2C",",").Replace("+5F","_") #decode DN string
$x500 = $x500.Split("@")[0]   #return only the DN part without @domain
Write-host -BackgroundColor Cyan -ForegroundColor Black $x500
