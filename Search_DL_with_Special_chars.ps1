write-host "getting all DLs..." -NoNewline
$DLArray = get-distributiongroup -resultsize Unlimited
write-host " Done"
$array = '~', '!', '@', '#', '$', '%', '^', '&', '(', ')', '.+', '=', '}', '{', '\', '/', '|', ';', ',', ':', '<', '>', '?', '"', '*'

foreach($item in $DLArray){
    [string]$name = $item.name
    $array | foreach{
		if ($name.IndexOf($_) -ge 0){
			Write-Host "$name;$_ is special char"
            Out-File -FilePath "C:\DATA\DL_Special_list.txt" -Encoding Default -InputObject "$name;$_" -NoClobber -Append

		}
	}

}
