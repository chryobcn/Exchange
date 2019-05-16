cls
Set-ADServerSettings -ViewEntireForest $true

$csv = Import-Csv "D:\Scripts\Test\tempcsv.csv" -Delimiter ","
$logfilename = "D:\Scripts\Test\SanofiExport.txt"
if(Test-Path $logfilename ) {$report = $logfilename } Else {$report = New-Item $logfilename -Type file }
Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): ** Start processing **"

foreach($item in $csv){
	$sourceMailbox = Get-Mailbox $item.alias -ErrorAction SilentlyContinue
	if($sourceMailbox -eq $null){write "Unable to find mailbox for $($item.alias)";Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): Unable to find mailbox for $($item.alias)";break} #End in mailbox is NULL
	if($sourceMailbox -is [Array]){write "Multiple mailboxes found for $($item.alias)";Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): Multiple mailboxes found for $($item.alias)";break} #End if multiple values
		
	$ProxyAddresses = $sourceMailbox.EmailAddresses | Where-Object {$_.PrefixString -eq "SMTP"}|Select-Object -ExpandProperty AddressString 
	$X400Addresses = $sourceMailbox.EmailAddresses | Where-Object {$_.PrefixString -eq "X400"}|Select-Object -ExpandProperty AddressString 	
	$X500Addresses = $sourceMailbox.EmailAddresses | Where-Object {$_.PrefixString -eq "X500"}|Select-Object -ExpandProperty AddressString 		
	$LegacyExAddress = $sourceMailbox.legacyexchangedn
	$SANAddress = $item.SAN
	Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): Mailbox $($item.alias)"
	Add-Content $report "  >EmailAddresses: $($sourceMailbox.EmailAddresses)"
	Add-Content $report "  >Legacyexchangedn: $LegacyExAddress"	
	
	Write-Host "Disabling mailbox $($sourceMailbox.alias)" -NoNewline
	Disable-Mailbox $sourceMailbox -Confirm:$false
	do{
		sleep -Milliseconds 500
		$MailboxCheck = Get-Mailbox $item.alias -ErrorAction SilentlyContinue
		Write-Host "." -NoNewline
	}while($MailboxCheck)
	
	Write-Host "`nEnabling MailUser $($sourceMailbox.alias)" -NoNewline
	enable-mailuser -Identity $item.alias -ExternalEmailAddress $SANAddress -Confirm:$false | Out-Null
		do{
		sleep -Milliseconds 500
		$MailUserCheck = Get-MailUser $item.alias -ErrorAction SilentlyContinue
		Write-Host "." -NoNewline
	}while(!($MailUserCheck))
	Set-Mailuser $item.alias -EmailAddressPolicyEnabled $false
	
	Write-Host "`nPress any key to continue ..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown" )
	
	Write-Host "Copying Source proxy addresses to target"	
	foreach ($ProxyAddress in $ProxyAddresses){ Set-MailUser $item.alias -EmailAddresses  @{Add=$ProxyAddress}}
	if($X400Addresses){
		Write-Host "Copying Source X400 addresses to target"
		foreach ($x400Address in $X400Addresses){ write $X400Address;Set-MailUser $item.alias -EmailAddresses  @{Add="X400:$X400Address"}}
	}
	Write-Host "Copying Source LegacyExchangedn address to target"
	write $LegacyExAddress
	Set-MailUser $item.alias -EmailAddresses @{Add="X500:$LegacyExAddress"}
	if($X500Addresses){
		Write-Host "Copying Source X500 addresses to target"		
		foreach ($X500Address in $X500Addresses){ write $X500Address;Set-MailUser $item.alias -EmailAddresses  @{Add="X500:$X500Address"}}
	}
	write "Done."
	Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): Mailbox $($item.alias) Done"
}
Add-Content $report "$((Get-Date).ToShortDateString()) $((Get-Date).ToShortTimeString()): ** End processing **"
