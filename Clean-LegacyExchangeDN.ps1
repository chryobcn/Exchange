[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=2)]
    [string] $FilePath ##File without headers
)

start-transcript -Path "transcript_Clean-LegacyExDN.txt"
$userarray = Get-Content $filepath
$i=1
foreach ($samaccountname in $userarray){

                $mailbox = Get-User $samaccountname -errorAction 0 | Get-Mailbox -ErrorAction 0
                If ($mailbox){
                    Write-Host "[$i/$($userarray.count)] Processing $($mailbox.name)"
                
                    $dn = $mailbox.LegacyExchangeDN               
                    if($dn -like "* "){
                                   $mailbox | set-mailbox -EmailAddresses @{add="x500:$($dn)"}
                
                                   "'$($dn)'"
                                   Set-ADUser $samaccountname -Replace @{legacyExchangeDN=$dn.Trim()} -Server eu.boehringer.com
                    }
                    Write-host "====================="
                    $i++
                }
}
stop-transcript
