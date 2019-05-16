# OS 07/03/2015
Set-ADServerSettings -ViewEntireForest $true
$date=Get-Date -Format "MMddyyyy-hh:mm"
$whoami=whoami.exe

Write-host " This script will remove SoftDeleted mailboxes in mailbox databases of your choice" -ForegroundColor Green
write-host " #################################################################################" -ForegroundColor Green
Write-Host ""

$filter=read-host("Enter the database filter you want to clean up, eg.: EU_01 or US_*")
$Databasescope=Get-MailboxDatabase | ? {$_.Name -like "*$filter*"}

foreach ($db in $Databasescope) {

      Write-Host "Processing database:" $db
      $Softdeleted=Get-MailboxStatistics -Database $db | where {$_.DisconnectReason -eq "SoftDeleted"}

      IF ([string]::IsNullOrEmpty($softdeleted))
      {
        Write-Host "No softdeleted mailbox found in:" $db -ForegroundColor Yellow
      }
      
      ELSE
      {
        foreach ($single_softdeleted in $Softdeleted) {
        write-host "Found" $single_softdeleted.DisplayName "DisconnectDate" $single_softdeleted.DisconnectDate "Reason" $single_softdeleted.DisconnectReason -ForegroundColor Green
        Remove-StoreMailbox -Database $db -Identity $single_softdeleted.mailboxguid -MailboxState SoftDeleted -Confirm:$false
        #Get-MailboxDatabase $db -Status | ft Name,AvailableNewMailboxSpace
        #Write-Host "the following mailboxes will be deleted:"
        #Echo $single_softdeleted.DisplayName
        }
      }
} 
Write-Host "Processing of: " $Databasescope " done" -ForegroundColor Green
foreach ($db in $Databasescope) {
Get-MailboxDatabase $db -Status | Select Name,AvailableNewMailboxSpace
}
