$mailboxtomodify = "be_merialra"
$PublishingEditor = "BI-szrBRUBelgiumMerial-RAPH_BE_WRITE"
$reviewer = "BI-szrBRUBelgiumMerial-RAPH_BE_READ"

$count1 = $null
ForEach($folder in (Get-MailboxFolderStatistics $mailboxtomodify | Where { $_.FolderPath.ToLower().StartsWith("/") -eq $True } ) )
{
$foldername = $mailboxtomodify + ":" + $folder.FolderID; 
$count1++; Write-Host "Foldercount:" $count1
if ($PublishingEditor -ne $Null) {Add-MailboxFolderPermission -erroraction silentlycontinue $foldername -User $PublishingEditor -AccessRights PublishingEditor }
if ($reviewer -ne $Null) {Add-MailboxFolderPermission -erroraction silentlycontinue $foldername -User $Reviewer -AccessRights Reviewer}
}
