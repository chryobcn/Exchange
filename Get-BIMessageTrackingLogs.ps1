Param( 
  [Parameter(Mandatory=$false,ParameterSetName='explicit',HelpMessage="Enter Start date as MM/DD/YYYY")][AllowNull()][AllowEmptyString()]
    [ValidateScript({If ($_ -match '(0[1-9]|1[012])[/](0[1-9]|[12][0-9]|3[01])[/](19|20)[0-9]{2}') { $True } Else { Throw "$_ is not a valid date. Follwoing format is mandatorial mm/dd/YYYY !" }})]$StartSearchDate = $Null,
  [Parameter(Mandatory=$false,ParameterSetName='explicit',HelpMessage="Enter End date as MM/DD/YYYY")][AllowNull()][AllowEmptyString()]
    [ValidateScript({If ($_ -match '(0[1-9]|1[012])[/](0[1-9]|[12][0-9]|3[01])[/](19|20)[0-9]{2}') { $True } Else { Throw "$_ is not a valid date. Follwoing format is mandatorial mm/dd/YYYY !" }})]$EndSearchDate = $Null,
  [Parameter(Mandatory=$false,ParameterSetName='explicit',HelpMessage="Enter a single email address")][AllowNull()][AllowEmptyString()]
     [ValidateScript({If ($_ -match "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$") { $True } Else { Throw "$_ is not a valid email address!" }})]$Sender = $Null,    
  [Parameter(Mandatory=$false,ParameterSetName='explicit',HelpMessage="Enter a single email address")][AllowNull()][AllowEmptyString()]
    [ValidateScript({If ($_ -match "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$") { $True } Else { Throw "$_ is not a valid email address!" }})]$Recipient = $Null,
  [Parameter(Mandatory=$false,ParameterSetName='explicit',HelpMessage="Enter part of the Subject text")][AllowNull()][AllowEmptyString()]$MessageSubject = $Null
)

<#
    Purpose: 
        Script to perform a message tracking logs to all Transport servers by using multiple psh instances. It opens a grid once completed. 
        Executing machine must have Exchange Tools installed.

    Notes: 
        Dates must have the following format mm/dd/yyyy, the hours are automatically set.         

    Version:
        1.0 - XR  10/11/2019  Initial Release

    Credits:
        Xavier Rodriguez Ruiz (IT INF - UCS)
#>

[datetime] $today = get-date 
[System.Collections.Arraylist]$global:messagetrackingresults = @() 
$timer = [diagnostics.stopwatch]::startnew() 
$TotalLogsCollected = @()
$messagetrackinglogresults = @() 
[int]$TotalLogsCollected = 0

[array]$hts = Get-TransportService | % {$_.name} #Save all transport servers in an array
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS + 3)

$ScriptBlock = {      
    Param( 
        $MTTargetServer = $NULL, 
        $MTstartsearchdate = $NULL, 
        $MTEndSearchDate = $NULL, 
        $MTTargetSender = $NULL,
        $MTMessageSubject = $NULL, 
        $MTTargetRecipient = $NULL        
    ) 
        Add-PSSnapin *Exchange*         
        get-messagetrackinglog -Server $MTTargetServer -Start $MTstartsearchdate -End $MTEndSearchDate -Sender $MTTargetSender -Recipients $MTTargetRecipient -MessageSubject $MTMessageSubject -resultsize unlimited  | select Timestamp,ClientIp,ClientHostname,ServerIp,ServerHostname,SourceContext,ConnectorId,Source,EventId,InternalMessageId,MessageIdNetworkMessageId,Recipients,RecipientStatus,TotalBytes,RecipientCount,RelatedRecipientAddress,Reference,MessageSubject,Sender,ReturnPath,Directionality,TenantId,OriginalClientIp,MessageInfo,MessageLatency,MessageLatencyType,EventData,TransportTrafficType,SchemaVersion        
} 

function time_pipeline { 
    param ( 
        [int]$increment  = 1000 
    ) 
    begin{$i=0;$timer = [diagnostics.stopwatch]::startnew();$previousSecond=0} 
    process { 
        $i++         
        if ($timer.elapsed.Seconds -notlike $previousSecond){ 
            Write-Progress -Activity "Summarizing Message Tracking Logs" -status “Processed $i records at $(if($timer.elapsed.Hours){"$($timer.elapsed.Hours) Hours "})$(if($timer.elapsed.Minutes){"$($timer.elapsed.Minutes) Mins "})$(if($timer.elapsed.Seconds){"$($timer.elapsed.Seconds) Seconds"})” -PercentComplete (($i / $messagetrackinglogresults.count)  * 100) 
            $previousSecond = $timer.elapsed.Seconds 
        } 
        $_ 
    } 
    end { 
        write-host “Processed $i log records in $(if($timer.elapsed.Hours){"$($timer.elapsed.Hours) Hours "})$(if($timer.elapsed.Minutes){"$($timer.elapsed.Minutes) Mins "})$(if($timer.elapsed.Seconds){"$($timer.elapsed.Seconds) Secs"}else{"$($timer.elapsed.milliseconds) ms"})” 
        Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec.`n" 
    } 
} 

$ProcessEmailStats = {   
    if ($_.eventid -ne $null){
        $Obj = new-object PSObject                
            $Obj | add-member -membertype NoteProperty -name "Timestamp" -value "$($_.Timestamp)"
            $Obj | add-member -membertype NoteProperty -name "ClientIp" -value "$($_.ClientIp)" 
            $Obj | add-member -membertype NoteProperty -name "ClientHostname" -value "$($_.ClientHostname)"
            $Obj | add-member -membertype NoteProperty -name "ServerIp" -value "$($_.ServerIp)" 
            $Obj | add-member -membertype NoteProperty -name "ServerHostname" -value "$($_.ServerHostname)"
            $Obj | add-member -membertype NoteProperty -name "SourceContext" -value "$($_.SourceContext)"
            $Obj | add-member -membertype NoteProperty -name "ConnectorId" -value "$($_.ConnectorId)"                                   
            $Obj | add-member -membertype NoteProperty -name "Source" -value "$($_.EventId)"
            $Obj | add-member -membertype NoteProperty -name "EventId" -value "$($_.ConnectorId)"
            $Obj | add-member -membertype NoteProperty -name "InternalMessageId" -value "$($_.InternalMessageId)"
            $Obj | add-member -membertype NoteProperty -name "MessageIdNetworkMessageId" -value "$($_.MessageIdNetworkMessageId)"
            $Obj | add-member -membertype NoteProperty -name "Recipients" -value "$($_.Recipients)"
            $Obj | add-member -membertype NoteProperty -name "RecipientStatus" -value "$($_.RecipientStatus)"
            $Obj | add-member -membertype NoteProperty -name "TotalBytes" -value "$($_.TotalBytes)"
            $Obj | add-member -membertype NoteProperty -name "RecipientCount" -value "$($_.RecipientCount)"
            $Obj | add-member -membertype NoteProperty -name "RelatedRecipientAddress" -value "$($_.RelatedRecipientAddress)"
            $Obj | add-member -membertype NoteProperty -name "Reference" -value "$($_.Reference)"
            $Obj | add-member -membertype NoteProperty -name "MessageSubject" -value "$($_.MessageSubject)"
            $Obj | add-member -membertype NoteProperty -name "Sender" -value "$($_.Sender)"
            $Obj | add-member -membertype NoteProperty -name "ReturnPath" -value "$($_.ReturnPath)"
            $Obj | add-member -membertype NoteProperty -name "Directionality" -value "$($_.Directionality)"
            $Obj | add-member -membertype NoteProperty -name "TenantId" -value "$($_.TenantId)"
            $Obj | add-member -membertype NoteProperty -name "OriginalClientIp" -value "$($_.OriginalClientIp)"
            $Obj | add-member -membertype NoteProperty -name "MessageInfo" -value "$($_.MessageInfo)"
            $Obj | add-member -membertype NoteProperty -name "MessageLatency" -value "$($_.MessageLatency)"
            $Obj | add-member -membertype NoteProperty -name "MessageLatencyType" -value "$($_.MessageLatencyType)"
            $Obj | add-member -membertype NoteProperty -name "EventData" -value "$($_.EventData)"
            $Obj | add-member -membertype NoteProperty -name "TransportTrafficType" -value "$($_.TransportTrafficType)"
            $Obj | add-member -membertype NoteProperty -name "SchemaVersion" -value "$($_.SchemaVersion)"            
        $messagetrackingresults += $Obj               
    }    
    return $messagetrackingresults 
}

write-host "`nSearch criteria: "
if($StartSearchDate){ $StartSearchDate = get-date $StartSearchDate; write-host "Start Date: $StartSearchDate" }
if($EndSearchDate){ $EndSearchDate = get-date $EndSearchDate; $EndSearchDate = (($EndSearchDate.AddHours(23)).AddMinutes(59)).AddSeconds(59); write-host "End Date: $EndSearchDate" }
if($Recipient) { Write-Host "Recipient: $Recipient" }
if($Sender) { Write-Host "Sender: $Sender" }
if($MessageSubject) { Write-Host "Message Subject: $MessageSubject"}

$Jobs = @() 
$RunspacePool.Open() 

write-host "`nStarting Search Instances for $($hts.count) Exchange Transport servers"
ForEach ($ht in $hts) {                 
    $Param = @{ 
        "MTTargetServer" = $ht;
        "MTstartsearchdate" = $StartSearchDate;
        "MTEndSearchDate" = $today; 
        "MTTargetSender" = $Sender; 
        "MTTargetRecipient" = $Recipient;   
        "MTMessageSubject" = $MessageSubject;
    } 
 
    $Job = [powershell]::Create().AddScript($ScriptBlock)
         
    foreach ($key in $Param.Keys) { 
        $Job.AddParameter($key,$Param.$key) | Out-Null 
    } 
 
    $Job.RunspacePool = $RunspacePool 
    $Jobs += New-Object PSObject -Property @{ 
        Pipe = $Job 
        Result = $Job.BeginInvoke() 
    } 
} 

Do { 
    Write-Progress -Activity "Waiting for Jobs to Complete" -status “Runtime: $(if($timer.elapsed.Hours){"$($timer.elapsed.Hours) Hours "})$(if($timer.elapsed.Minutes){"$($timer.elapsed.Minutes) Mins "})$(if($timer.elapsed.Seconds){"$($timer.elapsed.Seconds) Seconds"})” 
    Start-Sleep -Seconds 1 
} While ( $Jobs.Result.IsCompleted -contains $false ) 
 
ForEach ($Job in $Jobs) { 
    $messagetrackinglogresults += $Job.Pipe.EndInvoke($Job.Result) 
    $Job.Pipe.Dispose() 
}

$RunspacePool.Close() | Out-Null 
$RunspacePool.Dispose() | Out-Null 

Write-Host "Jobs completed!" 

$TotalLogsCollected += $messagetrackinglogresults.count 
 
Write-Host $messagetrackinglogresults.count " message tracking logs collected" 
Write-Host "Total Logs:" $TotalLogsCollected

$TotalRunTime = (get-date) - $today 
Write-Host "`nRun time was $(if($TotalRunTime.Hours){"$($TotalRunTime.Hours) Hours "})$(if($TotalRunTime.Minutes){"$($TotalRunTime.Minutes) Mins "})$(if($TotalRunTime.Seconds){"$($TotalRunTime.Seconds) Seconds"})" 

If ($messagetrackinglogresults.Count -gt 0) { 
    write-host "Concurrently generating tracking logs statistics" 
    $messagetrackingresults = $messagetrackinglogresults | time_pipeline 100 | %{ &$ProcessEmailStats }   
         
    Write-host "Messages matching criteria:" $messagetrackingresults.count
    write-host "Generation Complete." 
    $messagetrackinglogresults = @() 
} 

$messagetrackingresults | Out-GridView
