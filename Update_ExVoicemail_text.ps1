$UMPolicies = Get-UMMailboxPolicy

foreach($item in $UMPolicies){
    Write-Host "Working on $($item.name)"
    $WorkingPolicy = Get-UMMailboxPolicy $item.name
    $Logfile = "C:\Data\PSHScripts\ExO_$WorkingPolicy.txt"
    
    [String]$strUMEnabledText = $item.UMEnabledText
    Out-File -FilePath $Logfile -Encoding Default -InputObject "[UMEnabledText] = $strUMEnabledText" -NoClobber -Append; 

    [String]$strResetPINText = $item.ResetPINText
    Out-File -FilePath $Logfile -Encoding Default -InputObject "[ResetPINText] = $strResetPINText" -NoClobber -Append; 
    
    
#    $strUMEnabledText = $strUMEnabledText -replace "https://mybi17.eu.boehringer.com/sites/functional/fin/it/services/itservices/CommunicationServices/Pages/Voice%20Mail%20-%20Training.aspx","http://BI-Voicemail"
#    $strResetPINText = $strResetPINText -replace "https://mybi17.eu.boehringer.com/sites/functional/fin/it/services/itservices/CommunicationServices/Pages/Voice%20Mail%20-%20Training.aspx","http://BI-Voicemail"
    
    $strUMEnabledText = "<p> <font size=""4"" color=""black""> Welcome to <b>Boehringer Ingelheim Exchange Online Unified Messaging.</b> Unified Messaging is BI's global voicemail system. <BR> <BR> <BR> <b><a href=""http://BI-Voicemail"">Click here for TRAINING!</a></b> <BR> <BR> <BR> or paste the following into a browser address bar: <i>http://BI-Voicemail</i> </font> </p>"
    $strResetPINText = "<p> <font size=""4"" color=""black""> Welcome to <b>Boehringer Ingelheim Exchange Online Unified Messaging.</b> Unified Messaging is BI's global voicemail system. <BR> <BR> <BR> <b><a href=""http://BI-Voicemail"">Click here for TRAINING!</a></b> <BR> <BR> <BR> or paste the following into a browser address bar: <i>http://BI-Voicemail</i> </font> </p>"                                            

    Out-File -FilePath $Logfile -Encoding Default -InputObject "====================================================================" -NoClobber -Append; 
    Out-File -FilePath $Logfile -Encoding Default -InputObject "[NEW-UMEnabledText] = $strUMEnabledText" -NoClobber -Append;
    Out-File -FilePath $Logfile -Encoding Default -InputObject "[NEW-ResetPINText] = $strResetPINText" -NoClobber -Append; 

    set-UMMailboxPolicy $item.name -UMEnabledText $strUMEnabledText -ResetPINText $strResetPINText
    
    sleep -Seconds 5
} 
