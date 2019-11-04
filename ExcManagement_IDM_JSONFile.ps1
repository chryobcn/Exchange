[CmdletBinding(DefaultParameterSetName='options')]
param
(	        
    [Parameter(HelpMessage='Variable with Json file must be provided.',Mandatory=$true,ParameterSetName='explicit')] $jsonfile,
    [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateRange("DEV","PRO")][String]$Environment
)

<#
    Purpose: Interface script to perform actions required by SalePoint IDM system by parsing a JSON file

    Notes:
        * Don't use -confirm:$false in the string command as is not recognized the switch parameter, use -Force insted
        * Don't use boolean variables in the string command, convert them to Integer values: $true=1 and $false=0
        * Credentials must be stored by a new powershell session on same computer/ user logged by doing as the export maintains the Hash of the exporter:
            * $cred = get-credentials
            * Export-Clixml -InputObject $cred -Path C:\Data\x2bidso365.cred
        * Templates for RemMRSMail and AddMRSMail have not migrated as are decommisioned.

    Version:
        1.0a - XR - 19/10/2019 > Pre-Release

    Credits:
        Xavier Rodriguez Ruiz (IT INF - UCS)
#>

#===============================================================================
# Declarations
#===============================================================================
$jsonDataTable = $null
[string] $LogTimestamp = get-date -Format yyyyMMdd

[String] $errorpath = ""
[String] $completedpath = ""
[string] $VerboseLogfile = "C:\Data\temp\$LogTimestamp ExchangeLog.txt"

#Exchange connections
$global:o365opt = New-PSSessionOption -ProxyAccessType ie -IdleTimeout 300000 -SkipRevocationCheck #5minutes

if($Environment -eq "PRO"){ #PRO
    [String] $global:credPath = "C:\Data\x2bidso365.cred"  
    # $global:exchangeServer = ($server2k16="INHEXMB1601.eu.boehringer.com","INHEXMB1602.eu.boehringer.com","INHEXMB1603.eu.boehringer.com","INHEXMB1604.eu.boehringer.com","INHEXMB1607.eu.boehringer.com","INHEXMB1605.eu.boehringer.com","INHEXMB1608.eu.boehringer.com","INHEXMB1606.eu.boehringer.com","INHEXMB1617.eu.boehringer.com","INHEXMB1618.eu.boehringer.com","INHEXMB1619.eu.boehringer.com","INHEXMB1620.eu.boehringer.com","INHEXMB1621.eu.boehringer.com","INHEXMB1622.eu.boehringer.com","INHEXMB1624.eu.boehringer.com","INHEXMB1623.eu.boehringer.com") | get-random
    $global:exchangeServer = ($server="INHEXCH03.eu.boehringer.com","INHEXCH04.eu.boehringer.com","INHEXCH05.eu.boehringer.com","INHEXCH06.eu.boehringer.com","INHEXCH07.eu.boehringer.com","INHEXCH08.eu.boehringer.com","INHEXCH09.eu.boehringer.com","INHEXCH10.eu.boehringer.com","INHEXCH11.eu.boehringer.com","INHEXCH12.eu.boehringer.com","INHEXCH13.eu.boehringer.com","INHEXCH14.eu.boehringer.com","INHEXCH15.eu.boehringer.com","INHEXCH16.eu.boehringer.com") | get-random
    #Domain Controllers
    [String] $global:EUDC = "INHDC01.eu.boehringer.com"
    [String] $global:AMDC = "NAHDC01.am.boehringer.com"
    [String] $global:APDC = "SINDC02.ap.boehringer.com"
}else{ #DEV
    [String] $global:credPath = "C:\Data\x2bidso365.cred"
    $global:exchangeServer = "INHAS65418.eu.boehringer-dev.com" | Get-Random
    #Domain Controllers
    [String] $global:EUDC = "INHDC01DEV.eu.boehringer-dev.com"
    [String] $global:AMDC = "INHDCAM01DEV.am.boehringer-dev.com"
    [String] $global:APDC = "INHDCAP01DEV.ap.boehringer-dev.com"
}


#===============================================================================
# Functions
#===============================================================================
function Get-TimeStamp {    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)    
}

function fncWriteToLogFile {
    param(        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')] [String]$message,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')] [String]$LogFile
    )
<#
    Purpose: Writes the output on console and also in a log file
#>
    write-host $message
    Out-File -FilePath $Logfile -Encoding Default -InputObject "$message" -NoClobber -Append
}

Function Connect-EXOnpremise {
    param(        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$CmdList,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$exchangeServer,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$CredPath
    )
<#
    Purpose: Connects to Exchange OnPremise
#>

    try{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Connecting to Exchange OnPremise..." 
        $exadmin = Import-Clixml -Path $CredPath        
        $exsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://$exchangeServer/powershell") -Credential $exadmin 
        Import-Module (Import-PSSession $exsession -CommandName $CmdList -AllowClobber) -Global
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Connected."         

        return $exsession
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
        return "ERROR"
        exit
    }

}

Function Connect-EXOnline {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$o365opt,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$cmdList,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$CredPath
    )
<#
    Purpose: Connects to Exchange Online
#>
    $ExOSession = Get-PSSession -ErrorAction silentlycontinue | ?{$_.ComputerName -like "outlook.office365.com"}
    if(!$ExOSession){               
        try{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Connecting to Exchange Online..."             
            $exadmin = Import-Clixml -Path $CredPath
            $sessionName = "ExoSalePoint_" + $exadmin.Username + "$(get-date -f yyyyMMdd_HHmmss)"            
            $ExOSession = New-PSSession -name $sessionName -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $exadmin -Authentication "Basic" -AllowRedirection -SessionOption $o365opt   
            Import-Module (Import-PSSession -Session $ExOSession -AllowClobber -DisableNameChecking -CommandName $cmdList) -Global
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Connected." 

            return $ExOSession
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
            return "ERROR"
            exit
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Already connected to Exchange Online." 
        return $($ExOSession.name)
    }

}

Function fncOpAddAccount{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Alias,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Database,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $MailboxType,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $PrimarySmtpAddress,        
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $ManagedFolderMailboxPolicy       
    )

<#
    Purpose: Enables the Ex mailbox to an AD Account
#>
    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
        return "ERROR"        
    }

    if($objUser.RecipientTypeDetails -like "Remote*"){
       fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud and could not be created."        
       return "ERROR"
    }else{
        #Execute the command in Office 365 with the existing session
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList "Get-Mailbox"
        
        #Check if mailbox on ExO has been created due assigning a license before OnPrem mailbox
        $Mailbox =  Get-Mailbox -Identity $objUser.UserPrincipalName -ErrorAction SilentlyContinue

        #Connect to Ex to perform action
        [Array]$cmdList = "Enable-mailbox","Set-mailbox","Get-mailbox"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist

        if($Mailbox){
            fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud, cloud mailbox must be offboarded to onpremise."
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
            return "ERROR"            
        }else{
            #No mailbox created in the cloud due license missmatch
            try{
                if($PrimarySmtpAddress -ne ""){
                    if($MailboxType -eq "Regular"){
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox  -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                        $Enable = Enable-mailbox  -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;               
                    }else{
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                        $Enable = Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;                                     
                    }
                }else{ 
                    #Creating mailbox without specific primary smtp address
                    if($MailboxType -eq "Regular"){
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox  -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                        $Enable = Enable-mailbox  -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;                                    
                    }else{
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                        $Enable = Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;                                                                                         
                    }
                }
                   
                return "OK"
            }catch{
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Enable-Mailbox for $Identity has failed. Trying Set-Mailbox."                  
                $error.clear()
                try{
                    [String]$command =  "Set-mailbox -Type $MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -Force -ErrorAction Stop"
                    fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"                       
                    Invoke-Expression -Command $command 
                                         
                    return "OK"
                }catch{
                    $errmsg=[string]$_.Exception.Message
                    $error.clear()
                    if ($errmsg -notlike "*is already of the type*"){
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Error: $errmsg." 
                        return "ERROR"
                    }
                }
                finally {            
                    Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
                }
            }
            finally {            
                Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
            }
        }
    }
}

Function fncOpFixLegacyDN {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: Removes the tailspace on the ExchangeLegacyDN
#>
    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."         
        return "ERROR"        
    }

    if($objUser.RecipientTypeDetails -like "Remote*"){
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud and this operation is not allowed."           
        return "ERROR"
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){
        #Connect to Ex 
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-mailbox" 

        try{
            $target = Get-Mailbox -Identity $Identity -DomainController $DomainController | Select-Object SamAccountName,LegacyExchangeDN

            $userid = $target.SamAccountName
            $dn = $target.LegacyExchangeDN
			Import-Module ActiveDirectory -cmdlet Get-ADUser,Set-ADUser
            
            if($dn -like "* "){                      
				$dnrep = $dn.trim()

                [String]$command =  "Set-ADUser $userid -Replace @{legacyExchangeDN=$dnrep} -Server $DomainController -ErrorAction Stop"
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
                Invoke-Expression -Command $command  				                				                      
            }
            return "OK"   
        }
        catch{
            $errmsg=[string]$_.Exception.Message
            $error.clear()
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Error: $errmsg." 
            return "ERROR"
        }
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: Recipient Type $($objUser.RecipientTypeDetails) and this operation is not allowed."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}

Function fncOpAddExOSMTP{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: Removes the tailspace on the ExchangeLegacyDN
#>
    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."   
        return "ERROR"        
    }

    if($objUser.RecipientTypeDetails -like "Remote*"){
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud and this operation is not allowed." 
        return "ERROR"
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){
        #Connect to Ex 
        [Array]$cmdList = "Get-mailbox","Set-mailbox"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 

        try{
            $target = Get-Mailbox $Identity | select samaccountname,PrimarySmtpAddress,emailaddresses 

            if($target.emailaddresses -notlike "*@boehringer.mail.onmicrosoft.com") {
                $sPrimarySmtpAddress = [string]($target | select -ExpandProperty PrimarySmtpAddress)
                $sOnMicrosoftSmtpAddress = ($sPrimarySmtpAddress).split("@") | Select-Object -Index 0
                $sOnMicrosoftSmtpAddress = $sOnMicrosoftSmtpAddress+"@boehringer.mail.onmicrosoft.com"
                
                [String]$command = "Set-Mailbox -Identity $Identity -DomainController $DomainController -EmailAddresses @{add=smtp:$sOnMicrosoftSmtpAddress}"
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
                Invoke-Expression -Command $command                                  
            }
            return "OK"
        }catch{
            $errmsg=[string]$_.Exception.Message
            $error.clear()
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Error: $errmsg."   
            return "ERROR"
        }
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: Recipient Type $($objUser.RecipientTypeDetails) and this operation is not allowed."   
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}

function  fncOpAddAccountLegalHold{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$LitigationHoldEnabled,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: enables and set parameters for Litigation hold mailboxes
#>
    
    ##############
    ##############
    ##############
    #NO ACTION DONE FOR EXO MAILBOX. NEED TO CHECK
    ##############
    ##############
    ##############

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
           
    if(($objUser.RecipientTypeDetails -eq "UserMailbox") -or ($objUser.RecipientTypeDetails -eq "SharedMailbox")){
        #Connect to Ex 
        [Array]$cmdList = "Get-mailbox","Set-mailbox"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 
		
		try{
            [Boolean]$LitigationHoldEnabled = [System.Convert]::ToBoolean($LitigationHoldEnabled)                         
 
			if($LitigationHoldEnabled -eq $true){ # Enable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 1 -UseDatabaseQuotaDefaults 0 -IssueWarningQuota `"Unlimited`" -ProhibitSendReceiveQuota `"Unlimited`" -ProhibitSendQuota `"Unlimited`" -Force -ErrorAction Stop"
				fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command                 
			}else{ # Disable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 0 -Force -ErrorAction Stop"				
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command 
			}
                       
            return "OK"
		
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  			
			$Error.clear()
			return "ERROR";
		}
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."          
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}


function funcopSetSpecSetsDefault{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $OwaMailboxPolicy,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: set the OWA mailbox policy
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
    
    [String] $cmdList = "Set-CASMailbox"

    if($OwaMailboxPolicy -ne ""){
        if($objUser.RecipientTypeDetails -like "Remote*"){
            #Connect to Ex to perform action
            $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
            
            [String]$OptionalDC = ""
        }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){
            #Connect to Ex to perform action
            $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList
            
            [String]$OptionalDC = "-DomainController $DomainController"
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."              
            return "ERROR"
        }
        
        try{
            [String]$command = "Set-CASMailbox -Identity $objUser.UserPrincipalName -OwaMailboxPolicy $OwaMailboxPolicy $OptionalDC -Force -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"              
            Invoke-Expression -Command $command   
            return "OK"
                 
        }catch{
			fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
			$Error.clear()
			return "ERROR";
		}    

    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function funcopHidAccount{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$HiddenFromAddressListsEnabled,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: set the parameters when account is hidden
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Set-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Set-Mailbox" }
    elseif($objUser.RecipientTypeDetails -eq "MailUser"){ [String]$cmdlet = "Set-Mailuser" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    [String]$OptionalDC = "-DomainController $DomainController"
    [String] $cmdList = "Set-Mailbox,Set-RemoteMailbox,Set-MailUser"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   

    try{
        [Boolean]$HiddenFromAddressListsEnabled = [System.Convert]::ToBoolean($HiddenFromAddressListsEnabled)

        if($HiddenFromAddressListsEnabled -eq $true){
            [String]$Command = "$cmdlet -Identity $objUser.UserPrincipalName -HiddenFromAddressListsEnabled 1 -RequireSenderAuthenticationEnabled 1 -AcceptMessagesOnlyFromSendersOrMembers $objUser.UserPrincipalName $OptionalDC -Force -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command              
	    }else{
            [String]$Command = "$cmdlet -Identity $objUser.UserPrincipalName -HiddenFromAddressListsEnabled 0 -RequireSenderAuthenticationEnabled 0 -AcceptMessagesOnlyFromSendersOrMembers $null $OptionalDC -Force -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"            
            Invoke-Expression -Command $Command              
        }

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpAddSecMail {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $SecondaryMail,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: set proxyaddresses on a recipient
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Set-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Set-Mailbox" }
    elseif($objUser.RecipientTypeDetails -eq "MailUser"){ [String]$cmdlet = "Set-Mailuser" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    [String] $OptionalDC = "-DomainController $DomainController"
    [String] $cmdList = "Set-Mailbox,Set-RemoteMailbox,Set-MailUser"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   

    try{
        if($SecondaryMail -ne ""){
            [String]$Command = "$cmdlet -Identity $objUser.UserPrincipalName -EmailAddresses @{add=$SecondaryMail} $OptionalDC -Force -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command             
	    }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: SecondaryMail value is empty "                            
        }

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpRemSecMail {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $SecondaryMail,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: Removes proxyaddresses on a recipient
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Set-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Set-Mailbox" }
    elseif($objUser.RecipientTypeDetails -eq "MailUser"){ [String]$cmdlet = "Set-Mailuser" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    [String] $OptionalDC = "-DomainController $DomainController"
    [String] $cmdList = "Set-Mailbox,Set-RemoteMailbox,Set-MailUser"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   

    try{
        if($SecondaryMail -ne ""){
            [String]$Command = "$cmdlet -Identity $objUser.UserPrincipalName -EmailAddresses @{remove=$SecondaryMail} $OptionalDC -Force -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command             
	    }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: SecondaryMail value is empty "                            
        }

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpSetExtOoof{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $ExternalOofOptions,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: set Oof Audience
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
        
    [String] $cmdList = "Set-Mailbox"
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""
        #Connect to Ex to perform action
        $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  
        [String]$OptionalDC = "-DomainController $DomainController"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    try{
        [String]$Command = "Set-Mailbox -Identity $objUser.UserPrincipalName -ExternalOofOptions $ExternalOofOptions $OptionalDC -Force -ErrorAction Stop"       
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command              
	   
        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}


function fncOpSetSpecSets {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$UseDatabaseRetentionDefaults,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$SingleItemRecoveryEnabled,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $RetentionPolicy,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $RetainDeletedItemsUntilBackup,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $RetainDeletedItemsFor,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )

<#
    Purpose: set default settings for a mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [String] $cmdList = "Set-Mailbox"
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""
        #Connect to Ex to perform action
        $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  
        [String]$OptionalDC = "-DomainController $DomainController"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    try{
        if($RetentionPolicy -ne ""){
            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -RetentionPolicy $RetentionPolicy -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
		    Invoke-Expression -Command $command              
        }
        
        if($UseDatabaseRetentionDefaults -ne ""){
            [Boolean]$UseDatabaseRetentionDefaults = [System.Convert]::ToBoolean($UseDatabaseRetentionDefaults)
            $UseDatabaseRetentionDefaults = [int]$UseDatabaseRetentionDefaults 

            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -UseDatabaseRetentionDefaults $UseDatabaseRetentionDefaults -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
		    Invoke-Expression -Command $command             
        }
        
        if($SingleItemRecoveryEnabled -ne ""){
            [Boolean]$SingleItemRecoveryEnabled = [System.Convert]::ToBoolean($SingleItemRecoveryEnabled)
            $SingleItemRecoveryEnabled = [int]$SingleItemRecoveryEnabled 

            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -SingleItemRecoveryEnabled $SingleItemRecoveryEnabled -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
		    Invoke-Expression -Command $command 
        }
        
        if($RetainDeletedItemsUntilBackup -ne ""){
            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -RetainDeletedItemsUntilBackup $RetainDeletedItemsUntilBackup -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
		    Invoke-Expression -Command $command               
        }            

        if($RetainDeletedItemsFor -ne ""){
            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -RetainDeletedItemsFor $RetainDeletedItemsFor -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
		    Invoke-Expression -Command $command             
        }
	   
        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}


function fncOpSetPrimMail {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $PrimarySmtpAddress,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$EmailAddressPolicyEnabled,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )

<#
    Purpose: set primary emailaddress for a mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
   
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String] $cmdlet = "Set-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  [String] $cmdlet = "Set-Mailbox" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    [String]$OptionalDC = "-DomainController $DomainController"
    [String] $cmdList = "Set-Mailbox,Set-RemoteMailbox"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 

    try{
        if($PrimarySmtpAddress -ne ""){
            #Converting boolean values to Int to be accepted in a String
            [Boolean]$EmailAddressPolicyEnabled = [System.Convert]::ToBoolean($EmailAddressPolicyEnabled)
            $EmailAddressPolicyEnabled = [Int]$EmailAddressPolicyEnabled

            [String]$Command = "$cmdlet -Identity $objUser.UserPrincipalName -PrimarySmtpAddress $PrimarySmtpAddress -EmailAddressPolicyEnabled $EmailAddressPolicyEnabled -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }

}

function fncOpAddDiLRole {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DistributionGroup                                           
    )

<#
    Purpose: Add a member on a Distribution Group
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-DistributionGroup,Add-DistributionGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    # check OU of the DistributionList to retrieve the DC
	$currDC = $null
	try{
        $currServer = Get-DistributionGroup -Identity $DistributionGroup -ErrorAction Stop | Select-Object -ExpandProperty OrganizationalUnit
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }

	$currServerSplit= $currServer.split("/")
	$currDomain = $currServerSplit[0]
	If ($currDomain -eq "eu.boehringer.com"){ $currDC = $EUDC }
	elseif ($currDomain -eq "am.boehringer.com"){ $currDC = $AMDC }
	elseif ($currDomain -eq "ap.boehringer.com"){ $currDC = $APDC }
           	
    try{
	    $NewMember = $Identity.split("\");
	    $NewMemberAlias = $NewMember[1];
	    
        if($currDC -ne $null) {
            [String]$OptionalDC = "-DomainController $currDC"               	    	    
	    }else{
            [String]$OptionalDC = ""	    	    
	    }

        [String]$Command = "Add-DistributionGroupMember -Identity $DistributionGroup -Member $NewMemberAlias $OptionalDC -Force -bypassSecurityGroupManagerCheck -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike '*is already a member of the group*' )
	    {
		    fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
	    }else {            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info:  $NewMemberAlias is already member of DistributionGroup $DistributionGroup"  
            return "OK"
        }
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpRemDiLRole {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DistributionGroup                                           
    )

<#
    Purpose: Add a member on a Distribution Group
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-DistributionGroup,Remove-DistributionGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    # check OU of the DistributionList to retrieve the DC
	$currDC = $null
	try{
        $currServer = Get-DistributionGroup -Identity $DistributionGroup -ErrorAction Stop | Select-Object -ExpandProperty OrganizationalUnit
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }

	$currServerSplit= $currServer.split("/")
	$currDomain = $currServerSplit[0]
	If ($currDomain -eq "eu.boehringer.com"){ $currDC = $EUDC }
	elseif ($currDomain -eq "am.boehringer.com"){ $currDC = $AMDC }
	elseif ($currDomain -eq "ap.boehringer.com"){ $currDC = $APDC }
           	
    try{
	    $RemMember = $Identity.split("\");
	    $RemMemberAlias = $RemMember[1];
	    
        if($currDC -ne $null) {
            [String]$OptionalDC = "-DomainController $currDC"               	    	    
	    }else{
            [String]$OptionalDC = ""	    	    
	    }

        [String]$Command = "Remove-DistributionGroupMember -Identity $DistributionGroup -Member $RemMemberAlias $OptionalDC -Force -bypassSecurityGroupManagerCheck -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike "*isn't a member of the group*" )
	    {
		    fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
	    }else {            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: $NewMemberAlias is already not member of DistributionGroup $DistributionGroup"  
            return "OK"
        }
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpSetCalProc {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,        
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $Notes,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$ResourceDelegates,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$ForwardRequestsToDelegates,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$AutomateProcessing,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$BookInPolicy,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$DeleteAttachments,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$RemovePrivateProperty,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$DeleteSubject,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$DeleteNonCalendarItems,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$AllBookInPolicy,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$AllRequestInPolicy,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$AddAdditionalResponse,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$AdditionalResponse,       
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()]$DeleteComments                                                  
    )

<#
    Purpose: set parameters on a Room/Equipment Mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [String] $cmdList = "Set-CalendarProcessing,Get-Recipient"

    #Check location of mailbox  
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""
        #Connect to Ex to perform action
        $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  
        [String]$OptionalDC = "-DomainController $DomainController"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        return "ERROR"
    }

    try{
        if($AutomateProcessing -ne $null -and $AutomateProcessing -ne ""){
            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -AutomateProcessing $AutomateProcessing"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($ForwardRequestsToDelegates -ne $null -and $ForwardRequestsToDelegates -ne ""){
            [Boolean]$ForwardRequestsToDelegates = [System.Convert]::ToBoolean($ForwardRequestsToDelegates)
            $ForwardRequestsToDelegates = [Int]$ForwardRequestsToDelegates

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -ForwardRequestsToDelegates $ForwardRequestsToDelegates"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($DeleteAttachments -ne $null -and $DeleteAttachments -ne ""){
            [Boolean]$DeleteAttachments = [System.Convert]::ToBoolean($DeleteAttachments)
            $DeleteAttachments = [Int]$DeleteAttachments

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -DeleteAttachments $DeleteAttachments"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($DeleteComments -ne $null -and $DeleteComments -ne ""){
            [Boolean]$DeleteComments = [System.Convert]::ToBoolean($DeleteComments)
            $DeleteComments = [Int]$DeleteComments

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -DeleteComments $DeleteComments"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($RemovePrivateProperty -ne $null -and $RemovePrivateProperty -ne ""){
            [Boolean]$RemovePrivateProperty = [System.Convert]::ToBoolean($RemovePrivateProperty)
            $RemovePrivateProperty = [Int]$RemovePrivateProperty

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -RemovePrivateProperty $RemovePrivateProperty"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($DeleteSubject -ne $null -and $DeleteSubject -ne ""){
            [Boolean]$DeleteSubject = [System.Convert]::ToBoolean($DeleteSubject)
            $DeleteSubject = [Int]$DeleteSubject

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -DeleteSubject $DeleteSubject"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($DeleteNonCalendarItems -ne $null -and $DeleteNonCalendarItems -ne ""){
            [Boolean]$DeleteNonCalendarItems = [System.Convert]::ToBoolean($DeleteNonCalendarItems)
            $DeleteNonCalendarItems = [Int]$DeleteNonCalendarItems

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -DeleteNonCalendarItems $DeleteNonCalendarItems"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($AllBookInPolicy -ne $null -and $AllBookInPolicy -ne ""){
            [Boolean]$AllBookInPolicy = [System.Convert]::ToBoolean($AllBookInPolicy)
            $AllBookInPolicy = [Int]$AllBookInPolicy

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -AllBookInPolicy $AllBookInPolicy"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($AllRequestInPolicy -ne $null -and $AllRequestInPolicy -ne ""){
            [Boolean]$AllRequestInPolicy = [System.Convert]::ToBoolean($AllRequestInPolicy)
            $AllRequestInPolicy = [Int]$AllRequestInPolicy

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -AllRequestInPolicy $AllRequestInPolicy"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($AddAdditionalResponse -ne $null -and $AddAdditionalResponse -ne ""){ #For future implementation
            [Boolean]$AddAdditionalResponse = [System.Convert]::ToBoolean($AddAdditionalResponse)
            $AddAdditionalResponse = [Int]$AddAdditionalResponse

            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -AddAdditionalResponse $AddAdditionalResponse"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if($AdditionalResponse -ne $null -and $AdditionalResponse -ne ""){ #For future implementation
            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -AdditionalResponse $AdditionalResponse"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command
        }
        if(($ResourceDelegates -ne $null) -and ($ResourceDelegates -ne "")){                               
            $resourceDelegatesList = @()
            $ArResDelegates = $ResourceDelegates
            $ArResDelegates = $ArResDelegates.split(",")
            foreach($item in $ArResDelegates){                
                $DelegateType = Get-Recipient -Identity $item -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RecipientTypeDetails
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: Account $item is type $DelegateType" 
                if($DelegateType -like "*Mailbox"){
                    $resourceDelegatesList += $item
                }
            }
            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -ResourceDelegates $resourceDelegatesList"    
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command                             
        }
        #Workarround until BIDSG/SalePoint doesnt sent the BookInPolicy parameter
        $BookInPolicy = $ResourceDelegates
        #End Workarround
        if(($BookInPolicy -ne $null) -and ($BookInPolicy -ne "") -and ($Notes -eq "closed" -or $Notes -eq "moderated")){                             
            $BookInPolicyList = @()
			$ArBookInPolicy = $BookInPolicy
			$ArBookInPolicy = $ArBookInPolicy.split(",")
            foreach($item in $ArBookInPolicy){
                    $AccType = Get-Recipient -Identity $item -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RecipientTypeDetails
                    fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: Account $item is type $AccType" 
                    if($AccType -like "*Mailbox"){
                            $BookInPolicyList += $item
                    }
            }                                  
            [String]$Command = "Set-CalendarProcessing -Identity $Identity $OptionalDC -ErrorAction Stop -BookInPolicy $BookInPolicyList"    
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		    Invoke-Expression -Command $Command                                         
        }

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpRevokeEUUM {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )

<#
    Purpose: Removes Exchange UM for a mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [String] $cmdList = "Disable-UMMailbox"
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""
        #Connect to Ex to perform action
        $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  
        [String]$OptionalDC = "-DomainController $DomainController"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue         
        return "ERROR"
    }

    try{
        [String]$Command = "Disable-UMMailbox -Identity $objUser.UserPrincipalName $OptionalDC -ErrorAction Stop"    
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command 
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message
		$error.clear()
		if ($errmsg -notlike '*is already disabled for Unified Messaging*' )
		{
			fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
			return "ERROR";
		}         
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpAssignEUUM {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$UMMailboxPolicy,                                           
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Extensions, 
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$SIPResourceIdentifier, 
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Creates Exchange UM for a mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [String] $cmdList = "Enable-UMMailbox"
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""
        #Connect to Ex to perform action
        $ExSession = Connect-EXOnline -o365opt $o365opt -CredPath $credPath -cmdList $cmdList
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  
        [String]$OptionalDC = "-DomainController $DomainController"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList   
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
        [String]$Command = "Enable-UMMailbox -Identity $objUser.UserPrincipalName -UMMailboxPolicy $UMMailboxPolicy -Extensions $Extensions -SIPResourceIdentifier $SIPResourceIdentifier $OptionalDC -ErrorAction Stop"    
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command 
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message
		$error.clear()
		if ($errmsg -notlike '*is already enabled for Unified Messaging*' )
		{
			fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
			return "ERROR";
		}         
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }

}

function fncOpSetAlias {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Alias,                                           
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Assign the Alias for an Ex Object
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [String] $cmdList = "Set-Mailbox"
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String] $cmdList = "Set-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){  [String] $cmdList = "Set-Mailbox"  }
    elseif($objUser.RecipientTypeDetails -eq "MailUser"){  [String] $cmdList = "Set-MailUser" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    [String]$OptionalDC = "-DomainController $DomainController"
    [String] $cmdlet = $cmdList
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 

    try{
        [String]$command = "$cmdlet -Identity $objUser.UserPrincipalName -Alias $Alias -Force $OptionalDC -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
		Invoke-Expression -Command $command 
        
        Return "OK"    
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpSetUser {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Notes,                                           
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Set the Notes for an AD User
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User,Set-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "*Mailbox"){  
        try{
            [String]$command = "Set-User -Identity $objUser.UserPrincipalName -Notes $Notes -Force $OptionalDC -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
		    Invoke-Expression -Command $command 
        
            Return "OK"    
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		    $Error.clear()
		    return "ERROR";
        }
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }     
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}

Function fncOpAddRoomAccount{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Alias,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Database,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $MailboxType,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String] $PrimarySmtpAddress,        
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $ManagedFolderMailboxPolicy       
    )

<#
    Purpose: Enables the Ex room mailbox to an AD Account
#>
    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
        return "ERROR"        
    }

    if($objUser.RecipientTypeDetails -like "Remote*"){
       fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud and could not be created."        
       return "ERROR"
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account already enabled as a mailbox object."        
       return "ERROR"
    }else{
        #Execute the command in Office 365 with the existing session
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList "Get-Mailbox"
        
        #Check if mailbox on ExO has been created due assigning a license before OnPrem mailbox
        $Mailbox =  Get-Mailbox -Identity $objUser.UserPrincipalName -ErrorAction SilentlyContinue

        if($Mailbox){
            fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is on cloud, cloud mailbox must be offboarded to onpremise."
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
            return "ERROR"            
        }else{
            #No mailbox created in the cloud due license missmatch

            #Connect to Ex to perform action
            [Array]$cmdList = "Enable-mailbox","Set-mailbox","Get-mailbox"
            $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist

            try{
                if($PrimarySmtpAddress -ne ""){                       
                    fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                    $Enable = Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -PrimarySmtpAddress $PrimarySmtpAddress -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;                                                         
                }else{ 
                    #Creating mailbox without specific primary smtp address                    
                    fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop"
                    $Enable = Enable-mailbox -$MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -force -ErrorAction Stop;                                                                                                             
                }
                   
                return "OK"
            }catch{
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Enable-Mailbox for $Identity has failed. Trying Set-Mailbox."                  
                $error.clear()
                try{
                    [String]$command =  "Set-mailbox -Type $MailboxType -Identity $Identity -Alias $Alias -Database $Database -ManagedFolderMailboxPolicy $ManagedFolderMailboxPolicy -DomainController $DomainController -Force -ErrorAction Stop"
                    fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"                       
                    Invoke-Expression -Command $command 
                                         
                    return "OK"
                }catch{
                    $errmsg=[string]$_.Exception.Message
                    $error.clear()
                    if ($errmsg -notlike "*is already of the type*"){
                        fncWriteToLogFile -LogFile $VerboseLogfile -message "Error: $errmsg." 
                        return "ERROR"
                    }
                }
                finally {            
                    Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
                }
            }
            finally {            
                Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
            }
        }
    }
}

function fncOpAddMailEnabledAccount {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Alias,                                           
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$ExternalEmailAddress,                              
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Creates a Mail Enabled User
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User,Enable-MailUser"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
        return "ERROR"        
    }

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -notlike "*Mailbox"){  
        try{
            [String]$command = "Enable-MailUser -Identity $objUser.UserPrincipalName -Alias $Alias -ExternalEmailAddress $ExternalEmailAddress -PrimarySmtpAddress $ExternalEmailAddress -DomainController $DomainController -ErrorAction Stop"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
		    Invoke-Expression -Command $command 
        
            Return "OK"    
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		    $Error.clear()
		    return "ERROR";
        }
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }     
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

}

function fncOpAddManagBy {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,    
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DistributionGroup                                           
    )

<#
    Purpose: Add a member on a Distribution Group
            Needs to be confirmed by BIDS as on old template they used Set-CC_DistributionGroupManagedBy -bicn $NewMemberAlias -DGName "[DGName]" -Methode 'add' -DomainController "[DomainController]"
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-DistributionGroup,Add-DistributionGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    # check OU of the DistributionList to retrieve the DC
	$currDC = $null
	try{
        $currServer = Get-DistributionGroup -Identity $DistributionGroup -ErrorAction Stop | Select-Object -ExpandProperty OrganizationalUnit
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }

    $currServerSplit= $currServer.split("/")
	$currDomain = $currServerSplit[0]
	If ($currDomain -eq "eu.boehringer.com"){ $currDC = $EUDC }
	elseif ($currDomain -eq "am.boehringer.com"){ $currDC = $AMDC }
	elseif ($currDomain -eq "ap.boehringer.com"){ $currDC = $APDC }

    try{
	    $NewMember = $Identity.split("\");
	    $NewMemberAlias = $NewMember[1];
	    
        if($currDC -ne $null) {
            [String]$OptionalDC = "-DomainController $currDC"               	    	    
	    }else{
            [String]$OptionalDC = ""	    	    
	    }

        [String]$Command = "Add-DistributionGroupMember -Identity $DistributionGroup -Member $NewMemberAlias $OptionalDC -Force -bypassSecurityGroupManagerCheck -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike '*is already a member of the group*' )
	    {
		    fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
	    }else {            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info:  $NewMemberAlias is already member of DistributionGroup $DistributionGroup"  
            return "OK"
        }
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpRemManagBy {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,    
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DistributionGroup                                           
    )

<#
    Purpose: Removes a member on a Distribution Group
            Needs to be confirmed by BIDS as on old template they used Set-CC_DistributionGroupManagedBy -bicn $NewMemberAlias -DGName "[DGName]" -Methode 'remove' -DomainController "[DomainController]"
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-DistributionGroup,Remove-DistributionGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    # check OU of the DistributionList to retrieve the DC
	$currDC = $null
	try{
        $currServer = Get-DistributionGroup -Identity $DistributionGroup -ErrorAction Stop | Select-Object -ExpandProperty OrganizationalUnit
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }

    $currServerSplit= $currServer.split("/")
	$currDomain = $currServerSplit[0]
	If ($currDomain -eq "eu.boehringer.com"){ $currDC = $EUDC }
	elseif ($currDomain -eq "am.boehringer.com"){ $currDC = $AMDC }
	elseif ($currDomain -eq "ap.boehringer.com"){ $currDC = $APDC }

    try{
	    $NewMember = $Identity.split("\");
	    $NewMemberAlias = $NewMember[1];
	    
        if($currDC -ne $null) {
            [String]$OptionalDC = "-DomainController $currDC"               	    	    
	    }else{
            [String]$OptionalDC = ""	    	    
	    }

        [String]$Command = "Remove-DistributionGroupMember -Identity $DistributionGroup -Member $NewMemberAlias $OptionalDC -Force -bypassSecurityGroupManagerCheck -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $Command"  
		Invoke-Expression -Command $Command
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike "*isn't a member of the group*" ){
		    fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
	    }else{            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info:  $NewMemberAlias is already member of DistributionGroup $DistributionGroup"  
            return "OK"
        }
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpRemAccount {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Disables a mailbox or mailuser
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike "*couldn't be found on*" ){
		    fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
	    }else{            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: User $Identity Not Found"  
            return "OK"
        }      
    }
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Disable-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Disable-Mailbox" }
    elseif($objUser.RecipientTypeDetails -eq "MailUser*"){ [String]$cmdlet = "Disable-MailUser" }
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
    
    #Connect to Ex to perform action
    [Array]$cmdList = "Disable-MailUser","Disable-RemoteMailbox","Disable-mailbox"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
     
    try{
        [String]$command = "$cmdlet -Identity $objUser.UserPrincipalName -Force $OptionalDC -ErrorAction Stop"
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command" 
		Invoke-Expression -Command $command 
        
        Return "OK"    
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }        
}

Function fncOpSetMailFPe {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Set the calendar folder permissions
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }
        
    [Array]$cmdList = "Set-MailboxFolderPermission","Get-MailboxFolderStatistics"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif(($objUser.RecipientTypeDetails -eq "RoomMailbox") -or ($objUser.RecipientTypeDetails -eq "EquipmentMailbox")){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
        $Calendar = Get-MailboxFolderStatistics -Identity $Identity | Where{$_.foldertype -eq "Calendar"} 
		$Calendarfolder = $Identity + ":" + $Calendar.FolderID 

		[String]$command = "Set-MailboxFolderPermission -AccessRights LimitedDetails -User Default -Identity $calendarfolder $OptionalDC" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command 

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpAddAdmRole {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Member 
    )
<#
    Purpose: Adds an account to RoleGroup
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User,Add-RoleGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }

    try{           
        [String]$command = "Add-RoleGroupMember -Identity $Identity -Member $member -Force -bypassSecurityGroupManagerCheck" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command 

        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
        If ($errmsg -like "*already part of*") {
            # User is already part of that group
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Info: Success, because User $Member is already part of the Admin group." 

            return "OK"
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $errmsg."  
		    return "ERROR";
        }
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpRemAdmRole {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Member 
    )
<#
    Purpose: Removes an account from RoleGroup
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User,Remove-RoleGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }

    try{           
        [String]$command = "Remove-RoleGroupMember -Identity $Identity -Member $member -Force -bypassSecurityGroupManagerCheck" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command
        
        return "OK"
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
        If ($errmsg -like "*not part of*") {
            # User is already part of that group
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Info: Success, because User $Member is already not part of the Admin group." 

            return "OK"
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $errmsg."  
		    return "ERROR";
        }
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }

}

Function fncOpModEntitl {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$NewName 
    )
<#
    Purpose: Modifies the name of a Distribution Group
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User,Remove-RoleGroupMember"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{  
        $GUID = Get-DistributionGroup $Identity -ErrorAction Stop |Select-Object -ExpandProperty Guid | Select-Object -ExpandProperty Guid
        $Alias = $NewName -replace " ",""
        
        If($GUID -ne $null){                             
            [String]$command = "Set-DistributionGroup -Identity $GUID -Name $NewName -SamAccountName $Alias -Alias $Alias -Displayname $NewName" 
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
            Invoke-Expression -Command $command 
            
            return "OK"
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: Distribution Group $Identity Not found."  
		    return "ERROR"; 
        }
    }catch{        
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		return "ERROR";      
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpDelCheAcco {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Check if a mailbox has been disabled
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	              
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";              
    }
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Get-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Get-Mailbox" }    
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR - Action for recipienttype $($objUser.RecipientTypeDetails) not configured."
    }
    
    #Connect to Ex to perform action
    [Array]$cmdList = "Get-RemoteMailbox","Get-mailbox"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist

    try{        
        [String]$command = "$cmdlet -Identity $Identity -DomainController $DomainController -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        $mailbox = Invoke-Expression -Command $command 
        
        $mailboxisThere  = "true"
        $Mailboxtype = $Mailbox.recipienttypedetails
        $MailboxRetentionPolicy = $Mailbox.LitigationHoldEnabled
		$MailboxHiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
		$ExternalOOFOptions = $Mailbox.ExternalOOFOptions
		$PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
		$MRSAddress = $Mailbox|Select-Object -expandproperty emailaddresses|Where-Object {$_ -like "MRS*"} | select -First 1
		
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: $mailboxRetentionPolicy|$ExternalOOFOptions|$PrimarySmtpAddress|$MRSAddress|$mailboxisThere|$mailboxHiddenFromAddressListsEnabled|$Mailboxtype"       
        return "$mailboxRetentionPolicy|$ExternalOOFOptions|$PrimarySmtpAddress|$MRSAddress|$mailboxisThere|$mailboxHiddenFromAddressListsEnabled|$Mailboxtype";
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
        if ($errmsg -notlike "*couldn't be found on*" ){
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: Mailbox $Identity Not Found"  
            $OutError = "Success, Account Mailbox was found in Exchange"
			$Mailboxtype = "not Found in Exchange"
            return "ERROR - Mailbox: $Mailboxtype. Message: $OutError"	    
	    }else{  
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		    return "ERROR"; 
        }     
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function fncOpCheAcco {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Check if a mailbox has been enabled
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
	    if ($errmsg -notlike "*couldn't be found on*" ){
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: User $Identity Not Found"  
            $OutError = "Success, Account was not found in Active Directory"
			$Mailboxtype = "not Found in Exchange"
            return "OK - Mailbox: $Mailboxtype. Message: $OutError"	    
	    }else{            
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		    return "ERROR";
        }      
    }
     
    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ [String]$cmdlet = "Get-RemoteMailbox" }
    elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ [String]$cmdlet = "Get-Mailbox" }    
    else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR - Action for recipienttype $($objUser.RecipientTypeDetails) not configured."
    }
    
    #Connect to Ex to perform action
    [Array]$cmdList = "Get-RemoteMailbox","Get-mailbox"
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist

    try{
        
        [String]$command = "$cmdlet -Identity $Identity -DomainController $DomainController -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        $mailbox = Invoke-Expression -Command $command 
        
        $mailboxisThere  = "true"
        $Mailboxtype = $Mailbox.recipienttypedetails
        $MailboxRetentionPolicy = $Mailbox.LitigationHoldEnabled
		$MailboxHiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
		$ExternalOOFOptions = $Mailbox.ExternalOOFOptions
		$PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
		$MRSAddress = $Mailbox|Select-Object -expandproperty emailaddresses|Where-Object {$_ -like "MRS*"} | select -First 1
		
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: $mailboxRetentionPolicy|$ExternalOOFOptions|$PrimarySmtpAddress|$MRSAddress|$mailboxisThere|$mailboxHiddenFromAddressListsEnabled|$Mailboxtype"       
        return "$mailboxRetentionPolicy|$ExternalOOFOptions|$PrimarySmtpAddress|$MRSAddress|$mailboxisThere|$mailboxHiddenFromAddressListsEnabled|$Mailboxtype";
    }catch{
        $errmsg=[string]$_.Exception.Message	
	    $error.clear()
        if ($errmsg -notlike "*couldn't be found on*" ){
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Info: Mailbox $Identity Not Found"  
            $OutError = "Success, Account Mailbox was found in Exchange"
			$Mailboxtype = "not Found in Exchange"
            return "OK - Mailbox: $Mailboxtype. Message: $OutError"	    
	    }else{  
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		    return "ERROR"; 
        }     
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

function  fncOpAddLegHold{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$LitigationHoldEnabled,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: enables and set parameters for Litigation hold mailboxes
#>
    
    ##############
    ##############
    ##############
    #NO ACTION DONE FOR EXO MAILBOX. NEED TO CHECK
    ##############
    ##############
    ##############

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
           
    if(($objUser.RecipientTypeDetails -eq "UserMailbox") -or ($objUser.RecipientTypeDetails -eq "SharedMailbox")){
        #Connect to Ex 
        [Array]$cmdList = "Get-mailbox","Set-mailbox"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 
		
		try{
            [Boolean]$LitigationHoldEnabled = [System.Convert]::ToBoolean($LitigationHoldEnabled)                                                 

			if($LitigationHoldEnabled -eq $true){ # Enable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 1 -UseDatabaseQuotaDefaults 0 -IssueWarningQuota `"Unlimited`" -ProhibitSendReceiveQuota `"Unlimited`" -ProhibitSendQuota `"Unlimited`" -Force -ErrorAction Stop"
				fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command                 
			}else{ # Disable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 0 -Force -ErrorAction Stop"				
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command 
			}
                       
            return "OK"
		
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  			
			$Error.clear()
			return "ERROR";
		}
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."          
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}

function  fncOpRemLegHold{
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()]$LitigationHoldEnabled,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController                                           
    )
<#
    Purpose: Removes and set parameters for Litigation hold mailboxes
#>
    
    ##############
    ##############
    ##############
    #NO ACTION DONE FOR EXO MAILBOX. NEED TO CHECK
    ##############
    ##############
    ##############

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }
           
    if(($objUser.RecipientTypeDetails -eq "UserMailbox") -or ($objUser.RecipientTypeDetails -eq "SharedMailbox")){
        #Connect to Ex 
        [Array]$cmdList = "Get-mailbox","Set-mailbox"
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdList 
		
		try{
            [Boolean]$LitigationHoldEnabled = [System.Convert]::ToBoolean($LitigationHoldEnabled)                                                 

			if($LitigationHoldEnabled -eq $true){ # Enable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 1 -UseDatabaseQuotaDefaults 0 -IssueWarningQuota `"Unlimited`" -ProhibitSendReceiveQuota `"Unlimited`" -ProhibitSendQuota `"Unlimited`" -Force -ErrorAction Stop"
				fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command                 
			}else{ # Disable Litigation Hold
                [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -DomainController $DomainController -LitigationHoldEnabled 0 -Force -ErrorAction Stop"				
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
                Invoke-Expression -Command $command 
			}
                       
            return "OK"
		
        }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  			
			$Error.clear()
			return "ERROR";
		}
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."          
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }
}

function  fncOpSetResourceCapacity{
    param(        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,         
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $ResourceCustom,         
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $ResourceCapacity                                                  
    )
<#
    Purpose: Set parameters on Room Mailboxes
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [Array]$cmdList = "Set-Mailbox"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif(($objUser.RecipientTypeDetails -eq "RoomMailbox") -or ($objUser.RecipientTypeDetails -eq "EquipmentMailbox")){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
        if($ResourceCustom -ne "" -and $ResourceCustom -ne $null){
            $ResourceCustom = $ResourceCustom -replace '[()]', ""
	        $ResourceCustom = $ResourceCustom -replace "Flatscreen TV","FlatscreenTV"

            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -ResourceCustom $ResourceCustom $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }
    
        if($ResourceCapacity -ne "" -and $ResourceCapacity -ne $null){
            [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -ResourceCapacity $ResourceCapacity $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }
        
        return "OK"
    }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  			
			$Error.clear()
			return "ERROR";
		}
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
}

function  fncOpSetBasicCalProc {
    param(        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $Identity,         
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String] $DomainController,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $AllowConflicts,         
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $AllowRecurringMeetings,
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $ScheduleOnlyDuringWorkHours,       
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $EnforceSchedulingHorizon,       
        [Parameter(Mandatory=$false,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()] $ForwardRequestsToDelegates                                                 
    )
<#
    Purpose: Set parameters on Room Mailboxes
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{        
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
        return "ERROR"        
    }

    [Array]$cmdList = "Set-CalendarProcessing"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif(($objUser.RecipientTypeDetails -eq "RoomMailbox") -or ($objUser.RecipientTypeDetails -eq "EquipmentMailbox")){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
        if($AllowConflicts -ne "" -and $AllowConflicts -ne $null){
            [Boolean]$AllowConflicts = [System.Convert]::ToBoolean($AllowConflicts)
            $AllowConflicts = [Int]$AllowConflicts

            [String]$command = "Set-CalendarProcessing -Identity $objUser.UserPrincipalName -AllowConflicts $AllowConflicts $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }
    
        if($AllowRecurringMeetings -ne "" -and $AllowRecurringMeetings -ne $null){
            [Boolean]$AllowRecurringMeetings = [System.Convert]::ToBoolean($AllowRecurringMeetings)
            $AllowRecurringMeetings = [Int]$AllowRecurringMeetings

            [String]$command = "Set-CalendarProcessing -Identity $objUser.UserPrincipalName -AllowRecurringMeetings $AllowRecurringMeetings $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }

        if($ScheduleOnlyDuringWorkHours -ne "" -and $ScheduleOnlyDuringWorkHours -ne $null){
            [Boolean]$ScheduleOnlyDuringWorkHours = [System.Convert]::ToBoolean($ScheduleOnlyDuringWorkHours)
            $ScheduleOnlyDuringWorkHours = [Int]$ScheduleOnlyDuringWorkHours

            [String]$command = "Set-CalendarProcessing -Identity $objUser.UserPrincipalName -ScheduleOnlyDuringWorkHours $ScheduleOnlyDuringWorkHours $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }
        
        if($EnforceSchedulingHorizon -ne "" -and $EnforceSchedulingHorizon -ne $null){
            [Boolean]$EnforceSchedulingHorizon = [System.Convert]::ToBoolean($EnforceSchedulingHorizon)
            $EnforceSchedulingHorizon = [Int]$EnforceSchedulingHorizon

            [String]$command = "Set-CalendarProcessing -Identity $objUser.UserPrincipalName -EnforceSchedulingHorizon $EnforceSchedulingHorizon $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }

        if($ForwardRequestsToDelegates -ne "" -and $ForwardRequestsToDelegates -ne $null){
            [Boolean]$ForwardRequestsToDelegates = [System.Convert]::ToBoolean($ForwardRequestsToDelegates)
            $ForwardRequestsToDelegates = [Int]$ForwardRequestsToDelegates

            [String]$command = "Set-CalendarProcessing -Identity $objUser.UserPrincipalName -ForwardRequestsToDelegates $ForwardRequestsToDelegates $OptionalDC -Force -ErrorAction Stop"				
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"  
            Invoke-Expression -Command $command 
        }

        return "OK"
    }catch{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  			
			$Error.clear()
			return "ERROR";
		}
        finally {            
            Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        }
}

Function fncOpSetTimeZone {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$WorkingHoursTimeZone,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Set the timeZone for a Mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }
        
    [String]$cmdList = "Set-MailboxCalendarConfiguration"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
		[String]$command = "Set-MailboxCalendarConfiguration -Identity $objUser.UserPrincipalName -WorkingHoursTimeZone $WorkingHoursTimeZone $OptionalDC -Force -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command 

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpSetMBCalPro {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$WorkingHoursStartTime,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$WorkingHoursEndTime,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Set time attributes for a Mailbox
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }
        
    [String]$cmdList = "Set-MailboxCalendarConfiguration"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
		[String]$command = "Set-MailboxCalendarConfiguration -Identity $objUser.UserPrincipalName -WorkingHoursStartTime $WorkingHoursStartTime -WorkingHoursEndTime $WorkingHoursEndTime $OptionalDC -Force -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command 

        return "OK"
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }
    finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpSETSMRight {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$OrderParamWrite,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$OrderParamRead,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$OrderParamSendAs,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController 
    )

<#
    Purpose: Set permissions for service mailboxes
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "Get-User"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{
        $objUser = Get-User -Identity $Identity -DomainController $DomainController -ErrorAction Stop | Select-Object RecipientTypeDetails,UserPrincipalName        
     }catch{        
		fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg"  
		return "ERROR";	         
    }
        
    [Array]$cmdList = "Add-MailboxPermission","Get-MailboxFolderStatistics","Add-MailboxFolderPermission","Set-Mailbox","get-distributiongroup"

    #Check location of mailbox 
    if($objUser.RecipientTypeDetails -like "Remote*"){ 
        [String]$OptionalDC = ""

        #Connect to ExO to perform action 
        $ExOSession = Connect-EXOnline -CredPath $credPath -o365opt $o365opt -cmdList $cmdlist
    }elseif($objUser.RecipientTypeDetails -like "*Mailbox"){ 
        [String]$OptionalDC = "-DomainController $DomainController" 
        
        #Connect to Ex to perform action            
        $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList $cmdlist
    }else{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: the account is $($objUser.RecipientTypeDetails) and this operation is not configured."         
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
        return "ERROR"
    }

    try{
        #Retrieve UPN for onPremise object
        $userUPN = "$($objUser.UserPrincipalName)"
		$OrderParamWrite = get-distributiongroup $OrderParamWrite | select -expandproperty PrimarySmtpAddress
		$OrderParamRead = get-distributiongroup $OrderParamRead | select -expandproperty PrimarySmtpAddress
		$OrderParamSendAs = get-distributiongroup $OrderParamSendAs | select -expandproperty PrimarySmtpAddress
    }catch{
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception."  
		$Error.clear()
		return "ERROR";
    }

    try{ #Assign FullAccess Permissions
        [String]$command = "Add-MailboxPermission -Identity $objUser.UserPrincipalName -User $OrderParamWrite -AccessRights FullAccess -AutoMapping 0 $OptionalDC -Force -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command         
    }catch{
        $errmsg = [string]$_.Exception.Message
		$error.clear()          
        If ($errmsg -notlike "*existing permission entry was found*") {      
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg."           
            return "ERROR";
        }
    }
    
    try{ #Grant Send Of Behalf
       [String]$command = "Set-Mailbox -Identity $objUser.UserPrincipalName -GrantSendOnBehalfTo $OrderParamSendAs $OptionalDC -Force -ErrorAction Stop" 
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
        Invoke-Expression -Command $command         
    }catch{
        $errmsg = [string]$_.Exception.Message
		$error.clear()          
        If ($errmsg -notlike "*existing permission entry was found*") {      
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg."           
            return "ERROR";
        }
    }  
	
    try{ #Assign Reviewer permissions on Folder
       ForEach($folder in (Get-MailboxFolderStatistics $userUPN | Where{$_.FolderPath.ToLower().StartsWith("/") -eq $True})){
            $foldername = $userUPN + ":" + $folder.FolderID; 
            $foldright = $null
            $foldright = Get-MailboxFolderPermission -Identity $foldername -User $OrderParamRead -ErrorAction SilentlyContinue
            if(!$foldright){             
                [String]$command = "Add-MailboxFolderPermission $foldername -User $OrderParamRead -AccessRights Reviewer $OptionalDC | out-null" 
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
                Invoke-Expression -Command $command 
            }
        }        
    }catch{
        $errmsg = [string]$_.Exception.Message
		$error.clear()          
        If ($errmsg -notlike "*existing permission entry was found*") {      
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg."           
            return "ERROR";
        }
    } 	
              
    return "OK"              
    Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue    
}

Function fncOpAddEntitle {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$OrganizationalUnit,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DisplayName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][AllowNull()][AllowEmptyString()][String]$Type,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController
    )
<#
    Purpose: Creates the distribution group related to an entitlement
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "get-DistributionGroup,New-DistributionGroup"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{  
        $GUID = $null
        $GUID = Get-DistributionGroup $Identity -ErrorAction silentlycontinue |Select-Object -ExpandProperty Guid | Select-Object -ExpandProperty Guid
        
        If($GUID -eq $null){                             
            if($Type -eq $null -or $Type -eq ""){                
                [String]$command =  "New-DistributionGroup -Name $Identity -SamAccountName $Identity -OrganizationalUnit $OrganizationalUnit -Displayname $DisplayName -DomainController $DomainController  -ErrorAction Stop -Force"
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
                Invoke-Expression -Command $command 
            
                return "OK"
            }else{
                [String]$command =  "New-DistributionGroup -Name $Identity -SamAccountName $Identity -Type $Type -OrganizationalUnit $OrganizationalUnit -Displayname $DisplayName -DomainController $DomainController  -ErrorAction Stop -Force"
                fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
                Invoke-Expression -Command $command 
            
                return "OK"
            }
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Info: Distribution Group $Identity already exists."  
		    return "OK"; 
        }
    }catch{   
        $errmsg = [string]$_.Exception.Message
		$error.clear()          
        If ($errmsg -notlike "*already exists*") {      
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $errmsg."           
            return "ERROR";
        }               
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpRenEntitle {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DGNewName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController
    )
<#
    Purpose: Modifies the distribution group related to an entitlement
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "get-DistributionGroup,Set-DistributionGroup"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{  
        $GUID = $null
        $Alias = $DGNewName -replace " ",""
        $GUID = Get-DistributionGroup $Identity -ErrorAction silentlycontinue |Select-Object -ExpandProperty Guid | Select-Object -ExpandProperty Guid
        
        If($GUID -ne $null){                                         
            [String]$command =  "Set-DistributionGroup -Identity $GUID -SamAccountName $DGNewName -Displayname $DGNewName -Alias $Alias -DomainController $DomainController  -ErrorAction Stop -Force"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
            Invoke-Expression -Command $command 
            
            return "OK"
            
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Error: Distribution Group $Identity does not exists."  
		    return "ERROR"; 
        }
    }catch{   		         
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $_.Exception."   
        $error.clear()         
        return "ERROR";
                       
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

Function fncOpRenEntitle {
    param(
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$Identity,                                                        
        [Parameter(Mandatory=$true,ParameterSetName='explicit')][ValidateNotNullOrEmpty()][String]$DomainController
    )
<#
    Purpose: Modifies the distribution group related to an entitlement
#>

    #Connect to Ex to perform action
    $ExSession = Connect-EXOnpremise -CredPath $credPath -exchangeServer $exchangeServer -CmdList "get-DistributionGroup,Remove-DistributionGroup"
    Set-ADServerSettings -ViewEntireForest 1 -WA 0 

    try{  
        $GUID = $null
        $GUID = Get-DistributionGroup $Identity -ErrorAction silentlycontinue |Select-Object -ExpandProperty Guid | Select-Object -ExpandProperty Guid
        
        If($GUID -ne $null){                                         
            [String]$command =  "Remove-DistributionGroup -Identity $GUID -DomainController $DomainController -ErrorAction Stop -Force"
            fncWriteToLogFile -LogFile $VerboseLogfile -message "Command: $command"       
            Invoke-Expression -Command $command             
            return "OK"
            
        }else{
            fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Info: Distribution Group $Identity does not exists."  
		    return "OK"; 
        }
    }catch{   		         
        fncWriteToLogFile -LogFile $VerboseLogfile -message "Critical: $_.Exception."   
        $error.clear()         
        return "ERROR";
                       
    }finally {            
        Get-PSSession -ErrorAction silentlycontinue | Remove-PSSession -ErrorAction silentlycontinue
    }
}

#===============================================================================	
#===============================================================================
# Main
#===============================================================================
#===============================================================================

Start-Transcript -Path $VerboseLogfile -Append
$Global:ErrorActionPreference = 'Stop'
$StopWatch = [System.Diagnostics.StopWatch]::StartNew()

fncWriteToLogFile -LogFile $VerboseLogfile -message "<PRECOMMAND>"
fncWriteToLogFile -LogFile $VerboseLogfile -message $Jsonfile
fncWriteToLogFile -LogFile $VerboseLogfile -message "</PRECOMMAND>"

try{
    $JsonObject = ConvertFrom-Json $Jsonfile
}catch{
    fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: Reading JSon file." 
    fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Critical: $_.Exception." 
    return "ERROR"
    Stop-Transcript
    exit
}

#Validating if Json Object contains valid data
$JsonAttributes = Get-Member -InputObject $JsonObject.attributes -MemberType NoteProperty
if($JsonAttributes.count -lt 1){
    fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: JSon object contains less than 1 attribute."      
    return "ERROR"
    Stop-Transcript
    exit
}

switch($JsonObject.operation){
    "AddAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddAccount function."  
        $result = fncOpAddAccount -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Alias $JsonObject.attributes.alias -Database $JsonObject.attributes.Database -MailboxType $JsonObject.attributes.MailboxType -PrimarySmtpAddress $JsonObject.attributes.PrimarySmtpAddress -ManagedFolderMailboxPolicy $JsonObject.attributes.ManagedFolderMailboxPolicy
        break;
    }
    "FixLegacyDN" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing FixLegacyDN function."        
        $result = fncOpFixLegacyDN -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController
        break;
    }
    "AddExOSMTP" {  
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddExOSMTP function."          
        $result = fncOpAddExOSMTP -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController       
        break;
    }
    "AddAccountLegalHold" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddAccountLegalHold function."         
        $result = fncOpAddAccountLegalHold -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -LitigationHoldEnabled $JsonObject.attributes.LitigationHoldEnabled
        break;
    }
    "SetSpecSetsDefault" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetSpecSetsDefault function."        
        $result = funcopSetSpecSetsDefault -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -OwaMailboxPolicy $JsonObject.attributes.OwaMailboxPolicy
        break;
    }
    "HidAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing HidAccount function."        
        $result = funcopHidAccount -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -HiddenFromAddressListsEnabled $JsonObject.attributes.HiddenFromAddressListsEnabled
        break;
    }
    "AddSecMail" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing HidAccount function."        
        $result = fncOpAddSecMail -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -SecondaryMail $JsonObject.attributes.SecondaryMail
        break;
    }
    "RemSecMail" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing HidAccount function."        
        $result = fncOpRemSecMail -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -SecondaryMail $JsonObject.attributes.SecondaryMail
        break;
    }
    "SetExtOoof" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetExtOoof function."        
        $result = fncOpSetExtOoof -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -ExternalOofOptions $JsonObject.attributes.ExternalOofOptions
        break;    
    }
    "SetSpecSets" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetSpecSets function."   
        $result = fncOpSetSpecSets -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -UseDatabaseRetentionDefaults $JsonObject.attributes.UseDatabaseRetentionDefaults -SingleItemRecoveryEnabled $JsonObject.attributes.SingleItemRecoveryEnabled -RetentionPolicy $JsonObject.attributes.RetentionPolicy -RetainDeletedItemsUntilBackup $JsonObject.attributes.RetainDeletedItemsUntilBackup -RetainDeletedItemsFor $JsonObject.attributes.RetainDeletedItemsFor
        break;
    }
    "SetPrimMail" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetPrimMail function."   
        $result = fncOpSetPrimMail -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -PrimarySmtpAddress $JsonObject.attributes.PrimarySmtpAddress -EmailAddressPolicyEnabled $JsonObject.attributes.EmailAddressPolicyEnabled
        break;
    }
    "AddDiLRole" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddDiLRole function."   
        $result = fncOpAddDiLRole -Identity $JsonObject.attributes.nativeIdentity -DistributionGroup $JsonObject.attributes.DistributionGroup 
        break;
    }
    "RemDiLRole" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddDiLRole function."   
        $result = fncOpRemDiLRole -Identity $JsonObject.attributes.nativeIdentity -DistributionGroup $JsonObject.attributes.DistributionGroup 
        break;
    }
    "SetCalProc" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetCalProc function."   
        $result = fncOpSetCalProc  -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Notes $JsonObject.attributes.Notes -ResourceDelegates $JsonObject.attributes.ResourceDelegates -ForwardRequestsToDelegates $JsonObject.attributes.ForwardRequestsToDelegates -AutomateProcessing $JsonObject.attributes.AutomateProcessing  -BookInPolicy $JsonObject.attributes.BookInPolicy -DeleteAttachments $JsonObject.attributes.DeleteAttachments -RemovePrivateProperty $JsonObject.attributes.RemovePrivateProperty -DeleteSubject $JsonObject.attributes.DeleteSubject -DeleteNonCalendarItems $JsonObject.attributes.DeleteNonCalendarItems -AllBookInPolicy $JsonObject.attributes.AllBookInPolicy -AllRequestInPolicy $JsonObject.attributes.AllRequestInPolicy -AddAdditionalResponse $JsonObject.attributes.AddAdditionalResponse -AdditionalResponse $JsonObject.attributes.AdditionalResponse -DeleteComments $JsonObject.attributes.DeleteComments
        break;
    }
    "SynAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SynAccount function. No action required."           
        break;
    }
    "revokeEUUM" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing revokeEUUM function."   
        $result = fncOpRevokeEUUM -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController                
        break;
    }
    "assignEUUM" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing assignEUUM function."   
        $result = fncOpAssignEUUM -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -UMMailboxPolicy $JsonObject.attributes.UMMailboxPolicy -Extensions $JsonObject.attributes.Extensions -SIPResourceIdentifier $JsonObject.attributes.SIPResourceIdentifier        
        break;
    } 
    "setAlias" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing setAlias function."   
        $result = fncOpSetAlias -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Alias $JsonObject.attributes.Alias
        break;
    }
    "SetUser" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing setAlias function."   
        $result = fncOpSetUser -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Notes $JsonObject.attributes.Notes
        break;
    }
    "AddRoomAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddRoomAccount function."   
        $result = fncOpAddRoomAccount -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Alias $JsonObject.attributes.alias -Database $JsonObject.attributes.Database -MailboxType $JsonObject.attributes.MailboxType -PrimarySmtpAddress $JsonObject.attributes.PrimarySmtpAddress -ManagedFolderMailboxPolicy $JsonObject.attributes.ManagedFolderMailboxPolicy
        break;
    }
    "AddMailEnabledAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddMailEnabledAccount function."   
        $result = fncOpAddMailEnabledAccount -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -Alias $JsonObject.attributes.Alias -ExternalEmailAddress $JsonObject.attributes.ExternalEmailAddress
        break;
    }
    "AddManagBy" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddManagBy function."   
        $result = fncOpAddManagBy -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -DistributionGroup $JsonObject.attributes.DistributionGroup
        break;
    }
    "RemManagBy" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing RemManagBy function."   
        $result = fncOpRemManagBy -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -DistributionGroup $JsonObject.attributes.DistributionGroup
        break;
    }
    "RemAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing RemAccount function."   
        $result = fncOpRemAccount -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController
        break;
    }
    "SetMailFPe" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetMailFPe function."   
        $result = fncOpSetMailFPe -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController
        break;
    }
    "AddAdmRole" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddAdmRole function."   
        $result = fncOpAddAdmRole -Identity $JsonObject.attributes.nativeIdentity -Member $JsonObject.attributes.Member
        break;
    }
    "RemAdmRole" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing RemAdmRole function."   
        $result = fncOpRemAdmRole -Identity $JsonObject.attributes.nativeIdentity -Member $JsonObject.attributes.Member
        break;
    }
    "ModEntitl" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing ModEntitl function."   
        $result = fncOpModEntitl -Identity $JsonObject.attributes.DGName -NewName $JsonObject.attributes.DGNewName
        break;
    }
    "DelCheAcco" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing DelCheAcco function."   
        $result = fncOpDelCheAcco -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController
        break;
    }
    "CheAccount" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing CheAccount function."   
        $result = fncOpCheAcco -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController
        break;   
    }
    "AddLegHold" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddLegHold function."   
        $result = fncOpAddLegHold -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -LitigationHoldEnabled $JsonObject.attributes.LitigationHoldEnabled
        break;   
    }
    "RemLegHold" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing RemLegHold function."   
        $result = fncOpRemLegHold -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -LitigationHoldEnabled $JsonObject.attributes.LitigationHoldEnabled
        break;   
    }
    "SetResourceCapacity" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetResourceCapacity function."   
        $result = fncOpSetResourceCapacity -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -ResourceCustom $JsonObject.attributes.ResourceCustom -ResourceCapacity $JsonObject.attributes.ResourceCapacity
        break;   
    }
    "SetBasicCalProc" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetBasicCalProc function."   
        $result = fncOpSetBasicCalProc -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -AllowConflicts $JsonObject.attributes.AllowConflicts -AllowRecurringMeetings $JsonObject.attributes.AllowRecurringMeetings -ScheduleOnlyDuringWorkHours $JsonObject.attributes.ScheduleOnlyDuringWorkHours -EnforceSchedulingHorizon $JsonObject.attributes.EnforceSchedulingHorizon -ForwardRequestsToDelegates $JsonObject.attributes.ForwardRequestsToDelegates
        break;   
    }
    "SetTimeZone" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetTimeZone function."   
        $result = fncOpSetTimeZone -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -WorkingHoursTimeZone $JsonObject.attributes.WorkingHoursTimeZone
        break; 
    }
    "SetMBCalPro" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SetMBCalPro function."   
        $result = fncOpSetMBCalPro -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -WorkingHoursStartTime $JsonObject.attributes.WorkingHoursStartTime -WorkingHoursEndTime $JsonObject.attributes.WorkingHoursEndTime
        break; 
    }    
    "SETSMRight" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing SETSMRight function."   
        $result = fncOpSETSMRight -Identity $JsonObject.attributes.nativeIdentity -DomainController $JsonObject.attributes.domainController -OrderParamWrite $JsonObject.attributes.OrderParamWrite -OrderParamRead $JsonObject.attributes.OrderParamRead -OrderParamSendAs $JsonObject.attributes.OrderParamSendAs
        break; 
    }
    "AddEntitle" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing AddEntitle function."   
        $result = fncOpAddEntitle -Identity $JsonObject.attributes.Name -DomainController $JsonObject.attributes.domainController -OrganizationalUnit $JsonObject.attributes.OrganizationalUnit -Type $JsonObject.attributes.Type -DisplayName $JsonObject.attributes.DisplayName
        break;
    }
    "RenEntitle" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing RenEntitle function."   
        $result = fncOpRenEntitle -Identity $JsonObject.attributes.Name -DomainController $JsonObject.attributes.domainController -DGNewName $JsonObject.attributes.DGNewName
        break;
    }
    "DelEntitle" {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "$(Get-TimeStamp) Processing DelEntitle function."   
        $result = fncOpRenEntitle -Identity $JsonObject.attributes.Name -DomainController $JsonObject.attributes.domainController
        break;
    }

    # Default action when operation is not found
    default {
        fncWriteToLogFile -LogFile $VerboseLogfile -message "ERROR: Operation $($JsonObject.operation) set in Json file is not supported."        
        $result = "ERROR"
    }
}



$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
fncWriteToLogFile -LogFile $VerboseLogfile -message  $msg
$msg = $null
$StopWatch = $null
return $result
Stop-Transcript
