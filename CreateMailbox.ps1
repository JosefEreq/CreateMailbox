

<#

This script creates mailbox, sets license/services, audit and more in Exchange hybrid environments.
Controlled by the 4 CSV config files(License, Audit, SKUIds and settings).


Author 
Josef Ereq

Version 1.2




#> 

# - Import the AD Powershell-module.
import-module activedirectory

# - Set the path to the config file.
$ConfigFile = ".\CreateMailboxInput.csv"

# - Import the data from the config file.
$Config = import-csv $ConfigFile -Delimiter ";" -Encoding utf8


# - Set the credentials for connecting to Exchange, AADConnect server and AzureAD.

$EXcred = Import-CliXml -Path "C:\script\cred\ServiceAccount_EXO.cred"

$AADcred = Import-CliXml -Path "C:\script\cred\ServiceAccount_AAD.cred"

# - Connect to the AzureAD with the service account credentials.
Connect-MsolService -Credential $AADCred

# - Import the domain controller to use, from the config file.
$DomainController = ($Config | where {$_.name -eq "DomainController"}).value
# - Import the path to the logging folder, from the config file.
$LogPath = ($Config | where {$_.name -eq "LogPath"}).value
# - Import the path to the input-data file, from the config file, and replace any frontslashes with backslashes.
$ScriptDataFile = ($Config | where {$_.name -eq "ScriptDataFile"}).value
$ScriptDataFile = $ScriptDataFile -replace "\057","\"
# - Import the path to the work-data file, from the config file, and replace any frontslashes with backslashes.
$ScriptDataFileStart = ($Config | where {$_.name -eq "ScriptWorkFile"}).value
$ScriptDataFileStart = $ScriptDataFileStart -replace "\057","\"
# - Import the path to the audit-properties file, from the config file.
$AuditData = import-csv ($Config | where {$_.name -eq "AuditDataFile"}).value -Encoding utf8
# - Import the list licenses, from the config file.
$LicensSkuIDs = import-csv ($Config | where {$_.name -eq "LicensSkuIDsFile"}).value -Encoding utf8 -Delimiter ";"
# - Import the list of disabled services, from the config file.
$DisabledServices = import-csv ($Config | where {$_.name -eq "DisabledServicesFile"}).value -Encoding utf8 -Delimiter ";"
# - Import the groups to add each user in, from the config file.
$GroupsToAdd = ($Config | where {$_.name -eq "MailboxGroupMembershipAdd"}).value -split ","
# - Import the URL to the Exchange Admin for the on-premises server, from the config file.
$ExchangeServerURL = ($Config | where {$_.name -eq "ExchangeServerURL"}).value
# - Import the URL to the Exchange Online admin enviornment, from the config file.
$ExchangeOnlineURL = ($Config | where {$_.name -eq "ExchangeOnlineURL"}).value
# - Import the FQDN for the AzureAD Connect-server, from the config file.
$AADConnectHost = ($Config | where {$_.name -eq "AADConnectHost"}).value
# - Get the audit-properties from the input-file.
$AuditLogAgeLimit = ($AuditData | where {$_.property -eq "AuditLogAgeLimit"}).value
$AuditOwner = ($AuditData | where {$_.property -eq "AuditOwner"}).value -join ", "
$AuditDelegate = ($AuditData | where {$_.property -eq "AuditDelegate"}).value  -join ", "
$AuditAdmin = ($AuditData | where {$_.property -eq "AuditAdmin"}).value -join ", "
# - Get the attribute name for distinguishing user types.
$AttrUsrType = ($Config | where {$_.name -eq "AttributeUserType"}).value

# - Create sessions to the on-premises Exchange and Exchange Online.
$EXPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServerURL -Credential $EXcred -Authentication Kerberos
$EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeOnlineURL -Credential $EXcred -Authentication Basic -AllowRedirection

# - Test if the input-data file contains entries for mailbox creation.
$FileExist = Test-Path $ScriptDataFile

# - Create a function that triggers a AzureAD directory syncronization.
Function RunAADSync 
    {
    Param(
    [Parameter(Mandatory=$true,Position=0)]
    $HostSession)

    # - Check if a AzureAD-sync is running. If so, wait some time and check if it's still running, until it's done.
    $Syncjob = Invoke-Command -Session $HostSession -ScriptBlock {(Get-ADSyncScheduler -WarningAction Continue -ErrorAction Continue).SYNCCYCLEINPROGRESS}            
    while ($Syncjob -eq $true)
            {
            sleep -Seconds 10
            $Syncjob = Invoke-Command -Session $HostSession -ScriptBlock {(Get-ADSyncScheduler -WarningAction Continue -ErrorAction Continue).SYNCCYCLEINPROGRESS}
            } 
    # - Start a new sync.
    Invoke-Command -Session $HostSession -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta -WarningAction Continue -ErrorAction Continue}
    sleep -Seconds 15
    # - Check if a AzureAD-sync is running. If so, wait some time and check if it's still running, until it's done
    $Syncjob = Invoke-Command -Session $HostSession -ScriptBlock {(Get-ADSyncScheduler -WarningAction Continue -ErrorAction Continue).SYNCCYCLEINPROGRESS}            
    while ($Syncjob -eq $true)
            {
            sleep -Seconds 10
            $Syncjob = Invoke-Command -Session $HostSession -ScriptBlock {(Get-ADSyncScheduler -WarningAction Continue -ErrorAction Continue).SYNCCYCLEINPROGRESS}
            }  
    }

# - If workfile exists, run the script block.
if ($FileExist -eq "true")
    {
    # - Rename the input-data file to the file used when script is processing entries.
    rename-item $ScriptDataFile $ScriptDataFileStart -Force
    # - Get the content of the input-data file.
    $Entries = Get-Content $ScriptDataFileStart
    # - Loop trough each row in the input-data file.
    foreach ($Entry in $Entries)
        {
        # - Clear user properties to make sure they are not saved from previous loop.
        $name = ""
        $UsageLocation = ""
        $UPN = ""
        $aduser = $null

        # - Get the samaccountname part from the entriy.
		$name = $Entry.Substring(0, $Entry.indexof(";"))
        # - Get the usagelocation part from the entry.
        $UsageLocation = $Entry.Substring($Entry.indexof(";") + 1)
        # - Load the aduser into a variable.
	    $aduser = get-aduser $name -server $DomainController -Property $AttrUsrType
        # - Get the UPN for the user.
        $UPN = $ADUser.UserPrincipalName     
          
        # - Invoke a command to create a remote mailbox on the on-premises Exchange server, with the specified parameters.
	    Invoke-Command -Session $EXPSession -ArgumentList $name,$upn -ScriptBlock {Enable-RemoteMailbox -Identity $args[0]-RemoteRoutingAddress "$($args[0])@OFFICE365-domain.mail.onmicrosoft.com" -PrimarySmtpAddress $args[1]}
       	    
        # - Clear the variable user for storing the remote mailbox from any previous loop.        
        $RemoteMB = $null
        # - Load the remote mailbox into a variable.
        $RemoteMB = Invoke-Command -Session $EXPSession -ArgumentList $name -ScriptBlock {get-RemoteMailbox $args[0]}  
        # - Create a loop that loads the remote mailbox into a variable, until it is found.      
        $retryn = 0
        do
            {
            sleep -Seconds 10
            $retryn++
            $RemoteMB = Invoke-Command -Session $EXPSession -ArgumentList $name -ScriptBlock {get-RemoteMailbox $args[0]} 
            }
            while((!$RemoteMB) -and ($retryn -lt 20))
        # - Enable email address policy on the remote-mailbox.
        Invoke-Command -Session $EXPSession -ArgumentList $name,$true -ScriptBlock {Set-RemoteMailbox -Identity $args[0] -EmailAddressPolicyEnabled $args[1]}
        # - Loop trough each of the group to add this user to, and add the user to it.
        foreach($grp in $GroupsToAdd)
            {
            Add-ADGroupMember -server $DomainController -Identity $grp -Members $name
            }

        # - Trigger AzureAD directory synk function.
        $AADCSession = New-PSSession -ComputerName $AADConnectHost -Credential $AADcred -name AADC
        Invoke-Command -Session $AADCSession -ScriptBlock {Import-Module -Name 'ADSync'}
        RunAADSync $AADCSession
        Remove-PSSession $AADCSession

        # - Clear the variable used for storing the AzureAD user from any previous loops.
        $MsolUsr = $null
        # - Create a loop that loads the Azuread user into a variable, until it is found. 
        $retryn = 0
        do
            {     
            sleep -Seconds 10
            $retryn++       
            $MsolUsr = Get-msoluser -UserPrincipalName $upn
            }
            while((!$MsolUsr) -and ($retryn -lt 20))
        # - Set the usage location for the AzureAD user.
	    Set-MsolUser -UserPrincipalName $UPN -UsageLocation $UsageLocation
        
        # - Set the user type in a variable. This will be used for fetching the correct licenses. 
        $userType = $null
        $UserType = $aduser.($AttrUsrType)
        
        # - Set the license order to check, to it's initial value of 1.
        $LicOrder = 1

        # - Set the variable for signaling license assignment failure to its default value FALSE.
        $LicenseCheck = $false

        # - Create a loop that assigns the assign the user licenses in the order of availability, that runs as long as any of the license assignments fail.
        do
            {
            # - Get the applicable license for the user, from the SKUID-array, based on user attributes and license order.
            $UserLicenses = $Null
            $UserLicenses = ($LicensSkuIDs | where {($_.UserType -eq $UserType) -and ($_.order -eq $LicOrder)}).license   
            
            # - Set the variable used for signaling full slots of license, to its intial value of FALSE.
            $SKUFull = $false
                                        
            # - Loop trouch each of licenses for the user, and check if each one has a free slot.
            foreach ($SkuID in $UserLicenses)
                {          

                # - Get the license-SKU from Office365.
                $SKU = $null
                $SKU = Get-MsolAccountSku | where {$_.accountskuid -eq $SkuID}
                # - Check if the license has a available slot. It it doesnt, set the variable for signaling full license slots to TRUE.
                if(!($sku.ConsumedUnits -lt $sku.activeunits))
                    {
                    $SKUFull = $true
                    }    
                }
            # - If the variable for signaling full licenses is not TRUE, run the script block that sets the variable for signaling license assignment to TRUE.
            if($SKUFull -ne $true)
                {
                $LicenseCheck = $true
                }
            # - Check if the variable used for signaling license assignment is TRUE, if so, run the script block that assigns the licenses to the user.
            if($LicenseCheck -eq $true)
                {
                # - Loop trouch each of licenses and assign them to the user.
                foreach ($SkuID in $UserLicenses)
                    {                    
                    Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $SkuID -WarningAction Continue -ErrorAction Continue  
                    
                    # - Check if the license to assign have a service to disable.
                    $ServicesForDisable = $null
                    $ServicesForDisable = ($DisabledServices | where {$_.license -match $SkuID}).DisabledService

                    # - Set the license options variable to null, to clear data from any previous loops.         
                    $LicOptions = $null

                    # - If the variable for service to disable exist, run the script block.
                    if($ServicesForDisable)
                        {
                        # - Create a service plan for the license to assign, with the undesired services disabled.
                        $LicOptions = New-MsolLicenseOptions -AccountSkuId $SkuID -DisabledPlans $ServicesForDisable
                        # - Assign the service plan on the user.
                        Set-MsolUserLicense -UserPrincipalName $UPN -LicenseOptions $LicOptions
                        }   
                    }    
                }
            # - If the variable used for signaling license assignment is FALSE, run the script block that increase the number for what license order to test.
            else
                {
                $LicOrder++
                }       

            }while(($LicenseCheck -eq $false) -and ($LicOrder -lt 20))



        # - Trigger AzureAD directory synk function.
        $AADCSession = New-PSSession -ComputerName $AADConnectHost -Credential $AADcred -name AADC
        Invoke-Command -Session $AADCSession -ScriptBlock {Import-Module -Name 'ADSync'}
        RunAADSync $AADCSession
        Remove-PSSession $AADCSession

        # - Load the aduser to the varible once again, now that his/hers mail attribute has been populated.
        $aduser = get-aduser $name -server $DomainController -Properties mail
        # - Clear the variable used for storing the cloud mailbox from any previous loops.
        $CloudMB = $null
        # - Create a loop that loads the cloud mailbox into a variable, until it is found.
        $retryn = 0 
        do
            {
            sleep -Seconds 10
            $retryn++
            $CloudMB = Invoke-Command -Session $EXOSession -ArgumentList $ADUser.mail -ScriptBlock {get-mailbox -identity $args[0]}  
            }
            while((!$CloudMB) -and ($retryn -lt 20))

        # - Set audit properties on the mailbox
        Invoke-Command -Session $EXOSession -ArgumentList $($ADUser.mail),$true,$AuditLogAgeLimit,$AuditOwner,$AuditDelegate,$AuditAdmin -ScriptBlock {set-mailbox -Identity $args[0] -auditenabled $args[1] -AuditLogAgeLimit $args[2] -AuditOwner ($args[3]) -AuditDelegate ($args[4]) -AuditAdmin ($args[5])}
        # - Enable Litigation Hold on the mailbox
        Invoke-Command -Session $EXOSession -ArgumentList $($ADUser.mail),$true -ScriptBlock {set-mailbox -Identity $args[0] -LitigationHoldEnabled $args[1] -LitigationHoldDuration unlimited}
        
        # - Load todays date and time into a varaible, and output the UPN to the logfile that has this timestamp in its name.
        $DateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        "$($UPN)" | Out-File -Encoding utf8 -Force -FilePath "$($Logpath)\MailboxCreation_$($DateStamp).txt" 
        }
    # - Remove the processed workfile now that the all entries have been processed.
    remove-item $ScriptDataFileStart -Force
    }

Remove-PSSession $EXPSession
Remove-PSSession $EXOSession
 
