#------------------------------------------------------------------------------
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# Author Rui Pereira Tabares



#------------------------------------------------------------------------------
# PowerShell Functions
#------------------------------------------------------------------------------

if( -not $UserCredential ){
    write-host "[$((Get-Date).ToString("HH:mm:ss"))] Enter tenant admin credentials:"
    $UserCredential = Get-Credential -Message "Please specify O365 Global Admin Credentials"
}
write-host -ForegroundColor Red  "If you have never run this tool before Please verify if your admin has the following permissions:" 
Write-host -ForegroundColor Red "Go to https://protection.office.com/permissions"
Write-host -ForegroundColor Red "Add your admin account into the Ediscovery Manager and Compliance Administrator Permissions"
write-host -ForegroundColor Red "After adding your admin into those permissions wait 30 to 40 minutes to be effective"

$Loop = $true
While ($Loop)
{
write-host 
    write-host ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host -BackgroundColor Magenta  "Teams chat Export Tool"
    write-host ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host
                write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Green   'Please select your option           ' 
    write-host -ForegroundColor white '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 1)  Export teams chat data from a user mailbox  hosted online'
                write-host                                              ' 2)  Export teams chat data from a user mailbox hosted onprem  '
                write-host                                              ' 3)  Export teams chat data from a Team  '
    write-host
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Red 'End of PowerShell - Script menu ' 
    write-host -ForegroundColor  white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor Yellow            "4)  Exit the PowerShell script menu" 
    write-host
    $opt = Read-Host "Select an option [1-3]"
    write-host $opt
    switch ($opt) 
    {



        1
        {  
           
            ####Getting powershell sessions ########
             $Sessions = Get-PSSession
             #region Connecting to SCC
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security & Compliance Center if not already connected..." -foregroundColor Green
if ( -not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com") ) {
    write-host
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession -Credential $UserCredential
}
#endregion

#region Connecting to Exchange Online
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online if not already connected..." -foregroundColor Green
if ( $Sessions.ComputerName -notcontains "outlook.office365.com" ) {
    write-host -
    Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$False
}
#endregion

            

                

           

         ##### Online mailbox validation also works for O365 groups#####
                do{ 
                    $email = Read-Host "Enter an email address (Note: for groups you can obtain this from https://admin.microsoft.com/Adminportal/Home?source=applauncher#/groups)"
                    $emailAddress = $email

                $userMailbox=get-mailbox -identity $email  -ErrorAction SilentlyContinue
          
                    if(($userMailbox -eq $null))
                    {
                    write-host "Error:Please enter a valid online mailbox" -ForegroundColor red 

                     }

                } while(($userMailbox -eq $null))

                    #### saving searched user displayname
                    $DisplayName= $userMailbox.DisplayName 

                write-host "Performing Search on $DisplayName Teams chat history"
                    ##Starting search from user
                $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
                
             $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery "kind:microsoftteams AND kind:im" -ExchangeLocation $emailAddress
            Start-ComplianceSearch $searchName
                write-host "A Search from Teams chat history has started with the Seacrh Name" $searchName
                    ## loop until search completes

                     do{
                    Write-host "Waiting for search to complete..."
                     Start-Sleep -s 5
                     $complianceSearch = Get-ComplianceSearch $searchName
                     }while ($complianceSearch.Status -ne 'Completed')


                    if ($complianceSearch.Items -gt 0)
                    {
                        # Create a Complinace Search Action and wait for it to complete. The folders will be listed in the .Results parameter
                        $complianceSearchAction = New-ComplianceSearchaction -SearchName $searchName -Preview
                        do
                        {
                            Write-host "Waiting for search action to complete..."
                            Start-Sleep -s 5
                            $complianceSearchAction = Get-ComplianceSearchAction $searchActionName
                        }while ($complianceSearchAction.Status -ne 'Completed')
                        
                        
                        $results = Get-ComplianceSearch -Identity $searchName |select successresults
                        $results= $results -replace "@{SuccessResults={", "" -replace "}}",""
                        $results -match "size:","(\d+)"
                        $match= $matches[1]
                        $matchmb= $match/1Mb 
                        $matchGb= $match/1Gb 
                        Write-Host "------------------------"
                        Write-Host "Results"
                        Write-Host "------------------------"
                        Write-Host "$results"
                        Write-Host "------------------------"
                        Write-Host "Found Size"
                        Write-Host "$matchmb","Mb"
                        Write-Host "$matchGb","Gb"
                        Write-Host "________________________"
                        Write-Host -foregroundcolor green "Success"
                        Write-Host "________________________"
                        Write-Host 
                        Write-Host "Step 1: Go to Office 365 Security & Compliance->Searches"

    
Write-Host "Step 2: Check the search with name"

    
Write-Host "Step 3: Verify that you can view chat messages exported by clicking View Results"

    
Write-Host "Step 4: Click "Export results", use the default options and click Export"

    
Write-Host "Step 5: Now goto Office 365 Security & Compliance->Exports and click Refresh"

    
Write-Host "Step 6: Click the export with name"
  
Write-Host "Step 7:  Copy the Export Key and then click Download results and paste the Export key and specify the location where you want to download the exported chat messages."

    
Write-Host "Step 8: Click Start"

    
Write-Host "Step 9: When you see message download completed successfully then click on the link under Export location: to get to the exported files. Here you will find the exported PST of chat messages under Exchange folder <user or group email>.pst"

Write-Host " Step 10: Open Outlook application on your Windows PC and click File>Open&Export>Open Outlook data file (.pst) and provide location of the PST file from Step 9. In Outlook you will need to located the folder <user or group email> and under that TeamsMessagesData. You will see all the chat messages exported here now."
        }
        }
    


        2
        {   

            $Module=get-module

                if ( -not ($Module.Name -match "MSOnline") )
                  {
                 write-host
                write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Msolservice" -foregroundColor Green
                if ( !(Get-Module MSOnline -ListAvailable) -and !(Get-Module MSOnline) ) {
                Install-Module MSOnline -Force -ErrorAction Stop
                }
                Connect-msolservice
                }

           

             ####Getting powershell sessions ########
             $Sessions = Get-PSSession
             #region Connecting to SCC
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security & Compliance Center if not already connected..." -foregroundColor Green
if ( -not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com") ) {
    write-host
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession -Credential $UserCredential
}
#endregion

#region Connecting to Exchange Online
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online if not already connected..." -foregroundColor Green
if ( $Sessions.ComputerName -notcontains "outlook.office365.com" ) {
    write-host -
    Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$False
}
#endregion


                do{
                $UserPrincipalName = Read-Host "Enter User Principal Name (UPN)"
                    
                    

                $User=get-MsolUser -UserPrincipalName $UserPrincipalName  -ErrorAction SilentlyContinue
                    if($User -eq $null)
                    {
                    write-host "User must be synced to the cloud or there is a misstype on the UserPrincipalName" -ForegroundColor red 

                     }

                } while($User -eq $null)

             $ValidateExoLicenseE = $User | ? {$_.Licenses.ServiceStatus | ? {$_.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE" -and $_.ProvisioningStatus -eq "Success"}}
            $ValidateExoLicenseS = $User | ? {$_.Licenses.ServiceStatus | ? {$_.ServicePlan.ServiceName -eq "EXCHANGE_S_STANDARD" -and $_.ProvisioningStatus -eq "Success"}}
                
                if ( -not (($ValidateExoLicenseE.IsLicensed) -or ($ValidateExoLicenseS.IsLicensed))) 
                {

                    write-host "Error:User does not has an exchange online license" -ForegroundColor red 
                    write-host "See requirements on https://docs.microsoft.com/en-us/microsoft-365/compliance/search-cloud-based-mailboxes-for-on-premises-users" -ForegroundColor red 
                    Exit
                   
                }

            $OnpremValidation=  $User | Select-Object -Property DisplayName, UserPrincipalName, isLicensed, @{label='MailboxLocation';expression={switch ($_.MSExchRecipientTypeDetails) {1 {'Onprem'; break} 2147483648 {'Office365'; break} default {'Unknown'}}}}
            
            $ValidationDup=get-mailbox -identity $UserPrincipalName  -ErrorAction SilentlyContinue


            if ( -not (( $OnpremValidation.MailboxLocation -eq "Onprem") -and ($ValidationDup -eq $null)))
                { write-host "WARNING!!:MSExchRecipientTypeDetails is not in value 1, there is an exchange online mailbox for this user" -ForegroundColor yellow 
                  write-host "Please verify if there is not a duplicate mailbox in online for this user or you selected the wrong option" -ForegroundColor yellow 

                    Exit
                }

            $PrimarySmtp= (Get-Recipient -identity $UserPrincipalName).PrimarySmtpAddress

                 ##Starting search from user
                 $DisplayName = (Get-Recipient -identity $UserPrincipalName).Name
                $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
                
                
            $complianceSearch = New-ComplianceSearch "Redstone_Search" -ContentMatchQuery "kind:im" -ExchangeLocation $PrimarySmtp -IncludeUserAppContent $true -AllowNotFoundExchangeLocationsEnabled $true
             Start-ComplianceSearch $searchName
                write-host "A Search from Teams chat history has started with the Seacrh Name" $searchName
                    ## loop until search completes

                     do{
                    Write-host "Waiting for search to complete..."
                     Start-Sleep -s 5
                     $complianceSearch = Get-ComplianceSearch $searchName
                     }while ($complianceSearch.Status -ne 'Completed')


                    if ($complianceSearch.Items -gt 0)
                    {
                        # Create a Complinace Search Action and wait for it to complete. The folders will be listed in the .Results parameter
                        $complianceSearchAction = New-ComplianceSearchaction -SearchName $searchName -Preview
                        do
                        {
                            Write-host "Waiting for search action to complete..."
                            Start-Sleep -s 5
                            $complianceSearchAction = Get-ComplianceSearchAction $searchActionName
                        }while ($complianceSearchAction.Status -ne 'Completed')
                        
                        
                        $results = Get-ComplianceSearch -Identity $searchName |select successresults
                        $results= $results -replace "@{SuccessResults={", "" -replace "}}",""
                        $results -match "size:","(\d+)"
                        $match= $matches[1]
                        $matchmb= $match/1Mb 
                        $matchGb= $match/1Gb 
                        Write-Host "------------------------"
                        Write-Host "Results"
                        Write-Host "------------------------"
                        Write-Host "$results"
                        Write-Host "------------------------"
                        Write-Host "Found Size"
                        Write-Host "$matchmb","Mb"
                        Write-Host "$matchGb","Gb"
                        Write-Host "________________________"
                        Write-Host -foregroundcolor green "Success"
                        Write-Host "________________________"
                       Write-Host "Step 1: Go to Office 365 Security & Compliance->Searches"

    
Write-Host "Step 2: Check the search with name"

    
Write-Host "Step 3: Verify that you can view chat messages exported by clicking View Results"

    
Write-Host "Step 4: Click "Export results", use the default options and click Export"

    
Write-Host "Step 5: Now goto Office 365 Security & Compliance->Exports and click Refresh"

    
Write-Host "Step 6: Click the export with name"
  
Write-Host "Step 7:  Copy the Export Key and then click Download results and paste the Export key and specify the location where you want to download the exported chat messages."

    
Write-Host "Step 8: Click Start"

    
Write-Host "Step 9: When you see message download completed successfully then click on the link under Export location: to get to the exported files. Here you will find the exported PST of chat messages under Exchange folder <user or group email>.pst"

Write-Host " Step 10: Open Outlook application on your Windows PC and click File>Open&Export>Open Outlook data file (.pst) and provide location of the PST file from Step 9. In Outlook you will need to located the folder <user or group email> and under that TeamsMessagesData. You will see all the chat messages exported here now."
                        }
           
           }
               

        3   
        {  
           
            ####Getting powershell sessions ########
             $Sessions = Get-PSSession
             #region Connecting to SCC
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security & Compliance Center if not already connected..." -foregroundColor Green
if ( -not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com") ) {
    write-host
    if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
        Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession -Credential $UserCredential
}
#endregion

#region Connecting to Exchange Online
write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online if not already connected..." -foregroundColor Green
if ( $Sessions.ComputerName -notcontains "outlook.office365.com" ) {
    write-host -
    Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$False
}
#endregion


                

           

         ##### Online mailbox validation also works for O365 groups#####
                do{ 
                    $email = Read-Host "Enter an email address"
                    $emailAddress = $email

                $userMailbox=get-unifiedgroup -identity $email  -ErrorAction SilentlyContinue
          
                    if(($userMailbox -eq $null))
                    {
                    write-host "Error:Please enter a valid Group(team) Email adress" -ForegroundColor red 

                     }

                } while(($userMailbox -eq $null))

                    #### saving searched user displayname
                    $DisplayName= $userMailbox.DisplayName 

                write-host "Performing Search on $DisplayName Teams chat history"
                    ##Starting search from user
                $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
                
             $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery "kind:microsoftteams AND kind:im" -ExchangeLocation $emailAddress
            Start-ComplianceSearch $searchName
                write-host "A Search from Teams chat history has started with the Seacrh Name" $searchName
                    ## loop until search completes

                     do{
                    Write-host "Waiting for search to complete..."
                     Start-Sleep -s 5
                     $complianceSearch = Get-ComplianceSearch $searchName
                     }while ($complianceSearch.Status -ne 'Completed')


                    if ($complianceSearch.Items -gt 0)
                    {
                        # Create a Complinace Search Action and wait for it to complete. The folders will be listed in the .Results parameter
                        $complianceSearchAction = New-ComplianceSearchaction -SearchName $searchName -Preview
                        do
                        {
                            Write-host "Waiting for search action to complete..."
                            Start-Sleep -s 5
                            $complianceSearchAction = Get-ComplianceSearchAction $searchActionName
                        }while ($complianceSearchAction.Status -ne 'Completed')
                        
                        
                        $results = Get-ComplianceSearch -Identity $searchName |select successresults
                        $results= $results -replace "@{SuccessResults={", "" -replace "}}",""
                        $results -match "size:","(\d+)"
                        $match= $matches[1]
                        $matchmb= $match/1Mb 
                        $matchGb= $match/1Gb 
                        Write-Host "------------------------"
                        Write-Host "Results"
                        Write-Host "------------------------"
                        Write-Host "$results"
                        Write-Host "------------------------"
                        Write-Host "Found Size"
                        Write-Host "$matchmb","Mb"
                        Write-Host "$matchGb","Gb"
                        Write-Host "________________________"
                        Write-Host -foregroundcolor green "Success"
                        Write-Host "________________________"
                      Write-Host "Step 1: Go to Office 365 Security & Compliance->Searches"

    
Write-Host "Step 2: Check the search with name"

    
Write-Host "Step 3: Verify that you can view chat messages exported by clicking View Results"

    
Write-Host "Step 4: Click "Export results", use the default options and click Export"

    
Write-Host "Step 5: Now goto Office 365 Security & Compliance->Exports and click Refresh"

    
Write-Host "Step 6: Click the export with name"
  
Write-Host "Step 7:  Copy the Export Key and then click Download results and paste the Export key and specify the location where you want to download the exported chat messages."

    
Write-Host "Step 8: Click Start"

    
Write-Host "Step 9: When you see message download completed successfully then click on the link under Export location: to get to the exported files. Here you will find the exported PST of chat messages under Exchange folder <user or group email>.pst"

Write-Host " Step 10: Open Outlook application on your Windows PC and click File>Open&Export>Open Outlook data file (.pst) and provide location of the PST file from Step 9. In Outlook you will need to located the folder <user or group email> and under that TeamsMessagesData. You will see all the chat messages exported here now."
                        }
        }
        
   4
        {

$Loop = $true
Exit
} 
        }
       
}
