$disclaimer = @"
#------------------------------------------------------------------------------
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# Author Rui Pereira Tabares
#------------------------------------------------------------------------------
"@
Write-Host $disclaimer -ForegroundColor Yellow

#------------------------------------------------------------------------------
# PowerShell Functions
#------------------------------------------------------------------------------
# import required functions
. .\Assert-ServiceConnection.ps1
. .\Connect-OnlineServices.ps1

#region
####Getting powershell sessions ########
# Check current connection status, and connect if needed
$ServicesToConnect = Assert-ServiceConnection
# Connect to services if ArrayList is not empty
if ( $ServicesToConnect.Count ) { 
    $UserCredential = Get-Credential -Message "Enter Global admin Credentials"
    Connect-OnlineServices -Credential $UserCredential -Services $ServicesToConnect }
#endregion

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
    write-host                                              ' 1)  Export teams chat data from a user mailbox hosted online or A Team'
    write-host                                              ' 2)  Export teams chat data from a user hosted onprem  '
    write-host
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Red 'End of PowerShell - Script menu ' 
    write-host -ForegroundColor  white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor Yellow            "3)  Exit the PowerShell script menu" 
    write-host
    [int]$opt = Read-Host "Select an option [1-3]"
    write-host $opt
    switch ($opt) 
    {
        1 {  
            ##### Online mailbox validation also works for O365 groups#####
            do{ 
                $email = Read-Host "Enter an email address"
                $emailAddress = $email

                $userMailbox=get-mailbox -identity $email  -ErrorAction SilentlyContinue
                $GroupMailbox= Get-unifiedGroup -Identity $email -ErrorAction SilentlyContinue
                    if(($null -eq $userMailbox) -and ($null -eq $GroupMailbox))
                    {
                        write-host "Error:Please enter a valid online mailbox" -ForegroundColor red
                    }

            } while(($null -eq $userMailbox) -and ($null -eq $GroupMailbox))

            #### saving searched user displayname
            $DisplayName = $userMailbox.DisplayName 

            write-host "Performing Search on $DisplayName's Teams chat history"
            ##Starting search from user
            $searchName = "$(((Get-Date).ToString("HH:mm:ss"))) + $DisplayName"
            
            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery "kind:microsoftteams AND kind:im" -ExchangeLocation $emailAddress
            Start-ComplianceSearch $searchName
            write-host "A Search from Teams chat history has started with the Seacrh Name" $searchName
            ## loop until search completes

            do {
                Write-host "Waiting for search to complete..."
                Start-Sleep -s 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')

            if ($complianceSearch.Items -gt 0)
            {
                # Create a Complinace Search Action and wait for it to complete. The folders will be listed in the .Results parameter
                $complianceSearchAction = New-ComplianceSearchaction -SearchName $searchName -Preview
                do {
                    Write-host "Waiting for search action to complete..."
                    Start-Sleep -s 5
                    $complianceSearchAction = Get-ComplianceSearchAction $searchActionName
                } while ($complianceSearchAction.Status -ne 'Completed')
                
                $results = Get-ComplianceSearch -Identity $searchName |Select-Object successresults
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
                Write-Host "go to https://protection.office.com/#/contentsearch and export your PST. (Opening site automatically)"
                Start-Process "https://protection.office.com/#/contentsearch"
                write-host
                write-host
                Read-Host "Press Enter to get back to the menu..."
                write-host
                write-host
            }
        }
    
        2 {   
            do{
                $UserPrincipalName = Read-Host "Enter User Principal Name (UPN)"
                $User = Get-MsolUser -UserPrincipalName $UserPrincipalName  -ErrorAction SilentlyContinue
                if($null -eq $User) {
                    write-host "User must be synced to the cloud or there is a misstype on the UserPrincipalName" -ForegroundColor red 
                }
            } while ($null -eq $User)

            $ValidateExoLicenseE = $User | Where-Object {$_.Licenses.ServiceStatus | Where-Object {$_.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE" -and $_.ProvisioningStatus -eq "Success"}}
            $ValidateExoLicenseS = $User | Where-Object {$_.Licenses.ServiceStatus | Where-Object {$_.ServicePlan.ServiceName -eq "EXCHANGE_S_STANDARD" -and $_.ProvisioningStatus -eq "Success"}}
                
            if ( -not (($ValidateExoLicenseE.IsLicensed) -or ($ValidateExoLicenseS.IsLicensed))) 
            {
                write-host "Error:User does not has an exchange online license" -ForegroundColor red 
                write-host "See requirements on https://docs.microsoft.com/en-us/microsoft-365/compliance/search-cloud-based-mailboxes-for-on-premises-users" -ForegroundColor red 
                Exit
            }

            $OnpremValidation=  $User | Select-Object -Property DisplayName, UserPrincipalName, isLicensed, @{label='MailboxLocation';expression={switch ($_.MSExchRecipientTypeDetails) {1 {'Onprem'; break} 2147483648 {'Office365'; break} default {'Unknown'}}}}
            $ValidationDup = Get-Mailbox -identity $UserPrincipalName  -ErrorAction SilentlyContinue

            if ( -not (( $OnpremValidation.MailboxLocation -eq "Onprem") -and ($null -eq $ValidationDup)) )
            { 
                write-host "WARNING!!:MSExchRecipientTypeDetails is not in value 1, there is an exchange online mailbox for this user" -ForegroundColor yellow 
                write-host "Please verify if there is not a duplicate mailbox in online for this user or you selected the wrong option" -ForegroundColor yellow 

                Exit
            }

            $PrimarySmtp= (Get-Recipient -identity $UserPrincipalName).PrimarySmtpAddress

            ##Starting search from user
            $DisplayName = (Get-Recipient -identity $UserPrincipalName).Name
            $searchName = "$(((Get-Date).ToString("HH:mm:ss"))) + $DisplayName"
            ####Getting powershell sessions ########
            
            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery "kind:im" -ExchangeLocation $PrimarySmtp -IncludeUserAppContent $true -AllowNotFoundExchangeLocationsEnabled $true
            Start-ComplianceSearch $searchName
            write-host "A Search from Teams chat history has started with the Seacrh Name" $searchName
            ## loop until search completes

            do{
                Write-host "Waiting for search to complete..."
                Start-Sleep -s 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')


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
                
                $results = Get-ComplianceSearch -Identity $searchName |Select-Object successresults
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
                Write-Host "go to https://protection.office.com/#/contentsearch and export your PST. (Opening site automatically)"
                Start-Process "https://protection.office.com/#/contentsearch"
                write-host 
                write-host
                Read-Host "Press Enter to get back to the menu..."
                write-host
                write-host
            }
        }       

        3 {
            $Loop = $true
            Exit
        }  
    }   
}
