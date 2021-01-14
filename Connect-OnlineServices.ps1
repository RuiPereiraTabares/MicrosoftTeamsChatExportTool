function Connect-OnlineServices {
    <#
    .SYNOPSIS
    Connect to Online Services.

    .DESCRIPTION
    Use this function to connect to EXO, SCC, MicrosoftTeams, MS Online and AzureAD Online Services.

    .PARAMETER Credential
    Credential to use for the connection.

    .PARAMETER Services
    List of the desired services to connect to. Current available services: EXO, SCC, MicrosoftTeams, MSOnline, AzureAD, AzureADPreview, Azure.

    .PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

    .PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.
   
    .EXAMPLE
    PS C:\> Connect-OnlineServices -Credential $UserCredential -EXO -AzureAD
    Connects to Exchange and AzureAD Online Services with the passed User Credentials variable.
    
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "")]
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param(
        [PSCredential]
        $Credential = (Get-Credential -Message "Please specify O365 Global Admin Credentials"),

        [ValidateSet('EXO', 'SCC', 'MicrosoftTeams', 'MSOnline', 'AzureAD', 'AzureADPreview', 'Azure')]
        [String[]]
        $Services
    )
    if(-not $Credential){
        $Credential = Get-Credential -Message "Please specify O365 Global Admin Credentials"
    }
    if(-not $Credential){
        Write-host "Credentials entered are invalid." -ForegroundColor white -BackgroundColor red
        Exit           
    }

    Switch ( $Services ) {
        Azure {
            Invoke-PSFProtectedCommand -Action "Connecting to Azure" -Target "Azure" -ScriptBlock {
                Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Azure"
                Install-Module Azure -Force -ErrorAction Stop
                Import-Module Azure -ErrorAction Stop
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        AzureAD {
            Invoke-PSFProtectedCommand -Action "Connecting to AzureAD" -Target "AzureAD" -ScriptBlock {
                Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to AzureAD"
                if ( !(Get-Module AzureAD -ListAvailable) -and !(Get-Module AzureAD) ) {
                    Install-Module AzureAD -Force -ErrorAction Stop
                }
                try {
                    Import-module AzureAD
                    $null = Connect-AzureAD -Credential $Credential -ErrorAction Stop
                }
                catch {
                    if ( ($_.Exception.InnerException.InnerException.InnerException.InnerException.ErrorCode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                        Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Azure AD. Requesting to authenticate"
                        $null = Connect-AzureAD -AccountId $Credential.UserName.toString() -ErrorAction Stop
                    }
                    else {
                        return $_
                    }
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        AzureADPreview {
            Invoke-PSFProtectedCommand -Action "Connecting to AzureAD Preview" -Target "AzureAD" -ScriptBlock {
                Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to AzureAD Preview"
                if ( !(Get-Module AzureADPreview -ListAvailable) -and !(Get-Module AzureADPreview) ) {
                    Install-Module AzureADPreview -Force -ErrorAction Stop
                }
                try {
                    Import-module AzureADPreview
                    $null = Connect-AzureAD -Credential $Credential -ErrorAction Stop
                }
                catch {
                    if ( ($_.Exception.InnerException.InnerException.InnerException.InnerException.ErrorCode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                        Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Azure AD. Requesting to authenticate"
                        $null = Connect-AzureAD -AccountId $Credential.UserName.toString() -ErrorAction Stop
                    }
                    else {
                        return $_
                    }
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        MSOnline {
            Invoke-PSFProtectedCommand -Action "Connecting to MSOnline" -Target "MSOnline" -ScriptBlock {
                Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to MSOnline"
                if ( !(Get-Module MSOnline -ListAvailable) -and !(Get-Module MSOnline) ) {
                    Install-Module MSOnline -Force -ErrorAction Stop
                }
                try {
                    Import-Module MSOnline
                    Connect-MsolService -Credential $Credential -ErrorAction Stop
                }
                catch {
                    Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to MS Online. Requesting to authenticate"
                    Connect-MsolService -ErrorAction Stop
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        MicrosoftTeams {
            Invoke-PSFProtectedCommand -Action "Connecting to MicrosoftTeams" -Target "MicrosoftTeams" -ScriptBlock {
                Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to MicrosoftTeams"
                if ( !(Get-Module MicrosoftTeams -ListAvailable) -and !(Get-Module MicrosoftTeams) ) {
                    Install-Module MicrosoftTeams -Force -ErrorAction Stop
                }
                try {
                    #Connect to Microsoft Teams
                    $null = Connect-MicrosoftTeams -Credential $Credential -ErrorAction Stop
    
                    #Connection to Skype for Business Online and import into Ps session
                    $session = New-CsOnlineSession -Credential $Credential -ErrorAction Stop
                    $null = Import-PsSession $session
                }
                catch {
                    if ( ($_.Exception.InnerException.InnerException.InnerException.InnerException.ErrorCode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                        Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to MicrosoftTeams. Requesting to authenticate"
                        #Connect to Microsoft Teams
                        $null = Connect-MicrosoftTeams -ErrorAction Stop
    
                        #Connection to Skype for Business Online and import into Ps session
                        $session = New-CsOnlineSession -ErrorAction Stop
                        $null = Import-PsSession $session
                    }
                    else {
                        return $_
                    }
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        SCC {
            Invoke-PSFProtectedCommand -Action "Connecting to Security and Compliance" -Target "SCC" -ScriptBlock {
                Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security and Compliance"
                try {
                    Connect-IPPSSession -Credential $Credential -ErrorAction Stop -WarningAction SilentlyContinue
                }
                catch {
                    if ( ($_.Exception.InnerException.InnerException.InnerException.InnerException.ErrorCode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                        Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Security and Compliance. Requesting to authenticate"
                        Connect-IPPSSession -UserPrincipalName $Credential.Username.toString() -ErrorAction Stop -WarningAction SilentlyContinue
                    }
                    else {
                        return $_
                    }
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }

        EXO {
            Invoke-PSFProtectedCommand -Action "Connecting to Exchange Online" -Target "EXO" -ScriptBlock {
                Write-PSFHostColor -String "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online"
                try {
                    # Getting current PS Sessions
                    $Sessions = Get-PSSession
                    if ($Sessions.ComputerName -eq "outlook.office365.com") { return }
                    else { Connect-ExchangeOnline -Credential $Credential -ShowBanner:$False -ErrorAction Stop }
                }
                catch {
                    if ( ($_.Exception.InnerException.InnerException.InnerException.InnerException.ErrorCode | ConvertFrom-Json).error -eq 'interaction_required' ) {
                        Write-PSFHostColor -String  "[$((Get-Date).ToString("HH:mm:ss"))] Your account seems to be requiring MFA to connect to Exchange Online. Requesting to authenticate"
                        Connect-ExchangeOnline -UserPrincipalName $Credential.Username.toString() -ShowBanner:$False -ErrorAction Stop
                    }
                    else {
                        return $_
                    }
                }
            } -EnableException $true -PSCmdlet $PSCmdlet
        }
    }
}