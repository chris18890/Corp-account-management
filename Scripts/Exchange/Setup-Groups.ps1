[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)

#====================================================================
#group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Group,
        [Parameter(Mandatory)][string]$Member
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          User will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    # Returns:
    # Notes:
    #================================================================
    $Error.Clear()
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)"
    if ($checkGroup -ne $null) {
        Write-Host "Adding $Member to $Group"
        try {
            Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            Write-Host "Added $Member to $Group"
        }
        catch [Microsoft.ActiveDirectory.Management.ADException] {
            switch ($Error[0].Exception.ErrorCode) {
                1378 { # 'The specified object is already a member of the group'
                    Write-Host "'$Member' is already a member of group '$Group'" -ForegroundColor Green
                }
                default {
                    Write-Host "ERROR: An unexpected error occurred while attempting to add user '$Member' to a group:`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                }
            }
        }
    } else {
        Write-Host "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

#====================================================================
function Test-Cred {
    [CmdletBinding()]
    [OutputType([String])]
    Param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            'PSCredential'
        )]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credentials
    )
    #================================================================
    # Purpose:          Test Provided Credentials Prior to Use
    # Assumptions:      
    # Effects:
    # Inputs:           $Cred
    # $LogString:       
    # Calls:
    # Returns:         Authenticated or Unauthenticated
    # Notes:
    #================================================================
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
    if ($Credentials -eq $null) {
        try {
            $Credentials = Get-Credential "domain\$env:username" -ErrorAction Stop
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Warning "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
    }
    # Checking module
    try {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    } catch {
        $_.Exception.Message
        Continue
    }
    if (!$domain) {
        Write-Warning "Something went wrong"
    } else {
        if ($domain.name -ne $null) {
            return "Authenticated"
        } else {
            return "Not authenticated"
        }
    }
}
#====================================================================

$Domain = "$env:userdomain"
$SubDomain = "exchange" #enter subdomain
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$StaffGroup = "Staff"
$ITGroup = "IT"
$ITAdminGroup = "IT_Admin"
$Groups = @($StaffGroup, $ITGroup, "UG_Office365")
$AdminGroups = (Get-ADGroup -Filter * -searchbase "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -Properties *).Name
$ExServer = "$Domain-Exch.$DNSSuffix" #Remote Exchange PS session
$DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
Add-GroupMember -Group "Organization Management" -Member "Domain Admins"
Add-GroupMember -Group "Organization Management" -Member "UG_Level_3_Admins"
Add-GroupMember -Group "Server Management" -Member "RG_Server_Admins"
Add-GroupMember -Group "Recipient Management" -Member "RG_Account_Admins"
Add-GroupMember -Group "Help Desk" -Member "UG_Level_1_Admins"

#====================================================================
# Get user credentials for server connectivity (Non-MFA)
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-Host "Failed to validate credentials: $ErrorMsg "
    Pause
    Break
}
$CredCheck = $Cred | Test-Cred
if ($CredCheck -ne "Authenticated") {
    Write-Host "Credential validation failed - Script Terminating"
    pause
    Exit
}
#Connect to remote Exchange PowerShell
Write-Host "Connecting to remote Exchange PowerShell session... "
try {
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
    $ExConnected = $true
    Write-Host "connected."
    Write-Host "Importing Exchange session... "
    Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
    Write-Host "done."
} catch {
    $e = $_.Exception
    Write-Host $e
    $line = $_.InvocationInfo.ScriptLineNumber
    Write-Host $line
    $msg = $e.Message
    Write-Host $msg
    $Action = "Error Importing Exchange Session"
    Write-Host $Action
    Write-Host "failed."
    Write-Host "ERROR: $_" -ForegroundColor Red
}
if (!$ExSession) {
    Write-Host "Exchange session not connected Stopping Script"
    Exit
}

New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 80 -HostHeader "$subdomain.$EmailSuffix"

Get-ClientAccessServer -Identity $ExServer | Set-ClientAccessServer -AutoDiscoverServiceInternalUri "https://autodiscover.$EmailSuffix/Autodiscover/Autodiscover.xml"
Get-EcpVirtualDirectory -Server $ExServer | Set-EcpVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/ecp" -InternalUrl "https://$subdomain.$EmailSuffix/ecp"
Get-WebServicesVirtualDirectory -Server $ExServer | Set-WebServicesVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/EWS/Exchange.asmx" -InternalUrl "https://$subdomain.$EmailSuffix/EWS/Exchange.asmx"
Get-MapiVirtualDirectory -Server $ExServer | Set-MapiVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/mapi" -InternalUrl "https://$subdomain.$EmailSuffix/mapi"
Get-ActiveSyncVirtualDirectory -Server $ExServer | Set-ActiveSyncVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/Microsoft-Server-ActiveSync" -InternalUrl "https://$subdomain.$EmailSuffix/Microsoft-Server-ActiveSync"
Get-OabVirtualDirectory -Server $ExServer | Set-OabVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/OAB" -InternalUrl "https://$subdomain.$EmailSuffix/OAB"
Get-OwaVirtualDirectory -Server $ExServer | Set-OwaVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/owa" -InternalUrl "https://$subdomain.$EmailSuffix/owa"
Get-PowerShellVirtualDirectory -Server $ExServer | Set-PowerShellVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/powershell" -InternalUrl "https://$subdomain.$EmailSuffix/powershell"
Get-OutlookAnywhere -Server $ExServer | Set-OutlookAnywhere -ExternalHostname "$subdomain.$EmailSuffix" -InternalHostname "$subdomain.$EmailSuffix" -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true -DefaultAuthenticationMethod NTLM -SSLOffloading $false

# add accepted domain
New-AcceptedDomain -DomainName $EmailSuffix -DomainType Authoritative -Name $Domain -DomainController $DCHostName
# edit default address policy
Get-EmailAddressPolicy -Identity "Default Policy" | Set-EmailAddressPolicy -EnabledEmailAddressTemplates "SMTP:@$EmailSuffix" -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "Default Policy" -DomainController $DCHostName
# add new address policy
New-EmailAddressPolicy -Name "FirstName LastName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffGroup,$EndPath" -EnabledEmailAddressTemplates "SMTP:%g.%s@$EmailSuffix" -Priority 1 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "FirstName LastName" -DomainController $DCHostName
New-EmailAddressPolicy -Name "UserName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffGroup,$EndPath" -EnabledEmailAddressTemplates "SMTP:%m@$EmailSuffix" -Priority 2 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "UserName" -DomainController $DCHostName
foreach ($GroupName in $Groups) {
    Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
    Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
}
foreach ($GroupName in $AdminGroups) {
    Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
    Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName -HiddenFromAddressListsEnabled $true
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    Write-Host "Closed Exchange session."
}
#====================================================================
