#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$SubDomain
)

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
if (!$SubDomain) {
    $SubDomain = "exchange" #enter subdomain
}
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$StaffGroup = "Staff"
$MailGroups = (Get-ADGroup -Filter * -searchbase "OU=Groups,$EndPath" -Properties *).Name
$ExServer = "$Domain-Exch.$DNSSuffix" #Remote Exchange PS session
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain Exchange Post-Install Config Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_exchange_postinstall_config_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log ("=" * 80)
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""
$requiredGroups = @('Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

Add-GroupMember -Group "Organization Management" -Member "Domain Admins"
Add-GroupMember -Group "Organization Management" -Member "ADM_Role_Tier1_Level_3_Admins"
Add-GroupMember -Group "Server Management" -Member "ADM_Task_Server_Admins"
Add-GroupMember -Group "Recipient Management" -Member "ADM_Task_HiPriv_Account_Admins"
Add-GroupMember -Group "Recipient Management" -Member "ADM_Task_HiPriv_Group_Admins"
Add-GroupMember -Group "Recipient Management" -Member "ADM_Task_Standard_Account_Admins"
Add-GroupMember -Group "Recipient Management" -Member "ADM_Task_Standard_Group_Admins"
Add-GroupMember -Group "Recipient Management" -Member "ADM_Task_SER_Account_Admins"
Add-GroupMember -Group "Help Desk" -Member "ADM_Role_Tier1_Level_1_Admins"

#====================================================================
# Get user credentials for server connectivity (Non-MFA)
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-Log "Failed to validate credentials: $ErrorMsg "
    Read-Host -Prompt "Press Enter to exit"
    Break
}
#Connect to remote Exchange PowerShell
Write-Log "Connecting to remote Exchange PowerShell session... "
try {
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
    Write-Log "connected."
    Write-Log "Importing Exchange session... "
    Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
    Write-Log "done."
} catch {
    $e = $_.Exception
    Write-Log $e
    $line = $_.InvocationInfo.ScriptLineNumber
    Write-Log $line
    $msg = $e.Message
    Write-Log $msg
    $Action = "Error Importing Exchange Session"
    Write-Log $Action
    Write-Log "failed."
    Write-Log "ERROR: $_" -ForegroundColor Red
}
if (!$ExSession) {
    Write-Log "Exchange session not connected Stopping Script"
    Exit
}

if (-not (Get-WebBinding -Name "Default Web Site" -HostHeader "$subdomain.$EmailSuffix" -ErrorAction SilentlyContinue)) {
    New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 80 -HostHeader "$subdomain.$EmailSuffix"
} else {
    Write-Log "Web binding for '$subdomain.$EmailSuffix' already exists" -ForegroundColor Green
}

Write-Log "Setting Virtual Directories"
Get-ClientAccessServer -Identity $ExServer | Set-ClientAccessServer -AutoDiscoverServiceInternalUri "https://autodiscover.$EmailSuffix/Autodiscover/Autodiscover.xml"
Get-EcpVirtualDirectory -Server $ExServer | Set-EcpVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/ecp" -InternalUrl "https://$subdomain.$EmailSuffix/ecp"
Get-WebServicesVirtualDirectory -Server $ExServer | Set-WebServicesVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/EWS/Exchange.asmx" -InternalUrl "https://$subdomain.$EmailSuffix/EWS/Exchange.asmx"
Get-MapiVirtualDirectory -Server $ExServer | Set-MapiVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/mapi" -InternalUrl "https://$subdomain.$EmailSuffix/mapi"
Get-ActiveSyncVirtualDirectory -Server $ExServer | Set-ActiveSyncVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/Microsoft-Server-ActiveSync" -InternalUrl "https://$subdomain.$EmailSuffix/Microsoft-Server-ActiveSync"
Get-OabVirtualDirectory -Server $ExServer | Set-OabVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/OAB" -InternalUrl "https://$subdomain.$EmailSuffix/OAB"
Get-OwaVirtualDirectory -Server $ExServer | Set-OwaVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/owa" -InternalUrl "https://$subdomain.$EmailSuffix/owa" -InternalDownloadHostName "download.$EmailSuffix" -ExternalDownloadHostName "download.$EmailSuffix"
Get-PowerShellVirtualDirectory -Server $ExServer | Set-PowerShellVirtualDirectory -ExternalUrl "https://$subdomain.$EmailSuffix/powershell" -InternalUrl "https://$subdomain.$EmailSuffix/powershell"
Get-OutlookAnywhere -Server $ExServer | Set-OutlookAnywhere -ExternalHostname "$subdomain.$EmailSuffix" -InternalHostname "$subdomain.$EmailSuffix" -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $true -DefaultAuthenticationMethod NTLM -SSLOffloading $false

# Enable Download Domains
Write-Log "Enabling Download Domains"
Set-OrganizationConfig -EnableDownloadDomains $true

# Set TCP KeepAliveTime in Exchange Server
New-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\TcpIp\Parameters" -Name "KeepAliveTime" -PropertyType DWORD -Value 1800000 -Force

Write-Log "Setting accepted domains and address policies"
# add accepted domain
New-AcceptedDomain -DomainName $EmailSuffix -DomainType Authoritative -Name $Domain -DomainController $DCHostName | Set-AcceptedDomain -MakeDefault $true
# edit default address policy
Get-EmailAddressPolicy -Identity "Default Policy" | Set-EmailAddressPolicy -EnabledEmailAddressTemplates "SMTP:@$EmailSuffix" -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "Default Policy" -DomainController $DCHostName
# add new address policy
New-EmailAddressPolicy -Name "FirstName LastName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffGroup,$EndPath" -EnabledEmailAddressTemplates "SMTP:%g.%s@$EmailSuffix" -Priority 1 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "FirstName LastName" -DomainController $DCHostName
New-EmailAddressPolicy -Name "UserName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffGroup,$EndPath" -EnabledEmailAddressTemplates "SMTP:%m@$EmailSuffix" -Priority 2 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "UserName" -DomainController $DCHostName
foreach ($GroupName in $MailGroups) {
    try {
        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName -ErrorAction Stop;
        Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
    } catch {
        Write-Log "WARNING: Could not enable $GroupName — $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    Write-Log "Closed Exchange session."
}
#====================================================================
