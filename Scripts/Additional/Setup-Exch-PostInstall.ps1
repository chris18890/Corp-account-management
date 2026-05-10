#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$SubDomain
    , [string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$Domain = "$env:userdomain"
if (!$SubDomain) {
    $SubDomain = $Env.Exchange.Subdomain #enter subdomain
}
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$StaffOU = $Env.OUs.Staff
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$MailGroups = (Get-ADGroup -Filter * -Server $DCHostName -searchbase "OU=$($Env.OUs.Groups),$EndPath" -Properties *).Name
$ExServer = "$Domain-Exch.$DNSSuffix" #Remote Exchange PS session
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain Exchange Post-Install Config Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_exchange_postinstall_config_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

$requiredGroups = @('Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Organization Management" -Member "Domain Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Organization Management" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Server Management" -Member "$($Env.Groups.TaskPrefix)Server_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Recipient Management" -Member "$($Env.Groups.TaskPrefix)HiPriv_Account_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Recipient Management" -Member "$($Env.Groups.TaskPrefix)HiPriv_Group_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Recipient Management" -Member "$($Env.Groups.TaskPrefix)Standard_Account_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Recipient Management" -Member "$($Env.Groups.TaskPrefix)Standard_Group_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Recipient Management" -Member "$($Env.Groups.TaskPrefix)SER_Account_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Help Desk" -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"

#====================================================================
# Get user credentials for server connectivity (Non-MFA)
try {
    $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
} catch {
    $ErrorMsg = $_.Exception.Message
    Write-LogFile -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
    Read-Host -Prompt "Press Enter to exit"
    Exit
}
#Connect to remote Exchange PowerShell
Write-LogFile -LogFile $LogFile -LogString "Connecting to remote Exchange PowerShell session... "
try {
    $so = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos -AllowRedirection -SessionOption $so
    Write-LogFile -LogFile $LogFile -LogString "connected."
    Write-LogFile -LogFile $LogFile -LogString "Importing Exchange session... "
    Import-PSSession -Session $ExSession -DisableNameChecking -ErrorAction Stop -AllowClobber > $null
    Write-LogFile -LogFile $LogFile -LogString "done."
} catch {
    $e = $_.Exception
    Write-LogFile -LogFile $LogFile -LogString $e
    $line = $_.InvocationInfo.ScriptLineNumber
    Write-LogFile -LogFile $LogFile -LogString $line
    $msg = $e.Message
    Write-LogFile -LogFile $LogFile -LogString $msg
    $Action = "Error Importing Exchange Session"
    Write-LogFile -LogFile $LogFile -LogString $Action
    Write-LogFile -LogFile $LogFile -LogString "failed."
    Write-LogFile -LogFile $LogFile -LogString "ERROR: $_" -ForegroundColor Red
}
if (!$ExSession) {
    Write-LogFile -LogFile $LogFile -LogString "Exchange session not connected Stopping Script"
    Exit
}

if (-not (Get-WebBinding -Name "Default Web Site" -HostHeader "$subdomain.$EmailSuffix" -ErrorAction SilentlyContinue)) {
    New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 80 -HostHeader "$subdomain.$EmailSuffix"
} else {
    Write-LogFile -LogFile $LogFile -LogString "Web binding for '$subdomain.$EmailSuffix' already exists" -ForegroundColor Green
}

Write-LogFile -LogFile $LogFile -LogString "Setting Virtual Directories"
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
Write-LogFile -LogFile $LogFile -LogString "Enabling Download Domains"
Set-OrganizationConfig -EnableDownloadDomains $true

# Set TCP KeepAliveTime in Exchange Server
New-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\TcpIp\Parameters" -Name "KeepAliveTime" -PropertyType DWORD -Value 1800000 -Force

Write-LogFile -LogFile $LogFile -LogString "Setting accepted domains and address policies"
# add accepted domain
New-AcceptedDomain -DomainName $EmailSuffix -DomainType Authoritative -Name $Domain -DomainController $DCHostName | Set-AcceptedDomain -MakeDefault $true
# edit default address policy
Get-EmailAddressPolicy -Identity "Default Policy" | Set-EmailAddressPolicy -EnabledEmailAddressTemplates "SMTP:@$EmailSuffix" -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "Default Policy" -DomainController $DCHostName
# add new address policy
New-EmailAddressPolicy -Name "FirstName LastName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffOU,$EndPath" -EnabledEmailAddressTemplates "SMTP:%g.%s@$EmailSuffix" -Priority 1 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "FirstName LastName" -DomainController $DCHostName
New-EmailAddressPolicy -Name "UserName" -RecipientFilter "(RecipientType -eq 'UserMailbox' -or RecipientType -eq 'MailUser')" -RecipientContainer "OU=$StaffOU,$EndPath" -EnabledEmailAddressTemplates "SMTP:%m@$EmailSuffix" -Priority 2 -DomainController $DCHostName
Update-EmailAddressPolicy -Identity "UserName" -DomainController $DCHostName
foreach ($GroupName in $MailGroups) {
    try {
        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName -ErrorAction Stop;
        Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
    Remove-PsSession $ExSession
    $Cred.Password.Dispose()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-LogFile -LogFile $LogFile -LogString "Closed Exchange session."
}
#====================================================================
