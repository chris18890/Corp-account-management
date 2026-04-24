#Requires -RunAsAdministrator

# Execution Tier: Tier-0
# Mode: Standalone / No Shared Modules

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][ValidateSet('Azure','Local')][string]$Platform
    ,[string]$DomainSuffix
    ,[string]$IPNet1 = "10.71.104"
    ,[string]$Netmask = "21"
)

Set-StrictMode -Version Latest

If ($env:computername -ne "$Domain-EXCH") {
    $Machine = "$Domain-EXCH"
    rename-computer $Machine
}
If (!$IPNet1) {
    $IPNet1 = "10.71.104"
}
$IPNetSite = $IPNet1
$IPAddress = "$IPNetSite.7"
If (!$Netmask) {
    $Netmask = "21"
}
$GatewayAddress = "$IPNetSite.1"
$DNSAddress = @("$IPNetSite.2", "$IPNetSite.3")
if ($Platform -eq "Local") {
    Get-NetAdapter -Name Ethernet | Rename-NetAdapter -NewName LAN -PassThru
    New-NetIPAddress -IPAddress $IPAddress -InterfaceAlias "LAN" -DefaultGateway $GatewayAddress -AddressFamily IPv4 -PrefixLength $Netmask
    Set-DnsClientServerAddress -InterfaceAlias "LAN" -ServerAddresses $DNSAddress
}
Cscript c:\windows\system32\SCRegEdit.wsf /ar 0
Cscript c:\windows\system32\scregedit.wsf /cs 1
Cscript c:\windows\system32\scregedit.wsf /AU 4
Import-Module NetSecurity
Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-In)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-Out)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (SMB-In)" -enabled True
Set-NetFirewallRule -DisplayName "File and Printer Sharing (SMB-Out)" -enabled True
Set-NetFirewallRule -DisplayGroup "Windows Remote Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Volume Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Event Log Management" -enabled True
Set-NetFirewallRule -DisplayGroup "Remote Scheduled Tasks Management" -enabled True
$FeatureName = @("NET-Framework-45-Features", "NET-WCF-HTTP-Activation45", "NET-WCF-Pipe-Activation45", "Server-Media-Foundation", "RPC-over-HTTP-proxy", "RSAT-Clustering"
, "RSAT-Clustering-CmdInterface", "RSAT-Clustering-PowerShell", "WAS-Process-Model", "Web-Asp-Net45", "Web-Basic-Auth", "Web-IP-Security"
, "Web-Client-Auth", "Web-Digest-Auth", "Web-Dir-Browsing", "Web-Dyn-Compression", "Web-Http-Errors", "Web-Http-Logging"
, "Web-Http-Redirect", "Web-Http-Tracing", "Web-ISAPI-Ext", "Web-ISAPI-Filter", "Web-Metabase", "Web-Mgmt-Service", "Web-Net-Ext45"
, "Web-Request-Monitor", "Web-Server", "Web-Stat-Compression", "Web-Static-Content", "Web-Windows-Auth", "Web-WMI", "RSAT-ADDS", "rsat-ad-powershell")
$DesktopFeatureName = @("RSAT-Clustering-Mgmt", "Web-Mgmt-Console", "Windows-Identity-Foundation")
$FeatureToRemove = @("NET-WCF-MSMQ-Activation45", "MSMQ")
$regKey = "HKLM:\software\microsoft\windows nt\currentversion"
$Core = (Get-ItemProperty $regKey).InstallationType -eq "Server Core"
foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
        Write-Host $Feature "is already installed" -ForegroundColor Green
    } else {
        Write-Host "installing" $Feature
        install-windowsfeature -IncludeManagementTools $Feature
        Write-Host "installed" $Feature
    }
}
if (!$Core) {
    foreach ($Feature in $DesktopFeatureName){
        if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
            Write-Host $Feature "is already installed" -ForegroundColor Green
        } else {
            Write-Host "installing" $Feature
            install-windowsfeature -IncludeManagementTools $Feature
            Write-Host "installed" $Feature
        }
    }
}
foreach ($Feature in $FeatureToRemove){
    if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
        Remove-WindowsFeature $Feature
    }
}
if ((Get-CimInstance win32_computersystem).partofdomain -eq $false) {
    If (!$DomainSuffix) {
        $DomainSuffix = READ-HOST 'Enter a public FQDN- '
    }
    $DNSSuffix = "$Domain.$DomainSuffix"
    Add-Computer -DomainName "$DNSSuffix" -Restart
} else {
    if ((Get-CimInstance win32_computersystem).partofdomain -eq $true) {
        $EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
        $DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
        $ParentOU = "Domain Computers"
        $Location = "OU=$ParentOU,$EndPath"
        $ServerName = "$env:computername"
        $Member = "$env:username"
        $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
        $requiredGroups = @('Domain Admins')
        $groups = $requiredGroups | ForEach-Object {
            Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
        }
        if (-not $groups) {
            Write-Host "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        try {
            Get-ADComputer $ServerName -Server $DCHostName | Move-ADObject -TargetPath "ou=Servers,$Location" -Server $DCHostName
        } catch {
            $ex = $_.Exception
            Write-Host "ERROR: $($ex.Message)" -ForegroundColor Red
        }
        try {
            Add-ADGroupMember -Identity "Enterprise Admins" -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes 480) -Server $DCHostName
        } catch {
            $ex = $_.Exception
            Write-Host "ERROR: $($ex.Message)" -ForegroundColor Red
        }
        try {
            Add-ADGroupMember -Identity "Schema Admins" -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes 480) -Server $DCHostName
        } catch {
            $ex = $_.Exception
            Write-Host "ERROR: $($ex.Message)" -ForegroundColor Red
        }
        Restart-Computer
    }
}
