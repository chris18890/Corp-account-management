#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Execution Tier: Tier-0

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][ValidateSet('Azure','Local')][string]$Platform
    ,[string]$DomainSuffix
)

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$ServerName = "$env:computername"
$Netmask = "255.255.255.0"
$CDIR = "24"
$Router = "1"
$DNSServer2 = "2"
$DNSServer3 = "3"
$StartRange = "10"
$EndRange = "254"
$IPNet = @("10.71.104", "10.71.105")
$Site = @("Site1", "Site2")
$FeatureName = @("rsat-ad-powershell", "rsat-adds", "GPMC", "RSAT-DFS-Mgmt-Con", "rsat-dns-server", "rsat-dhcp", "WDS-AdminPack", "RSAT-ADCS", "UpdateServices-RSAT")
if ($Platform -eq "Local") {
    $FeatureName += @("routing", "dhcp")
    $IPAddress1 = $IPNet[0] + "." + $Router
    $IPAddress2 = $IPNet[0] + "." + $DNSServer2
    $IPAddress3 = $IPNet[0] + "." + $DNSServer3
    $ExternalInterface = "WAN"
    $InternalInterface = "LAN"
}
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain network setup Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_network_setup_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log ("=" * 80)
Write-Log "Processing commenced, running as user '$env:USERNAME'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log "Note - Script must be run TWICE on this server - first to do the domain join then again to set up DHCP"
Write-Log ("=" * 80)
Write-Log ""
Foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where InstallState -Eq Installed) {
        Write-Log "$Feature is already installed" -ForegroundColor Green
    } else {
        Write-Log "installing $Feature"
        install-windowsfeature -IncludeManagementTools $Feature
        Write-Log "installed $Feature"
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
        $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
        $SID500 = (Get-ADUser -Filter * -Server $DCHostName | Select-Object -Property SID,Name | Where-Object -Property SID -like "*-500").Name
        $requiredGroups = @("Domain Admins", "Enterprise Admins")
        $groups = $requiredGroups | ForEach-Object {
            Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
        }
        if (-not $groups) {
            Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        try {
            if ($Platform -eq "Local") {
                Set-DhcpServerv4Binding -BindingState $true -InterfaceAlias "$InternalInterface"
                Add-DhcpServerSecurityGroup -ComputerName $ServerName
                Restart-service dhcpserver
                Add-DhcpServerInDC "$ServerName" "$IPAddress1"
                try {
                    Add-LocalGroupMember -Group "DHCP Administrators" -Member "$Domain\ADM_Task_DHCP_Admins" -ErrorAction Stop
                } catch [Microsoft.PowerShell.Commands.MemberExistsException] {
                    Write-Log "$Domain\ADM_Task_DHCP_Admins is already a member of DHCP Administrators" -ForegroundColor Green
                }
                try {
                    Add-LocalGroupMember -Group "DHCP Users" -Member "$Domain\ADM_Task_DHCP_Users" -ErrorAction Stop
                } catch [Microsoft.PowerShell.Commands.MemberExistsException] {
                    Write-Log "$Domain\ADM_Task_DHCP_Users is already a member of DHCP Users" -ForegroundColor Green
                }
                Set-ItemProperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManager\Roles\12 -Name ConfigurationState -Value 2
            }
            Get-ADComputer -Identity $ServerName -Server $DCHostName | Move-ADObject -TargetPath "ou=Servers,$Location" -Server $DCHostName
            if ($Platform -eq "Local") {
                Set-DhcpServerv4OptionValue -Router "$IPAddress1"
                Set-DhcpServerv4OptionValue -OptionId 42 -value "$IPAddress2"
                Set-DhcpServerv4OptionValue -DnsServer "$IPAddress2", "$IPAddress3" -Force
                Set-DhcpServerv4OptionValue -DnsDomain "$DNSSuffix"
            }
            For ($i=0; $i -le $IPNet.Length -1; $i++){
                if ($Platform -eq "Local") {
                    $ScopeNetwork = $IPNet[$i] + ".0"
                    $ScopeName = "$ScopeNetwork/$CDIR"
                    $ScopeStart = $IPNet[$i] + ".$StartRange"
                    $ScopeEnd = $IPNet[$i] + ".$EndRange"
                    $ScopeNetmask = $Netmask
                    $ScopeDNS2 = $IPNet[$i] + ".$DNSServer2"
                    $ScopeDNS3 = $IPNet[$i] + ".$DNSServer3"
                    $ScopeRouter = $IPNet[$i] + ".$Router"
                }
                try {
                    Add-DnsServerPrimaryZone -NetworkID "$ScopeName" -ReplicationScope "Domain" -DynamicUpdate:Secure -ComputerName $DCHostName -ErrorAction Stop
                } catch [Microsoft.Management.Infrastructure.CimException] {
                    switch ($_.Exception.ErrorCode) {
                        9609 { # 'The specified object already exists'
                            Write-Log "Zone $ScopeName already exists" -ForegroundColor Green
                        }
                        default {
                            Write-Log "ERROR: An unexpected error occurred while attempting to create Zone $ScopeName `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                if ($Platform -eq "Local") {
                    try {
                        Add-DhcpServerv4Scope -Name $Site[$i] -StartRange "$ScopeStart" -EndRange "$ScopeEnd" -SubnetMask "$ScopeNetmask" -State Active -ErrorAction Stop
                    } catch [Microsoft.Management.Infrastructure.CimException] {
                        switch ($_.Exception.ErrorCode) {
                            20052 { # 'The specified object already exists'
                                Write-Log "Scope $Site[$i] already exists" -ForegroundColor Green
                            }
                            default {
                                Write-Log "ERROR: An unexpected error occurred while attempting to create Scope $Site[$i] `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                            }
                        }
                    }
                    Set-DhcpServerv4OptionValue -ScopeId "$ScopeNetwork" -DnsServer "$ScopeDNS2", "$ScopeDNS3" -Force -Router "$ScopeRouter"
                }
                try {
                    New-ADReplicationSite -Name $Site[$i] -Server $DCHostName
                } catch [Microsoft.ActiveDirectory.Management.ADException] {
                    switch ($_.Exception.ErrorCode) {
                        8305 { # 'The specified object already exists'
                            Write-Log "Site $Site[$i] already exists" -ForegroundColor Green
                        }
                        default {
                            Write-Log "ERROR: An unexpected error occurred while attempting to create Site $Site[$i] `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                try {
                    New-ADReplicationSubnet -Name "$ScopeName" -Site $Site[$i] -Server $DCHostName
                } catch [Microsoft.ActiveDirectory.Management.ADException] {
                    switch ($_.Exception.ErrorCode) {
                        8305 { # 'The specified object already exists'
                            Write-Log "Subnet $ScopeName already exists" -ForegroundColor Green
                        }
                        default {
                            Write-Log "ERROR: An unexpected error occurred while attempting to create Subnet $ScopeName `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -SitesIncluded @{Add=$Site[$i]} -Server $DCHostName
            }
            Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -Cost "10" -ReplicationFrequencyInMinutes "15" -Server $DCHostName
            Move-ADDirectoryServer -Identity $(($DCHostName -split '\.')[0]) -Site $Site[0] -Server $DCHostName
        } finally {
            Remove-ADGroupMember -Identity "Enterprise Admins" -Members $SID500 -Confirm:$False -Server $DCHostName
        }
        if ($Platform -eq "Local") {
            Install-RemoteAccess -VpnType RoutingOnly -Legacy
            cmd.exe /c "netsh routing ip nat install"
            cmd.exe /c "netsh routing ip nat add interface $ExternalInterface"
            cmd.exe /c "netsh routing ip nat set interface $ExternalInterface mode=full"
            cmd.exe /c "netsh routing ip nat add interface $InternalInterface"
        }
    }
}
