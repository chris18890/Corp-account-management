#Requires -RunAsAdministrator

# Execution Tier: Tier-0

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[Parameter(Mandatory)][ValidateSet('Azure','Local')][string]$Platform
    ,[string]$DomainSuffix
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$ServerName = "$env:computername"
$Netmask    = $Env.Network.SiteNetmask
$CDIR       = $Env.Network.SiteCidr
$Router     = $Env.Network.HostOffsetsLocal.RTR
$DNSServer2 = $Env.Network.HostOffsetsLocal.DC1
$DNSServer3 = $Env.Network.HostOffsetsLocal.DC2
$StartRange = $Env.Network.DhcpStart
$EndRange   = $Env.Network.DhcpEnd
$IPNet      = $Env.Network.SiteSubnets
$Site       = $Env.Network.SiteNames
$FeatureName = @("rsat-ad-powershell", "rsat-adds", "GPMC", "RSAT-DFS-Mgmt-Con", "rsat-dns-server", "rsat-dhcp", "WDS-AdminPack", "RSAT-ADCS", "UpdateServices-RSAT")
if ($Platform -eq "Local") {
    $FeatureName += @("routing", "dhcp")
    $IPAddress1 = $IPNet[0] + "." + $Router
    $IPAddress2 = $IPNet[0] + "." + $DNSServer2
    $IPAddress3 = $IPNet[0] + "." + $DNSServer3
    $ExternalInterface = $Env.Network.Interfaces.External
    $InternalInterface = $Env.Network.Interfaces.Internal
}
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain network setup Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_network_setup_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Note - Script must be run TWICE on this server - first to do the domain join then again to set up DHCP"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

Foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where-Object InstallState -Eq Installed) {
        Write-LogFile -LogFile $LogFile -LogString "$Feature is already installed" -ForegroundColor Green
    } else {
        Write-LogFile -LogFile $LogFile -LogString "installing $Feature"
        install-windowsfeature -IncludeManagementTools $Feature
        Write-LogFile -LogFile $LogFile -LogString "installed $Feature"
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
        $Location = "OU=$($Env.OUs.DomainComputers),$EndPath"
        $Member = "$env:username"
        $DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
        $SID500 = (Get-ADUser -Identity ('{0}-500' -f (Get-ADDomain).DomainSID) -Server $DCHostName).SamAccountName
        $requiredGroups = @("Domain Admins", "Enterprise Admins")
        $groups = $requiredGroups | ForEach-Object {
            Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $Member
        }
        if (-not $groups) {
            Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        try {
            if ($Platform -eq "Local") {
                Set-DhcpServerv4Binding -BindingState $true -InterfaceAlias "$InternalInterface"
                Add-DhcpServerSecurityGroup -ComputerName $ServerName
                Restart-service dhcpserver
                Add-DhcpServerInDC "$ServerName" "$IPAddress1"
                try {
                    Add-LocalGroupMember -Group "DHCP Administrators" -Member "$Domain\$($Env.Groups.TaskPrefix)DHCP_Admins" -ErrorAction Stop
                } catch [Microsoft.PowerShell.Commands.MemberExistsException] {
                    Write-LogFile -LogFile $LogFile -LogString "$Domain\$($Env.Groups.TaskPrefix)DHCP_Admins is already a member of DHCP Administrators" -ForegroundColor Green
                }
                try {
                    Add-LocalGroupMember -Group "DHCP Users" -Member "$Domain\$($Env.Groups.TaskPrefix)DHCP_Users" -ErrorAction Stop
                } catch [Microsoft.PowerShell.Commands.MemberExistsException] {
                    Write-LogFile -LogFile $LogFile -LogString "$Domain\$($Env.Groups.TaskPrefix)DHCP_Users is already a member of DHCP Users" -ForegroundColor Green
                }
                Set-ItemProperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ServerManager\Roles\12 -Name ConfigurationState -Value 2
            }
            Get-ADComputer -Identity $ServerName -Server $DCHostName | Move-ADObject -TargetPath "ou=$($Env.OUs.Servers),$Location" -Server $DCHostName
            if ($Platform -eq "Local") {
                Set-DhcpServerv4OptionValue -Router "$IPAddress1"
                Set-DhcpServerv4OptionValue -OptionId 42 -value "$IPAddress2"
                Set-DhcpServerv4OptionValue -DnsServer "$IPAddress2", "$IPAddress3" -Force
                Set-DhcpServerv4OptionValue -DnsDomain "$DNSSuffix"
            }
            For ($i=0; $i -le $IPNet.Length -1; $i++){
                $ScopeNetwork = $IPNet[$i] + ".0"
                $ScopeName = "$ScopeNetwork/$CDIR"
                if ($Platform -eq "Local") {
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
                            Write-LogFile -LogFile $LogFile -LogString "Zone $ScopeName already exists" -ForegroundColor Green
                        }
                        default {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR: An unexpected error occurred while attempting to create Zone $ScopeName `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                if ($Platform -eq "Local") {
                    try {
                        Add-DhcpServerv4Scope -Name $Site[$i] -StartRange "$ScopeStart" -EndRange "$ScopeEnd" -SubnetMask "$ScopeNetmask" -State Active -ErrorAction Stop
                    } catch [Microsoft.Management.Infrastructure.CimException] {
                        switch ($_.Exception.ErrorCode) {
                            20052 { # 'The specified object already exists'
                                Write-LogFile -LogFile $LogFile -LogString "Scope $($Site[$i]) already exists" -ForegroundColor Green
                            }
                            default {
                                Write-LogFile -LogFile $LogFile -LogString "ERROR: An unexpected error occurred while attempting to create Scope $($Site[$i]) `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
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
                            Write-LogFile -LogFile $LogFile -LogString "Site $($Site[$i]) already exists" -ForegroundColor Green
                        }
                        default {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR: An unexpected error occurred while attempting to create Site $($Site[$i]) `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                try {
                    New-ADReplicationSubnet -Name "$ScopeName" -Site $Site[$i] -Server $DCHostName
                } catch [Microsoft.ActiveDirectory.Management.ADException] {
                    switch ($_.Exception.ErrorCode) {
                        8305 { # 'The specified object already exists'
                            Write-LogFile -LogFile $LogFile -LogString "Subnet $ScopeName already exists" -ForegroundColor Green
                        }
                        default {
                            Write-LogFile -LogFile $LogFile -LogString "ERROR: An unexpected error occurred while attempting to create Subnet $ScopeName `n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                        }
                    }
                }
                Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -SitesIncluded @{Add=$Site[$i]} -Server $DCHostName
            }
            Set-ADReplicationSiteLink -Identity "DEFAULTIPSITELINK" -Cost "10" -ReplicationFrequencyInMinutes "15" -Server $DCHostName
            try {
                Move-ADDirectoryServer -Identity $(($DCHostName -split '\.')[0]) -Site $Site[0] -Server $DCHostName
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                # Already moved - first run already moved it
                if ($_.Exception.Message -notmatch "already exists in target container") {
                    throw
                }
            }
        } finally {
            try {
                Remove-ADGroupMember -Identity "Enterprise Admins" -Members $SID500 -Confirm:$False -Server $DCHostName -ErrorAction Stop
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                # Already not a member - first run already removed it
                if ($_.Exception.Message -notmatch "not a member") {
                    throw
                }
            }
        }
        if ($Platform -eq "Local") {
            Install-RemoteAccess -VpnType RoutingOnly -Legacy
            Get-NetAdapter -Name "Ethernet 2" | Rename-NetAdapter -NewName $ExternalInterface -PassThru
            cmd.exe /c "netsh routing ip nat install"
            cmd.exe /c "netsh routing ip nat add interface $ExternalInterface"
            cmd.exe /c "netsh routing ip nat set interface $ExternalInterface mode=full"
            cmd.exe /c "netsh routing ip nat add interface $InternalInterface"
        }
    }
}
