#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Domain
    ,[string]$DomainSuffix
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

$ServerName = "$env:computername"
$FeatureName = @("NET-Framework-45-Features", "NET-WCF-HTTP-Activation45", "NET-WCF-Pipe-Activation45", "Server-Media-Foundation", "RPC-over-HTTP-proxy", "RSAT-Clustering"
, "RSAT-Clustering-CmdInterface", "RSAT-Clustering-PowerShell", "WAS-Process-Model", "Web-Asp-Net45", "Web-Basic-Auth", "Web-IP-Security"
, "Web-Client-Auth", "Web-Digest-Auth", "Web-Dir-Browsing", "Web-Dyn-Compression", "Web-Http-Errors", "Web-Http-Logging"
, "Web-Http-Redirect", "Web-Http-Tracing", "Web-ISAPI-Ext", "Web-ISAPI-Filter", "Web-Metabase", "Web-Mgmt-Service", "Web-Net-Ext45"
, "Web-Request-Monitor", "Web-Server", "Web-Stat-Compression", "Web-Static-Content", "Web-Windows-Auth", "Web-WMI", "RSAT-ADDS", "rsat-ad-powershell")
$DesktopFeatureName = @("RSAT-Clustering-Mgmt", "Web-Mgmt-Console", "Windows-Identity-Foundation")
$FeatureToRemove = @("NET-WCF-MSMQ-Activation45", "MSMQ")
$regKey = "HKLM:\software\microsoft\windows nt\currentversion"
$Core = (Get-ItemProperty $regKey).InstallationType -eq "Server Core"
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain Exchange Prereq Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_exchange_prereq_log-$(Get-Date -Format 'yyyyMMdd')"
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
Write-LogFile -LogFile $LogFile -LogString "Note - Script must be run TWICE on this server - first to do the domain join then again to set up Exchange Prereqs"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

foreach ($Feature in $FeatureName){
    if (Get-WindowsFeature -Name $Feature | Where-Object InstallState -Eq Installed) {
        Write-LogFile -LogFile $LogFile -LogString "$Feature is already installed" -ForegroundColor Green
    } else {
        Write-LogFile -LogFile $LogFile -LogString "installing $Feature"
        install-windowsfeature -IncludeManagementTools $Feature
        Write-LogFile -LogFile $LogFile -LogString "installed $Feature"
    }
}
if (!$Core) {
    foreach ($Feature in $DesktopFeatureName){
        if (Get-WindowsFeature -Name $Feature | Where-Object InstallState -Eq Installed) {
            Write-LogFile -LogFile $LogFile -LogString "$Feature is already installed" -ForegroundColor Green
        } else {
            Write-LogFile -LogFile $LogFile -LogString "installing $Feature"
            install-windowsfeature -IncludeManagementTools $Feature
            Write-LogFile -LogFile $LogFile -LogString "installed $Feature"
        }
    }
}
foreach ($Feature in $FeatureToRemove){
    if (Get-WindowsFeature -Name $Feature | Where-Object InstallState -Eq Installed) {
        Write-LogFile -LogFile $LogFile -LogString "removing $Feature"
        Remove-WindowsFeature $Feature
        Write-LogFile -LogFile $LogFile -LogString "removed $Feature"
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
        $requiredGroups = @('Domain Admins', 'Enterprise Admins', 'Schema Admins')
        if (-not (Test-IsMemberOf -Sam $env:USERNAME -GroupNames $requiredGroups -DCHostName $DCHostName)) {
            Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        try {
            Get-ADComputer $ServerName -Server $DCHostName | Move-ADObject -TargetPath "ou=Servers,$Location" -Server $DCHostName
        } catch [Microsoft.ActiveDirectory.Management.ADException] {
            # Already moved - first run already moved it
            if ($_.Exception.Message -notmatch "already exists in target container") {
                throw
            }
        }
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Enterprise Admins" -Member $Member -TimeSpan $Env.Security.MaxElevationMinutes
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Schema Admins" -Member $Member -TimeSpan $Env.Security.MaxElevationMinutes
        Write-LogFile -LogFile $LogFile -LogString "Intalling IIS URL Rewrite"
        Start-Process -FilePath msiexec "\\$DNSSuffix\Store\Software\rewrite_amd64_en-US.msi" -Argumentlist "/i /Quiet" -wait
        Write-LogFile -LogFile $LogFile -LogString "Intalling Visual C++ Runtime"
        Start-Process -Filepath "\\$DNSSuffix\Store\Software\vcredist_x64_2013.exe" -Argumentlist "/Q" -wait
        Write-LogFile -LogFile $LogFile -LogString "Intalling Unified Communications Managed API"
        Start-Process -Filepath ".\UCMARedist\Setup.exe" -Argumentlist "/Q" -wait
        Write-LogFile -LogFile $LogFile -LogString "Launching Exchange Setup with the /PrepareAD switch"
        try {
            .\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"$Domain"
            if ($LASTEXITCODE -ne 0) { throw "Exchange PrepareAD failed - exit code $LASTEXITCODE." }
        } finally {
            try {
                Remove-ADGroupMember -Identity "Enterprise Admins" -Members $Member -Confirm:$False -Server $DCHostName
                Write-LogFile -LogFile $LogFile -LogString "Removed $Member from Enterprise Admins"
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                # Already not a member - first run already removed it
                if ($_.Exception.Message -notmatch "not a member") {
                    throw
                }
            }
            try {
                Remove-ADGroupMember -Identity "Schema Admins" -Members $Member -Confirm:$False -Server $DCHostName
                Write-LogFile -LogFile $LogFile -LogString "Removed $Member from Schema Admins"
            } catch [Microsoft.ActiveDirectory.Management.ADException] {
                # Already not a member - first run already removed it
                if ($_.Exception.Message -notmatch "not a member") {
                    throw
                }
            }
        }
        Restart-Computer
    }
}
