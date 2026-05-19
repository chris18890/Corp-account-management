#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

# Example Syntax:
# Powershell .\CreateLocalAdminGroups.ps1 -ComputerOU Servers

[CmdletBinding()]
Param(
    [string]$ComputerOU
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$EnvConfig = Get-EnvironmentConfig

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$Location = "OU=$($EnvConfig.OUs.DomainComputers),$EndPath"
$GroupsOU = "OU=$($EnvConfig.OUs.LocalAdminGroups),OU=$($EnvConfig.OUs.Administration),$EndPath"
$GroupCategory = "Security"
$GroupScope = "DomainLocal"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = $PSScriptRoot
# Set variables
$ScriptTitle = "$Domain Local Admin Group Maintenance Script"
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_Local_Admin_group_log-$(Get-Date -Format 'yyyyMMdd')"
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

if(-not($ComputerOU)) {
    $ComputerOU = Read-Host -Prompt "You must provide a DistinguishedName for the Computers OU - e.g. OU=Servers,$Location"
}
if ($ComputerOU -notmatch '^OU=') {
    $ComputerOU = "OU=$ComputerOU,$Location"
}
try {
    $null = Get-ADOrganizationalUnit -Identity $ComputerOU -Server $DCHostName -ErrorAction Stop
} catch {
    Write-LogFile -LogFile $LogFile -LogString "ERROR: ComputerOU '$ComputerOU' does not exist or is not accessible: $_" -ForegroundColor Red
    throw "ComputerOU '$ComputerOU' is not a valid OU on $DCHostName."
}
foreach ($computer in (Get-ADComputer -SearchBase $ComputerOU -Filter *)) {
    $CompName = $computer.Name
    $DomainLocalGroupName = "$($EnvConfig.Groups.TaskPrefix)Local_Admin_"+$CompName
    $GroupDesc = "User Group: Local Admin Users for $CompName"
    if (-not (Get-ADGroup -Filter "sAMAccountName -eq '$DomainLocalGroupName'" -Server $DCHostName)) {
        New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $DomainLocalGroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path $GroupsOU -GroupDescription $GroupDesc
    }
}

# Clean up groups for machines that have been removed
#Get a list of all Local Admin Groups and report on the number
$DomainLocalList = Get-ADGroup -Filter {name -like "$($EnvConfig.Groups.TaskPrefix)Local_Admin_*"} -SearchBase $GroupsOU -Server $DCHostName | Select-Object name
Write-LogFile -LogFile $LogFile -LogString "Total Local Admin groups before cleaning: $($DomainLocalList.count)"
foreach ($c in $DomainLocalList) {
    # Create server name variable to check for existence.
    $trimname = $c.name.replace("$($EnvConfig.Groups.TaskPrefix)Local_Admin_","")
    # See if the computer exists, and remove the group if not
    try {
        Get-ADComputer $trimname -Server $DCHostName | Out-Null
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $GroupName = "CN=$($c.Name),$GroupsOU"
        try {
            Set-ADObject -Identity $GroupName -ProtectedFromAccidentalDeletion $false -Server $DCHostName -ErrorAction Stop
            Remove-ADGroup -Identity $c.Name -Confirm:$false -Server $DCHostName -ErrorAction Stop
            Write-LogFile -LogFile $LogFile -LogString "Group $($c.Name) deleted"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR deleting group $($c.Name) : $_" -ForegroundColor Red
        }
    }
}

#Get a list of all Local Admin Groups and report on the number after cleaning
$DomainLocalList = Get-ADGroup -Filter {name -like "$($EnvConfig.Groups.TaskPrefix)Local_Admin_*"} -SearchBase $GroupsOU -Server $DCHostName | Select-Object name
Write-LogFile -LogFile $LogFile -LogString "Total Local Admin groups after cleaning: $($DomainLocalList.count)"
