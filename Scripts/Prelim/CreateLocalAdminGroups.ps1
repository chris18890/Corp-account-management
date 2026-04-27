#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

# Example Syntax:
# Powershell .\CreateLocalAdminGroups.ps1 -ComputerOU Servers

[CmdletBinding()]
Param(
    [string]$ComputerOU
)

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GroupsOU = "OU=Local_Admin_Groups,OU=Administration,$EndPath"
$GroupCategory = "Security"
$GroupScope = "DomainLocal"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain Local Admin Group Maintenance Script"
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_Local_Admin_group_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log ("=" * 80)
    Write-Log "Log file is '$LogFile'"
    Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log "DC being used is '$DCHostName'"
    Write-Log "Script path is '$ScriptPath'"
    Write-Log "$ScriptTitle"
    Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log ("=" * 80)
    Write-Log ""
}

if(-not($ComputerOU)) {
    $ComputerOU = Read-Host -Prompt "You must provide a DistinguishedName for the Computers OU - e.g. OU=Servers,$Location"
}
if ($ComputerOU -notmatch '^OU=') {
    $ComputerOU = "OU=$ComputerOU,$Location"
}

foreach ($computer in (Get-ADComputer -SearchBase $ComputerOU -Filter *)) {
    $CompName = $computer.Name
    $DomainLocalGroupName = "ADM_Task_Local_Admin_"+$CompName
    $GroupDesc = "User Group: Local Admin Users for $CompName"
    if (-not (Get-ADGroup -Filter "sAMAccountName -eq '$DomainLocalGroupName'" -Server $DCHostName)) {
        New-DomainGroup -GroupName $DomainLocalGroupName -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path $GroupsOU -GroupDescription $GroupDesc
    }
}

# Clean up groups for machines that have been removed
#Get a list of all Local Admin Groups and report on the number
$DomainLocalList = Get-ADGroup -Filter {name -like "ADM_Task_Local_Admin_*"} -SearchBase $GroupsOU -Server $DCHostName | select name
Write-Log "Total Local Admin groups before cleaning: $($DomainLocalList.count)"
foreach ($c in $DomainLocalList) {
    #Create server name variable to check for existence.
    $trimname = $c.name.replace("ADM_Task_Local_Admin_","")
    #See if the computer exists, and remove the group if not
    try {
        Get-ADComputer $trimname -Server $DCHostName
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $GroupName = "CN=$($c.Name),$GroupsOU"
        Set-ADObject -Identity $GroupName -protectedFromAccidentalDeletion $False -Server $DCHostName
        Remove-AdGroup $c.Name -Confirm:$false -Server $DCHostName
        Write-Log "Group $($c.Name) deleted"
    }
}

#Get a list of all Local Admin Groups and report on the number after cleaning
$DomainLocalList = Get-ADGroup -Filter {name -like "ADM_Task_Local_Admin_*"} -SearchBase $GroupsOU -Server $DCHostName | select name
Write-Log "Total Local Admin groups after cleaning: $($DomainLocalList.count)"
