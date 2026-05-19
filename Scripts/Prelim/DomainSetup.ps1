#Requires -Modules ActiveDirectory, GroupPolicy
#Requires -RunAsAdministrator

# Execution Tier: Tier-0

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    ,[Parameter(Mandatory)][string]$Drive
    ,[string]$LogFile
    ,[switch]$SkipGPOs
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$rootdse = Get-ADRootDSE
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = $Env.OUs.DomainComputers
$Location = "OU=$ParentOU,$EndPath"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$SID500 = (Get-ADUser -Identity ('{0}-500' -f (Get-ADDomain).DomainSID) -Server $DCHostName).SamAccountName
$GPOLocation = Join-Path $PSScriptRoot "GPOs"
$ImportedGPOs = @()
$UserScriptsLocation = Join-Path (Split-Path $PSScriptRoot -Parent) "Users"
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain Bootstrap Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_domain_setup_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
#====================================================================

if (-not $SkipGPOs) {
    if (-not (Test-Path $GPOLocation)) {
        $msg = "GPO backup directory not found at '$GPOLocation'. Populate the directory, or rerun with -SkipGPOs to bypass GPO imports."
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        throw $msg
    }
}

#====================================================================
# Drive where all the folders will be created
#====================================================================
$Drive = $Drive.TrimEnd(':') + ':'
$RootShare = $Env.Shares.Root
#====================================================================

#====================================================================
# Group Variables
#====================================================================
$GroupsOU = $Env.OUs.Groups
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = $Env.Groups.Staff
$StaffOU = $Env.OUs.Staff
$AdministrationOU = $Env.OUs.Administration
$ITGroup = $Env.Groups.IT
$ITAdminGroup = $Env.Groups.ITAdmin
$O365LicenseGroup = $Env.Groups.O365License
$DNSOperatorsGroup = "$($Env.Groups.TaskPrefix)DNS_Operators"
$DNSReadOnlyGroup = "$($Env.Groups.TaskPrefix)DNS_ReadOnly"
$HiPrivAccountAdminGroup = "$($Env.Groups.TaskPrefix)HiPriv_Account_Admins"
$HiPrivGroupAdminGroup = "$($Env.Groups.TaskPrefix)HiPriv_Group_Admins"
$InstallerGroup = "$($Env.Groups.TaskPrefix)Installers"
$LocalAdminGroupAdminGroup = "$($Env.Groups.TaskPrefix)Local_Admin_Group_Admins"
$UserPasswordDelegationGroup = "$($Env.Groups.TaskPrefix)Password_Admins"
$SERAccessAdminGroup = "$($Env.Groups.TaskPrefix)SER_Access_Admins"
$SERAccountAdminGroup = "$($Env.Groups.TaskPrefix)SER_Account_Admins"
$ServiceAccountAdminGroup = "$($Env.Groups.TaskPrefix)Service_Account_Admins"
$StandardAccountAdminGroup = "$($Env.Groups.TaskPrefix)Standard_Account_Admins"
$StandardGroupAdminGroup = "$($Env.Groups.TaskPrefix)Standard_Group_Admins"
$EquipmentAccountsOU = "$($Env.OUs.EquipmentMailboxAccounts),OU=$AdministrationOU"
$RoomAccountsOU = "$($Env.OUs.RoomMailboxAccounts),OU=$AdministrationOU"
$SharedAccountsOU = "$($Env.OUs.SharedMailboxAccounts),OU=$AdministrationOU"
$EquipmentGroupsOU = "$($Env.OUs.EquipmentMailboxAccess),OU=$GroupsOU"
$RoomGroupsOU = "$($Env.OUs.RoomMailboxAccess),OU=$GroupsOU"
$SharedGroupsOU = "$($Env.OUs.SharedMailboxAccess),OU=$GroupsOU"
#====================================================================

Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "
$requiredGroups = @('Domain Admins')
if (-not (Test-IsMemberOf -Sam $env:USERNAME -GroupNames $requiredGroups -DCHostName $DCHostName)) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Add additional UPN suffix
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Adding $EmailSuffix as an additional UPN suffix"
$forest = Get-ADForest
if ($forest.UPNSuffixes -notcontains $EmailSuffix) {
    Get-ADForest | Set-ADForest -UPNSuffixes @{add = $EmailSuffix} -Server $DCHostName
}
#====================================================================

#====================================================================
# Prevent standard users from creating computer accounts
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Preventing standard users from creating computer accounts"
Set-ADDomain (Get-ADDomain).distinguishedname -Replace @{"ms-ds-MachineAccountQuota"="0"} -Server $DCHostName
#====================================================================

#====================================================================
# Enable PAM feature to use temporal group membership
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Enabling PAM feature to use temporal group membership"
Enable-ADOptionalFeature "Privileged Access Management Feature" -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False -Server $DCHostName
#====================================================================

#====================================================================
# Enable AD Recycle Bin
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Enabling AD Recycle Bin"
Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False -Server $DCHostName
#====================================================================

#====================================================================
# Protect Domain Controllers OU
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Protecting Domain Controllers OU"
Set-ADOrganizationalUnit -Identity "OU=Domain Controllers,$EndPath" -ProtectedFromAccidentalDeletion $true -Server $DCHostName
#====================================================================

Write-LogFile -LogFile $LogFile -LogString "Creating user OUs & Groups"
#====================================================================
# Staff OU & group creation
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $StaffOU -Path $EndPath -OUDescription "Top level OU for Staff User objects"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $GroupsOU -Path $EndPath -OUDescription "Top level OU for Group objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $StaffGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Org-wide group for all staff users"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $ITGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Departmental group holding all IT accounts"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $O365LicenseGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Used to assign Office365 licenses"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $AdministrationOU -Path $EndPath -OUDescription "Top level OU for IT Admin User & Group objects"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.EquipmentMailboxAccess)" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to equipment mailboxes"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.RoomMailboxAccess)" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to room mailboxes"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.SharedMailboxAccess)" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to shared mailboxes"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.EquipmentMailboxAccounts)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are equipment mailbox recipient types"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.HiPrivGroups)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that control Hi-Priv access"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.HiPrivAccounts)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are Hi-Priv accounts"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.LocalAdminGroups)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that give local admin on individual devices"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.RoomMailboxAccounts)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are room mailbox recipient types"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.SharedMailboxAccounts)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are shared mailbox recipient types"
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName "$($Env.OUs.ServiceAccounts)" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are service accounts"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $ITAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Group holding all IT Admin accounts"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Administration_OU_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Computer_OU_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_GPO_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create & link GPOs"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Group_OU_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Site_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Subnet_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Subnet objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_Transport_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Transport objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)AD_User_OU_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)Desktop_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local Admin on all computers in the Desktop, Laptop, & VM OUs"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)DFS_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control NTFS permissions on all DFS folders & have access to the DFS management console"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)DHCP_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Administrators"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)DHCP_Users" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Users"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $DNSOperatorsGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Edit access to DNS zones"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $DNSReadOnlyGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have read only access to the DNS service"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $HiPrivAccountAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $($Env.OUs.HiPrivAccounts) OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $HiPrivGroupAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $($Env.OUs.HiPrivGroups) OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $InstallerGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members have permission to join & move computer objects in $ParentOU & Sub OUs"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $LocalAdminGroupAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $($Env.OUs.LocalAdminGroups) OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $UserPasswordDelegationGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can reset passwords of users in the $StaffOU OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)Server_Admins" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local admin on all computers in $ParentOU & Sub OUs via GPO"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $SERAccessAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can edit the membership of sh_, eq_, & ro_ groups"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $SERAccountAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit sh_, eq_, & ro_ accounts"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $ServiceAccountAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $($Env.OUs.ServiceAccounts) OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $StandardAccountAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $StaffOU OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $StandardGroupAdminGroup -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $GroupsOU OU"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.TaskPrefix)WDS_Deploy_Servers" -GroupCategory $GroupCategory -GroupScope "DomainLocal" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Servers folder"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins" -GroupCategory $GroupCategory -GroupScope "Global" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Level 1 admins - desktop"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins" -GroupCategory $GroupCategory -GroupScope "Global" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Level 2 admins - junior server (Tier 1)"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins" -GroupCategory $GroupCategory -GroupScope "Global" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Level 3 admins - senior server (Tier 1)"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.RolePrefix)Tier0_Level_2_Admins" -GroupCategory $GroupCategory -GroupScope "Global" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Level 2 admins - junior server (Tier 0)"
New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName "$($Env.Groups.RolePrefix)Tier0_Level_3_Admins" -GroupCategory $GroupCategory -GroupScope "Global" -O365 "N" -HiddenFromAddressListsEnabled $true -Path "OU=$($Env.OUs.HiPrivGroups),OU=$AdministrationOU,$EndPath" -GroupDescription "Level 3 admins - senior server (Tier 0)"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StaffGroup -Member $ITGroup
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Remote Desktop Users" -Member "$($Env.Groups.RolePrefix)Tier0_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Server Operators" -Member "$($Env.Groups.RolePrefix)Tier0_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Administrators" -Member "$($Env.Groups.RolePrefix)Tier0_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Remote Desktop Users" -Member "$($Env.Groups.RolePrefix)Tier0_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "Server Operators" -Member "$($Env.Groups.RolePrefix)Tier0_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $ITAdminGroup -Member $SID500
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Administration_OU_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Computer_OU_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_GPO_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Group_OU_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Site_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Subnet_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_Transport_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)AD_User_OU_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)Desktop_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)Desktop_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)Desktop_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)DHCP_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)DHCP_Users" -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)DHCP_Users" -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)DFS_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $DNSOperatorsGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $DNSOperatorsGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $DNSReadOnlyGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $HiPrivAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $HiPrivGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $InstallerGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $InstallerGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $InstallerGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $LocalAdminGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $LocalAdminGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $LocalAdminGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $UserPasswordDelegationGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $UserPasswordDelegationGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $UserPasswordDelegationGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)Server_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)Server_Admins" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $SERAccessAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_1_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $SERAccessAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $SERAccessAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $SERAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $SERAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StandardAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StandardAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StandardGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StandardGroupAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $ServiceAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $ServiceAccountAdminGroup -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)WDS_Deploy_Servers" -Member "$($Env.Groups.RolePrefix)Tier1_Level_2_Admins"
Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group "$($Env.Groups.TaskPrefix)WDS_Deploy_Servers" -Member "$($Env.Groups.RolePrefix)Tier1_Level_3_Admins"
#====================================================================

Write-LogFile -LogFile $LogFile -LogString "Creating shares"
#====================================================================
# Create main Share
#====================================================================
$ShareName = $RootShare
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory
        } else {
            Write-LogFile -LogFile $LogFile -LogString "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
        Write-LogFile -LogFile $LogFile -LogString "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-LogFile -LogFile $LogFile -LogString "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "$($Env.Groups.TaskPrefix)DFS_Admins"
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "$($Env.Groups.TaskPrefix)DFS_Admins" -Force
} else {
    Write-LogFile -LogFile $LogFile -LogString "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
# Create Profiles Share
#====================================================================
$ShareName = $Env.Shares.Profiles
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory -force
            $Acl = Get-Acl "$Drive\$ShareName"
            $isProtected = $true
            $preserveInheritance = $false
            $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StaffGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StandardAccountAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("$($Env.Groups.TaskPrefix)DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl "$Drive\$ShareName" $Acl
        } else {
            Write-LogFile -LogFile $LogFile -LogString "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
        Write-LogFile -LogFile $LogFile -LogString "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-LogFile -LogFile $LogFile -LogString "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "$($Env.Groups.TaskPrefix)DFS_Admins"
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "$($Env.Groups.TaskPrefix)DFS_Admins" -Force
} else {
    Write-LogFile -LogFile $LogFile -LogString "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
# Create IT Group Share
#====================================================================
$ShareName = $ITGroup
if (!(TEST-PATH "\\$DNSSuffix\$RootShare\$ShareName")) {
    New-Item "\\$DNSSuffix\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$DNSSuffix\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ShareName,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("$($Env.Groups.TaskPrefix)DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
    Robocopy $UserScriptsLocation "\\$DNSSuffix\$RootShare\$ShareName\User_Scripts" /e
} else {
    Write-LogFile -LogFile $LogFile -LogString "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
}
$ShareName = $Env.Shares.Software
if (!(TEST-PATH "\\$DNSSuffix\$RootShare\$ShareName")) {
    New-Item "\\$DNSSuffix\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$DNSSuffix\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Authenticated Users","ReadAndExecute","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("$($Env.Groups.TaskPrefix)DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
} else {
    Write-LogFile -LogFile $LogFile -LogString "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

Write-LogFile -LogFile $LogFile -LogString "Creating computer OUs"
#====================================================================
# Create default computers OU
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $ParentOU -Path $EndPath -OUDescription "Top level OU for Computer objects"
#====================================================================

#====================================================================
# Create OU for servers
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $Env.OUs.Servers -Path $Location
#====================================================================

#====================================================================
# Create OU for desktops
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $Env.OUs.Desktops -Path $Location
#====================================================================

#====================================================================
# Create OU for laptops
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $Env.OUs.Laptops -Path $Location
#====================================================================

#====================================================================
# Create OU for VMs
#====================================================================
New-ADOU -LogFile $LogFile -DCHostName $DCHostName -OUName $Env.OUs.VMs -Path $Location
#====================================================================

#====================================================================
# Redirect default computer location & delegate permissions
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Creating Permission delegations"

$GuidMap           = Get-ADSchemaGuidMap     -Server $DCHostName
$ExtendedRightsMap = Get-ADExtendedRightsMap -Server $DCHostName
$DelegationCommon = @{
    BaseDN            = $EndPath
    GuidMap           = $GuidMap
    ExtendedRightsMap = $ExtendedRightsMap
}

redircmp $Location
Grant-ComputerJoinDelegation -AdminGroupName $InstallerGroup -TargetOU $ParentOU @DelegationCommon
Grant-OUDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Computer_OU_Admins" -TargetOU $ParentOU @DelegationCommon
Grant-PasswordResetDelegation -AdminGroupName $UserPasswordDelegationGroup -TargetOU $StaffOU @DelegationCommon
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $EquipmentGroupsOU @DelegationCommon
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $RoomGroupsOU @DelegationCommon
Grant-GroupMembershipEditDelegation -AdminGroupName $SERAccessAdminGroup -TargetOU $SharedGroupsOU @DelegationCommon
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentAccountsOU @DelegationCommon
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentGroupsOU @DelegationCommon
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomAccountsOU @DelegationCommon
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomGroupsOU @DelegationCommon
Grant-UserDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedAccountsOU @DelegationCommon
Grant-GroupDelegation -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedGroupsOU @DelegationCommon
Grant-UserDelegation -AdminGroupName $HiPrivAccountAdminGroup -TargetOU "$($Env.OUs.HiPrivAccounts),OU=$AdministrationOU" @DelegationCommon
Grant-GroupDelegation -AdminGroupName $HiPrivGroupAdminGroup -TargetOU "$($Env.OUs.HiPrivGroups),OU=$AdministrationOU" @DelegationCommon
Grant-UserDelegation -AdminGroupName $StandardAccountAdminGroup -TargetOU $StaffOU @DelegationCommon
Grant-GroupDelegation -AdminGroupName $StandardGroupAdminGroup -TargetOU $GroupsOU @DelegationCommon
Grant-UserDelegation -AdminGroupName $ServiceAccountAdminGroup -TargetOU "$($Env.OUs.ServiceAccounts),OU=$AdministrationOU" @DelegationCommon
Grant-GroupDelegation -AdminGroupName $LocalAdminGroupAdminGroup -TargetOU "$($Env.OUs.LocalAdminGroups),OU=$AdministrationOU" @DelegationCommon
Grant-OUDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Administration_OU_Admins" -TargetOU $AdministrationOU @DelegationCommon
Grant-OUDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Group_OU_Admins" -TargetOU $GroupsOU @DelegationCommon
Grant-OUDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_User_OU_Admins" -TargetOU $StaffOU @DelegationCommon
Grant-ADObjectPermissionDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Site_Admins" -TargetDN "CN=Sites,CN=Configuration" @DelegationCommon
Grant-ADObjectPermissionDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Subnet_Admins" -TargetDN "CN=Subnets,CN=Sites,CN=Configuration" @DelegationCommon
Grant-ADObjectPermissionDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_Transport_Admins" -TargetDN "CN=Inter-Site Transports,CN=Sites,CN=Configuration" @DelegationCommon
Grant-GPOPermissionDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_GPO_Admins" @DelegationCommon
Grant-GPOCreationDelegation -AdminGroupName "$($Env.Groups.TaskPrefix)AD_GPO_Admins" -TargetDN "CN=Policies,CN=System" @DelegationCommon
Grant-DNSOperatorsPermissionDelegation -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,CN=System" @DelegationCommon
Grant-DNSOperatorsPermissionDelegation -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones" @DelegationCommon
Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,CN=System" @DelegationCommon
Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones" @DelegationCommon
$DNSZones = Get-ADObject -Filter * -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$EndPath" -SearchScope 1
foreach ($DNSZone in $DNSZones) {
    $DNSZoneName = $DNSZone.Name
    Grant-DNSReadOnlyPermissionDelegation -AdminGroupName $DNSReadOnlyGroup -TargetDN "DC=$DNSZoneName,CN=MicrosoftDNS,DC=DomainDnsZones" @DelegationCommon
}
#====================================================================

#====================================================================
# Set up LAPS
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Setting up LAPS"
Update-LapsADSchema -Confirm:$False
Set-LapsADComputerSelfPermission -Identity $Location
Set-LapsADReadPasswordPermission -Identity "OU=$($Env.OUs.Desktops),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=$($Env.OUs.Desktops),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=$($Env.OUs.Laptops),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=$($Env.OUs.Laptops),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=$($Env.OUs.Servers),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Server_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=$($Env.OUs.Servers),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Server_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=$($Env.OUs.VMs),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=$($Env.OUs.VMs),$Location" -AllowedPrincipals "$Domain\$($Env.Groups.TaskPrefix)Desktop_Admins"
#====================================================================

#====================================================================
# Secure built-in admin account
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Securing built-in admin account"
try {
    Move-ADObject -Identity $(Get-ADUser -Identity $SID500).DistinguishedName -TargetPath "OU=$($Env.OUs.HiPrivAccounts),OU=$AdministrationOU,$EndPath" -Server $DCHostName
} catch [Microsoft.ActiveDirectory.Management.ADException] {
    # Already moved - first run already moved it
    if ($_.Exception.Message -notmatch "already exists in target container") {
        throw
    }
}
Set-ADAccountControl -Identity $SID500 -AccountNotDelegated $True -Server $DCHostName
try {
    Remove-ADGroupMember -Identity "Schema Admins" -Members $SID500 -Server $DCHostName -Confirm:$False -ErrorAction Stop
} catch [Microsoft.ActiveDirectory.Management.ADException] {
    # Already not a member - first run already removed it
    if ($_.Exception.Message -notmatch "not a member") {
        throw
    }
}
#====================================================================

#====================================================================
# Import GPOs
#====================================================================
Write-LogFile -LogFile $LogFile -LogString "Creating & linking GPOs"
$GPOName = "Default Domain Policy"
try {
    Set-GPLink -name $GPOName -target $EndPath -enforced yes -ErrorAction Stop -Server $DCHostName
    Write-LogFile -LogFile $LogFile -LogString "Enforced $GPOName on $EndPath"
} catch {
    throw
}
$GPOName = "Default Domain Controllers Policy"
try {
    Set-GPLink -name $GPOName -target "ou=Domain Controllers,$EndPath" -enforced yes -ErrorAction Stop -Server $DCHostName
    Write-LogFile -LogFile $LogFile -LogString "Enforced $GPOName on $EndPath"
} catch {
    throw
}

# =============================================================================
# GPO imports and links
# =============================================================================
# Each Import-GPO call is wrapped in its own `if (-not $SkipGPOs)` block
# rather than one consolidating wrapper around all of them. This is a
# deliberate readability choice: the wrapping pattern is right there to
# imitate for any future GPO import added below, so the 21st import is
# almost guaranteed to be gated correctly by copy-paste rather than by
# the maintainer remembering to extend a single distant wrapper.
#
# AST regression net: ScriptBody.Tests.ps1 has 'gates every Import-GPO
# call behind -not $SkipGPOs' which walks every Import-GPO command and
# verifies it has an IfStatementAst ancestor mentioning $SkipGPOs - so
# both wrapping styles pass the test, but the per-call style fails more
# loudly when accidentally dropped (one missing wrapper is one ungated
# call, not all of them).
#
# Do not consolidate without first checking the AST test still pins the
# invariant on the new shape.
# =============================================================================

if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "Logon Policy" -TargetName "Logon Policy" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "Logon Policy" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "TLS" -TargetName "TLS" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "TLS" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "CWDIllegalInDllSearch" -TargetName "CWDIllegalInDllSearch" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "CWDIllegalInDllSearch" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "DisableNullSessionEnumeration" -TargetName "DisableNullSessionEnumeration" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "DisableNullSessionEnumeration" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "EnableSMBSigning" -TargetName "EnableSMBSigning" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "EnableSMBSigning" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "EnforceNLAandTLSforRDP" -TargetName "EnforceNLAandTLSforRDP" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "EnforceNLAandTLSforRDP" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "GroupPolicyHardening_MS150-011" -TargetName "GroupPolicyHardening_MS150-011" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "GroupPolicyHardening_MS150-011" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "Kerberos_Armouring" -TargetName "Kerberos_Armouring" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "Kerberos_Armouring" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "LSAProtection" -TargetName "LSAProtection" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "LSAProtection" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "NTLMv2" -TargetName "NTLMv2" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "NTLMv2" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "LDAP_signing_requirements" -TargetName "LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "LDAP_signing_requirements" -GPOTarget $EndPath
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "DC_AppLocker_Disable_Browsers" -TargetName "DC_AppLocker_Disable_Browsers" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "DC_AppLocker_Disable_Browsers" -GPOTarget "ou=Domain Controllers,$EndPath"
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "DC_Auditing" -TargetName "DC_Auditing" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "DC_Auditing" -GPOTarget "ou=Domain Controllers,$EndPath"
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "DC_Disable_Print_Spooler" -TargetName "DC_Disable_Print_Spooler" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "DC_Disable_Print_Spooler" -GPOTarget "ou=Domain Controllers,$EndPath"
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "DC_LDAP_signing_requirements" -TargetName "DC_LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "DC_LDAP_signing_requirements" -GPOTarget "ou=Domain Controllers,$EndPath"
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "$($Env.Groups.TaskPrefix)Server_Admins as members of Local admins" -TargetName "$($Env.Groups.TaskPrefix)Server_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "$($Env.Groups.TaskPrefix)Server_Admins as members of Local admins" -GPOTarget $Location
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "LAPS" -TargetName "LAPS" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "LAPS" -GPOTarget $Location
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "$($Env.Groups.TaskPrefix)Desktop_Admins as members of Local admins" -TargetName "$($Env.Groups.TaskPrefix)Desktop_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "$($Env.Groups.TaskPrefix)Desktop_Admins as members of Local admins" -GPOTarget "OU=$($Env.OUs.Desktops),$Location"
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "$($Env.Groups.TaskPrefix)Desktop_Admins as members of Local admins" -GPOTarget "OU=$($Env.OUs.Laptops),$Location"
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "$($Env.Groups.TaskPrefix)Desktop_Admins as members of Local admins" -GPOTarget "OU=$($Env.OUs.VMs),$Location"
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "IT Desktop Prefs" -TargetName "IT Desktop Prefs" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "IT Desktop Prefs" -GPOTarget "OU=$StaffOU,$EndPath"
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "IT Desktop Prefs" -GPOTarget "OU=$($Env.OUs.HiPrivAccounts),OU=$AdministrationOU,$EndPath"
    try {
        Write-LogFile -LogFile $LogFile -LogString "Updating permissions on IT Desktop Prefs"
        Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Server $DCHostName -Replace
        Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITGroup -TargetType Group -Server $DCHostName
        Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITAdminGroup -TargetType Group -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Updated permissions on IT Desktop Prefs"
    } catch {
        throw
    }
}
if (-not $SkipGPOs) {
    $ImportedGPOs += Import-GPO -BackupGpoName "CM visual help" -TargetName "CM visual help" -path $GPOLocation -CreateIfNeeded -Server $DCHostName
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "CM visual help" -GPOTarget "OU=$StaffOU,$EndPath"
    Add-GPOLink -LogFile $LogFile -DCHostName $DCHostName -GPOName "CM visual help" -GPOTarget "OU=$($Env.OUs.HiPrivAccounts),OU=$AdministrationOU,$EndPath"
    try {
        Write-LogFile -LogFile $LogFile -LogString "Updating permissions on CM visual help"
        Set-GPPermission -Name "CM visual help" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Server $DCHostName -Replace
        Set-GPPermission -Name "CM visual help" -PermissionLevel GpoApply -TargetName $SID500 -TargetType User -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Updated permissions on CM visual help"
    } catch {
        throw
    }
}
foreach ($GPOName in $ImportedGPOs) {
    try {
        Write-LogFile -LogFile $LogFile -LogString "Updating permissions on $($GPOName.DisplayName)"
        Set-GPPermission -Name $GPOName.DisplayName -PermissionLevel GpoEditDeleteModifySecurity -TargetName "$($Env.Groups.TaskPrefix)AD_GPO_Admins" -TargetType Group -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Updated permissions on $($GPOName.DisplayName)"
    } catch {
        throw
    }
}
#====================================================================
