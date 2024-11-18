[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)
#====================================================================
#Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$ServerName = "$env:computername"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$rootdse = Get-ADRootDSE
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = "c:\scripts\prelim\gpos"
$UserScriptsLocation = "c:\scripts\users"
#====================================================================

#====================================================================
#Drive where all the folders will be created
#====================================================================
$Drive = "E:"
$RootShare = "Store"
#====================================================================

#====================================================================
#Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
$InstallerGroup = "ADM_Task_Installers"
$ITGroup = "IT"
$ITAdminGroup = "IT_Admin"
$UserPasswordDelegationGroup = "ADM_Task_Password_Admins"
$SERAccessAdminGroup = "ADM_Task_SER_Access_Admins"
$SERAccountAdminGroup = "ADM_Task_SER_Account_Admins"
$HiPrivAccountAdminGroup = "ADM_Task_HiPriv_Account_Admins"
$HiPrivGroupAdminGroup = "ADM_Task_HiPriv_Group_Admins"
$StandardAccountAdminGroup = "ADM_Task_Standard_Account_Admins"
$StandardGroupAdminGroup = "ADM_Task_Standard_Group_Admins"
$LocalAdminAdminGroup = "ADM_Task_Local_Admin_Admins"
$ServiceAccountAdminGroup = "ADM_Task_Service_Account_Admins"
$DNSOperatorsGroup = "ADM_Task_DNS_Operators"
$DNSReadOnlyGroup = "ADM_Task_DNS_ReadOnly"
$SharedGroupsOU = "Shared_Mailbox_Access,OU=$GroupsOU"
$EquipmentGroupsOU = "Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "Room_Mailbox_Access,OU=$GroupsOU"
$SharedAccountsOU = "Shared_Mailbox_Accounts,OU=Administration"
$EquipmentAccountsOU = "Equipment_Mailbox_Accounts,OU=Administration"
$RoomAccountsOU = "Room_Mailbox_Accounts,OU=Administration"
$SID500 = "$Domain-Admin"
#====================================================================

#====================================================================
#OU creation function
#====================================================================
Function Create-ADOU {
    [CmdletBinding()]
    param(
        [string]$OUName,[String]$Path,[String]$OUDescription
    )
    $Error.Clear()
    try {
        New-ADOrganizationalUnit -Name $OUName -Path $Path -ProtectedFromAccidentalDeletion:$true -Description $OUDescription
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        Write-Host "'$OUName' already exists" -ForegroundColor Green
    }
}
#====================================================================

#====================================================================
#group creation function
#====================================================================
Function Create-ADGroup {
    [CmdletBinding()]
    param(
        [string]$GroupName,[String]$Path,[String]$GroupDescription
    )
    $Error.Clear()
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -OtherAttributes:@{mail="$GroupName@$EmailSuffix"} -Path $Path -SamAccountName $GroupName -Description $GroupDescription
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        Write-Host "'$GroupName' already exists" -ForegroundColor Green
    }
}
#====================================================================

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
            Add-ADGroupMember -Identity $Group -Members $Member
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
#GPO link function
#====================================================================
function Link-GPO {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GPOName,
        [Parameter(Mandatory)][string]$GPOTarget
    )
    #================================================================
    # Purpose:          To link a GPO to an OU
    # Assumptions:      Parameters have been set correctly
    # Effects:          GPO will be linked to the OU
    # Inputs:           $GPOName - Name of GPO as set before calling the function
    #                   $GPOTarget - OU where GPO will be linked
    # Returns:
    # Notes:
    #================================================================
    $Error.Clear()
    try {
        New-GPLink -name $GPOName -target $GPOTarget -LinkEnabled Yes -enforced yes -ErrorAction Stop
    } catch {
        Write-Host "GPLink already exists" -ForegroundColor Green
    }
}
#====================================================================

#====================================================================
#Delegate permission on computer objects to a group
#====================================================================
function Delegate-Computer-Join {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild","Allow","All",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild","Allow","All",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow","Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteDacl","Allow","Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self","Allow",$ExtendedRightsMap["Validated write to DNS host name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self","Allow",$ExtendedRightsMap["Validated write to service principal name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,"Allow",$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,"Allow",$ExtendedRightsMap["Change Password"],"Descendents",$GuidMap["computer"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permission on group objects to a group
#====================================================================
function Delegate-Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild","Allow","All",$GuidMap["group"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild","Allow","All",$GuidMap["group"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow","Descendents",$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate modify membership permission on group objects to a group
#====================================================================
function Delegate-Group-Membership-Edit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow",$GuidMap["member"],'Descendents',$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate change password permission on user objects to a group
#====================================================================
function Delegate-Password-Reset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight","Allow",$ExtendedRightsMap["Reset Password"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow",$GuidMap["pwdLastSet"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow",$GuidMap["lockoutTime"],'Descendents',$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permission on user objects to a group
#====================================================================
function Delegate-User {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild","Allow","All",$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild","Allow","All",$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty","Allow","Descendents",$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight","Allow",$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permissions to DNS Operators
#====================================================================
function Delegate-DNSOperatorsPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] "GenericRead"
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] "GenericExecute"
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] "GenericWrite"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceType))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permissions to DNS ReadOnly Group
#====================================================================
function Delegate-DNSReadOnlyPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] "GenericRead"
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] "GenericExecute"
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] "GenericWrite"
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] "GenericAll"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $adRightsRC = [System.DirectoryServices.ActiveDirectoryRights] "ReadControl"
    $adRightsWO = [System.DirectoryServices.ActiveDirectoryRights] "WriteOwner"
    $adRightsWD = [System.DirectoryServices.ActiveDirectoryRights] "WriteDacl"
    $adRightsDT = [System.DirectoryServices.ActiveDirectoryRights] "DeleteTree"
    $adRightsDEL = [System.DirectoryServices.ActiveDirectoryRights] "Delete"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $AccessControlTypeDeny = [System.Security.AccessControl.AccessControlType] "Deny"
    $inheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWO,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWD,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDT,$AccessControlTypeDeny,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDEL,$AccessControlTypeDeny,$inheritanceType))
    $Acl.RemoveAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsRC,$AccessControlTypeDeny,$inheritanceType))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permissions to AD Sites, Subnets, and Transports admins
#====================================================================
function Delegate-ADObjectPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] "GenericAll"
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] "DeleteChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGA,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceType))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Delegate permissions to GPO admins
#====================================================================
function Delegate-GPO-Permissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsRP = [System.DirectoryServices.ActiveDirectoryRights] "ReadProperty"
    $adRightsWP = [System.DirectoryServices.ActiveDirectoryRights] "WriteProperty"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceType = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Logging)"],$inheritanceType))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Planning)"],$inheritanceType))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
#Add additional UPN suffix
#====================================================================
Get-ADForest | Set-ADForest -UPNSuffixes @{add="$EmailSuffix"}
#====================================================================

#====================================================================
#Enable PAM feature to use temporal group membership
#====================================================================
Enable-ADOptionalFeature "Privileged Access Management Feature" -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False
#====================================================================

#====================================================================
#Enable AD Recycle Bin
#====================================================================
Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False
#====================================================================

Write-Host "Creating user OUs & Groups"
#====================================================================
#Staff OU & group creation
#====================================================================
Create-ADOU -OUName $StaffGroup -Path $EndPath -OUDescription "Top level OU for User objects"
Create-ADOU -OUName $GroupsOU -Path $EndPath -OUDescription "Top level OU for Group objects"
Create-ADGroup -GroupName $StaffGroup -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Org-wide group for all users"
Create-ADGroup -GroupName $ITGroup -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Departmental group holding all IT accounts"
Create-ADGroup -GroupName "License_Office365" -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Used to assign Office365 licenses"
Create-ADOU -OUName "Administration" -Path $EndPath -OUDescription "Top level OU for IT Admin User & Group objects"
Create-ADOU -OUName "Equipment_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to equipment mailboxes"
Create-ADOU -OUName "Room_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to room mailboxes"
Create-ADOU -OUName "Shared_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to shared mailboxes"
Create-ADOU -OUName "Equipment_Mailbox_Accounts" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for User objects that are equipment mailbox recipient types"
Create-ADOU -OUName "Hi_Priv_Groups" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for Group objects that control Hi-Priv access"
Create-ADOU -OUName "Hi_Priv_Accounts" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for User objects that are Hi-Priv accounts"
Create-ADOU -OUName "Local_Admin_Groups" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for Group objects that give local admin on individual devices"
Create-ADOU -OUName "Room_Mailbox_Accounts" -Path "OU=Administration,$EndPath" -OUDescription "IT OAdmin OU for User objects that are room mailbox recipient types"
Create-ADOU -OUName "Shared_Mailbox_Accounts" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for User objects that are shared mailbox recipient types"
Create-ADOU -OUName "Service_Accounts" -Path "OU=Administration,$EndPath" -OUDescription "IT Admin OU for User objects that are service accounts"
Create-ADGroup -GroupName $ITAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Group holding all IT Admin accounts"
Create-ADGroup -GroupName "ADM_Task_ADGPO_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create and link GPOs"
Create-ADGroup -GroupName "ADM_Task_ADSite_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_ADSubnet_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Full Control over AD Subnet objects"
Create-ADGroup -GroupName "ADM_Task_ADTransport_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Full Control over AD Transport objects"
Create-ADGroup -GroupName "ADM_Task_Desktop_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members are added to Local Admin on all computers in the Desktop, Laptop, & VM OUs"
Create-ADGroup -GroupName "ADM_Task_DFS_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Full Control NTFS permissions on all DFS folders & have access to DFS console"
Create-ADGroup -GroupName "ADM_Task_DHCP_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members are members of DHCP Administrators"
Create-ADGroup -GroupName "ADM_Task_DHCP_Users" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members are members of DHCP Users"
Create-ADGroup -GroupName "ADM_Task_DNS_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members are indirect members of the DNSAdmins BuiltIn group"
Create-ADGroup -GroupName $DNSOperatorsGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Edit access to DNS zones"
Create-ADGroup -GroupName $DNSReadOnlyGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have read only access to the DNS service"
Create-ADGroup -GroupName $HiPrivAccountAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Hi_Priv_Accounts OU"
Create-ADGroup -GroupName $HiPrivGroupAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Hi_Priv_Groups OU"
Create-ADGroup -GroupName $InstallerGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have permission to join & move computer objects in $ParentOU & Sub OUs"
Create-ADGroup -GroupName $LocalAdminAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Local_Admin_Groups OU"
Create-ADGroup -GroupName $UserPasswordDelegationGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can reset passwords of users in the $StaffGroup OU"
Create-ADGroup -GroupName "ADM_Task_Server_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members are added to Local admin on all computers in $ParentOU & Sub OUs via GPO, are indirect members of the Server Operators BuiltIn group"
Create-ADGroup -GroupName $SERAccessAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can edit the membership of sh_, eq_, & ro_ groups"
Create-ADGroup -GroupName $SERAccountAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit sh_, eq_, & ro_ accounts"
Create-ADGroup -GroupName $ServiceAccountAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Service_Accounts OU"
Create-ADGroup -GroupName $StandardAccountAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $StaffGroup OU"
Create-ADGroup -GroupName $StandardGroupAdminGroup -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $GroupsOU OU"
Create-ADGroup -GroupName "ADM_Task_Subscription_Contributors" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Contributor permissions on the subscription"
Create-ADGroup -GroupName "ADM_Task_Subscription_Owners" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have Owner permissions on the subscription"
Create-ADGroup -GroupName "ADM_Task_Subscription_User_Access_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members have User Access Admin permissions on the subscription"
Create-ADGroup -GroupName "ADM_Task_WDS_Deploy_Clients" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Clients folder"
Create-ADGroup -GroupName "ADM_Task_WDS_Deploy_Servers" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Servers folder"
Create-ADGroup -GroupName "ADM_Role_Level_1_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Level 1 admins - desktop, includes all level 2 admins"
Create-ADGroup -GroupName "ADM_Role_Level_2_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Level 2 admins - junior server, includes all level 3 admins"
Create-ADGroup -GroupName "ADM_Role_Level_3_Admins" -Path "OU=Hi_Priv_Groups,OU=Administration,$EndPath" -GroupDescription "Level 3 admins - senior server"
Add-GroupMember -group $StaffGroup -Member $ITGroup
Add-GroupMember -group "Remote Desktop Users" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group "Server Operators" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group "DnsAdmins" -Member "ADM_Task_DNS_Admins"
Add-GroupMember -group $ITAdminGroup -Member $SID500
Add-GroupMember -group "ADM_Task_ADGPO_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_ADSite_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_ADSubnet_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_ADTransport_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_DFS_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DNS_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $DNSReadOnlyGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $DNSOperatorsGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $HiPrivAccountAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $HiPrivGroupAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $LocalAdminAdminGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $LocalAdminAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $LocalAdminAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $UserPasswordDelegationGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Server_Admins" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Server_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $SERAccessAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $SERAccountAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $SERAccountAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $StandardAccountAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $StandardAccountAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $StandardGroupAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $StandardGroupAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $ServiceAccountAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $ServiceAccountAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Subscription_Contributors" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Subscription_Contributors" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Clients" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Servers" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_WDS_Deploy_Servers" -Member "ADM_Role_Level_3_Admins"
Get-ADUser $SID500 | Move-ADObject -TargetPath "OU=Hi_Priv_Accounts,OU=Administration,$EndPath"
#====================================================================

Write-Host "Creating Permission delegations"
#====================================================================
#Permission delegation creation
#====================================================================
$guidMap = @{}
$GuidMapParams = @{
    SearchBase = ($rootdse.SchemaNamingContext)
    LDAPFilter = "(schemaidguid=*)"
    Properties = ("lDAPDisplayName", "schemaIDGUID")
}
Get-ADObject @GuidMapParams | ForEach-Object { $guidMap[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID }
$ExtendedRightsMap = @{}
$ExtendedMapParams = @{
    SearchBase = ($rootdse.ConfigurationNamingContext)
    LDAPFilter = "(&(objectclass=controlAccessRight)(rightsguid=*))"
    Properties = ("displayName", "rightsGuid")
}
Get-ADObject @ExtendedMapParams | ForEach-Object { $ExtendedRightsMap[$_.displayName] = [System.GUID]$_.rightsGuid }
Delegate-Password-Reset -AdminGroupName $UserPasswordDelegationGroup -TargetOU $StaffGroup
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $EquipmentGroupsOU
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $RoomGroupsOU
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $SharedGroupsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentAccountsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomAccountsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedAccountsOU
Delegate-User -AdminGroupName $HiPrivAccountAdminGroup -TargetOU "Hi_Priv_Accounts,OU=Administration"
Delegate-Group -AdminGroupName $HiPrivGroupAdminGroup -TargetOU "Hi_Priv_Groups,OU=Administration"
Delegate-User -AdminGroupName $StandardAccountAdminGroup -TargetOU $StaffGroup
Delegate-Group -AdminGroupName $StandardGroupAdminGroup -TargetOU $GroupsOU
Delegate-User -AdminGroupName $ServiceAccountAdminGroup -TargetOU "Service_Accounts,OU=Administration"
Delegate-Group -AdminGroupName $LocalAdminAdminGroup -TargetOU "Local_Admin_Groups,OU=Administration"
Delegate-DNSOperatorsPermissions -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Delegate-DNSOperatorsPermissions -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_ADSite_Admins" -TargetDN "CN=Sites,CN=Configuration"
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_ADSubnet_Admins" -TargetDN "CN=Subnets,CN=Sites,CN=Configuration"
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_ADTransport_Admins" -TargetDN "CN=Inter-Site Transports,CN=Sites,CN=Configuration"
Delegate-GPO-Permissions -AdminGroupName "ADM_Task_ADGPO_Admins"
$DNSZones = Get-ADObject -Filter * -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$EndPath" -SearchScope 1
foreach ($DNSZone in $DNSZones) {
    $DNSZoneName = $DNSZone.Name
    Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "DC=$DNSZoneName,CN=MicrosoftDNS,DC=DomainDnsZones"
}
#====================================================================

Write-Host "Creating shares"
#====================================================================
#Create main Share
#====================================================================
$ShareName = $RootShare
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(TEST-PATH "\\$ServerName\$ShareName")) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory
        } else {
            Write-Host "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "authenticated users"
        Write-Host "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-Host "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "ADM_Task_DFS_Admins"
    New-DfsReplicationGroup -GroupName "$ShareName" | New-DfsReplicatedFolder -FolderName "$ShareName" -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName "$ShareName" -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Host "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
#Create Profiles Share
#====================================================================
$ShareName = "Profiles"
if (!(TEST-PATH "\\$DNSSuffix\$ShareName")) {
    if (!(TEST-PATH "\\$ServerName\$ShareName")) {
        if (!(TEST-PATH "$Drive\$ShareName")) {
            New-Item "$Drive\$ShareName" -type directory -force
            $Acl = Get-Acl "$Drive\$ShareName"
            $isProtected = $true
            $preserveInheritance = $false
            $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StaffGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl "$Drive\$ShareName" $Acl
        } else {
            Write-Host "$Drive\$ShareName already exists" -ForegroundColor Green
        }
        New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "authenticated users"
        Write-Host "Pausing for 60 seconds after creating share $ShareName"
        Start-Sleep -s 60
    } else {
        Write-Host "\\$ServerName\$ShareName already exists" -ForegroundColor Green
    }
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "ADM_Task_DFS_Admins"
    New-DfsReplicationGroup -GroupName "$ShareName" | New-DfsReplicatedFolder -FolderName "$ShareName" -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName "$ShareName" -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Host "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
#Create IT Group Share
#====================================================================
$ShareName = $ITGroup
if (!(TEST-PATH "\\$Domain\$RootShare\$ShareName")) {
    New-Item "\\$Domain\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$Domain\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ShareName,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$Domain\$RootShare\$ShareName" $Acl
    Robocopy -e -s $UserScriptsLocation "\\$Domain\$RootShare\$ShareName\User_Scripts"
} else {
    Write-Host "\\$Domain\$RootShare\$ShareName already exists" -ForegroundColor Green
}
$ShareName = "Software"
if (!(TEST-PATH "\\$Domain\$RootShare\$ShareName")) {
    New-Item "\\$Domain\$RootShare\$ShareName" -type directory
    $Acl = Get-Acl "\\$Domain\$RootShare\$ShareName"
    $isProtected = $true
    $preserveInheritance = $false
    $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Authenticated Users","ReadAndExecute","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ITAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$Domain\$RootShare\$ShareName" $Acl
} else {
    Write-Host "\\$Domain\$RootShare\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

Write-Host "Creating computer OUs"
#====================================================================
# Create default computers OU
#====================================================================
Create-ADOU -OUName $ParentOU -Path $EndPath -OUDescription "Top level OU for Computer objects"
#====================================================================

#====================================================================
# Create OU for servers
#====================================================================
Create-ADOU -OUName "Servers" -Path $Location
#====================================================================

#====================================================================
# Create OU for desktops
#====================================================================
Create-ADOU -OUName "Desktops" -Path $Location
#====================================================================

#====================================================================
# Create OU for laptops
#====================================================================
Create-ADOU -OUName "Laptops" -Path $Location
#====================================================================

#====================================================================
# Create OU for VMs
#====================================================================
Create-ADOU -OUName "VMs" -Path $Location
#====================================================================

#====================================================================
#Redirect default computer location
#====================================================================
redircmp "$Location"
Delegate-Computer-Join -AdminGroupName $InstallerGroup -TargetOU "$ParentOU"
#====================================================================

Write-Host "Creating GPOs"
#====================================================================
#Import GPOs
#====================================================================
try {
    Set-GPLink -Name "Default Domain Policy" -Target "$EndPath" -Enforced Yes -ErrorAction Stop
} catch {
    Write-Host "GPLink already exists" -ForegroundColor Green
}
try {
    Set-GPLink -Name "Default Domain Controllers Policy" -Target "ou=Domain Controllers,$EndPath" -Enforced Yes -ErrorAction Stop
} catch {
    Write-Host "GPLink already exists" -ForegroundColor Green
}
Import-GPO -BackupGpoName "Logon Policy" -TargetName "Logon Policy" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "Logon Policy" -GPOTarget "$EndPath"
Import-GPO -BackupGpoName "TLS" -TargetName "TLS" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "TLS" -GPOTarget "$EndPath"
Import-GPO -BackupGpoName "ADM_Task_Server_Admins as members of Local admins" -TargetName "ADM_Task_Server_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "ADM_Task_Server_Admins as members of Local admins" -GPOTarget "$Location"
Import-GPO -BackupGpoName "ADM_Task_Desktop_Admins as members of Local admins" -TargetName "ADM_Task_Desktop_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Desktops,$Location"
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Laptops,$Location"
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=VMs,$Location"
Import-GPO -BackupGpoName "IT Desktop Prefs" -TargetName "IT Desktop Prefs" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=Hi_Priv_Accounts,OU=Administration,$EndPath"
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Replace
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName "$ITGroup" -TargetType Group
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName "$ITAdminGroup" -TargetType Group
Import-GPO -BackupGpoName "CM visual help" -TargetName "CM visual help" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=Hi_Priv_Accounts,OU=Administration,$EndPath"
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Replace
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoApply -TargetName $SID500 -TargetType User
Import-GPO -BackupGpoName "Deploy Firefox" -TargetName "Deploy Firefox" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy Firefox" -GPOTarget "OU=Desktops,$Location"
Link-GPO -GPOName "Deploy Firefox" -GPOTarget "OU=Laptops,$Location"
Link-GPO -GPOName "Deploy Firefox" -GPOTarget "OU=VMs,$Location"
Import-GPO -BackupGpoName "Deploy Notepad++" -TargetName "Deploy Notepad++" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy Notepad++" -GPOTarget "$Location"
Import-GPO -BackupGpoName "Deploy PuTTY" -TargetName "Deploy PuTTY" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy PuTTY" -GPOTarget "OU=Desktops,$Location"
Link-GPO -GPOName "Deploy PuTTY" -GPOTarget "OU=Laptops,$Location"
Link-GPO -GPOName "Deploy PuTTY" -GPOTarget "OU=VMs,$Location"
Set-GPPermission -All -PermissionLevel GpoEditDeleteModifySecurity -TargetName "ADM_Task_ADGPO_Admins" -TargetType Group
#====================================================================
