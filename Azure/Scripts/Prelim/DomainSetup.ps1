[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
)
#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$ServerName = "$env:computername"
$SID500 = "$env:username"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$rootdse = Get-ADRootDSE
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GPOLocation = "c:\scripts\prelim\gpos"
$UserScriptsLocation = "c:\scripts\users"
#====================================================================

#====================================================================
# Drive where all the folders will be created
#====================================================================
$Drive = "F:"
$RootShare = "Store"
#====================================================================

#====================================================================
# Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
$AdministrationOU = "Administration"
$ITGroup = "IT"
$ITAdminGroup = "IT_Admin"
$O365LicenseGroup = "License_Office365"
$DNSOperatorsGroup = "ADM_Task_DNS_Operators"
$DNSReadOnlyGroup = "ADM_Task_DNS_ReadOnly"
$HiPrivAccountAdminGroup = "ADM_Task_HiPriv_Account_Admins"
$HiPrivGroupAdminGroup = "ADM_Task_HiPriv_Group_Admins"
$InstallerGroup = "ADM_Task_Installers"
$LocalAdminGroupAdminGroup = "ADM_Task_Local_Admin_Group_Admins"
$UserPasswordDelegationGroup = "ADM_Task_Password_Admins"
$SERAccessAdminGroup = "ADM_Task_SER_Access_Admins"
$SERAccountAdminGroup = "ADM_Task_SER_Account_Admins"
$ServiceAccountAdminGroup = "ADM_Task_Service_Account_Admins"
$StandardAccountAdminGroup = "ADM_Task_Standard_Account_Admins"
$StandardGroupAdminGroup = "ADM_Task_Standard_Group_Admins"
$EquipmentAccountsOU = "Equipment_Mailbox_Accounts,OU=$AdministrationOU"
$RoomAccountsOU = "Room_Mailbox_Accounts,OU=$AdministrationOU"
$SharedAccountsOU = "Shared_Mailbox_Accounts,OU=$AdministrationOU"
$EquipmentGroupsOU = "Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "Room_Mailbox_Access,OU=$GroupsOU"
$SharedGroupsOU = "Shared_Mailbox_Access,OU=$GroupsOU"
#====================================================================

#====================================================================
# OU creation function
#====================================================================
function Create-ADOU {
    [CmdletBinding()]
    param(
        [string]$OUName,[String]$Path,[String]$OUDescription
    )
    $Error.Clear()
    Write-Host "Creating OU $OUName"
    try {
        New-ADOrganizationalUnit -Name $OUName -Path $Path -ProtectedFromAccidentalDeletion:$true -Description $OUDescription
        Write-Host "Created OU $OUName"
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        switch ($Error[0].Exception.Message) {
            "The specified Organizational Unit already exists"{
                Write-Host "'$OUName' already exists" -ForegroundColor Green
            }
            default {
                Write-Host "ERROR: An unexpected error occurred while attempting to create OU '$OUName':`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
            }
        }
    }
}
#====================================================================

#====================================================================
# Group creation function
#====================================================================
function Create-ADGroup {
    [CmdletBinding()]
    param(
        [string]$GroupName,[String]$GroupScope,[String]$Path,[String]$GroupDescription
    )
    $Error.Clear()
    Write-Host "Creating Group $GroupName"
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -OtherAttributes:@{mail="$GroupName@$EmailSuffix"} -Path $Path -SamAccountName $GroupName -Description $GroupDescription
        Set-ADObject -Identity "CN=$GroupName,$Path" -protectedFromAccidentalDeletion $True
        Write-Host "Created $GroupName"
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        switch ($Error[0].Exception.Message) {
            "The specified group already exists"{
                Write-Host "'$GroupName' already exists" -ForegroundColor Green
            }
            default {
                Write-Host "ERROR: An unexpected error occurred while attempting to create group '$GroupName' to a group:`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
            }
        }
    }
}
#====================================================================

#====================================================================
# Group addition function
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
            switch ($Error[0].Exception.Message) {
                "The specified object is already a member of the group" {
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
# GPO link function
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
    Write-Host "Linking $GPOName to $GPOTarget"
    try {
        New-GPLink -name $GPOName -target $GPOTarget -LinkEnabled Yes -enforced yes -Order 1 -ErrorAction Stop
        Write-Host "Linked $GPOName to $GPOTarget"
    } catch {
        switch ($Error[0].Exception.Message) {
            "The specified GPLink already exists"{
                Write-Host "GPLink already exists" -ForegroundColor Green
            }
            default {
                Write-Host "ERROR: An unexpected error occurred while attempting to link GPO '$GPOName' to OU '$GPOTarget':`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
            }
        }
    }
}
#====================================================================

#====================================================================
# Delegate permission on computer objects to a group
#====================================================================
function Delegate-Computer-Join {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["computer"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["computer"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteDacl",$AccessControlTypeAllow,"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self",$AccessControlTypeAllow,$ExtendedRightsMap["Validated write to DNS host name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"Self",$AccessControlTypeAllow,$ExtendedRightsMap["Validated write to service principal name"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["computer"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap["Change Password"],"Descendents",$GuidMap["computer"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on group objects to a group
#====================================================================
function Delegate-Group {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["group"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["group"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate modify membership permission on group objects to a group
#====================================================================
function Delegate-Group-Membership-Edit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["member"],'Descendents',$GuidMap["group"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate change password permission on user objects to a group
#====================================================================
function Delegate-Password-Reset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["pwdLastSet"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,$GuidMap["lockoutTime"],'Descendents',$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],'Descendents',$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on user objects to a group
#====================================================================
function Delegate-User {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["user"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["user"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["user"]))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["user"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on OU objects to a group
#====================================================================
function Delegate-OU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetOU
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $Acl = Get-Acl "AD:\OU=$TargetOU,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"CreateChild",$AccessControlTypeAllow,$GuidMap["organizationalUnit"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"DeleteChild",$AccessControlTypeAllow,$GuidMap["organizationalUnit"],"All"))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"WriteProperty",$AccessControlTypeAllow,"Descendents",$GuidMap["organizationalUnit"]))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS Operators
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
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS ReadOnly Group
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
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGW,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWO,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsWD,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDT,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDEL,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.RemoveAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsRC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to AD Sites, Subnets, and Transports admins
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
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsGA,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to GPO admins
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
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "All"
    $Acl = Get-Acl "AD:\$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPLink"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap["gPOptions"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Logging)"],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,"ExtendedRight",$AccessControlTypeAllow,$ExtendedRightsMap["Generate Resultant Set of Policy (Planning)"],$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to create GPOs
#====================================================================
function Delegate-GPO-Creation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName,
        [Parameter(Mandatory)][string]$TargetDN
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AdminGroupSIDidentity = [System.Security.Principal.IdentityReference] $AdminGroupSID
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] "CreateChild"
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] "Allow"
    $inheritanceTypeNone = [System.DirectoryServices.ActiveDirectorySecurityInheritance] "None"
    $Acl = Get-Acl "AD:\$TargetDN,$EndPath"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSIDidentity,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeNone))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Add additional UPN suffix
#====================================================================
Write-Host "Adding $EmailSuffix as ana dditional UPN suffix"
Get-ADForest | Set-ADForest -UPNSuffixes @{add="$EmailSuffix"}
#====================================================================

#====================================================================
# Prevent standard users from creating computer accounts
#====================================================================
Write-Host "Preventing standard users from creating computer accounts"
Set-ADDomain (Get-ADDomain).distinguishedname -Replace @{"ms-ds-MachineAccountQuota"="0"}
#====================================================================

#====================================================================
# Enable PAM feature to use temporal group membership
#====================================================================
Write-Host "Enabling PAM feature to use temporal group membership"
Enable-ADOptionalFeature "Privileged Access Management Feature" -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False
#====================================================================

#====================================================================
# Enable AD Recycle Bin
#====================================================================
Write-Host "Enabling AD Recycle Bin"
Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target $DNSSuffix -Confirm:$False
#====================================================================

#====================================================================
# Protect Domain Controllers OU
#====================================================================
Write-Host "Protecting Domain Controllers OU"
Set-ADOrganizationalUnit -Identity "OU=Domain Controllers,$EndPath" -ProtectedFromAccidentalDeletion $true
#====================================================================

Write-Host "Creating user OUs & Groups"
#====================================================================
# Staff OU & group creation
#====================================================================
Create-ADOU -OUName $StaffGroup -Path $EndPath -OUDescription "Top level OU for Staff User objects"
Create-ADOU -OUName $GroupsOU -Path $EndPath -OUDescription "Top level OU for Group objects"
Create-ADGroup -GroupName $StaffGroup -GroupScope $GroupScope -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Org-wide group for all staff users"
Create-ADGroup -GroupName $ITGroup -GroupScope $GroupScope -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Departmental group holding all IT accounts"
Create-ADGroup -GroupName $O365LicenseGroup -GroupScope $GroupScope -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Used to assign Office365 licenses"
Create-ADOU -OUName "Administration" -Path $EndPath -OUDescription "Top level OU for IT Admin User & Group objects"
Create-ADOU -OUName "Equipment_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to equipment mailboxes"
Create-ADOU -OUName "Room_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to room mailboxes"
Create-ADOU -OUName "Shared_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT Admin OU for Group objects that grant access to shared mailboxes"
Create-ADOU -OUName "Equipment_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are equipment mailbox recipient types"
Create-ADOU -OUName "Hi_Priv_Groups" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that control Hi-Priv access"
Create-ADOU -OUName "Hi_Priv_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are Hi-Priv accounts"
Create-ADOU -OUName "Local_Admin_Groups" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for Group objects that give local admin on individual devices"
Create-ADOU -OUName "Room_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are room mailbox recipient types"
Create-ADOU -OUName "Shared_Mailbox_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are shared mailbox recipient types"
Create-ADOU -OUName "Service_Accounts" -Path "OU=$AdministrationOU,$EndPath" -OUDescription "IT Admin OU for User objects that are service accounts"
Create-ADGroup -GroupName $ITAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Group holding all IT Admin accounts"
Create-ADGroup -GroupName "ADM_Task_AD_Administration_OU_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_AD_Computer_OU_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_AD_GPO_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create and link GPOs"
Create-ADGroup -GroupName "ADM_Task_AD_Group_OU_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_AD_Site_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_AD_Subnet_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Subnet objects"
Create-ADGroup -GroupName "ADM_Task_AD_Transport_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Transport objects"
Create-ADGroup -GroupName "ADM_Task_AD_User_OU_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control over AD Site objects"
Create-ADGroup -GroupName "ADM_Task_Desktop_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local Admin on all computers in the Desktop, Laptop, & VM OUs"
Create-ADGroup -GroupName "ADM_Task_DFS_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Full Control NTFS permissions on all DFS folders & have access to DFS console"
Create-ADGroup -GroupName "ADM_Task_DHCP_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Administrators"
Create-ADGroup -GroupName "ADM_Task_DHCP_Users" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are members of DHCP Users"
Create-ADGroup -GroupName "ADM_Task_DNS_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are indirect members of the DNSAdmins BuiltIn group"
Create-ADGroup -GroupName $DNSOperatorsGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Edit access to DNS zones"
Create-ADGroup -GroupName $DNSReadOnlyGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have read only access to the DNS service"
Create-ADGroup -GroupName $HiPrivAccountAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Hi_Priv_Accounts OU"
Create-ADGroup -GroupName $HiPrivGroupAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Hi_Priv_Groups OU"
Create-ADGroup -GroupName $InstallerGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have permission to join & move computer objects in $ParentOU & Sub OUs"
Create-ADGroup -GroupName $LocalAdminGroupAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the Local_Admin_Groups OU"
Create-ADGroup -GroupName $UserPasswordDelegationGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can reset passwords of users in the $StaffGroup OU"
Create-ADGroup -GroupName "ADM_Task_Server_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members are added to Local admin on all computers in $ParentOU & Sub OUs via GPO, are indirect members of the Server Operators BuiltIn group"
Create-ADGroup -GroupName $SERAccessAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can edit the membership of sh_, eq_, & ro_ groups"
Create-ADGroup -GroupName $SERAccountAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit sh_, eq_, & ro_ accounts"
Create-ADGroup -GroupName $ServiceAccountAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the Service_Accounts OU"
Create-ADGroup -GroupName $StandardAccountAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit accounts in the $StaffGroup OU"
Create-ADGroup -GroupName $StandardGroupAdminGroup -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members can create/delete/edit groups in the $GroupsOU OU"
Create-ADGroup -GroupName "ADM_Task_Subscription_Contributors" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Contributor permissions on the subscription"
Create-ADGroup -GroupName "ADM_Task_Subscription_Owners" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have Owner permissions on the subscription"
Create-ADGroup -GroupName "ADM_Task_Subscription_User_Access_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Members have User Access Admin permissions on the subscription"
Create-ADGroup -GroupName "ADM_Role_Level_1_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 1 admins - desktop"
Create-ADGroup -GroupName "ADM_Role_Level_2_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 2 admins - junior server"
Create-ADGroup -GroupName "ADM_Role_Level_3_Admins" -GroupScope $GroupScope -Path "OU=Hi_Priv_Groups,OU=$AdministrationOU,$EndPath" -GroupDescription "Level 3 admins - senior server"
Add-GroupMember -group $StaffGroup -Member $ITGroup
Add-GroupMember -group "Remote Desktop Users" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group "Server Operators" -Member "ADM_Task_Server_Admins"
Add-GroupMember -group "DnsAdmins" -Member "ADM_Task_DNS_Admins"
Add-GroupMember -group $ITAdminGroup -Member $SID500
Add-GroupMember -group "ADM_Task_AD_Administration_OU_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Computer_OU_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_GPO_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Group_OU_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Site_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Subnet_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_Transport_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_AD_User_OU_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_Desktop_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group "ADM_Task_DHCP_Users" -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group "ADM_Task_DFS_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group "ADM_Task_DNS_Admins" -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $DNSOperatorsGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $DNSReadOnlyGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $HiPrivAccountAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $HiPrivGroupAdminGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $InstallerGroup -Member "ADM_Role_Level_3_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Level_1_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Level_2_Admins"
Add-GroupMember -group $LocalAdminGroupAdminGroup -Member "ADM_Role_Level_3_Admins"
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
#====================================================================

Write-Host "Creating shares"
#====================================================================
# Create main Share
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
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Host "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
}
#====================================================================

#====================================================================
# Create Profiles Share
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
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StandardAccountAdminGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
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
    New-DfsReplicationGroup -GroupName $ShareName | New-DfsReplicatedFolder -FolderName $ShareName -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName $ServerName
    Set-DfsrMembership -GroupName $ShareName -FolderName $ShareName -ContentPath "$Drive\$ShareName" -ComputerName $ServerName -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName $ShareName -AccountName "ADM_Task_DFS_Admins" -Force
} else {
    Write-Host "\\$DNSSuffix\$ShareName already exists" -ForegroundColor Green
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
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
    Robocopy -e -s $UserScriptsLocation "\\$DNSSuffix\$RootShare\$ShareName\User_Scripts"
} else {
    Write-Host "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
}
$ShareName = "Software"
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
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("ADM_Task_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "\\$DNSSuffix\$RootShare\$ShareName" $Acl
} else {
    Write-Host "\\$DNSSuffix\$RootShare\$ShareName already exists" -ForegroundColor Green
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
# Redirect default computer location & delegate permissions
#====================================================================
Write-Host "Creating Permission delegations"
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
redircmp $Location
Delegate-Computer-Join -AdminGroupName $InstallerGroup -TargetOU $ParentOU
Delegate-OU -AdminGroupName "ADM_Task_AD_Computer_OU_Admins" -TargetOU $ParentOU
Delegate-Password-Reset -AdminGroupName $UserPasswordDelegationGroup -TargetOU $StaffGroup
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $EquipmentGroupsOU
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $RoomGroupsOU
Delegate-Group-Membership-Edit -AdminGroupName $SERAccessAdminGroup -TargetOU $SharedGroupsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentAccountsOU
Delegate-Group -AdminGroupName $SERAccountAdminGroup -TargetOU $EquipmentGroupsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomAccountsOU
Delegate-Group -AdminGroupName $SERAccountAdminGroup -TargetOU $RoomGroupsOU
Delegate-User -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedAccountsOU
Delegate-Group -AdminGroupName $SERAccountAdminGroup -TargetOU $SharedGroupsOU
Delegate-User -AdminGroupName $HiPrivAccountAdminGroup -TargetOU "Hi_Priv_Accounts,OU=$AdministrationOU"
Delegate-Group -AdminGroupName $HiPrivGroupAdminGroup -TargetOU "Hi_Priv_Groups,OU=$AdministrationOU"
Delegate-User -AdminGroupName $StandardAccountAdminGroup -TargetOU $StaffGroup
Delegate-Group -AdminGroupName $StandardGroupAdminGroup -TargetOU $GroupsOU
Delegate-User -AdminGroupName $ServiceAccountAdminGroup -TargetOU "Service_Accounts,OU=$AdministrationOU"
Delegate-Group -AdminGroupName $LocalAdminGroupAdminGroup -TargetOU "Local_Admin_Groups,OU=$AdministrationOU"
Delegate-OU -AdminGroupName "ADM_Task_AD_Administration_OU_Admins" -TargetOU $AdministrationOU
Delegate-OU -AdminGroupName "ADM_Task_AD_Group_OU_Admins" -TargetOU $GroupsOU
Delegate-OU -AdminGroupName "ADM_Task_AD_User_OU_Admins" -TargetOU $StaffGroup
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_AD_Site_Admins" -TargetDN "CN=Sites,CN=Configuration"
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_AD_Subnet_Admins" -TargetDN "CN=Subnets,CN=Sites,CN=Configuration"
Delegate-ADObjectPermissions -AdminGroupName "ADM_Task_AD_Transport_Admins" -TargetDN "CN=Inter-Site Transports,CN=Sites,CN=Configuration"
Delegate-GPO-Permissions -AdminGroupName "ADM_Task_AD_GPO_Admins"
Delegate-GPO-Creation -AdminGroupName "ADM_Task_AD_GPO_Admins" -TargetDN "CN=Policies,CN=System"
Delegate-DNSOperatorsPermissions -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Delegate-DNSOperatorsPermissions -AdminGroupName $DNSOperatorsGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,CN=System"
Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "CN=MicrosoftDNS,DC=DomainDnsZones"
$DNSZones = Get-ADObject -Filter * -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$EndPath" -SearchScope 1
foreach ($DNSZone in $DNSZones) {
    $DNSZoneName = $DNSZone.Name
    Delegate-DNSReadOnlyPermissions -AdminGroupName $DNSReadOnlyGroup -TargetDN "DC=$DNSZoneName,CN=MicrosoftDNS,DC=DomainDnsZones"
}
#====================================================================

#====================================================================
# Set up LAPS
#====================================================================
Write-Host "Setting up LAPS"
Update-LapsADSchema -Confirm:$False
Set-LapsADComputerSelfPermission -Identity $Location
Set-LapsADReadPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Desktops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Laptops,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\ADM_Task_Server_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=Servers,$Location" -AllowedPrincipals "$Domain\ADM_Task_Server_Admins"
Set-LapsADReadPasswordPermission -Identity "OU=VMs,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
Set-LapsADResetPasswordPermission -Identity "OU=VMs,$Location" -AllowedPrincipals "$Domain\ADM_Task_Desktop_Admins"
#====================================================================

#====================================================================
# Secure built-in admin account
#====================================================================
Write-Host "Securing built-in admin account"
Get-ADUser $SID500 | Move-ADObject -TargetPath "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"
Set-ADAccountControl -Identity $SID500 -AccountNotDelegated $True
Remove-ADGroupMember -Identity "Schema Admins" -Members $SID500 -Confirm:$False
#====================================================================

#====================================================================
# Import GPOs
#====================================================================
Write-Host "Creating & linking GPOs"
try {
    Set-GPLink -Name "Default Domain Policy" -Target $EndPath -Enforced Yes -ErrorAction Stop
} catch {
    Write-Host "GPLink already exists" -ForegroundColor Green
}
try {
    Set-GPLink -Name "Default Domain Controllers Policy" -Target "ou=Domain Controllers,$EndPath" -Enforced Yes -ErrorAction Stop
} catch {
    Write-Host "GPLink already exists" -ForegroundColor Green
}
Import-GPO -BackupGpoName "Logon Policy" -TargetName "Logon Policy" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "Logon Policy" -GPOTarget $EndPath
Import-GPO -BackupGpoName "TLS" -TargetName "TLS" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "TLS" -GPOTarget $EndPath
Import-GPO -BackupGpoName "CWDIllegalInDllSearch" -TargetName "CWDIllegalInDllSearch" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "CWDIllegalInDllSearch" -GPOTarget $EndPath
Import-GPO -BackupGpoName "DisableNullSessionEnumeration" -TargetName "DisableNullSessionEnumeration" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "DisableNullSessionEnumeration" -GPOTarget $EndPath
Import-GPO -BackupGpoName "EnableSMBSigning" -TargetName "EnableSMBSigning" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "EnableSMBSigning" -GPOTarget $EndPath
Import-GPO -BackupGpoName "EnforceNLAandTLSforRDP" -TargetName "EnforceNLAandTLSforRDP" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "EnforceNLAandTLSforRDP" -GPOTarget $EndPath
Import-GPO -BackupGpoName "GroupPolicyHardening_MS150-011" -TargetName "GroupPolicyHardening_MS150-011" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "GroupPolicyHardening_MS150-011" -GPOTarget $EndPath
Import-GPO -BackupGpoName "Kerberos_Armouring" -TargetName "Kerberos_Armouring" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "Kerberos_Armouring" -GPOTarget $EndPath
Import-GPO -BackupGpoName "LSAProtection" -TargetName "LSAProtection" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "LSAProtection" -GPOTarget $EndPath
Import-GPO -BackupGpoName "NTLMv2" -TargetName "NTLMv2" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "NTLMv2" -GPOTarget $EndPath
Import-GPO -BackupGpoName "LDAP_signing_requirements" -TargetName "LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "LDAP_signing_requirements" -GPOTarget $EndPath
Import-GPO -BackupGpoName "DC_AppLocker_Disable_Browsers" -TargetName "DC_AppLocker_Disable_Browsers" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "DC_AppLocker_Disable_Browsers" -GPOTarget "ou=Domain Controllers,$EndPath"
Import-GPO -BackupGpoName "DC_Auditing" -TargetName "DC_Auditing" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "DC_Auditing" -GPOTarget "ou=Domain Controllers,$EndPath"
Import-GPO -BackupGpoName "DC_Disable_Print_Spooler" -TargetName "DC_Disable_Print_Spooler" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "DC_Disable_Print_Spooler" -GPOTarget "ou=Domain Controllers,$EndPath"
Import-GPO -BackupGpoName "DC_LDAP_signing_requirements" -TargetName "DC_LDAP_signing_requirements" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "DC_LDAP_signing_requirements" -GPOTarget "ou=Domain Controllers,$EndPath"
Import-GPO -BackupGpoName "ADM_Task_Server_Admins as members of Local admins" -TargetName "ADM_Task_Server_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "ADM_Task_Server_Admins as members of Local admins" -GPOTarget $Location
Import-GPO -BackupGpoName "LAPS" -TargetName "LAPS" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "LAPS" -GPOTarget $Location
Import-GPO -BackupGpoName "ADM_Task_Desktop_Admins as members of Local admins" -TargetName "ADM_Task_Desktop_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Desktops,$Location"
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=Laptops,$Location"
Link-GPO -GPOName "ADM_Task_Desktop_Admins as members of Local admins" -GPOTarget "OU=VMs,$Location"
Import-GPO -BackupGpoName "IT Desktop Prefs" -TargetName "IT Desktop Prefs" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Replace
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITGroup -TargetType Group
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName $ITAdminGroup -TargetType Group
Import-GPO -BackupGpoName "CM visual help" -TargetName "CM visual help" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group -Replace
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoApply -TargetName $SID500 -TargetType User
Import-GPO -BackupGpoName "Deploy Notepad++" -TargetName "Deploy Notepad++" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy Notepad++" -GPOTarget $Location
Set-GPPermission -All -PermissionLevel GpoEditDeleteModifySecurity -TargetName "ADM_Task_AD_GPO_Admins" -TargetType Group
#====================================================================
