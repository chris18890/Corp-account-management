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
$Drive = "D:"
$RootShare = "Store"
#====================================================================

#====================================================================
#Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
$InstallerGroup = "Installers"
$ITGroup = "IT"
$ITAdminGroup = "IT_Admin"
$UserPasswordDelegationOU = "OU=$StaffGroup,$EndPath"
$UserPasswordDelegationGroup = "RG_Password_Admins"
$SID500 = "Administrator"
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
Create-ADOU -OUName "Groups" -Path $EndPath -OUDescription "Top level OU for Group objects"
Create-ADGroup -GroupName $StaffGroup -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Org-wide group for all users"
Create-ADGroup -GroupName $ITGroup -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Departmental group holding all IT accounts"
Create-ADGroup -GroupName "UG_Office365" -Path "OU=$GroupsOU,$EndPath" -GroupDescription "Used to assign Office365 licenses"
Create-ADOU -OUName $ITGroup -Path $EndPath -OUDescription "Top level OU for IT User & Group objects"
Create-ADOU -OUName "Shared_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT OU for Group objects that grant access to shared mailboxes"
Create-ADOU -OUName "Equipment_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT OU for Group objects that grant access to equipment mailboxes"
Create-ADOU -OUName "Room_Mailbox_Access" -Path "OU=$GroupsOU,$EndPath" -OUDescription "IT OU for Group objects that grant access to room mailboxes"
Create-ADOU -OUName "Equipment_Mailbox_Accounts" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for User objects that are equipment mailbox recipient types"
Create-ADOU -OUName "Hi_Priv_Groups" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for Group objects that control Hi-Priv access"
Create-ADOU -OUName "Hi_Priv_Accounts" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for User objects that are Hi-Priv accounts"
Create-ADOU -OUName "Room_Mailbox_Accounts" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for User objects that are room mailbox recipient types"
Create-ADOU -OUName "Shared_Mailbox_Accounts" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for User objects that are shared mailbox recipient types"
Create-ADOU -OUName "Service_Accounts" -Path "OU=$ITGroup,$EndPath" -OUDescription "IT OU for User objects that are service accounts"
Create-ADGroup -GroupName $ITAdminGroup -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Group holding all IT Admin accounts"
Create-ADGroup -GroupName "RG_Account_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are indirect members of the Account Operators BuiltIn group"
Create-ADGroup -GroupName "RG_Desktop_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are added to Local admin on all computers in the Desktop & Laptop OUs"
Create-ADGroup -GroupName "RG_DFS_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members have Full Control NTFS permissions on all DFS folder & have access to DFS console"
Create-ADGroup -GroupName "RG_DHCP_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are members of DHCP Administrators"
Create-ADGroup -GroupName "RG_DHCP_Users" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are members of DHCP Users"
Create-ADGroup -GroupName "RG_DNS_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are indirect members of the DNSAdmins BuiltIn group"
Create-ADGroup -GroupName "RG_$InstallerGroup" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members have permission to join & move computer objects in $ParentOU & Sub OUs"
Create-ADGroup -GroupName "UG_Level_1_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Level 1 admins - desktop, includes all level 2 admins"
Create-ADGroup -GroupName "UG_Level_2_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Level 2 admins - junior server, includes all level 3 admins"
Create-ADGroup -GroupName "UG_Level_3_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Level 3 admins - senior server"
Create-ADGroup -GroupName "$UserPasswordDelegationGroup" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members can reset passwords of users in the $StaffGroup OU"
Create-ADGroup -GroupName "RG_Server_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members are added to Local admin on all computers in $ParentOU & Sub OUs via GPO, are indirect members of the Server Operators BuiltIn group"
Create-ADGroup -GroupName "RG_Subscription_Contributors" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members have Contributor permissions on the subscription"
Create-ADGroup -GroupName "RG_Subscription_Owners" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members have Owner permissions on the subscription"
Create-ADGroup -GroupName "RG_Subscription_User_Access_Admins" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members have User Access Admin permissions on the subscription"
Create-ADGroup -GroupName "RG_WDS_Deploy_Clients" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Clients folder"
Create-ADGroup -GroupName "RG_WDS_Deploy_Servers" -Path "OU=Hi_Priv_Groups,OU=$ITGroup,$EndPath" -GroupDescription "Members can use WDS to deploy images in the Servers folder"
Add-GroupMember -group $StaffGroup -Member $ITGroup
Add-GroupMember -group $ITGroup -Member $ITAdminGroup
Add-GroupMember -group "Account Operators" -Member "RG_Account_Admins"
Add-GroupMember -group "Remote Desktop Users" -Member "RG_Server_Admins"
Add-GroupMember -group "Server Operators" -Member "RG_Server_Admins"
Add-GroupMember -group "DnsAdmins" -Member "RG_DNS_Admins"
Add-GroupMember -group $ITAdminGroup -Member $SID500
Add-GroupMember -group "RG_Account_Admins" -Member "UG_Level_2_Admins"
Add-GroupMember -group "RG_Desktop_Admins" -Member "UG_Level_1_Admins"
Add-GroupMember -group "RG_$InstallerGroup" -Member "UG_Level_1_Admins"
Add-GroupMember -group "RG_DHCP_Admins" -Member "UG_Level_3_Admins"
Add-GroupMember -group "RG_DHCP_Users" -Member "UG_Level_2_Admins"
Add-GroupMember -group "RG_DFS_Admins" -Member "UG_Level_3_Admins"
Add-GroupMember -group "RG_DNS_Admins" -Member "UG_Level_3_Admins"
Add-GroupMember -group "$UserPasswordDelegationGroup" -Member "UG_Level_1_Admins"
Add-GroupMember -group "RG_Server_Admins" -Member "UG_Level_2_Admins"
Add-GroupMember -group "RG_Subscription_Contributors" -Member "UG_Level_2_Admins"
Add-GroupMember -group "RG_WDS_Deploy_Servers" -Member "UG_Level_2_Admins"
Add-GroupMember -group "RG_WDS_Deploy_Clients" -Member "UG_Level_1_Admins"
Add-GroupMember -group "UG_Level_1_Admins" -Member "UG_Level_2_Admins"
Add-GroupMember -group "UG_Level_2_Admins" -Member "UG_Level_3_Admins"
Get-ADUser $SID500 | Move-ADObject -TargetPath "OU=Hi_Priv_Accounts,OU=$ITGroup,$EndPath"
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
$UserPasswordDelegationGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $UserPasswordDelegationGroup).SID
$ResetUserPasswordACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $UserPasswordDelegationGroupSID,"ExtendedRight","Allow",$ExtendedRightsMap["Reset Password"],'Descendents',$GuidMap["user"]
$ReadForceChangeUserPasswordACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $UserPasswordDelegationGroupSID,"ReadProperty","Allow",$GuidMap["pwdLastSet"],'Descendents',$GuidMap["user"]
$WriteForceChangeUserPasswordACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $UserPasswordDelegationGroupSID,"WriteProperty","Allow",$GuidMap["pwdLastSet"],'Descendents',$GuidMap["user"]
$ReadUnlockUserAccountACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $UserPasswordDelegationGroupSID,"ReadProperty","Allow",$GuidMap["lockoutTime"],'Descendents',$GuidMap["user"]
$WriteUnlockUserAccountACE = New-Object System.DirectoryServices.ActiveDirectoryAccessRule $UserPasswordDelegationGroupSID,"WriteProperty","Allow",$GuidMap["lockoutTime"],'Descendents',$GuidMap["user"]
$Acl = Get-Acl "AD:\$UserPasswordDelegationOU"
$Acl.AddAccessRule($ResetUserPasswordACE)
$Acl.AddAccessRule($ReadForceChangeUserPasswordACE)
$Acl.AddAccessRule($WriteForceChangeUserPasswordACE)
$Acl.AddAccessRule($ReadUnlockUserAccountACE)
$Acl.AddAccessRule($WriteUnlockUserAccountACE)
$Acl | Set-Acl
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
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "RG_DFS_Admins"
    New-DfsReplicationGroup -GroupName "$ShareName" | New-DfsReplicatedFolder -FolderName "$ShareName" -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName "$ShareName" -AccountName "RG_DFS_Admins"
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
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("RG_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
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
    New-DfsnRoot -TargetPath "\\$ServerName\$ShareName" -Type DomainV2 -Path "\\$DNSSuffix\$ShareName" -GrantAdminAccounts "RG_DFS_Admins"
    New-DfsReplicationGroup -GroupName "$ShareName" | New-DfsReplicatedFolder -FolderName "$ShareName" -DfsnPath "\\$DNSSuffix\$ShareName" | Add-DfsrMember -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -PrimaryMember $True -StagingPathQuotaInMB 16384 -Force
    Grant-DfsrDelegation -GroupName "$ShareName" -AccountName "RG_DFS_Admins"
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
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("RG_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
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
    $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("RG_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
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
# Create OU for laptops
#====================================================================
Create-ADOU -OUName "Laptops" -Path $Location
#====================================================================

#====================================================================
# Create OU for desktops
#====================================================================
Create-ADOU -OUName "Desktops" -Path $Location
#====================================================================

#====================================================================
#Redirect default computer location
#====================================================================
redircmp "$Location"
$InstallerGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup "RG_$InstallerGroup").SID
$Acl = Get-ACL "AD:\$Location"
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"CreateChild","Allow","All",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"DeleteChild","Allow","All",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"ReadProperty","Allow","Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"WriteProperty","Allow","Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"ReadControl","Allow","Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"WriteDacl","Allow","Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"Self","Allow",$ExtendedRightsMap["Validated write to DNS host name"],"Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $InstallerGroupSID,"Self","Allow",$ExtendedRightsMap["Validated write to service principal name"],"Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $InstallerGroupSID,"Allow",$ExtendedRightsMap["Reset Password"],"Descendents",$GuidMap["computer"]))
$Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $InstallerGroupSID,"Allow",$ExtendedRightsMap["Change Password"],"Descendents",$GuidMap["computer"]))
$Acl | Set-ACL
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
Import-GPO -BackupGpoName "RG_Server_Admins as members of Local admins" -TargetName "RG_Server_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "RG_Server_Admins as members of Local admins" -GPOTarget "$Location"
Import-GPO -BackupGpoName "RG_Desktop_Admins as members of Local admins" -TargetName "RG_Desktop_Admins as members of Local admins" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "RG_Desktop_Admins as members of Local admins" -GPOTarget "OU=Desktops,$Location"
Link-GPO -GPOName "RG_Desktop_Admins as members of Local admins" -GPOTarget "OU=Laptops,$Location"
Import-GPO -BackupGpoName "IT Desktop Prefs" -TargetName "IT Desktop Prefs" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "IT Desktop Prefs" -GPOTarget "OU=Hi_Priv_Accounts,OU=IT,$EndPath"
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel None -TargetName "Authenticated Users" -TargetType Group
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName "$ITGroup" -TargetType Group
Set-GPPermission -Name "IT Desktop Prefs" -PermissionLevel GpoApply -TargetName "$ITAdminGroup" -TargetType Group
Import-GPO -BackupGpoName "CM visual help" -TargetName "CM visual help" -path $GPOLocation -CreateIfNeeded
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=$StaffGroup,$EndPath"
Link-GPO -GPOName "CM visual help" -GPOTarget "OU=Hi_Priv_Accounts,OU=IT,$EndPath"
Set-GPPermission -Name "CM visual help" -PermissionLevel None -TargetName "Authenticated Users" -TargetType Group
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoRead -TargetName "Authenticated Users" -TargetType Group
Set-GPPermission -Name "CM visual help" -PermissionLevel GpoApply -TargetName $SID500 -TargetType User
Import-GPO -BackupGpoName "Deploy Firefox" -TargetName "Deploy Firefox" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy Firefox" -GPOTarget "$Location"
Import-GPO -BackupGpoName "Deploy Notepad++" -TargetName "Deploy Notepad++" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy Notepad++" -GPOTarget "$Location"
Import-GPO -BackupGpoName "Deploy PuTTY" -TargetName "Deploy PuTTY" -path $GPOLocation -MigrationTable "$GPOLocation\admins.migtable" -CreateIfNeeded
Link-GPO -GPOName "Deploy PuTTY" -GPOTarget "$Location"
#====================================================================
