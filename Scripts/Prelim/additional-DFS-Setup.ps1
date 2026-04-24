#Requires -RunAsAdministrator

# Mode: Standalone / No Shared Modules

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Drive
    ,[string]$Domain
    ,[string]$ServerName
)

Set-StrictMode -Version Latest

#=========================================
#Domain Names in ADS & DNS format, and main OU name
#=========================================
If (!$Domain) {
    $Domain = "$env:userdomain"
}
If (!$ServerName) {
    $ServerName = "$env:computername"
}
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
#=========================================

#=========================================
#Drive where all the folders will be created
#=========================================
$Drive = $Drive.TrimEnd(':') + ':'
$RootShare = "Store"
#=========================================

#=========================================
#Group Variables
#=========================================
$StaffGroup = "Staff"
#=========================================

#=========================================
#Create main store Share
#=========================================
$ShareName = $RootShare
if (!(TEST-PATH "\\$ServerName\$ShareName")) {
    if (!(TEST-PATH "$Drive\$ShareName")) {
        New-Item "$Drive\$ShareName" -type directory -force
    } else {
        Write-Host "$Drive\$ShareName already exists" -ForegroundColor Green
    }
    New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
    Write-Host "Pausing for 60 seconds after creating share $ShareName"
    Start-Sleep -s 60
} else {
    Write-Host "\\$ServerName\$ShareName already exists" -ForegroundColor Green
}
if ((TEST-PATH "\\$ServerName\$ShareName")) {
    New-DfsnRootTarget -TargetPath "\\$ServerName\$ShareName" -Path "\\$DNSSuffix\$ShareName"
    Add-DfsrMember -GroupName "$ShareName" -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -StagingPathQuotaInMB 16384 -Force
    Add-DfsrConnection -GroupName "$ShareName" -SourceComputerName "$Domain-DC1" -DestinationComputerName "$ServerName"
}
#=========================================

#=========================================
#Create Profiles Share
#=========================================
$ShareName = "Profiles"
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
    New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "Administrators", "SYSTEM" -ChangeAccess "authenticated users"
    Write-Host "Pausing for 60 seconds after creating share $ShareName"
    Start-Sleep -s 60
} else {
    Write-Host "\\$ServerName\$ShareName already exists" -ForegroundColor Green
}
if ((TEST-PATH "\\$ServerName\$ShareName")) {
    New-DfsnRootTarget -TargetPath "\\$ServerName\$ShareName" -Path "\\$DNSSuffix\$ShareName"
    Add-DfsrMember -GroupName "$ShareName" -ComputerName "$ServerName"
    Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -StagingPathQuotaInMB 16384 -Force
    Add-DfsrConnection -GroupName "$ShareName" -SourceComputerName "$Domain-DC1" -DestinationComputerName "$ServerName"
}
#=========================================
