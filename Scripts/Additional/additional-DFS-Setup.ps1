#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Drive
    ,[string]$Domain
    ,[string]$ServerName
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

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
$RootShare = $Env.Shares.Root
#=========================================

#=========================================
#Group Variables
#=========================================
$StaffGroup = $Env.Groups.Staff
#=========================================

#=========================================
#Create main store Share
#=========================================
$ShareName = $RootShare
if (!(Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
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
$ShareName = $Env.Shares.Profiles
if (!(Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    if (!(TEST-PATH "$Drive\$ShareName")) {
        New-Item "$Drive\$ShareName" -type directory -force
        $Acl = Get-Acl "$Drive\$ShareName"
        $isProtected = $true
        $preserveInheritance = $false
        $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
        $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($StaffGroup,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
        $Acl.SetAccessRule($Ar)
        $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("$($Env.Groups.TaskPrefix)DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
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
