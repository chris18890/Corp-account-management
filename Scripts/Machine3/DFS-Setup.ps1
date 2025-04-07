#=========================================
#Domain Names in ADS & DNS format, and main OU name
#=========================================
$Domain="$env:userdomain"
$ServerName="$env:computername"
$DNSSuffix=(Get-ADDomain -Identity $Domain).DNSRoot
#=========================================

#=========================================
#Drive where all the folders will be created
#=========================================
$Drive = "D:"
$RootShare = "Store"
#=========================================

#=========================================
#Group Variables
#=========================================
$StaffGroup="Staff"
#=========================================

#=========================================
#Create main store Share
#=========================================
$ShareName = $RootShare
if (!(TEST-PATH "$Drive\$ShareName")) {
    New-Item "$Drive\$ShareName" -type directory -force
} else {
    Write-Host "$Drive\$ShareName already exists" -ForegroundColor Green
}
New-SmbShare -Name $ShareName -Path "$Drive\$ShareName" -FullAccess "authenticated users"
Write-Host "Pausing for 60 seconds after creating share $ShareName"
Start-Sleep -s 60
New-DfsnRootTarget -TargetPath "\\$ServerName\$ShareName" -Path "\\$DNSSuffix\$ShareName"
Add-DfsrMember -GroupName "$ShareName" -ComputerName "$ServerName"
Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -StagingPathQuotaInMB 16384 -Force
Add-DfsrConnection -GroupName "$ShareName" -SourceComputerName "$Domain-DC1" -DestinationComputerName "$ServerName"
#=========================================

#=========================================
#Create Profiles Share
#=========================================
$ShareName = "Profiles"
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
New-DfsnRootTarget -TargetPath "\\$ServerName\$ShareName" -Path "\\$DNSSuffix\$ShareName"
Add-DfsrMember -GroupName "$ShareName" -ComputerName "$ServerName"
Set-DfsrMembership -GroupName "$ShareName" -FolderName "$ShareName" -ContentPath "$Drive\$ShareName" -ComputerName "$ServerName" -StagingPathQuotaInMB 16384 -Force
Add-DfsrConnection -GroupName "$ShareName" -SourceComputerName "$Domain-DC1" -DestinationComputerName "$ServerName"
#=========================================
