# Example Syntax:
# Powershell .\CreateLocalAdminGroups.ps1 -ComputerOU Servers

[CmdletBinding()]
Param(
    [string]$ComputerOU
)

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$ParentOU = "Domain Computers"
$Location = "OU=$ParentOU,$EndPath"
$GroupsOU = "OU=Local_Admin_Groups,OU=Administration,$EndPath"

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

if(-not($ComputerOU)) {
    $ComputerOU = Read-Host -Prompt "You must provide a DistinguishedName for the Computers OU - e.g. OU=Servers,$Location"
}
$ComputerOU = "OU=$ComputerOU,$Location"

Import-module activedirectory
foreach ($computer in (Get-ADComputer -SearchBase $ComputerOU -Filter *)) {
    $CompName = $computer.Name
    $GroupName = "ADM_Task_Local_Admin_"+$CompName
    $GroupDesc = "User Group: Local Admin Users for $CompName"
    Create-ADGroup -GroupName $GroupName -GroupScope DomainLocal -Path $GroupsOU -GroupDescription $GroupDesc
}

# Clean up groups for machines that have been removed
#Get a list of all Local Admin Groups and report on the number
$list = Get-ADGroup -Filter {name -like "ADM_Task_Local_Admin_*"} -SearchBase $GroupsOU | select name
"Total Local Admin groups before cleaning:"
$list.count

foreach ($c in $List){
    #Split to get the computer name only and remove the trailing character
    $splitname = @($c -split('_'))
    $trimname = $splitname[2].Trim("}")
    
    #See if the computer exists, and remove the group if not
    try {Get-ADComputer $trimname}
    catch {Remove-AdGroup $c.Name -Confirm:$false }
}

#Get a list of all Local Admin Groups and report on the number after cleaning
$list = Get-ADGroup -Filter {name -like "ADM_Task_Local_Admin_*"} -SearchBase $GroupsOU | select name
"Total Local Admin groups after cleaning:"
$list.count
