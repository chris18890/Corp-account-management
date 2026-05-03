#Requires -Modules ActiveDirectory, Microsoft.Graph
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [string]$UserName,[string]$EmailSuffix
    , [string]$Manager,[string]$Company,[string]$Dept
    , [string]$FirstName,[string]$LastName
    , [string]$LogFile,[int]$PasswordLength
)

Set-StrictMode -Version Latest
Add-Type -Assembly System.Web

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain Global Admin IT User Creation Script"
if (!$PasswordLength) {
    $PasswordLength = 20
}
if (!$Company) {
    $Company = $Domain
}
if (!$Dept) {
    $Dept = "IT"
}
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_new_Global_Admin_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
    Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
    Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
    Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString " "
}
$requiredGroups = @('ADM_Task_HiPriv_Account_Admins', 'ADM_Task_HiPriv_Group_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
$Roles = @("Global Administrator")
#====================================================================
if (-not (Get-MgContext)) {
    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory","User.ReadWrite.All"
}
if (!$FirstName) {
    $FirstName = READ-HOST 'Enter First Name - '
    $FirstName = $FirstName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from First name for Office 365 compliance. Note that \ is escaped to \\
}
if (!$LastName) {
    $LastName = READ-HOST 'Enter Last Name - '
    $LastName = $LastName.Trim() -replace '[?@\\+]', [String]::Empty # Strip out illegal characters from LastName for Office 365 compliance. Note that \ is escaped to \\
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter Username - '
    $UserName = $UserName.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
    $UserNameGlobalAdmin = "ga." + $UserName
} else {
    $UserNameGlobalAdmin = "ga." + $UserName
}
if (!$Manager) {
    $Manager = READ-HOST 'Enter manager username - '
    $Manager = $Manager.Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty # Strip out illegal characters from User ID
}
$DisplayName = "$LastName, $FirstName (Global Admin)"
$UserPrincipalName = "$UserNameGlobalAdmin@$EmailSuffix"
$Existing = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
if ($Existing) {
    Write-Log -LogFile $LogFile -LogString "User $UserPrincipalName already exists – aborting" -ForegroundColor Red
    return
} else {
    $PasswordProfile = @{
        Password = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
        ForceChangePasswordNextSignIn = $true
    }
    New-MgUser -PasswordProfile $PasswordProfile -AccountEnabled:$false -CompanyName $Company -Department $Dept -DisplayName $DisplayName -GivenName $FirstName -MailNickname $UserNameGlobalAdmin -Surname $LastName -UsageLocation "GB" -UserPrincipalName $UserPrincipalName
    $PasswordProfile.Password = $null
    $ManagerUPN = Get-MgUser -UserID "$Manager@$EmailSuffix"
    if ($ManagerUPN) {
        $ManagerId = $ManagerUPN.Id
        $ManagerRef = @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$ManagerId"}
        Set-MgUserManagerByRef -UserId $UserPrincipalName -BodyParameter $ManagerRef
        Write-Log -LogFile $LogFile -LogString "Manager added successfully."
    }
}
$MgUserGlobalAdmin = Get-MgUser -UserID "$UserNameGlobalAdmin@$EmailSuffix"
foreach ($roleName in $Roles) {
    Write-Log -LogFile $LogFile -LogString "Assigning roles for $UserNameGlobalAdmin"
    try {
        $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$roleName'"
        if (-not $roleDefinition) {
            throw "Role '$roleName' not found in Microsoft Graph"
        }
    } catch {
        Write-Log -LogFile $LogFile -LogString "Failed to find role: $_" -ForegroundColor Red
        throw
    }
    $ExistingAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($MgUserGlobalAdmin.Id)'" -All
    if ($ExistingAssignments.RoleDefinitionId -notcontains $roleDefinition.Id) {
        try {
            New-MgRoleManagementDirectoryRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $MgUserGlobalAdmin.Id
        } catch {
            throw
        }
    }
}
