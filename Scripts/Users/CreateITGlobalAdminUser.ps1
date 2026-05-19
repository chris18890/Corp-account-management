#Requires -Modules ActiveDirectory, Microsoft.Graph
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [Parameter(Mandatory)][string]$UserName
    , [string]$FirstName,[string]$LastName
    , [string]$Dept,[string]$Company,[string]$Manager
    , [string]$LogFile,[int]$PasswordLength
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = $PSScriptRoot
# Set variables
$ScriptTitle = "$Domain Global Admin IT User Creation Script"
if (!$PasswordLength) {
    $PasswordLength = $Env.Security.PasswordLength
}
if (!$Company) {
    $Company = $Domain
}
if (!$Dept) {
    $Dept = $Env.Groups.IT
}
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_new_Global_Admin_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

$requiredGroups = @("$($Env.Groups.TaskPrefix)HiPriv_Account_Admins", "$($Env.Groups.TaskPrefix)HiPriv_Group_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
$Roles = @("Global Administrator")
#====================================================================
if (-not (Get-MgContext)) {
    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-LogFile -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory","User.ReadWrite.All"
}
if (!$FirstName) {
    $FirstName = ConvertTo-SafeName (READ-HOST 'Enter First Name - ')
}
if (!$LastName) {
    $LastName = ConvertTo-SafeName (READ-HOST 'Enter Last Name - ')
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter Username - '
}
$UserNameGlobalAdmin = ConvertTo-SafeSamAccountName $UserName -Prefix 'ga.'
if (!$Manager) {
    $Manager = ConvertTo-SafeSamAccountName (READ-HOST 'Enter manager username - ')
}
$DisplayName = "$LastName, $FirstName (Global Admin)"
$UserPrincipalName = "$UserNameGlobalAdmin@$EmailSuffix"
$Existing = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction Stop
if ($Existing) {
    Write-LogFile -LogFile $LogFile -LogString "User $UserPrincipalName already exists - aborting" -ForegroundColor Red
    return
} else {
    $PasswordProfile = @{
        Password = New-Password -Length $PasswordLength
        ForceChangePasswordNextSignIn = $true
    }
    New-MgUser -PasswordProfile $PasswordProfile -AccountEnabled:$false -CompanyName $Company -Department $Dept -DisplayName $DisplayName -GivenName $FirstName -MailNickname $UserNameGlobalAdmin -Surname $LastName -UsageLocation $Env.Locale.UsageLocation -UserPrincipalName $UserPrincipalName
    $PasswordProfile.Password = $null
    $ManagerUPN = Get-MgUser -UserID "$Manager@$EmailSuffix" -ErrorAction SilentlyContinue
    if ($ManagerUPN) {
        $ManagerId = $ManagerUPN.Id
        $ManagerRef = @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$ManagerId"}
        Set-MgUserManagerByRef -UserId $UserPrincipalName -BodyParameter $ManagerRef
        Write-LogFile -LogFile $LogFile -LogString "Manager added successfully."
    } else {
        Write-LogFile -LogFile $LogFile -LogString "WARNING: Manager '$Manager' not found in Entra - account created without manager" -ForegroundColor Yellow
    }
}
$MgUserGlobalAdmin = Get-MgUser -UserID "$UserNameGlobalAdmin@$EmailSuffix"
Write-LogFile -LogFile $LogFile -LogString "Assigning roles for $UserNameGlobalAdmin"
foreach ($roleName in $Roles) {
    try {
        $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$roleName'" -ErrorAction Stop
        if (-not $roleDefinition) {
            throw "Role '$roleName' not found in Microsoft Graph"
        }
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "Failed to find role: $roleName" -ForegroundColor Red
        throw
    }
    $ExistingAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($MgUserGlobalAdmin.Id)'" -All
    if ($ExistingAssignments.RoleDefinitionId -notcontains $roleDefinition.Id) {
        try {
            New-MgRoleManagementDirectoryRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $MgUserGlobalAdmin.Id
            Write-LogFile -LogFile $LogFile -LogString "Added $roleName to $UserNameGlobalAdmin"
        } catch {
            throw
        }
    }
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Processing for '$DisplayName' ($UserNameGlobalAdmin) complete"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "
