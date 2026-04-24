#Requires -Modules ActiveDirectory, Microsoft.Graph
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [string]$UserName,[string]$EmailSuffix
    , [string]$Manager,[string]$Company,[string]$Dept
    , [string]$FirstName,[string]$LastName
    , [string]$LogFile,[int]$PasswordLength
)

Add-Type -Assembly System.Web

#====================================================================
# Set up logging
#====================================================================
function Write-Log {
    param([string]$LogString,[string]$ForegroundColor)
    #================================================================
    # Purpose:          To write a string with a date and time stamp to a log file
    # Assumptions:      $LogFile set with path to log file to write to
    # Effects:
    # Inputs:
    # $LogString:       String to write to log file
    # Calls:
    # Returns:
    # Notes:
    #================================================================
    "$(Get-Date -Format 'G') $LogString" | Out-File -Filepath $LogFile -Append -Encoding UTF8
    if ($ForegroundColor) {
        Write-Host $LogString -ForegroundColor $ForegroundColor
    } else {
        Write-Host $LogString
    }
}
#====================================================================

#====================================================================
# Test password against password policy
#====================================================================
function Test-Password {
    param([string]$Password)
    #================================================================
    # Purpose:          Test password against password policy
    # Assumptions:      Password has been generated with enough characters for required groups
    # Effects:          Password should be valid
    # Inputs:           $Password
    # Calls:            Write-Log function
    # Returns:
    # Notes:            There are 4 requirements in the current policy, but this could change in future
    #================================================================
    $TestsPassed = 0
    if ($Password.length -ge ($PasswordLength)) {$TestsPassed ++} # Must be >= 15 characters in length
    if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain a lowercase letter
    if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain an uppercase letter
    if ($Password -cmatch "[0-9]") {$TestsPassed ++} # Must contain a digit
    #if (-Not($Password -notmatch "[a-zA-Z0-9]")) {$TestsPassed ++} # Must contain a special character
    if ($TestsPassed -ge 4) {
        Write-Log "Password validated"
        Write-Log ""
    } else {
        Write-Log ("-" * 80) -ForegroundColor Red
        Write-Log "ERROR: Password does not comply with the password policy, script terminating" -ForegroundColor Red
        Write-Log ("-" * 80) -ForegroundColor Red
        return
    }
}
#====================================================================

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
    Write-Log "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_new_Global_Admin_user_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log ("=" * 80)
    Write-Log "Log file is '$LogFile'"
    Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log "Script path is '$ScriptPath'"
    Write-Log "$ScriptTitle"
    Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log ("=" * 80)
    Write-Log ""
}
$requiredGroups = @('ADM_Task_HiPriv_Account_Admins', 'ADM_Task_HiPriv_Group_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
$Roles = @("Global Administrator")
#====================================================================
if (-not (Get-MgContext)) {
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
    Write-Log "User $UserPrincipalName already exists – aborting" -ForegroundColor Red
    return
} else {
    $PasswordProfile = @{
        Password = [Web.Security.Membership]::GeneratePassword($PasswordLength,4)
        ForceChangePasswordNextSignIn = $true
    }
    New-MgUser -PasswordProfile $PasswordProfile -AccountEnabled:$false -CompanyName $Company -Department $Dept -DisplayName $DisplayName -GivenName $FirstName -MailNickname $UserNameGlobalAdmin -Surname $LastName -UsageLocation "GB" -UserPrincipalName $UserPrincipalName
    $ManagerUPN = Get-MgUser -UserID "$Manager@$EmailSuffix"
    if ($ManagerUPN) {
        $ManagerId = $ManagerUPN.Id
        $ManagerRef = @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$ManagerId"}
        Set-MgUserManagerByRef -UserId $UserPrincipalName -BodyParameter $ManagerRef
        Write-Log "Manager added successfully."
    }
}
$MgUserGlobalAdmin = Get-MgUser -UserID "$UserNameGlobalAdmin@$EmailSuffix"
foreach ($roleName in $Roles) {
    Write-Log "Assigning roles for $UserNameGlobalAdmin"
    $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$roleName'"
    $ExistingAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($MgUserGlobalAdmin.Id)'" -All
    if ($ExistingAssignments.RoleDefinitionId -notcontains $roleDefinition.Id) {
        New-MgRoleManagementDirectoryRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $MgUserGlobalAdmin.Id
    }
}
