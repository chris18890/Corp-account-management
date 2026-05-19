#Requires -Modules ActiveDirectory, Microsoft.Graph
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserName
    ,[Parameter(Mandatory)][string]$EmailSuffix
    ,[Parameter(Mandatory)][ValidateSet("Cloud","Global")][string]$Tier
    ,[string]$Reason
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

#====================================================================
# Variables
#====================================================================
$Domain = "$env:userdomain"
$DCHostName = (Get-ADDomain).PDCEmulator
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain Cloud Admin Disable Script"
$Prefix = switch ($Tier) {
    "Cloud"  { "ca." }
    "Global" { "ga." }
}
$AccountName = "$Prefix$UserName"
$UPN = "$AccountName@$EmailSuffix"
$AuditFile = Join-Path $ScriptPath "LogFiles\cloud-admin-elevations.csv"
$TaskName = "CloudAdminAutoDisable_$AccountName"
$LogPath = "$ScriptPath\LogFiles"
if (!(Test-Path $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -Type Directory -Force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_cloud_admin_disable_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}

#====================================================================
# Log header
#====================================================================
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

#====================================================================
# Authorisation
#====================================================================
$requiredGroups = @("$($Env.Groups.TaskPrefix)HiPriv_Account_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Connect to Graph
#====================================================================
if (-not (Get-MgContext)) {
    Connect-MgGraph -NoWelcome -Scopes "User.ReadWrite.All"
}

#====================================================================
# Look up account
#====================================================================
try {
    $MgUser = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction Stop
} catch {
    Write-LogFile -LogFile $LogFile -LogString "Cloud admin account '$UPN' not found in Entra ID. Aborting." -ForegroundColor Red
    throw
}
if (-not $MgUser) {
    Write-LogFile -LogFile $LogFile -LogString "Cloud admin account '$UPN' not found in Entra ID. Aborting." -ForegroundColor Red
    throw "Account '$UPN' not found"
}
$Outcome = "Failed"   # default
$DisableTime = Get-Date
try {
    #====================================================================
    # Disable account
    #====================================================================
    if (!$MgUser.AccountEnabled) {
        Write-LogFile -LogFile $LogFile -LogString "Account '$UPN' is already disabled. Cleaning up any scheduled task." -ForegroundColor Green
        $Outcome = "AlreadyDisabled"
    } else {
        Write-LogFile -LogFile $LogFile -LogString "Disabling '$UPN'"
        try {
            Update-MgUser -UserId $MgUser.Id -AccountEnabled:$false -ErrorAction Stop
            Write-LogFile -LogFile $LogFile -LogString "Disabled '$UPN'"
            $Outcome = "Success"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "ERROR disabling account: $($_.Exception.Message)" -ForegroundColor Red
            $Outcome = "Failed"
            throw
        }
    }
    
    #====================================================================
    # Unregister any pending auto-disable task
    #====================================================================
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($ExistingTask) {
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
            Write-LogFile -LogFile $LogFile -LogString "Unregistered pending auto-disable task '$TaskName'"
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: could not unregister scheduled task '$TaskName': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-LogFile -LogFile $LogFile -LogString "No pending auto-disable task to unregister"
    }
} finally {
    #====================================================================
    # Append to audit trail CSV
    #====================================================================
    if (!$Reason) { $Reason = "Manual disable" }
    $AuditRow = [pscustomobject]@{
        Timestamp       = $DisableTime.ToString("yyyy-MM-dd HH:mm:ss")
        Action          = "Disable"
        Operator        = "$Domain\$env:USERNAME"
        Account         = $UPN
        Tier            = $Tier
        DurationMinutes = ""
        DisableAt       = ""
        Reason          = $Reason
        Outcome         = $Outcome
    }
    if (Test-Path $AuditFile) {
        $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation -Append
    } else {
        $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation
    }
}

Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Done."
Write-LogFile -LogFile $LogFile -LogString " "
Write-LogFile -LogFile $LogFile -LogString "Cloud admin '$UPN' is now disabled." -ForegroundColor Green
Write-LogFile -LogFile $LogFile -LogString " "
