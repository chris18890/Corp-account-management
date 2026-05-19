#Requires -Modules ActiveDirectory, Microsoft.Graph
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserName
    ,[Parameter(Mandatory)][string]$EmailSuffix
    ,[Parameter(Mandatory)][ValidateSet("Cloud","Global")][string]$Tier
    ,[ValidateRange(0, 480)][int]$DurationMinutes = 0
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
$ScriptTitle = "$Domain Cloud Admin Enable Script"
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
    $LogFileName = $Domain + "_cloud_admin_enable_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}

#====================================================================
# Apply duration default & cap
#====================================================================
$MaxMinutes = $Env.Security.MaxElevationMinutes
if ($DurationMinutes -eq 0) { $DurationMinutes = $MaxMinutes }

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

if ($DurationMinutes -gt $MaxMinutes) {
    Write-LogFile -LogFile $LogFile -LogString "Requested duration $DurationMinutes exceeds max of $MaxMinutes minutes; capping to $MaxMinutes" -ForegroundColor Yellow
    $DurationMinutes = $MaxMinutes
}

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
# Capture reason (audit requirement)
#====================================================================
if (!$Reason) {
    $Reason = Read-Host "Reason for elevation (recorded in audit log)"
}
if ([string]::IsNullOrWhiteSpace($Reason)) {
    Write-LogFile -LogFile $LogFile -LogString "No reason supplied. Aborting." -ForegroundColor Red
    throw "No reason supplied. Elevation requires a reason for audit."
}

#====================================================================
# Connect to Graph
#====================================================================
if (-not (Get-MgContext)) {
    Connect-MgGraph -NoWelcome -Scopes "User.ReadWrite.All"
}

#====================================================================
# Look up the account
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
if ($MgUser.AccountEnabled) {
    Write-LogFile -LogFile $LogFile -LogString "Account '$UPN' is already enabled. Refreshing the auto-disable window." -ForegroundColor Yellow
}

#====================================================================
# Enable and schedule auto-disable
#====================================================================
$EnableTime  = Get-Date
$DisableTime = $EnableTime.AddMinutes($DurationMinutes)

Write-LogFile -LogFile $LogFile -LogString "Enabling '$UPN' for $DurationMinutes minutes (until $DisableTime). Reason: $Reason"
$Outcome = "Failed"   # default
try {
    Update-MgUser -UserId $MgUser.Id -AccountEnabled:$true -ErrorAction Stop
    # Replace any existing auto-disable task for this account
    Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false
    $DisableScript = Join-Path $ScriptPath "Disable-CloudAdmin.ps1"
    if (!(Test-Path $DisableScript)) {
        throw "Disable-CloudAdmin.ps1 not found at '$DisableScript' - cannot register auto-disable"
    }
    $Action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$DisableScript`" -UserName $UserName -EmailSuffix $EmailSuffix -Tier $Tier -Reason 'Auto-disable after $DurationMinutes-minute elevation window'"
    $Trigger   = New-ScheduledTaskTrigger -Once -At $DisableTime
    $Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest -LogonType S4U
    $Settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Principal $Principal `
        -Settings $Settings `
        -Description "Auto-disable cloud admin '$AccountName' after $DurationMinutes-minute elevation window" `
        -ErrorAction Stop | Out-Null
    Write-LogFile -LogFile $LogFile -LogString "Auto-disable scheduled task '$TaskName' registered for $DisableTime"
    $Outcome = "Success"
} catch {
    # Roll back: if anything failed after the Update-MgUser, disable again
    Write-LogFile -LogFile $LogFile -LogString "ERROR during enable / scheduling: $($_.Exception.Message)" -ForegroundColor Red
    Write-LogFile -LogFile $LogFile -LogString "Rolling back: disabling '$UPN'" -ForegroundColor Yellow
    try {
        Update-MgUser -UserId $MgUser.Id -AccountEnabled:$false -ErrorAction Stop
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ROLLBACK FAILED - '$UPN' may still be enabled with no auto-disable scheduled" -ForegroundColor Red
        $Outcome = "RollbackFailed"
    }
    throw
} finally {
    #====================================================================
    # Append to audit trail CSV
    #====================================================================
    $AuditRow = [pscustomobject]@{
        Timestamp       = $EnableTime.ToString("yyyy-MM-dd HH:mm:ss")
        Action          = "Enable"
        Operator        = "$Domain\$env:USERNAME"
        Account         = $UPN
        Tier            = $Tier
        DurationMinutes = $DurationMinutes
        DisableAt       = $DisableTime.ToString("yyyy-MM-dd HH:mm:ss")
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
Write-LogFile -LogFile $LogFile -LogString "Cloud admin '$UPN' enabled until $DisableTime" -ForegroundColor Green
Write-LogFile -LogFile $LogFile -LogString "Manual early disable: .\Disable-CloudAdmin.ps1 -UserName $UserName -EmailSuffix $EmailSuffix -Tier $Tier" -ForegroundColor Cyan
Write-LogFile -LogFile $LogFile -LogString " "
