#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName
    ,[Parameter(Mandatory)][string]$UserName
    ,[Parameter(Mandatory)][ValidateSet("E","R")][string]$UserAction
    ,[ValidateSet("P","T")][string]$TempOrPerm
    ,[ValidateRange(1, [int]::MaxValue)][int]$TimeSpan
    ,[string]$Reason
    ,[string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking

function Write-ElevationAuditRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AuditFile,
        [Parameter(Mandatory)][datetime]$Timestamp,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Operator,
        [Parameter(Mandatory)][string]$Account,
        [Parameter(Mandatory)][string]$Group,
        [AllowNull()][AllowEmptyString()][string]$DurationMinutes,
        [AllowNull()][AllowEmptyString()][string]$Reason,
        [Parameter(Mandatory)][string]$Outcome,
        [Parameter(Mandatory)][string]$LogFile
    )
    $AuditRow = [pscustomobject]@{
        Timestamp       = $Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
        Action          = $Action
        Operator        = $Operator
        Account         = $Account
        Group           = $Group
        DurationMinutes = $DurationMinutes
        Reason          = $Reason
        Outcome         = $Outcome
    }
    try {
        if (Test-Path $AuditFile) {
            $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation -Append
        } else {
            $AuditRow | Export-Csv -Path $AuditFile -NoTypeInformation
        }
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "WARNING: could not write audit row to '$AuditFile': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$Env = Get-EnvironmentConfig
$Domain = "$env:userdomain"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain User Elevation Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
$AuditFile = Join-Path $LogPath "on-prem-elevations.csv"
if (!$LogFile) {
    $LogFileName = $Domain + "_user_elevation_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
}
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-LogFile -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-LogFile -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-LogFile -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-LogFile -LogFile $LogFile -LogString "$ScriptTitle"
Write-LogFile -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
Write-LogFile -LogFile $LogFile -LogString " "

#====================================================================
# Authorisation tiers
#   - Protected groups (forest-admin + Tier-0 role) may only be
#     modified by a high-tier admin, in EITHER direction (E or R).
#   - Standard group admins may act on everything else.
#====================================================================
$ProtectedGroups = @(
    'Domain Admins'
    'Enterprise Admins'
    'Schema Admins'
    "$($Env.Groups.RolePrefix)Tier0_Level_3_Admins"
)
$StandardAdminGroups = @("$($Env.Groups.TaskPrefix)Standard_Group_Admins", "$($Env.Groups.TaskPrefix)SER_Access_Admins", "$($Env.Groups.TaskPrefix)Local_Admin_Group_Admins", 'Domain Admins')
$HighTierAdminGroups = @("$($Env.Groups.TaskPrefix)HiPriv_Group_Admins", 'Domain Admins')
$InvokerIsHighTier = Test-IsMemberOf -Sam $env:USERNAME -GroupNames $HighTierAdminGroups -DCHostName $DCHostName
$InvokerIsStandard = Test-IsMemberOf -Sam $env:USERNAME -GroupNames $StandardAdminGroups -DCHostName $DCHostName
if (-not ($InvokerIsHighTier -or $InvokerIsStandard)) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    $UnauthorisedAuditReason = if ([string]::IsNullOrWhiteSpace($Reason)) {
        "Unauthorised execution attempt before reason capture"
    } else {
        $Reason
    }
    Write-ElevationAuditRow -AuditFile $AuditFile -Timestamp (Get-Date) -Action $(switch ($UserAction.ToUpper()) {
        "E"     { "Elevate" }
        "R"     { "Remove" }
        default { $UserAction }
    }) -Operator "$Domain\$env:USERNAME" -Account $UserName -Group $GroupName -DurationMinutes "" -Reason $UnauthorisedAuditReason -Outcome "Denied" -LogFile $LogFile
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Capture reason for audit
#====================================================================
if (!$Reason) {
    $Reason = Read-Host "Reason for elevation/removal (recorded in audit log)"
}
if ([string]::IsNullOrWhiteSpace($Reason)) {
    Write-LogFile -LogFile $LogFile -LogString "No reason supplied. Aborting." -ForegroundColor Red
    Write-ElevationAuditRow -AuditFile $AuditFile -Timestamp (Get-Date) -Action $(switch ($UserAction.ToUpper()) {
        "E"     { "Elevate" }
        "R"     { "Remove" }
        default { $UserAction }
    }) -Operator "$Domain\$env:USERNAME" -Account $UserName -Group $GroupName -DurationMinutes "" -Reason "No reason supplied" -Outcome "Rejected" -LogFile $LogFile
    throw "No reason supplied. Elevation requires a reason for audit."
}

#====================================================================
# Action - wrapped in try/finally so the audit row is written even on
# early return or thrown exception. Denials are audited too.
#====================================================================
$ActionTime = Get-Date
$Outcome = "Failed" # Defaults; flipped on the success / decision paths
$AuditDuration = "" # Populated for temporal elevations
try {
    # Resolve the target group first so protected-group detection cannot be
    # bypassed by passing an alternate identity form, such as a distinguishedName.
    $TargetGroupObj = Get-ADGroup -Identity $GroupName -Server $DCHostName -ErrorAction Stop
    $TargetIsProtected = $ProtectedGroups -contains $TargetGroupObj.Name
    #================================================================
    # Protected-group gates (apply to BOTH elevate and remove)
    #================================================================
    if ($TargetIsProtected) {
        if (-not $InvokerIsHighTier) {
            Write-LogFile -LogFile $LogFile -LogString "DENIED: '$($TargetGroupObj.Name)' is a protected/Tier-0 group; only high-tier admins may modify it." -ForegroundColor Red
            $Outcome = "Denied"
            return
        }
    }
    switch ($UserAction.ToUpper()) {
        "E" {
            if (!$TempOrPerm) {
                $TempOrPerm = READ-HOST 'Enter a duration;  P for permanent, T for temporary - '
            }
            $TempOrPerm = $TempOrPerm.ToUpper()
            # Only Tier-0 / Level-3 accounts may be placed into a forest-admin group
            if ($TargetIsProtected) {
                if (-not (Test-IsMemberOf -Sam $UserName -GroupNames @("$($Env.Groups.RolePrefix)Tier0_Level_3_Admins") -DCHostName $DCHostName)) {
                    Write-LogFile -LogFile $LogFile -LogString "DENIED: tried to add a non-Tier-0 / non-Level-3 account ($UserName) to $GroupName." -ForegroundColor Red
                    $Outcome = "Denied"
                    return
                }
            }
            switch ($TempOrPerm.ToUpper()) {
                "P" {
                    Write-LogFile -LogFile $LogFile -LogString "Adding $UserName to $GroupName with no time limit, manual removal required. Reason: $Reason" -ForegroundColor Yellow
                    try {
                        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $UserName
                        $Outcome = "Success"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR: $($_.Exception.Message)" -ForegroundColor Red
                        $Outcome = "Failed"
                    }
                }
                "T" {
                    if (!$TimeSpan) {
                        Write-LogFile -LogFile $LogFile -LogString "No timespan specified, using default of 60 minutes"
                        $TimeSpan = 60
                    }
                    if ($TimeSpan -gt $Env.Security.MaxElevationMinutes) {
                        Write-LogFile -LogFile $LogFile -LogString "Requested timespan is longer than $($Env.Security.MaxElevationMinutes) minutes / $($Env.Security.MaxElevationMinutes / 60) hours, please rerun with a lower value" -ForegroundColor Red
                        $AuditDuration = $TimeSpan
                        $Outcome = "Rejected"
                        return
                    }
                    $AuditDuration = $TimeSpan
                    Write-LogFile -LogFile $LogFile -LogString "Adding $UserName to $GroupName for $TimeSpan minutes. Reason: $Reason" -ForegroundColor Yellow
                    try {
                        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $UserName -TimeSpan $TimeSpan
                        $Outcome = "Success"
                    } catch {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR: $($_.Exception.Message)" -ForegroundColor Red
                        $Outcome = "Failed"
                    }
                }
                default {
                    Write-LogFile -LogFile $LogFile -LogString "No valid option specified, quitting" -ForegroundColor Red
                    $Outcome = "Rejected"
                    return
                }
            }
        }
        "R" {
            Write-LogFile -LogFile $LogFile -LogString "Removing $UserName from $GroupName. Reason: $Reason" -ForegroundColor Yellow
            try {
                $Outcome = Remove-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $UserName
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "ERROR: $($_.Exception.Message)" -ForegroundColor Red
                $Outcome = "Failed"
            }
        }
    }
} finally {
    #================================================================
    # Append audit trail row regardless of outcome
    #================================================================
    $AuditAction = switch ($UserAction.ToUpper()) {
        "E"     { "Elevate" }
        "R"     { "Remove" }
        default { $UserAction }
    }
    Write-ElevationAuditRow -AuditFile $AuditFile -Timestamp $ActionTime -Action $AuditAction -Operator "$Domain\$env:USERNAME" -Account $UserName -Group $GroupName -DurationMinutes $AuditDuration -Reason $Reason -Outcome $Outcome -LogFile $LogFile
}
#====================================================================
