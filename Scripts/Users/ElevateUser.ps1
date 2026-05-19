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

$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Group_Admins", "$($Env.Groups.TaskPrefix)HiPriv_Group_Admins", 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
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
    throw "No reason supplied. Elevation requires a reason for audit."
}

#====================================================================
# Action - wrapped in try/finally so the audit row is written even on
# early return or thrown exception.
#====================================================================
$ActionTime = Get-Date
$Outcome = "Failed"          # Defaults; flipped on the success paths
$AuditDuration = ""          # Populated for temporal elevations
try {
    switch ($UserAction.ToUpper()) {
        "E" {
            if (!$TempOrPerm) {
                $TempOrPerm = READ-HOST 'Enter a duration;  P for permanent, T for temporary - '
            }
            $TempOrPerm = $TempOrPerm.ToUpper()
            if (!(Get-ADOptionalFeature -Filter "Name -eq 'Privileged Access Management Feature'" -Server $DCHostName)) {
                Write-LogFile -LogFile $LogFile -LogString "ERROR: AD PAM feature not enabled" -ForegroundColor Red
                throw
            }
            if ($GroupName -eq "Domain Admins" -or $GroupName -eq "Enterprise Admins" -or $GroupName -eq "Schema Admins") {
                $ProtectedGroups = @("$($Env.Groups.RolePrefix)Tier0_Level_3_Admins")
                $groups = $ProtectedGroups | ForEach-Object {
                    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $UserName
                }
                if (-not $groups) {
                    Write-LogFile -LogFile $LogFile -LogString "Tried to add a non-tier 0, non-level 3 account to $GroupName. Aborting." -ForegroundColor Red
                    throw "Tried to add a non-tier 0, non-level 3 account to $GroupName. Aborting."
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
            try {
                Remove-ADGroupMember -Identity $GroupName -Members $UserName -Server $DCHostName -Confirm:$false
                Write-LogFile -LogFile $LogFile -LogString "Removed $UserName from $GroupName. Reason: $Reason"
                $Outcome = "Success"
            } catch {
                $ex = $_.Exception
                Write-LogFile -LogFile $LogFile -LogString "ERROR: $($ex.Message)" -ForegroundColor Red
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
    $AuditRow = [pscustomobject]@{
        Timestamp       = $ActionTime.ToString("yyyy-MM-dd HH:mm:ss")
        Action          = $AuditAction
        Operator        = "$Domain\$env:USERNAME"
        Account         = $UserName
        Group           = $GroupName
        DurationMinutes = $AuditDuration
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
#====================================================================
