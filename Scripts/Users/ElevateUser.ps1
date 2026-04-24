#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName,
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][ValidateSet("E","R")][string]$UserAction,
    [ValidateSet("P","T")][string]$TempOrPerm,
    [int]$TimeSpan
)

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
# Group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Group
        , [Parameter(Mandatory)][string]$Member
        , [int]$TimeSpan
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          Member will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    #                   $TimeSpan - number of minutes to add temporal memebership for
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)" -Server $DCHostName
    if ($null -ne $checkGroup) {
        $checkMember = Get-ADObject -LDAPFilter "(SAMAccountName=$Member)" -Server $DCHostName
        if (-not $checkMember) {
            Write-Log "'$Member' does not exist" -ForegroundColor Red
            return
        }
        Write-Log "Adding $Member to $Group"
        try {
            if ($TimeSpan) {
                Write-Log "Adding $Member to $Group for $TimeSpan minutes" -ForegroundColor Yellow
                Add-ADGroupMember -Identity $Group -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes $TimeSpan) -Server $DCHostName
            } else {
                Write-Log "Adding $Member to $Group with no time limit, manual removal required" -ForegroundColor Yellow
                Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            }
            Write-Log "Added $Member to $Group"
        } catch {
            $ex = $_.Exception
            if ($ex.Message -match "already a member") {
                Write-Log "'$Member' is already a member of group '$Group'" -ForegroundColor Green
            } else {
                throw
            }
        }
    } else {
        Write-Log "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

$Domain = "$env:userdomain"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain User Elevation Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_user_elevation_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""

switch ($UserAction.ToUpper()) {
    "E" {
        if (!$TempOrPerm) {
            $TempOrPerm = READ-HOST 'Enter a duration;  P for permanent, T for temporary - '
        }
        $TempOrPerm = $TempOrPerm.ToUpper()
        if (!(Get-ADOptionalFeature -Filter "Name -eq 'Privileged Access Management Feature'" -Server $DCHostName)) {
            $TempOrPerm = "P"
            Write-Log "AD PAM feature not enabled, defaulting to permanent" -ForegroundColor Red
        }
        switch ($TempOrPerm.ToUpper()) {
            "P" {
                Add-GroupMember -Group $GroupName -Member $UserName
            }
            "T" {
                if (!$TimeSpan) {
                    Write-Log "No timespan specified, using default of 60 minutes"
                    $TimeSpan = 60
                }
                if ($TimeSpan -gt 480) {
                    Write-Log "Requested timespan is longer than 480 minutes / 8 hours, please rerun with a lower value"
                    return
                }
                Add-GroupMember -Group $GroupName -Member $UserName -TimeSpan $TimeSpan
            }
            default {
                Write-Log "No valid option specified, quitting"
                return
            }
        }
    }
    "R" {
        try {
            Remove-ADGroupMember -Identity $GroupName -Members $UserName -Server $DCHostName -Confirm:$false
            Write-Log "Removed $UserName from $GroupName"
        } catch {
            $ex = $_.Exception
            Write-Log "ERROR: $($ex.Message)" -ForegroundColor Red
        }
    }
}
#====================================================================
