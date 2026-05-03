#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName,
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][ValidateSet("E","R")][string]$UserAction,
    [ValidateSet("P","T")][string]$TempOrPerm,
    [int]$TimeSpan
)

Set-StrictMode -Version Latest

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

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

Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "
$requiredGroups = @('ADM_Task_Standard_Group_Admins', 'ADM_Task_HiPriv_Group_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
switch ($UserAction.ToUpper()) {
    "E" {
        if (!$TempOrPerm) {
            $TempOrPerm = READ-HOST 'Enter a duration;  P for permanent, T for temporary - '
        }
        $TempOrPerm = $TempOrPerm.ToUpper()
        if (!(Get-ADOptionalFeature -Filter "Name -eq 'Privileged Access Management Feature'" -Server $DCHostName)) {
            $TempOrPerm = "P"
            Write-Log -LogFile $LogFile -LogString "AD PAM feature not enabled, defaulting to permanent" -ForegroundColor Red
        }
        switch ($TempOrPerm.ToUpper()) {
            "P" {
                Write-Log -LogFile $LogFile -LogString "Adding $UserName to $GroupName with no time limit, manual removal required" -ForegroundColor Yellow
                Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $UserName
            }
            "T" {
                if (!$TimeSpan) {
                    Write-Log -LogFile $LogFile -LogString "No timespan specified, using default of 60 minutes"
                    $TimeSpan = 60
                }
                if ($TimeSpan -gt 480) {
                    Write-Log -LogFile $LogFile -LogString "Requested timespan is longer than 480 minutes / 8 hours, please rerun with a lower value"
                    return
                }
                Write-Log -LogFile $LogFile -LogString "Adding $UserName to $GroupName for $TimeSpan minutes" -ForegroundColor Yellow
                Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $GroupName -Member $UserName -TimeSpan $TimeSpan
            }
            default {
                Write-Log -LogFile $LogFile -LogString "No valid option specified, quitting"
                return
            }
        }
    }
    "R" {
        try {
            Remove-ADGroupMember -Identity $GroupName -Members $UserName -Server $DCHostName -Confirm:$false
            Write-Log -LogFile $LogFile -LogString "Removed $UserName from $GroupName"
        } catch {
            $ex = $_.Exception
            Write-Log -LogFile $LogFile -LogString "ERROR: $($ex.Message)" -ForegroundColor Red
        }
    }
}
#====================================================================
