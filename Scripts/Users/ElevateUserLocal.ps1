#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

# Mode: Standalone / No Shared CorpAdmin Module

# Deliberately standalone (no CorpAdmin import, no Write-LogFile, no audit CSV).
# This script runs during bootstrap before the domain is fully provisioned,
# so it cannot depend on the shared module. The negative contract for these
# absences is enforced in ScriptParameters.Tests.ps1.

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][ValidateSet("E","R")][string]$UserAction,
    [string]$GroupName,
    [string]$Domain
)

Set-StrictMode -Version Latest

if (!$Domain) {
    $Domain = "$env:userdomain"
}
if (!$GroupName) {
    $GroupName = "Administrators"
}
if (-not (Get-LocalGroup -Name $GroupName -ErrorAction SilentlyContinue)) {
    throw "Local group '$GroupName' does not exist"
}
$UserAction=$UserAction.ToUpper()
switch ($UserAction) {
    "E" {
        if (-not (Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -Server $Domain)) {
            throw "User '$Domain\$UserName' does not exist"
        }
        if ((Get-LocalGroupMember -Group $GroupName | Where-Object {$_.Name -eq "$Domain\$UserName"})) {
            throw "User '$Domain\$UserName' is already a member of this group"
        }
        try {
            Add-LocalGroupMember -Group $GroupName -Member "$Domain\$UserName" -ErrorAction Stop
        } catch {
            $ex = $_.Exception
            Write-Host "ERROR: $($ex.Message)" -ForegroundColor Red
        }
    }
    "R" {
        if (-not (Get-LocalGroupMember -Group $GroupName | Where-Object {$_.Name -eq "$Domain\$UserName"})) {
            throw "User '$Domain\$UserName' is not a member of this group"
        }
        try {
            Remove-LocalGroupMember -Group $GroupName -Member "$Domain\$UserName" -ErrorAction Stop
        } catch {
            $ex = $_.Exception
            Write-Host "ERROR: $($ex.Message)" -ForegroundColor Red
        }
    }
}
