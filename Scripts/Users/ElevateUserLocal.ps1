#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

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
$UserAction=$UserAction.ToUpper()
switch ($UserAction) {
    "E" {
        if (-not (Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -Server $Domain)) {
            throw "User '$Domain\$UserName' does not exist"
        }
        try {
            Add-LocalGroupMember -Group $GroupName -Member "$Domain\$UserName" -ErrorAction Stop
        } catch {
            $ex = $_.Exception
            if ($ex.Message -match "already a member") {
                Write-Host "'$UserName' is already a member of '$GroupName'" -ForegroundColor Yellow
            } else {
                throw
            }
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
