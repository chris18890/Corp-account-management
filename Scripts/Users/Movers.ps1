#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [string]$LogFile
)

Set-StrictMode -Version Latest

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

#====================================================================
# Environment
#====================================================================
$Domain = "$env:USERDOMAIN"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$AdministrationOU = "Administration"
$UsersOU = "Staff"
$StaffOU = "OU=$UsersOU,$EndPath"
$HiPrivOU = "OU=Hi_Priv_Accounts,OU=$AdministrationOU,$EndPath"

# ADConnect & Exchange settings
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain User Mover Script"
# File locations
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\movers.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_mover_log-$(Get-Date -Format 'yyyyMMdd')"
    $LogIndex = 0
    while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
        $LogIndex ++
    }
    $LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
    Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
    Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
    Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
    Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
    Write-Log -LogFile $LogFile -LogString ("=" * 80)
    Write-Log -LogFile $LogFile -LogString " "
}

$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Import CSV
#====================================================================
$Movers = @(Import-Csv $UserInputFile)
if ($Movers -isnot [Array]) {$Movers = @($Movers)}
$RequiredHeaders = @(
    "USERNAME","OLD_DEPT","NEW_DEPT","NEW_MANAGER"
)
$Headers = ($Movers | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -inotcontains $h) {
        throw "movers.csv missing required column '$h'"
    }
}
foreach ($Mover in $Movers) {
    $UserName  = $Mover.USERNAME.Trim().ToLower()
    $OldDept   = $Mover.OLD_DEPT.Trim()
    $NewDept   = $Mover.NEW_DEPT.Trim()
    $NewMgrSam = $Mover.NEW_MANAGER.Trim().ToLower()
    Write-Log -LogFile $LogFile -LogString "Processing mover for $UserName"
    
    #================================================================
    # Load user
    #================================================================
    $User = Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -SearchBase $StaffOU -Properties Department, Enabled, DistinguishedName -Server $DCHostName
    if (-not $User) {
        Write-Log -LogFile $LogFile -LogString "User $UserName not found in Staff OU - skipping" -ForegroundColor Yellow
        continue
    }
    try {
        #============================================================
        # Safety: ensure Tier-R only
        #============================================================
        if ($User.DistinguishedName -like "*$HiPrivOU*") {
            Write-Log -LogFile $LogFile -LogString "User $UserName is Hi-Priv - mover not permitted" -ForegroundColor Yellow
            continue
        }
        if (-not $User.Enabled) {
            Write-Log -LogFile $LogFile -LogString "User $UserName is disabled - skipping" -ForegroundColor Yellow
            continue
        }
        
        #============================================================
        # Update Department attribute
        # New Departmental group handling
        # Convention: department group name = exact dept string
        #============================================================
        if ($User.Department -ne $NewDept) {
            Set-ADUser -Identity $User -Department $NewDept -Server $DCHostName
            Write-Log -LogFile $LogFile -LogString "Department updated: $($User.Department) → $NewDept"
        } else {
            Write-Log -LogFile $LogFile -LogString "Department already set to $NewDept"
        }
        if (Get-ADGroup -Filter "Name -eq '$NewDept'" -ErrorAction SilentlyContinue -Server $DCHostName) {
            Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $NewDept -Member $UserName
            Write-Log -LogFile $LogFile -LogString "Added $UserName to $NewDept"
        } else {
            Write-Log -LogFile $LogFile -LogString "Department group $NewDept does not exist"
            throw
        }
        
        #============================================================
        # Old Departmental group handling
        #============================================================
        if ($OldDept) {
            if ($User.Department -eq $OldDept) {
                if (Get-ADGroup -Filter "Name -eq '$OldDept'" -ErrorAction SilentlyContinue  -Server $DCHostName) {
                    try {
                        Remove-ADGroupMember -Identity $OldDept -Members $UserName -Confirm:$false -ErrorAction SilentlyContinue -Server $DCHostName
                        Write-Log -LogFile $LogFile -LogString "Removed $UserName from $OldDept"
                    } catch {
                        $ex = $_.Exception
                        Write-Log -LogFile $LogFile -LogString "ERROR: $($ex.Message)" -ForegroundColor Red
                    }
                }
            }
        }
        
        #============================================================
        # Update manager
        #============================================================
        if ($NewMgrSam) {
            $Mgr = Get-ADUser -Filter "sAMAccountName -eq '$NewMgrSam'" -ErrorAction SilentlyContinue -Server $DCHostName
            if ($Mgr) {
                Set-ADUser -Identity $User -Manager $Mgr.DistinguishedName -Server $DCHostName
                Write-Log -LogFile $LogFile -LogString "Manager set to $NewMgrSam"
            } else {
                Write-Log -LogFile $LogFile -LogString "Manager $NewMgrSam not found"
            }
        }
        Write-Log -LogFile $LogFile -LogString "Mover completed for $UserName"
        Write-Log -LogFile $LogFile -LogString ("=" * 80)
        Write-Log -LogFile $LogFile -LogString " "
    } catch {
        Write-Log -LogFile $LogFile -LogString "ERROR processing $UserName : $_" -ForegroundColor Red
        continue
    }
}
Write-Log -LogFile $LogFile -LogString "Departmental mover completed"
