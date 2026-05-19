#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [string]$LogFile
)

Set-StrictMode -Version Latest

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

#====================================================================
# Environment
#====================================================================
$Domain = "$env:USERDOMAIN"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$StaffOU = "OU=$($Env.OUs.Staff),$EndPath"
$HiPrivOU = "OU=$($Env.OUs.HiPrivAccounts),OU=$($Env.OUs.Administration),$EndPath"

# ADConnect & Exchange settings
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
# Get containing folder for script to locate supporting files
$ScriptPath = $PSScriptRoot
# Set variables
$ScriptTitle = "$Domain User Mover Script"
# File locations
$LogPath = "$ScriptPath\LogFiles"
$UserInputFile = "$ScriptPath\movers.csv"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_mover_log-$(Get-Date -Format 'yyyyMMdd')"
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

$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Account_Admins", "$($Env.Groups.TaskPrefix)Standard_Group_Admins")
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
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
    Write-LogFile -LogFile $LogFile -LogString "Processing mover for $UserName"
    
    #================================================================
    # Load user
    #================================================================
    $User = Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -SearchBase $StaffOU -Properties Department, Enabled, DistinguishedName -Server $DCHostName
    if (-not $User) {
        Write-LogFile -LogFile $LogFile -LogString "User $UserName not found in Staff OU - skipping" -ForegroundColor Yellow
        continue
    }
    try {
        #============================================================
        # Safety: ensure Tier-R only
        #============================================================
        if ($User.DistinguishedName -like "*$HiPrivOU*") {
            Write-LogFile -LogFile $LogFile -LogString "User $UserName is Hi-Priv - mover not permitted" -ForegroundColor Yellow
            continue
        }
        if (-not $User.Enabled) {
            Write-LogFile -LogFile $LogFile -LogString "User $UserName is disabled - skipping" -ForegroundColor Yellow
            continue
        }
        
        #============================================================
        # New department group must exist BEFORE any mutation.
        # Convention: department group name = exact dept string.
        #
        # Verifying first means a typo'd or not-yet-created NEW_DEPT
        # aborts this row cleanly via continue, leaving the user
        # entirely in their original state (old Department attribute,
        # old group membership). A re-run after the group is created
        # then starts from a clean, consistent state.
        #
        # The previous ordering wrote the Department attribute first
        # and threw on a missing group - that left the attribute set
        # to a dept the user was never added to, and the idempotent
        # "already set" guard masked the half-completed row on re-run.
        #============================================================
        if (-not (Get-ADGroup -Filter "Name -eq '$NewDept'" -ErrorAction SilentlyContinue -Server $DCHostName)) {
            Write-LogFile -LogFile $LogFile -LogString "Department group $NewDept does not exist - skipping $UserName, no changes made" -ForegroundColor Yellow
            continue
        }
        
        #============================================================
        # Add to new department group FIRST (add-before-remove).
        # A failure anywhere after this point leaves the user
        # over-entitled (in both old and new dept groups) rather than
        # under-entitled (in neither) - the safe failure direction.
        # Add-GroupMember is idempotent (swallows "already a member"),
        # so this is safe on re-run.
        #============================================================
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $NewDept -Member $UserName
        Write-LogFile -LogFile $LogFile -LogString "Added $UserName to $NewDept"
        
        #============================================================
        # Update Department attribute only after the group join is
        # secured. Skip the write when it already matches to avoid
        # no-op writes against the DC and keep the audit log clean.
        #============================================================
        if ($User.Department -ne $NewDept) {
            Set-ADUser -Identity $User -Department $NewDept -Server $DCHostName
            Write-LogFile -LogFile $LogFile -LogString "Department updated: $($User.Department) set to $NewDept"
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Department already set to $NewDept"
        }
        
        #============================================================
        # Old department group cleanup - remove LAST, only when
        # OldDept differs from NewDept.
        #============================================================
        if ($OldDept -and $OldDept -ne $NewDept) {
            if (Get-ADGroup -Filter "Name -eq '$OldDept'" -ErrorAction SilentlyContinue -Server $DCHostName) {
                try {
                    Remove-ADGroupMember -Identity $OldDept -Members $UserName -Confirm:$false -ErrorAction Stop -Server $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Removed $UserName from $OldDept"
                } catch [Microsoft.ActiveDirectory.Management.ADException] {
                    if ($_.Exception.Message -notmatch "not a member") {
                        Write-LogFile -LogFile $LogFile -LogString "ERROR: $($_.Exception.Message)" -ForegroundColor Red
                        throw
                    }
                    Write-LogFile -LogFile $LogFile -LogString "$UserName was not a member of $OldDept - skipping" -ForegroundColor Green
                }
            } else {
                Write-LogFile -LogFile $LogFile -LogString "Old department group '$OldDept' does not exist - skipping removal" -ForegroundColor Yellow
            }
        }
        
        #============================================================
        # Update manager
        #============================================================
        if ($NewMgrSam) {
            $Mgr = Get-ADUser -Filter "sAMAccountName -eq '$NewMgrSam'" -ErrorAction SilentlyContinue -Server $DCHostName
            if ($Mgr) {
                Set-ADUser -Identity $User -Manager $Mgr.DistinguishedName -Server $DCHostName
                Write-LogFile -LogFile $LogFile -LogString "Manager set to $NewMgrSam"
            } else {
                Write-LogFile -LogFile $LogFile -LogString "Manager $NewMgrSam not found"
            }
        }
        Write-LogFile -LogFile $LogFile -LogString "Mover completed for $UserName"
        Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
        Write-LogFile -LogFile $LogFile -LogString " "
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR processing $UserName : $_" -ForegroundColor Red
        continue
    }
}
Write-LogFile -LogFile $LogFile -LogString "Departmental mover completed"
