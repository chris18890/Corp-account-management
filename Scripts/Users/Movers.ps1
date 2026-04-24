#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

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
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          User will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
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
            Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
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
$LogFileName = $Domain + "_mover_log-$(Get-Date -Format 'yyyyMMdd')"
$LogIndex = 0
while (Test-Path "$LogPath\$($LogFileName)_$LogIndex.log") {
    $LogIndex ++
}
$LogFile = "$LogPath\$($LogFileName)_$LogIndex.log"
#====================================================================

Write-Log ("=" * 80)
Write-Log "Log file is '$LogFile'"
Write-Log ("=" * 80)
Write-Log "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log "DC being used is '$DCHostName'"
Write-Log "Script path is '$ScriptPath'"
Write-Log "$ScriptTitle"
Write-Log "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log ("=" * 80)
Write-Log ""

$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_Standard_Group_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
# Import CSV
#====================================================================
$Rows = @(Import-Csv $UserInputFile)
$RequiredHeaders = @(
    "USERNAME","OLD_DEPT","NEW_DEPT","NEW_MANAGER"
)
$Headers = ($Rows | Select-Object -First 1).PSObject.Properties.Name
foreach ($h in $RequiredHeaders) {
    if ($Headers -inotcontains $h) {
        throw "movers.csv missing required column '$h'"
    }
}
foreach ($Row in $Rows) {
    $UserName  = $Row.USERNAME.Trim().ToLower()
    $OldDept   = $Row.OLD_DEPT.Trim()
    $NewDept   = $Row.NEW_DEPT.Trim()
    $NewMgrSam = $Row.NEW_MANAGER.Trim().ToLower()
    Write-Log "Processing mover for $UserName"
    
    #================================================================
    # Load user
    #================================================================
    $User = Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -SearchBase $StaffOU -Properties Department, Enabled, DistinguishedName -Server $DCHostName
    if (-not $User) {
        Write-Log "User $UserName not found in Staff OU - skipping" -ForegroundColor Yellow
        continue
    }
    
    #================================================================
    # Safety: ensure Tier-R only
    #================================================================
    if ($User.DistinguishedName -like "*$HiPrivOU*") {
        Write-Log "User $UserName is Hi-Priv - mover not permitted" -ForegroundColor Yellow
        continue
    }
    if (-not $User.Enabled) {
        Write-Log "User $UserName is disabled - skipping" -ForegroundColor Yellow
        continue
    }
    
    #================================================================
    # Update Department attribute
    #================================================================
    if ($User.Department -ne $NewDept) {
        Set-ADUser -Identity $User -Department $NewDept -Server $DCHostName
        Write-Log "Department updated: $($User.Department) → $NewDept"
    } else {
        Write-Log "Department already set to $NewDept"
    }
    
    #================================================================
    # Departmental group handling
    # Convention: department group name = exact dept string
    #================================================================
    if ($OldDept) {
        if (Get-ADGroup -Filter "Name -eq '$OldDept'" -ErrorAction SilentlyContinue  -Server $DCHostName) {
            Remove-ADGroupMember -Identity $OldDept -Members $UserName -Confirm:$false -ErrorAction SilentlyContinue -Server $DCHostName
            Write-Log "Removed $UserName from $OldDept"
        }
    }
    if (Get-ADGroup -Filter "Name -eq '$NewDept'" -ErrorAction SilentlyContinue -Server $DCHostName) {
        Add-GroupMember -Group $NewDept -Member $UserName
        Write-Log "Added $UserName to $NewDept"
    } else {
        Write-Log "Department group $NewDept does not exist"
    }
    
    #================================================================
    # Update manager
    #================================================================
    if ($NewMgrSam) {
        $Mgr = Get-ADUser -Filter "sAMAccountName -eq '$NewMgrSam'" -ErrorAction SilentlyContinue -Server $DCHostName
        if ($Mgr) {
            Set-ADUser -Identity $User -Manager $Mgr.DistinguishedName -Server $DCHostName
            Write-Log "Manager set to $NewMgrSam"
        } else {
            Write-Log "Manager $NewMgrSam not found"
        }
    }
    Write-Log "Mover completed for $UserName"
    Write-Log ("=" * 80)
    Write-Log ""
}
Write-Log "Departmental mover completed"
