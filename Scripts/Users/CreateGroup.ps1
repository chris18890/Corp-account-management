#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName
    , [Parameter(Mandatory)][ValidateSet("S","H")][string]$GroupType
    , [Parameter(Mandatory)][ValidateSet("E","H","N")][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$Description
)

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

#====================================================================
# Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
#====================================================================

#====================================================================
# Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
#====================================================================

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain Group Creation Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_group_creation_log-$(Get-Date -Format 'yyyyMMdd')"
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

$requiredGroups = @('ADM_Task_Standard_Group_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
if ($O365 -eq "E" -or $O365 -eq "H") {
    # Get user credentials for server connectivity (Non-MFA)
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log "Failed to validate credentials: $ErrorMsg "
        Read-Host -Prompt "Press Enter to exit"
        Break
    }
    #Connect to remote Exchange PowerShell
    Write-Log "Connecting to remote Exchange PowerShell session... "
    try {
        $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
        Write-Log "connected."
        Write-Log "Importing Exchange session... "
        Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
        Write-Log "done."
    } catch {
        $e = $_.Exception
        Write-Log $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log $line
        $msg = $e.Message
        Write-Log $msg
        $Action = "Error Importing Exchange Session"
        Write-Log $Action
        Write-Log "failed."
        Write-Log "ERROR: $_" -ForegroundColor Red
    }
    if (!$ExSession) {
        Write-Log "Exchange session not connected Stopping Script"
        Exit
    }
}
#====================================================================

#====================================================================
# Group creation
#====================================================================
if (!$Description) {
    $Description = READ-HOST 'Enter group description - '
}
switch ($GroupType.ToUpperInvariant()) {
    "S" {
        $OU = $GroupsOU
        $OUPath = "OU=$OU,$EndPath"
        Write-Log "Creating group $GroupName and adding to $StaffGroup"
        New-DomainGroup -GroupName $GroupName -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $false -Path $OUPath -GroupDescription $Description
        Add-GroupMember -Group $StaffGroup -Member $GroupName
        Write-Log "Group $GroupName created in location $OUPath and added to $StaffGroup"
    }
    "H" {
        $requiredGroups = @('ADM_Task_HiPriv_Group_Admins')
        $groups = $requiredGroups | ForEach-Object {
            Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
        }
        if (-not $groups) {
            Write-Log "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        $OU = "Hi_Priv_Groups"
        $OUPath = "OU=$OU,OU=Administration,$EndPath"
        Write-Log "Creating Hi-Priv group $GroupName" -ForegroundColor Yellow
        New-DomainGroup -GroupName $GroupName -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path $OUPath -GroupDescription $Description
        Write-Log "Hi-Priv group $GroupName created in location $OUPath and protectedFromAccidentalDeletion"
    }
    default {
        throw "Invalid UserType. Use 'S' (Staff) or 'H' (High Privilege)."
    }
}
#====================================================================

#====================================================================
if ($O365 -eq "E" -or $O365 -eq "H") {
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PsSession $ExSession
        Write-Log "Closed Exchange session."
    }
}
#====================================================================
