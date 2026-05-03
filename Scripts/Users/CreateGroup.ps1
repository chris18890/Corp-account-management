#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName
    , [Parameter(Mandatory)][ValidateSet("S","H")][string]$GroupType
    , [Parameter(Mandatory)][ValidateSet("E","H","N")][string]$O365
    , [Parameter(Mandatory)][string]$EmailSuffix
    , [string]$Description
)

Set-StrictMode -Version Latest

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

Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString "Log file is '$LogFile'"
Write-Log -LogFile $LogFile -LogString "Processing commenced, running as user '$Domain\$env:USERNAME'"
Write-Log -LogFile $LogFile -LogString "DC being used is '$DCHostName'"
Write-Log -LogFile $LogFile -LogString "Script path is '$ScriptPath'"
Write-Log -LogFile $LogFile -LogString "$ScriptTitle"
Write-Log -LogFile $LogFile -LogString "Errors and warnings will be displayed below. See the log file '$LogFile' for further details of these"
Write-Log -LogFile $LogFile -LogString ("=" * 80)
Write-Log -LogFile $LogFile -LogString " "

$requiredGroups = @('ADM_Task_Standard_Group_Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

#====================================================================
if ($O365 -eq "E" -or $O365 -eq "H") {
    # Get user credentials for server connectivity (Non-MFA)
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
        Read-Host -Prompt "Press Enter to exit"
        Exit
    }
    #Connect to remote Exchange PowerShell
    Write-Log -LogFile $LogFile -LogString "Connecting to remote Exchange PowerShell session... "
    try {
        $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
        Write-Log -LogFile $LogFile -LogString "connected."
        Write-Log -LogFile $LogFile -LogString "Importing Exchange session... "
        Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
        Write-Log -LogFile $LogFile -LogString "done."
    } catch {
        $e = $_.Exception
        Write-Log -LogFile $LogFile -LogString $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Log -LogFile $LogFile -LogString $line
        $msg = $e.Message
        Write-Log -LogFile $LogFile -LogString $msg
        $Action = "Error Importing Exchange Session"
        Write-Log -LogFile $LogFile -LogString $Action
        Write-Log -LogFile $LogFile -LogString "failed."
        Write-Log -LogFile $LogFile -LogString "ERROR: $_" -ForegroundColor Red
    }
    if (!$ExSession) {
        Write-Log -LogFile $LogFile -LogString "Exchange session not connected Stopping Script"
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
        Write-Log -LogFile $LogFile -LogString "Creating group $GroupName and adding to $StaffGroup"
        New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $GroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 $O365 -HiddenFromAddressListsEnabled $false -Path $OUPath -GroupDescription $Description
        Add-GroupMember -LogFile $LogFile -DCHostName $DCHostName -Group $StaffGroup -Member $GroupName
        Write-Log -LogFile $LogFile -LogString "Group $GroupName created in location $OUPath and added to $StaffGroup"
    }
    "H" {
        $requiredGroups = @('ADM_Task_HiPriv_Group_Admins')
        $groups = $requiredGroups | ForEach-Object {
            Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
        }
        if (-not $groups) {
            Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
            throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
        }
        $OU = "Hi_Priv_Groups"
        $OUPath = "OU=$OU,OU=Administration,$EndPath"
        Write-Log -LogFile $LogFile -LogString "Creating Hi-Priv group $GroupName" -ForegroundColor Yellow
        New-DomainGroup -LogFile $LogFile -DCHostName $DCHostName -GroupName $GroupName -GroupCategory $GroupCategory -GroupScope $GroupScope -O365 "N" -HiddenFromAddressListsEnabled $true -Path $OUPath -GroupDescription $Description
        Write-Log -LogFile $LogFile -LogString "Hi-Priv group $GroupName created in location $OUPath and protectedFromAccidentalDeletion"
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
        $Cred.Password.Dispose()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Log -LogFile $LogFile -LogString "Closed Exchange session."
    }
}
#====================================================================
