#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("S","H")][string]$UserType
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName "microsoft.visualbasic" -ErrorAction Stop

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$AdministrationOU = "Administration"
$UsersOU = "Staff"
switch ($UserType.ToUpperInvariant()) {
    "S" {
        $OU = $UsersOU
    }
    "H" {
        $OU = "Hi_Priv_Accounts,OU=$AdministrationOU"
    }
    default {
        throw "Invalid UserType. Use 'S' (Staff) or 'H' (High Privilege)."
    }
}
$OUPath = "OU=$OU,$EndPath"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set variables
$ScriptTitle = "$Domain Password Change Script"
$PasswordLength = 20 # Number of characters per password
# File locations
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_password_change_log-$(Get-Date -Format 'yyyyMMdd')"
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

#get all enabled user accounts in the OU
$User = Get-ADUser -filter "enabled -eq 'true'" -SearchBase $OUPath -Properties Name,SamAccountName,GivenName,Surname,DistinguishedName,Department -Server $DCHostName |
Out-GridView -title "Select a user account or cancel" -OutputMode Single
if ($User) {
    #prompt for the new password
    $Password = (Get-Credential -Message "Please enter new password for $($User.SamAccountName)" -User "$Domain\$($User.SamAccountName)").getNetworkCredential().password 
    #only continue is there is text for the password
    if ($Password -match "^\w") {
        #Validate Password against Password Policy
        try {
            Test-Password -LogFile $LogFile -Password $Password -PasswordLength $PasswordLength
        } catch {
            [microsoft.visualbasic.interaction]::Msgbox(
                "Password does not meet policy requirements. Please try again.",
                "OKOnly,Exclamation",
                $ScriptTitle
            ) | Out-Null
            return
        }
        #convert to secure string
        $NewPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Password = $null
        #define a hash table of parameter values to splat to
        #Set-ADAccountPassword
        $paramHash = @{
            Identity = $User.SamAccountName
            NewPassword = $NewPassword
            Reset = $True
            Passthru = $True
            ErrorAction = "Stop"
        }
        try {
            $output = Set-ADAccountPassword @paramHash -Server $DCHostName |
            Set-ADUser -ChangePasswordAtLogon $True -PassThru -Server $DCHostName |
            Get-ADuser -Properties PasswordLastSet,PasswordExpired,WhenChanged -Server $DCHostName |
            Out-String
            #display user in a message box
            $message = $output
            $button = "OKOnly"
            $icon = "Information"
            [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$ScriptTitle) | Out-Null
        } catch {
            #display error in a message box
            $message =  "Failed to reset password for $($User.SamAccountName). $($_.Exception.Message)"
            $button = "OKOnly"
            $icon = "Exclamation"
            [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$ScriptTitle) | Out-Null
        }
        $NewPassword = $null
    } #if plain text password
}
