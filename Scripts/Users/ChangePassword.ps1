#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest

[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("S","H")][string]$UserType
)

Add-Type -AssemblyName "microsoft.visualbasic" -ErrorAction Stop

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

#get all enabled user accounts in the OU
$User = Get-ADUser -filter "enabled -eq 'true'" -SearchBase $OUPath -Properties Name,SamAccountName,GivenName,Surname,DistinguishedName,Department -Server $DCHostName |
Out-GridView -title "Select a user account or cancel" -OutputMode Single
if ($User) {
    #prompt for the new password
    $Password = (Get-Credential -Message "Please enter new password for $($User.SamAccountName)" -User "$Domain\$($User.SamAccountName)").getNetworkCredential().password 
    #only continue is there is text for the password
    if ($Password -match "^\w") {
        #Validate Password against Password Policy
        #There are 4 requirements in current policy - this could change in future
        $TestsPassed = 0        #Counter for number of tests passed by Password
        if ($Password.length -ge $PasswordLength) {$TestsPassed ++} # Must be >= 20 chars
        if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain at least one lowercase letter (a-z)
        if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain at least one uppercase letter (A-Z)
        if ($Password -cmatch "[0-9]") {$TestsPassed ++} # Must contain at least one number (0-9)
        if ($Password -match "[^a-zA-Z0-9]") {$TestsPassed ++} # Must contain a special character
        if ($TestsPassed -ge 5) {
            Write-Verbose "Password validated"
        } else {
            Write-Host "ERROR: Password does not comply with the password policy, script`nterminating" -ForegroundColor Red
            Write-Host ("-" * 80) -ForegroundColor Red
            exit
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
