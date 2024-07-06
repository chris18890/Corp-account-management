[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserType
)
switch ($UserType) {
    "S" {
        $OU = "Staff"
    }
    "H" {
        $OU = "Hi_Priv_Accounts,OU=IT"
    }
}
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$OUPath = "OU=$OU,$EndPath"
#get all enabled user accounts in the OU
$user = Get-ADUser -filter "enabled -eq 'true'" -SearchBase $OUPath -Properties * |
Select Name,SamAccountname,GivenName,Surname,DistinguishedName,Department |
Out-GridView -title "Select a user account or cancel" -OutputMode Single
if ($user) {
    #prompt for the new password
    $prompt = "Enter the user's SAMAccountname"
    $Title = "Reset Password"
    $Default = $null
    Add-Type -AssemblyName "microsoft.visualbasic" -ErrorAction Stop
    $prompt = "Enter the user's new password"
    $Plaintext =[microsoft.visualbasic.interaction]::InputBox($Prompt,$Title,$Default)
    #only continue is there is text for the password
    if ($plaintext -match "^\w") {
        #Validate Password against Password Policy
        #There are 4 requirements in current policy - this could change in future
        $TestsPassed = 0        #Counter for number of tests passed by Password
        #TEST 1: Password length must be at least 7 chars
        if ($plaintext.length -ge 7) {
            $TestsPassed ++
        }
        #TEST 2: Password must contain at least one lowercase letter (a-z)
        if ($plaintext -cmatch "[a-z]") {
            $TestsPassed ++
        }
        #TEST 3: Password must contain at least one uppercase letter (A-Z)
        if ($plaintext -cmatch "[A-Z]") {
            $TestsPassed ++
        }
        #TEST 4: Password must contain at least one number (0-9)
        if ($plaintext -match "[0-9]") {
            $TestsPassed ++
        }
        # Must contain a special character (not currently required)
        #if (-Not($Password -notmatch "[a-zA-Z0-9]")) {
        #   $TestsPassed ++
        #}
        if ($TestsPassed -ge 4) {
            Write-Verbose "Password validated"
        } else {
            Write-Host "ERROR: Password '$plaintext' does not comply with the password policy, script`nterminating" -ForegroundColor Red
            Write-Host ("-" * 80) -ForegroundColor Red
            exit
        }
        #convert to secure string
        $NewPassword = ConvertTo-SecureString -String $Plaintext -AsPlainText -Force
        #define a hash table of parameter values to splat to
        #Set-ADAccountPassword
        $paramHash = @{
            Identity = $User.SamAccountname
            NewPassword = $NewPassword
            Reset = $True
            Passthru = $True
            ErrorAction = "Stop"
        }
        try {
            $output = Set-ADAccountPassword @paramHash |
            Set-ADUser -ChangePasswordAtLogon $True -PassThru |
            Get-ADuser -Properties PasswordLastSet,PasswordExpired,WhenChanged |
            Out-String
            #display user in a message box
            $message = $output
            $button = "OKOnly"
            $icon = "Information"
            [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$title) | Out-Null
        } catch {
            #display error in a message box
            $message =  "Failed to reset password for $Username. $($_.Exception.Message)"
            $button = "OKOnly"
            $icon = "Exclamation"
            [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$title) | Out-Null
        }
    } #if plain text password
}
