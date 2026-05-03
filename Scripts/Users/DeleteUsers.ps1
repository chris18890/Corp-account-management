#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$O365
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$ModulePath = (Split-Path $PSScriptRoot -Parent)
. $ModulePath\helpers.ps1

# windows GUI components
function Read-MultiLineInputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText) {
<#
    .SYNOPSIS
    Prompts the user with a multi-line input box and returns the text they enter, or null if they cancelled the prompt.
     
    .DESCRIPTION
    Prompts the user with a multi-line input box and returns the text they enter, or null if they cancelled the prompt.
     
    .PARAMETER Message
    The message to display to the user explaining what text we are asking them to enter.
     
    .PARAMETER WindowTitle
    The text to display on the prompt window's title.
     
    .PARAMETER DefaultText
    The default text to show in the input box.
     
    .EXAMPLE
    $userText = Read-MultiLineInputDialog "Input some text please:" "Get User's Input"
     
    Shows how to create a simple prompt to get mutli-line input from a user.
     
    .EXAMPLE
    # Setup the default multi-line address to fill the input box with.
    $defaultAddress = @'
    John Doe
    123 St.
    Some Town, SK, Canada
    A1B 2C3
    '@
     
    $address = Read-MultiLineInputDialog "Please enter your full address, including name, street, city, and postal code:" "Get User's Address" $defaultAddress
    if ($address -eq $null)
    {
        Write-Error "You pressed the Cancel button on the multi-line input box."
    }
     
    Prompts the user for their address and stores it in a variable, pre-filling the input box with a default multi-line address.
    If the user pressed the Cancel button an error is written to the console.
     
    .EXAMPLE
    $inputText = Read-MultiLineInputDialog -Message "If you have a really long message you can break it apart`nover two lines with the powershell newline character:" -WindowTitle "Window Title" -DefaultText "Default text for the input box."
     
    Shows how to break the second parameter (Message) up onto two lines using the powershell newline character (`n).
    If you break the message up into more than two lines the extra lines will be hidden behind or show ontop of the TextBox.
     
    .NOTES
    Name: Show-MultiLineInputDialog
    Author: Daniel Schroeder (originally based on the code shown at http://technet.microsoft.com/en-us/library/ff730941.aspx)
    Version: 1.0
#>
    # Create the Label.
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message
    # Create the TextBox used to capture the user's text.
    $textBox = New-Object System.Windows.Forms.TextBox 
    $textBox.Location = New-Object System.Drawing.Size(10,40) 
    $textBox.Size = New-Object System.Drawing.Size(575,200)
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $false
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $DefaultText
    $textBox.anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    # Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(415,250)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okbutton.anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(510,250)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(610,320)
    $form.FormBorderStyle = 'Sizable'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $true
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
    # Return the text that the user entered.
    return $form.Tag
}

function Show-MessageBox {
<#
    .SYNOPSIS
    Displays a MessageBox asking for a Yes/No repsonse, returning $true for Yes and $false for No
     
    .DESCRIPTION
    Displays a MessageBox asking for a Yes/No repsonse
     
    .PARAMETER WinTitle
    Displays a MessageBox asking for a Yes/No repsonse, returning $true for Yes and $false for No
     
    .PARAMETER MsgText
    The text in the MessageBox
     
    .EXAMPLE
    if (Show-MessageBox "MessageBox Title" "Press [Yes] or [No].") {
        "yes pressed"
     } else {
        "no pressed"
    }
    
    .NOTES
    Name: Show-MessageBox
    Author: Chris Murray
    Version: 1.0
#>
    param(
        [parameter(Mandatory = $false)][String]$WinTitle = 'PowerShell Script',
        [parameter(Mandatory = $false)]$MsgText = 'Do you really want to continue ?'
    )
    $result = [Windows.Forms.MessageBox]::Show($MsgText, $WinTitle, [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question)
    if ($result -eq [Windows.Forms.DialogResult]::Yes) {
        Return $true
    } else {
        Return $false
    }
}

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$GroupsOU = "Groups"
$EquipmentGroupsOU = "OU=Equipment_Mailbox_Access,OU=$GroupsOU"
$RoomGroupsOU = "OU=Room_Mailbox_Access,OU=$GroupsOU"
$SharedGroupsOU = "OU=Shared_Mailbox_Access,OU=$GroupsOU"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptTitle = "$Domain User Account Deletion Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item "$LogPath" -type directory -force
}
$LogFileName = $Domain + "_user_deletion_log-$(Get-Date -Format 'yyyyMMdd')"
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
$requiredGroups = @('ADM_Task_Standard_Account_Admins', 'ADM_Task_SER_Account_Admins', 'ADM_Task_HiPriv_Account_Admins', 'Domain Admins')
$groups = $requiredGroups | ForEach-Object {
    Get-ADGroup -Filter "Name -eq '$_'" -Server $DCHostName | Get-ADGroupMember -Server $DCHostName -Recursive | Where-Object SamAccountName -eq $env:USERNAME
}
if (-not $groups) {
    Write-Log -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}

# Ask for some work
$UserNames = Read-MultiLineInputBoxDialog "Type or paste in a list of IDs to delete" $ScriptTitle ""
if ([string]::IsNullOrWhiteSpace($UserNames)) {
    Write-Log -LogFile $LogFile -LogString "You pressed the Cancel button, exiting."
    Break
}
# Strip out whitespace and punctuation to leave a space separated list
$UserNames = $UserNames -replace '\s+', ' '
$UserNames = $UserNames -replace '[^A-Za-z0-9.-]', ' '
# Spin through the list dealing with them 
foreach ($UserName in $UserNames -split '\s+') {
    # Set up the log entry
    $t = Get-Date -Format "HH:mm:ss"
    $d = Get-Date -Format "dd/MM/yyyy"
    $u = [Environment]::UserName.ToUpper()
    Write-Log -LogFile $LogFile -LogString " "
    Write-Log -LogFile $LogFile -LogString "-----------------------"
    Write-Log -LogFile $LogFile -LogString "Account $UserName deleted at $t on $d by $u"
    Write-Log -LogFile $LogFile -LogString "-----------------------"
    #################
    Write-Log -LogFile $LogFile -LogString "Processing $UserName"# load the account from the AD
    Write-Log -LogFile $LogFile -LogString "-----------------------"
    $user = Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -Properties * -Server $DCHostName
    #$user
    # check it was found
    if (!$user) {
        Write-Log -LogFile $LogFile -LogString " > ID $UserName not found"
        Write-Log -LogFile $LogFile -LogString "-----------------------"
        continue # skip to the next ID
    }
    Write-Log -LogFile $LogFile -LogString (" > This account was for " + $user.displayname)
    #################
    Write-Log -LogFile $LogFile -LogString " - Ensuring $UserName disabled"
    if ($user.enabled -eq $true) {
        Write-Log -LogFile $LogFile -LogString "   - Account $UserName is not disabled!"
        if (Show-MessageBox $ScriptTitle "Account $UserName is not disabled. Are you sure you want to continue?") {
            Write-Log -LogFile $LogFile -LogString "   > Proceeding at user request, disabling the account $UserName"
            Set-ADUser -Identity $user -Enabled $false -Server $DCHostName
        } else {
            Write-Log -LogFile $LogFile -LogString "Skipping $UserName at user request"
            continue # skip to the next ID
        }
    }
    #################
    Write-Log -LogFile $LogFile -LogString " - Doing other things"
    #################
    Write-Log -LogFile $LogFile -LogString " - Removing Home Drive"
    $HomeDrive = $user.HomeDirectory
    if ($HomeDrive) {
        Write-Log -LogFile $LogFile -LogString "   > $HomeDrive"
        if (Test-Path $HomeDrive) {
            Write-Log -LogFile $LogFile -LogString "   - HomeDrive $HomeDrive for $UserName exists"
            if (Show-MessageBox $ScriptTitle "HomeDrive $HomeDrive for $UserName exists. Are you sure you want to continue?") {
                Write-Log -LogFile $LogFile -LogString "   > Proceeding at user request, deleting home directory $HomeDrive for $UserName"
                Remove-Item -Recurse -Force $HomeDrive
            } else {
                Write-Log -LogFile $LogFile -LogString "Skipping deletion of $HomeDrive for $UserName at user request"
            }
        }
        Write-Log -LogFile $LogFile -LogString " - Home Drive was at $HomeDrive"
    } else {
        Write-Log -LogFile $LogFile -LogString " - No Home Drive configured"
    }
    #check to see if there is a sh_ group for the account and remove it if so
    $SharedGroupName = "sh_$UserName"
    try {
        Set-ADObject -Identity "CN=$SharedGroupName,$SharedGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $SharedGroupName -Confirm:$false -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString " - Removing group $SharedGroupName for shared account $UserName"
    } catch {
        Write-Log -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $EquipmentGroupName = "eq_$UserName"
    try {
        Set-ADObject -Identity "CN=$EquipmentGroupName,$EquipmentGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $EquipmentGroupName -Confirm:$false -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString " - Removing group $EquipmentGroupName for equipment account $UserName"
    } catch {
        Write-Log -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $RoomGroupName = "ro_$UserName"
    try {
        Set-ADObject -Identity "CN=$RoomGroupName,$RoomGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $RoomGroupName -Confirm:$false -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString " - Removing group $RoomGroupName for room account $UserName"
    } catch {
        Write-Log -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    try {
        Set-ADObject -Identity $User -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADUser $UserName -Confirm:$false -Server $DCHostName
        Write-Log -LogFile $LogFile -LogString " - Deleting account $UserName"
    } catch {
        Write-Log -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    Write-Log -LogFile $LogFile -LogString "-----------------------"
}   # foreach id
if ($O365) {
    Write-Log -LogFile $LogFile -LogString " "
    Write-Log -LogFile $LogFile -LogString "Starting AzureAD Sync"
    Import-Module ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
}
Write-Log -LogFile $LogFile -LogString "Finished."
Write-Log -LogFile $LogFile -LogString ("=" * 80)
