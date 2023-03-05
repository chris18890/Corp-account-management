[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$O365
)
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$O365 = $O365.ToUpper()
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
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
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
    $form.FormBorderStyle = 'FixedSingle'
    $form.FormBorderStyle = 'Sizable'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
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
    Displays a MessageBox asking for a Yes/No repsonse, returning $True for Yes and $False for No
     
    .DESCRIPTION
    Displays a MessageBox asking for a Yes/No repsonse
     
    .PARAMETER WinTitle
    Displays a MessageBox asking for a Yes/No repsonse, returning $True for Yes and $False for No
     
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
        [parameter(
        Mandatory=$False)]
        [String]$WinTitle = 'PowerShell Script',
        [parameter(
        Mandatory=$False)]
        $MsgText = 'Do you really want to continue ?'
    )
    $result = [Windows.Forms.MessageBox]::Show($MsgText, $WinTitle, [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question)
    if ($result -eq [Windows.Forms.DialogResult]::Yes) {
        Return $true
    } else {
        Return $false
    }
}
$time = (Get-Date).toString("HH:mm")
Write-Host "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-"
Write-Host "User Account Deletion"
Write-Host "Script started at $time"
# Ask for some work
$UserNames = Read-MultiLineInputBoxDialog "Type or paste in a list of IDs to delete" $progTitle ""
if ($UserNames -eq $null) {
    Write-Host "You pressed the Cancel button, exiting."
    Break
}
# Strip out whitespace and punctuation to leave a space separated list
$UserNames = $UserNames -replace '\s+', ' '
#$UserNames = $UserNames -replace '\W+', ' '
# Spin through the list dealing with them 
foreach ($UserName in $UserNames -split '\s+') {
    # Set up the log entry
    $t = Get-Date -Format "HH:mm:ss"
    $d = Get-Date -Format "dd/MM/yyyy"
    $u = [Environment]::UserName.ToUpper()
    Write-Host ""
    Write-Host "-----------------------"
    Write-Host "Account $UserName archived at $t on $d by $u"
    Write-Host "-----------------------"
    #################
    Write-Host "Processing $UserName"# load the account from the AD
    Write-Host "-----------------------"
    $user = Get-ADUser -Filter {cn -eq $UserName} -Properties *
    #$user
    # check it was found
    if ($user.sAMAccountName -ne $UserName) {
        Write-Host " > ID $UserName not found"
        Write-Host "-----------------------"
        continue # skip to the next ID
    }
    Write-Host (" > This account was for " + $user.displayname)
    #################
    Write-Host " - Ensuring $UserName disabled"
    if ($user.enabled -eq $True) {
        Write-Host "   - Account $UserName is not disabled!"
        if (Show-MessageBox $progTitle "Account $UserName is not disabled. Are you sure you want to continue?") {
            Write-Host "   > Proceeding at user request, disabling the account $UserName"
            Set-ADUser -Identity $user -Enabled $false
        } else {
            Write-Host "Skipping $UserName at user request"
            continue # skip to the next ID
        }
    }
    #################
    Write-Host " - Doing other things"
    #################
    Write-Host " - Removing Home Drive"
    $HomeDrive = $user.HomeDirectory
    Write-Host "   > $HomeDrive"
    if (Test-Path $HomeDrive) {
        Remove-Item -Recurse -Force $HomeDrive
    }
    Write-Host " - Home Drive was at $HomeDrive"
    #check to see if there is a sh_ group for the account and remove it if so
    $SharedGroupName = "sh_$UserName"
    try {
        Remove-ADGroup $SharedGroupName -Confirm:$False
        if ($?) {
            Write-Host " - Removing group $SharedGroupName for shared account $UserName"
        }
    } catch {
        #nothing to do
    }
    $EquipmentGroupName = "eq_$UserName"
    try {
        Remove-ADGroup $EquipmentGroupName -Confirm:$False
        if ($?) {
            Write-Host " - Removing group $EquipmentGroupName for shared account $UserName"
        }
    } catch {
        #nothing to do
    }
    $RoomGroupName = "ro_$UserName"
    try {
        Remove-ADGroup $RoomGroupName -Confirm:$False
        if ($?) {
            Write-Host " - Removing group $RoomGroupName for shared account $UserName"
        }
    } catch {
        #nothing to do
    }
    try {
        Remove-ADUser $UserName -Confirm:$False
        if ($?) {
            Write-Host " - Deleting account $UserName"
        }
    } catch {
        #nothing to do
    }
    Write-Host "-----------------------"
}   # foreach id
if ($O365 -eq "Y") {
    Write-Host ""
    Write-Host "Starting AzureAD Sync"
    Import-Module ADSync
    Start-ADSyncSyncCycle -PolicyType Delta
}
Write-Host "Finished."
Write-Host "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-"
