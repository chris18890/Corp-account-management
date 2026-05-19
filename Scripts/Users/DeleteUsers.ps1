#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [switch]$O365
    ,[string]$LogFile
    ,[string]$EmailSuffix
    ,[string[]]$UserName            # explicit account list (array; pass (Get-ADUser...).SamAccountName)
    ,[string]$InputCsv              # reviewed file with a SamAccountName column
    ,[switch]$RemoveHomeDirectory   # home-dir removal is now an explicit policy decision
    ,[switch]$NonInteractive
    ,[string]$O365EmailSuffix
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules\CorpAdmin\CorpAdmin.psd1') -Force -DisableNameChecking
$Env = Get-EnvironmentConfig

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

$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$GroupsOU = $Env.OUs.Groups
$EquipmentGroupsOU = "OU=$($Env.OUs.EquipmentMailboxAccess),OU=$GroupsOU"
$RoomGroupsOU = "OU=$($Env.OUs.RoomMailboxAccess),OU=$GroupsOU"
$SharedGroupsOU = "OU=$($Env.OUs.SharedMailboxAccess),OU=$GroupsOU"
$DCHostName = (Get-ADDomain).PDCEmulator # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
$ScriptPath = $PSScriptRoot
$ScriptTitle = "$Domain User Account Deletion Script"
$LogPath = "$ScriptPath\LogFiles"
if (!(TEST-PATH $LogPath)) {
    Write-Host "Creating log folder"
    New-Item $LogPath -type directory -force
}
if (!$LogFile) {
    $LogFileName = $Domain + "_user_deletion_log-$(Get-Date -Format 'yyyyMMdd')"
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

$requiredGroups = @("$($Env.Groups.TaskPrefix)Standard_Account_Admins", "$($Env.Groups.TaskPrefix)SER_Account_Admins", "$($Env.Groups.TaskPrefix)HiPriv_Account_Admins", 'Domain Admins')
if (-not (Test-IsMemberOf -Sam $env:USERNAME -GroupNames $requiredGroups -DCHostName $DCHostName)) {
    Write-LogFile -LogFile $LogFile -LogString "Invoker is not authorised to run this script. Required privileges not present. Aborting." -ForegroundColor Red
    throw "Invoker is not authorised to run this script. Required privileges not present. Aborting."
}
$Interactive = [Environment]::UserInteractive -and -not $NonInteractive

# Resolve the work list: explicit input wins; the paste dialog is an interactive-only
# fallback; a headless run with nothing supplied stops clean (no Server-Core GUI dependency).
if ($InputCsv) {
    $rawList = (Import-Csv $InputCsv).SamAccountName -join ' '
} elseif ($UserName) {
    $rawList = $UserName -join ' '
} elseif ($Interactive) {
    $rawList = Read-MultiLineInputBoxDialog -Message "Type or paste in a list of IDs to delete" -WindowTitle $ScriptTitle ""
    if ([string]::IsNullOrWhiteSpace($rawList)) {
        Write-LogFile -LogFile $LogFile -LogString "You pressed the Cancel button, exiting."
        Break
    }
} else {
    throw "No -UserName or -InputCsv supplied in a non-interactive session: the paste dialog is unavailable (and absent on Server Core). Re-run with an explicit account list."
}
# Strip out whitespace and punctuation to leave a space separated list
$UserNames = $rawList -replace '\s+', ' '
$UserNames = $UserNames -replace '[^A-Za-z0-9.-]', ' '
# Spin through the list dealing with them
$onPremAccounts = [System.Collections.Generic.List[string]]::new()
foreach ($entry in $UserNames -split '\s+') {
    if ([string]::IsNullOrWhiteSpace($entry)) { continue }
    $isAdmin  = $entry -match '^admin\.'
    $isDa     = $entry -match '^da\.'
    $isCloud  = $entry -match '^(ca|ga)\.'
    if ($isCloud) {
        # Cloud account - on-prem branch has nothing to do with it.
        continue
    } elseif ($isAdmin -or $isDa) {
        # Caller gave a specific prefixed account; add only that one.
        $onPremAccounts.Add($entry)
    } else {
        $onPremAccounts.Add($entry)
        $onPremAccounts.Add("admin.$entry")
        $onPremAccounts.Add("da.$entry")
    }
}
foreach ($UserName in $onPremAccounts) {
    # Set up the log entry
    $t = Get-Date -Format "HH:mm:ss"
    $d = Get-Date -Format "dd/MM/yyyy"
    $u = [Environment]::UserName.ToUpper()
    Write-LogFile -LogFile $LogFile -LogString " "
    Write-LogFile -LogFile $LogFile -LogString "-----------------------"
    Write-LogFile -LogFile $LogFile -LogString "Account $UserName deleted at $t on $d by $u"
    Write-LogFile -LogFile $LogFile -LogString "-----------------------"
    #################
    Write-LogFile -LogFile $LogFile -LogString "Processing $UserName"# load the account from the AD
    Write-LogFile -LogFile $LogFile -LogString "-----------------------"
    $user = Get-ADUser -Filter "sAMAccountName -eq '$UserName'" -Properties * -Server $DCHostName
    #$user
    # check it was found
    if (!$user) {
        Write-LogFile -LogFile $LogFile -LogString " > ID $UserName not found"
        Write-LogFile -LogFile $LogFile -LogString "-----------------------"
        continue # skip to the next ID
    }
    Write-LogFile -LogFile $LogFile -LogString (" > This account was for " + $user.displayname)
    #################
    # Unconditional deletion gate. Previously the only Yes/No that could skip an
    # account was the enabled-check below, so an account that was ALREADY disabled
    # flowed straight through to Remove-ADUser with no confirmation. This fires for
    # every account regardless of state, and honours -WhatIf / -Confirm.
    $state = if ($user.enabled) { 'STILL ENABLED' } else { 'already disabled' }
    if (-not $PSCmdlet.ShouldProcess("$UserName ($($user.displayname)) [$state]", 'Permanently delete account and associated resources')) {
        Write-LogFile -LogFile $LogFile -LogString " > $UserName not deleted (WhatIf or declined)"
        continue
    }
    #################
    Write-LogFile -LogFile $LogFile -LogString " - Ensuring $UserName disabled"
    if ($user.enabled -eq $true) {
        Write-LogFile -LogFile $LogFile -LogString "   - Account $UserName still enabled; disabling before deletion"
        Set-ADUser -Identity $user -Enabled $false -Server $DCHostName
    }
    Write-LogFile -LogFile $LogFile -LogString " - Doing other things"
    #################
    Write-LogFile -LogFile $LogFile -LogString " - Removing Home Drive"
    $HomeDrive = $user.HomeDirectory
    if ($HomeDrive) {
        Write-LogFile -LogFile $LogFile -LogString "   > $HomeDrive"
        if (Test-Path $HomeDrive) {
            Write-LogFile -LogFile $LogFile -LogString "   - HomeDrive $HomeDrive for $UserName exists"
            if (-not $RemoveHomeDirectory) {
                Write-LogFile -LogFile $LogFile -LogString "   - -RemoveHomeDirectory not set; leaving $HomeDrive in place"
            } elseif ($PSCmdlet.ShouldProcess($HomeDrive, "Remove home directory for $UserName")) {
                Write-LogFile -LogFile $LogFile -LogString "   > Deleting home directory $HomeDrive for $UserName"
                Remove-Item -Recurse -Force $HomeDrive
            }
        }
        Write-LogFile -LogFile $LogFile -LogString " - Home Drive was at $HomeDrive"
    } else {
        Write-LogFile -LogFile $LogFile -LogString " - No Home Drive configured"
    }
    #check to see if there is a delegation group for the account and remove it if so
    $SharedGroupName = "$($Env.Groups.SharedAccessPrefix)$UserName"
    try {
        Set-ADObject -Identity "CN=$SharedGroupName,$SharedGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $SharedGroupName -Confirm:$false -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString " - Removing group $SharedGroupName for shared account $UserName"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $EquipmentGroupName = "$($Env.Groups.EquipmentAccessPrefix)$UserName"
    try {
        Set-ADObject -Identity "CN=$EquipmentGroupName,$EquipmentGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $EquipmentGroupName -Confirm:$false -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString " - Removing group $EquipmentGroupName for equipment account $UserName"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $RoomGroupName = "$($Env.Groups.RoomAccessPrefix)$UserName"
    try {
        Set-ADObject -Identity "CN=$RoomGroupName,$RoomGroupsOU,$EndPath" -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADGroup $RoomGroupName -Confirm:$false -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString " - Removing group $RoomGroupName for room account $UserName"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    try {
        Set-ADObject -Identity $User -protectedFromAccidentalDeletion $false -Server $DCHostName
        Remove-ADUser $UserName -Confirm:$false -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString " - Deleting account $UserName"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString " - Note: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    Write-LogFile -LogFile $LogFile -LogString "-----------------------"
}   # foreach id
if ($O365) {
    $AzureADConnect = "$Domain-RTR.$DNSSuffix"
    if (!$EmailSuffix) {
        $EmailSuffix = READ-HOST 'Enter domain suffix - '
    }
    if (!$O365EmailSuffix) {
        $O365EmailSuffix = READ-HOST 'Enter "onmicrosoft.com" domain - '
    }
    if ($O365EmailSuffix -notmatch '\.onmicrosoft\.com$') {
        $O365EmailSuffix = "$O365EmailSuffix.onmicrosoft.com"
    }
    if ([string]::IsNullOrWhiteSpace($EmailSuffix)) {
        throw "EmailSuffix parameter is required (e.g. 'company.com')"
    }
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-LogFile -LogFile $LogFile -LogString "Failed to validate credentials: $ErrorMsg "
        Read-Host -Prompt "Press Enter to exit"
        Exit
    }
    Write-LogFile -LogFile $LogFile -LogString " "
    Write-LogFile -LogFile $LogFile -LogString "Starting AzureAD Sync"
    Invoke-ADSync -LogFile $LogFile -Cred $Cred -AzureADConnect $AzureADConnect -O365EmailSuffix $O365EmailSuffix
    if (Get-PSSession) {
        Write-LogFile -LogFile $LogFile -LogString "Cleaning up PSSessions"
        Get-PSSession | Remove-PSSession
        $Cred.Password.Dispose()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-LogFile -LogFile $LogFile -LogString "Microsoft.Graph module not installed"
        throw "Microsoft.Graph module not installed"
    }
    $Connected = $false
    try {
        Write-LogFile -LogFile $LogFile -LogString "Connecting to Microsoft Graph"
        Import-Module -Name Microsoft.Graph.Authentication
        Import-Module -Name Microsoft.Graph.Users
        Import-Module -Name Microsoft.Graph.Identity.Governance
        Connect-MgGraph -NoWelcome -Scopes "RoleManagement.ReadWrite.Directory", "User.ReadWrite.All"
        Write-LogFile -LogFile $LogFile -LogString "Connected to Microsoft Graph"
        $Connected = $true
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg
        $Action = "Failed to connect to MS Graph on import session"
        Write-LogFile -LogFile $LogFile -LogString $Action
    }
    if ($Connected -eq $true) {
        Write-LogFile -LogFile $LogFile -LogString "Removing cloud admin accounts"
        $CloudAccounts = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $UserNames -split '\s+') {
            if ([string]::IsNullOrWhiteSpace($entry)) { continue }
            $isCa     = $entry -match '^ca\.'
            $isGa     = $entry -match '^ga\.'
            $isOnPrem = $entry -match '^(admin|da)\.'
            if ($isOnPrem) {
                # On-prem account - cloud branch has nothing to do with it.
                continue
            } elseif ($isCa -or $isGa) {
                # Caller gave a specific prefixed cloud account; add only that one.
                $CloudAccounts.Add("$entry@$EmailSuffix")
            } else {
                $CloudAccounts.Add("ca.$entry@$EmailSuffix")
                $CloudAccounts.Add("ga.$entry@$EmailSuffix")
            }
        }
        foreach ($CloudUserPrincipalName in $CloudAccounts) {
            try {
                $MgUser = Get-MgUser -Filter "userPrincipalName eq '$CloudUserPrincipalName'" -ErrorAction Stop
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "Cloud admin account '$CloudUserPrincipalName' not found in Entra ID. Skipping." -ForegroundColor Yellow
                continue
            }
            if (-not $MgUser) {
                Write-LogFile -LogFile $LogFile -LogString "Cloud admin account '$CloudUserPrincipalName' not found in Entra ID. Skipping." -ForegroundColor Yellow
                continue
            }
            if (-not $PSCmdlet.ShouldProcess($CloudUserPrincipalName, 'Remove role assignments and delete cloud account')) {
                Write-LogFile -LogFile $LogFile -LogString " > $CloudUserPrincipalName not deleted (WhatIf or declined)"
                continue
            }
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($MgUser.Id)'" -All -ErrorAction Stop
            foreach ($assignment in $assignments) {
                Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $assignment.Id -ErrorAction Stop
                Write-LogFile -LogFile $LogFile -LogString "Removed role assignment $($assignment.RoleDefinitionId) from $CloudUserPrincipalName"
            }
            try {
                Write-LogFile -LogFile $LogFile -LogString "Removing $CloudUserPrincipalName"
                Remove-MgUserByUserPrincipalName -UserPrincipalName $CloudUserPrincipalName
            } catch {
                Write-LogFile -LogFile $LogFile -LogString "Account '$CloudUserPrincipalName' not removed"
                continue
            }
        }
        Write-LogFile -LogFile $LogFile -LogString " "
        Write-LogFile -LogFile $LogFile -LogString "Cloud admin account removal complete"
        Disconnect-MgGraph
    }
}
Write-LogFile -LogFile $LogFile -LogString "Finished."
Write-LogFile -LogFile $LogFile -LogString ("=" * 80)
