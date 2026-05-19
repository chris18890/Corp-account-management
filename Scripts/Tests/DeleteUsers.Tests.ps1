#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Coverage for the family's most destructive script. The script is GUI-driven
# (Read-MultiLineInputBoxDialog) and confirmation is via Show-MessageBox rather
# than the standard -Confirm pipeline, so behavioural tests would require
# running the GUI. Coverage is therefore parameter-contract via Get-Command
# and source-pattern via the AST + raw text.

# Convention: every regex pattern is single-quoted. PowerShell's double-quoted
# strings interpolate $variable refs and don't escape backslashes (PowerShell
# uses backtick `, not \), which makes embedding regex patterns containing
# either $ or \ a magnet for subtle bugs. Single-quoted strings sidestep both
# - what you write is what the regex engine sees. Embedded ' becomes ''.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $script:scriptPath  = Join-Path $script:scriptsRoot 'Users\DeleteUsers.ps1'
    $script:source      = Get-Content $script:scriptPath -Raw
    $tokens = $null
    $errors = $null
    $script:ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:scriptPath, [ref]$tokens, [ref]$errors
    )
    function Script:Get-ParamAttr {
        param($Cmd, [string]$Name, [Type]$Type)
        $Cmd.Parameters[$Name].Attributes | Where-Object { $_ -is $Type }
    }
    function Script:Test-IsMandatory {
        param($Cmd, [string]$Name)
        $attrs = Get-ParamAttr $Cmd $Name ([System.Management.Automation.ParameterAttribute])
        $null -ne ($attrs | Where-Object Mandatory)
    }
}

Describe 'DeleteUsers.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command $script:scriptPath
    }
    It 'exposes -O365 as a switch' {
        $script:cmd.Parameters['O365'].SwitchParameter | Should -BeTrue
    }
    It 'exposes -LogFile, -EmailSuffix, -O365EmailSuffix' {
        $script:cmd.Parameters.Keys | Should -Contain 'LogFile'
        $script:cmd.Parameters.Keys | Should -Contain 'EmailSuffix'
        $script:cmd.Parameters.Keys | Should -Contain 'O365EmailSuffix'
    }
    It 'does not expose -Username (input is via Read-MultiLineInputBoxDialog)' {
        $script:cmd.Parameters.Keys | Should -Not -Contain 'Username'
    }
    It 'declares CmdletBinding' {
        # AST check because Get-Command doesn't surface the CmdletBinding attribute presence
        $cb = $script:ast.ParamBlock.Attributes |
            Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }
        $cb | Should -Not -BeNullOrEmpty
    }
}

Describe 'DeleteUsers.ps1 required modules' {
    It 'requires ActiveDirectory'    { $script:source | Should -Match '#Requires\s+-Modules\s+ActiveDirectory' }
    It 'requires RunAsAdministrator' { $script:source | Should -Match '#Requires\s+-RunAsAdministrator' }
    It 'imports CorpAdmin from Modules\CorpAdmin\CorpAdmin.psd1' {
        $script:source | Should -Match 'Modules.CorpAdmin.CorpAdmin\.psd1'
    }
    It 'declares Set-StrictMode -Version Latest' {
        $script:source | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
}

Describe 'DeleteUsers.ps1 authorization gate' {
    It 'checks invoker membership in the account-admin task groups' {
        $script:source | Should -Match 'Standard_Account_Admins'
        $script:source | Should -Match 'SER_Account_Admins'
        $script:source | Should -Match 'HiPriv_Account_Admins'
        $script:source | Should -Match '''Domain Admins'''
    }
    It 'throws when the invoker is not in an authorised group' {
        $script:source | Should -Match '(?ms)not\s+authorised.*?throw'
    }
}

Describe 'DeleteUsers.ps1 input dialog and sanitisation' {
    It 'reads the username list interactively' {
        $script:source | Should -Match 'Read-MultiLineInputBoxDialog'
    }
    It 'collapses whitespace to single spaces' {
        # matches:  -replace '\s+', ' '
        $script:source | Should -Match '-replace\s+''\\s\+'',\s*'' '''
    }
    It 'strips punctuation other than dot and hyphen' {
        # matches:  [^A-Za-z0-9.-]
        $script:source | Should -Match '\[\^A-Za-z0-9\.\-\]'
    }
    It 'skips empty entries after splitting' {
        # matches:  [string]::IsNullOrWhiteSpace($entry)
        $script:source | Should -Match '\[string\]::IsNullOrWhiteSpace\(\$entry\)'
    }
}

Describe 'DeleteUsers.ps1 on-prem prefix routing' {
    # The classifier handles four cases:
    #   admin.X -> preserve, da.X -> preserve, ca./ga.X -> cloud branch,
    #   bare X  -> expand to { X, admin.X, da.X }
    It "preserves an 'admin.' prefix without re-expansion" {
        # matches:  '^admin\.'   (10 chars including outer quotes and the backslash)
        $script:source | Should -Match '''\^admin\\\.'''
    }
    It "preserves a 'da.' prefix without re-expansion" {
        # matches:  '^da\.'
        $script:source | Should -Match '''\^da\\\.'''
    }
    It 'routes ca./ga. accounts to the cloud branch (continues past on-prem)' {
        # matches:  '^(ca|ga)\.'
        $script:source | Should -Match '''\^\(ca\|ga\)\\\.'''
    }
    It 'expands bare names into the three on-prem tiers' {
        # matches:  admin.$entry   and   da.$entry
        $script:source | Should -Match 'admin\.\$entry'
        $script:source | Should -Match 'da\.\$entry'
    }
}

Describe 'DeleteUsers.ps1 destructive-action confirmation' {
    It 'declares SupportsShouldProcess (so -WhatIf / -Confirm work)' {
        $cb = $script:ast.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }
        $cb.NamedArguments.ArgumentName | Should -Contain 'SupportsShouldProcess'
    }
    It 'uses Show-MessageBox prompts for per-account confirmation' {
        ([regex]::Matches($script:source, 'Show-MessageBox')).Count | Should -BeGreaterOrEqual 2
    }
    It 'inspects the Enabled property before deletion' {
        $script:source | Should -Match '\$user\.[Ee]nabled\s*-eq\s*\$true'
    }
    It 'gates EVERY deletion unconditionally, before (and independent of) the enabled check' {
        # Regression guard for the already-disabled bypass: the ShouldProcess +
        # Show-MessageBox gate must precede the enabled handling, so an account
        # that is already disabled still requires confirmation.
        $script:source | Should -Match '(?s)\$PSCmdlet\.ShouldProcess[\s\S]*?Show-MessageBox[\s\S]*?\$user\.[Ee]nabled\s*-eq\s*\$true'
    }
    It 'skips the account (continue) when ShouldProcess declines or -WhatIf is set' {
        $script:source | Should -Match '(?s)if\s*\(-not\s+\$PSCmdlet\.ShouldProcess[\s\S]*?continue'
    }
    It 'still disables a left-enabled account before deleting it' {
        $script:source | Should -Match '(?s)\$user\.[Ee]nabled\s*-eq\s*\$true[\s\S]*?Set-ADUser\s+-Identity\s+\$user\s+-Enabled\s+\$false'
    }
    It 'gates both on-prem and cloud deletions behind ShouldProcess' {
        ([regex]::Matches($script:source, '\$PSCmdlet\.ShouldProcess')).Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'DeleteUsers.ps1 home-directory cleanup' {
    It 'checks Test-Path before Remove-Item' {
        $script:source | Should -Match '(?s)Test-Path\s+\$HomeDrive.*?Remove-Item'
    }
    It 'gates Remove-Item behind a MessageBox prompt' {
        $script:source | Should -Match '(?s)HomeDrive.*?Show-MessageBox.*?Remove-Item'
    }
}

Describe 'DeleteUsers.ps1 related delegation-group cleanup' {
    It 'unprotects the shared-access delegation group before deletion' {
        $script:source | Should -Match 'SharedAccessPrefix'
        $script:source | Should -Match '(?s)protectedFromAccidentalDeletion\s+\$false.*?Remove-ADGroup'
    }
    It 'unprotects the equipment-access delegation group' {
        $script:source | Should -Match 'EquipmentAccessPrefix'
    }
    It 'unprotects the room-access delegation group' {
        $script:source | Should -Match 'RoomAccessPrefix'
    }
    It 'unprotects the user object before Remove-ADUser' {
        $script:source | Should -Match '(?s)Set-ADObject.*?protectedFromAccidentalDeletion.*?Remove-ADUser'
    }
}

Describe 'DeleteUsers.ps1 logging' {
    It 'logs via Write-LogFile throughout' {
        ([regex]::Matches($script:source, 'Write-LogFile')).Count | Should -BeGreaterThan 20
    }
    It 'constructs a date-stamped log path under \LogFiles' {
        $script:source | Should -Match 'LogFiles'
        $script:source | Should -Match 'Get-Date\s+-Format\s+''yyyyMMdd'''
    }
    It 'avoids overwriting a same-day log file by indexing' {
        $script:source | Should -Match '\$LogIndex'
        $script:source | Should -Match 'Test-Path\s+"\$LogPath'
    }
}

Describe 'DeleteUsers.ps1 cloud (Graph) branch' {
    It 'checks for the Microsoft.Graph module before connecting' {
        $script:source | Should -Match 'Get-Module\s+-ListAvailable\s+-Name\s+Microsoft\.Graph'
    }
    It 'connects with RoleManagement.ReadWrite.Directory and User.ReadWrite.All scopes' {
        $script:source | Should -Match 'Connect-MgGraph'
        $script:source | Should -Match 'RoleManagement\.ReadWrite\.Directory'
        $script:source | Should -Match 'User\.ReadWrite\.All'
    }
    It 'removes directory-role assignments before removing the user' {
        $assignment = ($script:source -split "`n" |
            Select-String 'Remove-MgRoleManagementDirectoryRoleAssignment' |
            Select-Object -First 1).LineNumber
        $userRemove = ($script:source -split "`n" |
            Select-String 'Remove-MgUserByUserPrincipalName' |
            Select-Object -First 1).LineNumber
        $assignment | Should -Not -BeNullOrEmpty
        $userRemove | Should -Not -BeNullOrEmpty
        $assignment | Should -BeLessThan $userRemove
    }
    It 'removes the cloud user via Remove-MgUserByUserPrincipalName -UserPrincipalName' {
        $script:source | Should -Match 'Remove-MgUserByUserPrincipalName\s+-UserPrincipalName'
    }
    It 'disconnects Microsoft Graph when finished' {
        $script:source | Should -Match 'Disconnect-MgGraph'
    }
    It 'cleans up PSSessions and disposes the credential password' {
        $script:source | Should -Match 'Get-PSSession\s*\|\s*Remove-PSSession'
        $script:source | Should -Match '\$Cred\.Password\.Dispose'
    }
}

Describe 'DeleteUsers.ps1 cloud-branch on-prem prefix skip' {
    # ------------------------------------------------------------------------
    # The cloud branch must detect admin./da. inputs and skip them, because
    # those accounts belong to the on-prem deletion branch above. The fix
    # mirrors the on-prem branch's $isCloud = -match '^(ca|ga)\.' check.
    # ------------------------------------------------------------------------
    It "cloud branch detects on-prem prefixes via -match '^(admin|da)\\.' " {
        # Matches the literal:  $isOnPrem = $entry -match '^(admin|da)\.'
        $script:source | Should -Match '\$isOnPrem\s*=\s*\$entry\s+-match\s+''\^\(admin\|da\)\\\.'''
    }
    It 'cloud branch continues on the on-prem-prefix skip path' {
        # The $isOnPrem branch must `continue`, not throw or fall through.
        $script:source | Should -Match '(?ms)\$isOnPrem[\s\S]*?continue'
    }
    It 'on-prem prefix check sits inside the cloud-branch foreach over $UserNames' {
        # Defence against someone hoisting the isOnPrem check out of the
        # cloud branch (which would also affect the on-prem branch and
        # break its three-tier expansion).
        $script:source | Should -Match '(?ms)if\s*\(\$O365\)[\s\S]*?\$isOnPrem'
    }
}

Describe 'DeleteUsers.ps1 cloud-branch preserves ca./ga. prefixes (regression guard)' {
    # ------------------------------------------------------------------------
    # The two same-tier cases. These guard against a refactor that
    # accidentally removes the preserve-this-prefix arm and re-expands
    # an already-prefixed cloud account into 'ca.ca.foo' / 'ga.ga.foo'.
    # ------------------------------------------------------------------------
    It "cloud branch detects 'ca.' prefix via -match '^ca\\.' " {
        $script:source | Should -Match '\$isCa\s*=\s*\$entry\s+-match\s+''\^ca\\\.'''
    }
    It "cloud branch detects 'ga.' prefix via -match '^ga\\.' " {
        $script:source | Should -Match '\$isGa\s*=\s*\$entry\s+-match\s+''\^ga\\\.'''
    }
    It 'preserves already-prefixed cloud accounts (no re-expansion)' {
        # The $isCa -or $isGa branch should append a single UPN, not two.
        $script:source | Should -Match '(?ms)\$isCa\s+-or\s+\$isGa[\s\S]*?\$CloudAccounts\.Add\("\$entry@\$EmailSuffix"\)'
    }
}

Describe 'DeleteUsers.ps1 cloud-branch bare-name expansion (regression guard)' {
    # ------------------------------------------------------------------------
    # Bare names expand to BOTH ca. and ga. tiers - so deleting 'potter'
    # removes both ca.potter and ga.potter from Entra. Guards against a
    # refactor that drops one of the two Adds.
    # ------------------------------------------------------------------------
    It 'expands bare names into ca.<entry>@<EmailSuffix>' {
        $script:source | Should -Match '\$CloudAccounts\.Add\("ca\.\$entry@\$EmailSuffix"\)'
    }
    It 'expands bare names into ga.<entry>@<EmailSuffix>' {
        $script:source | Should -Match '\$CloudAccounts\.Add\("ga\.\$entry@\$EmailSuffix"\)'
    }
}

Describe 'DeleteUsers.ps1 on-prem branch cloud-prefix skip - regression guard' {
    # ------------------------------------------------------------------------
    # The mirror property: the on-prem branch already detects ca./ga. and
    # skips because those belong to the cloud branch. Pinning this here
    # protects against someone "simplifying" the on-prem branch by
    # removing its skip and reintroducing the inverse bug.
    # ------------------------------------------------------------------------
    It 'on-prem branch detects ca./ga. cloud prefixes via -match' {
        $script:source | Should -Match '\$isCloud\s*=\s*\$entry\s+-match\s+''\^\(ca\|ga\)\\\.'''
    }
    It 'on-prem branch continues on the cloud-prefix skip path' {
        $script:source | Should -Match '(?ms)\$isCloud[\s\S]*?continue'
    }
}

Describe 'DeleteUsers.ps1 branch symmetry' {
    # ------------------------------------------------------------------------
    # Both branches must use the same pattern: combined other-tier flag
    # for skip, separate same-tier flags for preserve, bare-name expansion
    # in else. If the two branches diverge structurally, the next
    # maintainer is going to have a bad time.
    # ------------------------------------------------------------------------
    It 'on-prem branch uses $isAdmin + $isDa for same-tier preservation' {
        $script:source | Should -Match '\$isAdmin\s*=\s*\$entry\s+-match\s+''\^admin\\\.'''
        $script:source | Should -Match '\$isDa\s*=\s*\$entry\s+-match\s+''\^da\\\.'''
    }
    It 'cloud branch uses $isCa + $isGa for same-tier preservation' {
        $script:source | Should -Match '\$isCa\s*=\s*\$entry'
        $script:source | Should -Match '\$isGa\s*=\s*\$entry'
    }
    It 'both branches use a List[string] for the accumulator' {
        # On-prem uses $onPremAccounts, cloud uses $CloudAccounts;
        # both should be generic List[string] for symmetric semantics.
        $script:source | Should -Match '\$onPremAccounts\s*=\s*\[System\.Collections\.Generic\.List\[string\]\]::new\(\)'
        $script:source | Should -Match '\$CloudAccounts\s*=\s*\[System\.Collections\.Generic\.List\[string\]\]::new\(\)'
    }
}
