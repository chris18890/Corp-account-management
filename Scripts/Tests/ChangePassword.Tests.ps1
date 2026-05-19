#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Body-structure tests for ChangePassword.ps1.
#
# The parameter contract and -SearchScope Subtree check live in
# ScriptParameters.Tests.ps1. This file pins the rest of the load-bearing
# body logic:
#
#  - VisualBasic assembly load (for the MsgBox UI)
#  - UserType S -> Staff OU, UserType H -> HiPrivAccounts under Administration
#  - Default switch arm throws (defense in depth against ValidateSet bypass)
#  - Password policy gate: Test-Password called BEFORE Set-ADAccountPassword
#  - ChangePasswordAtLogon $True applied after the password reset
#  - PassThru piping chain through Set-ADUser -> Get-ADUser -> Out-String
#  - Out-String widened so the DistinguishedName isn't clipped in the dialog
#  - GUI feedback via VisualBasic MsgBox on both success and failure
#  - Success dialog SURFACES the generated temp password to the operator
#    (without it, ChangePasswordAtLogon leaves an unusable account)
#  - $Password is retained until the success dialog, then both it and the
#    SecureString copy are cleared in a finally block on every exit path
#  - Audit CSV write is guarded so a failed append falls back to the log
#    rather than masking the reset outcome
#  - Invoker authorisation gating: a top-level Test-IsMemberOf gate (Standard
#    tier or Domain Admins) BEFORE the switch, PLUS a stricter HiPriv-tier
#    precheck inside the H clause (privilege-escalation guard); the S clause
#    carries no extra gate, since the top-level gate is sufficient for Staff
#
# Convention: single-quoted regex patterns throughout.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $script:scriptPath  = Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'
    $script:source      = Get-Content $script:scriptPath -Raw
    $tokens = $null; $errors = $null
    $script:ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:scriptPath, [ref]$tokens, [ref]$errors
    )
}

Describe 'ChangePassword.ps1 required modules and hardening' {
    It 'requires ActiveDirectory'    { $script:source | Should -Match '#Requires\s+-Modules\s+ActiveDirectory' }
    It 'requires RunAsAdministrator' { $script:source | Should -Match '#Requires\s+-RunAsAdministrator' }
    It 'imports CorpAdmin from Modules\CorpAdmin\CorpAdmin.psd1' {
        $script:source | Should -Match 'Modules.CorpAdmin.CorpAdmin\.psd1'
    }
    It 'declares Set-StrictMode -Version Latest' {
        $script:source | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It 'loads the VisualBasic assembly for MsgBox UI' {
        $script:source | Should -Match 'Add-Type\s+-AssemblyName\s+"microsoft\.visualbasic"'
    }
}

Describe 'ChangePassword.ps1 UserType OU routing' {
    BeforeAll {
        # Find the UserType switch.
        $script:utSwitch = $script:ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst]
        }, $true) | Where-Object {
            $_.Condition.Extent.Text -match '\$UserType'
        } | Select-Object -First 1
        $script:utClauses = @{}
        if ($script:utSwitch) {
            foreach ($clause in $script:utSwitch.Clauses) {
                $script:utClauses[$clause.Item1.Value] = $clause.Item2
            }
        }
    }
    It 'UserType switch has clauses S and H (plus a default)' {
        ($script:utClauses.Keys | Sort-Object) -join ',' | Should -Be 'H,S'
        $script:utSwitch.Default | Should -Not -BeNullOrEmpty
    }
    It 'default clause throws (defense in depth even with ValidateSet on -UserType)' {
        $throws = $script:utSwitch.Default.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
        }, $true)
        $throws.Count | Should -BeGreaterOrEqual 1
    }
    It "S clause targets the Staff OU" {
        $script:utClauses['S'].Extent.Text | Should -Match '\$Env\.OUs\.Staff'
    }
    It 'H clause targets HiPrivAccounts under Administration' {
        $hBody = $script:utClauses['H'].Extent.Text
        $hBody | Should -Match '\$Env\.OUs\.HiPrivAccounts'
        $hBody | Should -Match '\$Env\.OUs\.Administration'
    }
}

# =============================================================================
# Invoker authorisation gating (parity with CreateGroup.ps1)
# =============================================================================
# ChangePassword.ps1 gates execution in two places:
#   1. A top-level gate (BEFORE the UserType switch) requires the invoker to be
#      a member of the Standard-tier admin group or Domain Admins. It applies to
#      EVERY run, regardless of UserType.
#   2. An additional, stricter precheck INSIDE the H switch clause requires the
#      HiPriv-tier admin group (or Domain Admins) before a high-privilege
#      account's password can be reset.
# The asymmetry is load-bearing: resetting a HiPriv account must demand a
# strictly higher privilege than resetting a Staff account. Both gates log the
# refusal via Write-LogFile and then throw. Losing the top-level gate lets any
# caller through; losing the H-arm precheck is a privilege-escalation surface (a
# Standard_*_Admins member could reset a HiPriv account's password). The S
# clause intentionally carries NO extra precheck - the top-level gate suffices
# for Staff resets, mirroring CreateGroup.ps1's S/H precheck asymmetry.
# =============================================================================
Describe 'ChangePassword.ps1 invoker authorisation gating' {
    BeforeAll {
        $cpPath = Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'
        $script:source = Get-Content $cpPath -Raw
        $script:ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $cpPath, [ref]$null, [ref]$null
        )
        # Every Test-IsMemberOf invocation in the script.
        $script:memberChecks = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Test-IsMemberOf'
        }, $true)
        # The UserType switch and its clauses (keyed by clause value).
        $script:utSwitch = $script:ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst]
        }, $true) | Where-Object {
            $_.Condition.Extent.Text -match '\$UserType'
        } | Select-Object -First 1
        $script:utClauses = @{}
        if ($script:utSwitch) {
            foreach ($clause in $script:utSwitch.Clauses) {
                $script:utClauses[$clause.Item1.Value] = $clause.Item2
            }
        }
        # The script text that precedes the switch - the top-level gate lives
        # here, so scoping to it disambiguates the top-level requiredGroups /
        # log+throw from the H-arm copy further down.
        $script:preSwitchText = if ($script:utSwitch) {
            $script:source.Substring(0, $script:utSwitch.Extent.StartOffset)
        } else { '' }
        # Classify the membership checks by position: top-level runs before the
        # switch; the H-arm precheck sits inside the H clause's extent.
        $script:topLevelCheck = $script:memberChecks | Where-Object {
            $script:utSwitch -and
            $_.Extent.StartLineNumber -lt $script:utSwitch.Extent.StartLineNumber
        } | Select-Object -First 1
        $script:hExtent = if ($script:utClauses.ContainsKey('H')) {
            $script:utClauses['H'].Extent
        } else { $null }
        $script:hArmCheck = $script:memberChecks | Where-Object {
            $script:hExtent -and
            $_.Extent.StartLineNumber -ge $script:hExtent.StartLineNumber -and
            $_.Extent.EndLineNumber   -le $script:hExtent.EndLineNumber
        } | Select-Object -First 1
    }
    It 'gates the invoker with Test-IsMemberOf at least twice (top-level + H-arm)' {
        $script:memberChecks.Count | Should -BeGreaterOrEqual 2
    }
    Context 'Top-level gate (applies to every UserType)' {
        It 'runs BEFORE the UserType switch' {
            $script:topLevelCheck | Should -Not -BeNullOrEmpty
        }
        It 'passes -Sam $env:USERNAME, -GroupNames $requiredGroups and -DCHostName $DCHostName' {
            $t = $script:topLevelCheck.Extent.Text
            $t | Should -Match '-Sam\s+\$env:USERNAME'
            $t | Should -Match '-GroupNames\s+\$requiredGroups'
            $t | Should -Match '-DCHostName\s+\$DCHostName'
        }
        It 'requires the Standard-tier account admins plus the additional Domain Admins' {
            # Pins the top-level requiredGroups assignment, scoped to the text
            # before the switch so it can't accidentally match the H-arm copy.
            # Uses *_Account_Admins, consistent with the sibling account-lifecycle
            # scripts (DeleteUsers, Enable/Disable-CloudAdmin).
            $script:preSwitchText | Should -Match '\$requiredGroups\s*=\s*@\(\s*"\$\(\$Env\.Groups\.TaskPrefix\)Standard_Account_Admins"\s*,\s*''Domain Admins''\s*\)'
        }
        It 'logs the refusal and throws when the invoker is not authorised' {
            $script:preSwitchText | Should -Match '(?s)if\s*\(\s*-not\s*\(Test-IsMemberOf.*?Write-LogFile.*?not authorised.*?throw'
        }
    }
    Context 'H-arm precheck (privilege-escalation guard)' {
        It 'the H switch clause carries its OWN Test-IsMemberOf precheck' {
            $script:hArmCheck | Should -Not -BeNullOrEmpty
        }
        It 'the H precheck requires the HiPriv-tier admin group or Domain Admins' {
            $hBody = $script:utClauses['H'].Extent.Text
            $hBody | Should -Match '\$requiredGroups\s*=\s*@\(\s*"\$\(\$Env\.Groups\.TaskPrefix\)HiPriv_Account_Admins"\s*,\s*''Domain Admins''\s*\)'
        }
        It 'the H precheck logs the refusal and throws' {
            $hBody = $script:utClauses['H'].Extent.Text
            $hBody | Should -Match '(?s)if\s*\(\s*-not\s*\(Test-IsMemberOf.*?Write-LogFile.*?not authorised.*?throw'
        }
        It 'demands a STRICTLY higher tier than the top-level gate (HiPriv, not Standard)' {
            # The whole point of the second gate: a Standard_*_Admins member must
            # not be able to reset a HiPriv account. The H-arm group token must
            # be the HiPriv tier and must NOT be the Standard tier.
            $hBody = $script:utClauses['H'].Extent.Text
            $hBody | Should -Match 'HiPriv_\w+_Admins'
            $hBody | Should -Not -Match 'Standard_\w+_Admins'
        }
        It 'runs the precheck BEFORE resolving the HiPriv OU (fail closed)' {
            # The gate must throw before $OU is set, so an unauthorised caller
            # never reaches the OU-routing / user-selection that follows.
            $hBody = $script:utClauses['H'].Extent.Text
            $gateIdx = $hBody.IndexOf('Test-IsMemberOf')
            $ouIdx   = $hBody.IndexOf('$OU')
            $gateIdx | Should -BeGreaterOrEqual 0
            $ouIdx   | Should -BeGreaterOrEqual 0
            $gateIdx | Should -BeLessThan $ouIdx
        }
    }
    Context 'S clause carries no extra gate (top-level gate suffices for Staff resets)' {
        It 'the S switch clause does NOT contain its own Test-IsMemberOf precheck' {
            $sBody = $script:utClauses['S'].Extent.Text
            $sBody | Should -Not -Match 'Test-IsMemberOf'
        }
    }
}

Describe 'ChangePassword.ps1 user selection' {
    It 'filters Get-ADUser by enabled=$true' {
        $script:source | Should -Match 'Get-ADUser\s+-filter\s+"enabled\s+-eq\s+''true''"'
    }
    It 'restricts the search to the resolved OUPath' {
        $script:source | Should -Match '-SearchBase\s+\$OUPath'
    }
    It 'uses Out-GridView with -OutputMode Single for interactive selection' {
        $script:source | Should -Match 'Out-GridView\s+-title.+-OutputMode\s+Single'
    }
}

Describe 'ChangePassword.ps1 password policy gate' {
    BeforeAll {
        # Find Test-Password and Set-ADAccountPassword calls; the policy
        # check MUST come before the reset.
        $script:testPwdCall = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Test-Password'
        }, $true) | Select-Object -First 1
        $script:setPwdCall  = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Set-ADAccountPassword'
        }, $true) | Select-Object -First 1
    }
    It 'calls Test-Password before Set-ADAccountPassword' {
        $script:testPwdCall | Should -Not -BeNullOrEmpty
        $script:setPwdCall  | Should -Not -BeNullOrEmpty
        $script:testPwdCall.Extent.StartLineNumber | Should -BeLessThan $script:setPwdCall.Extent.StartLineNumber
    }
    It 'passes -PasswordLength from environment.psd1' {
        $script:source | Should -Match '\$PasswordLength\s*=\s*\$Env\.Security\.PasswordLength'
    }
    It 'surfaces a policy-violation message and returns (no reset attempted)' {
        # On policy failure, the script should NOT call Set-ADAccountPassword.
        $script:source | Should -Match '(?s)try\s*\{\s*Test-Password.*?\}\s*catch\s*\{[\s\S]*?Show-OperatorMessage[\s\S]*?return'
    }
    It 'rejects empty / whitespace passwords up front' {
        $script:source | Should -Match 'IsNullOrWhiteSpace\(\$Password\)'
    }
}

Describe 'ChangePassword.ps1 reset pipeline' {
    BeforeAll {
        $script:source = Get-Content (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1') -Raw
    }
    It 'pipes Set-ADAccountPassword through Set-ADUser -ChangePasswordAtLogon $True' {
        # Forcing a change at next logon is load-bearing - admins setting
        # one-time passwords for users must not leave the value sticky.
        $script:source | Should -Match '(?s)Set-ADAccountPassword.*?\|\s*Set-ADUser\s+-ChangePasswordAtLogon\s+\$True'
    }
    It 'pipes through Get-ADUser to surface PasswordLastSet / PasswordExpired / WhenChanged' {
        $script:source | Should -Match '(?s)Set-ADUser.*?\|\s*Get-ADuser\s+-Properties\s+PasswordLastSet,PasswordExpired,WhenChanged'
    }
    It 'targets the PDC emulator (avoids replication races)' {
        $script:source | Should -Match '\$DCHostName\s*=\s*\(Get-ADDomain\)\.PDCEmulator'
        $script:source | Should -Match 'Set-ADAccountPassword[\s\S]*?-Server\s+\$DCHostName'
    }
    It 'surfaces the result on success via the interactivity-guarded helper' {
        $script:source | Should -Match '(?s)Set-ADAccountPassword[\s\S]*?Show-OperatorMessage'
    }
    It 'surfaces the generated temporary password to the operator on success' {
        # Load-bearing: the reset sets ChangePasswordAtLogon $True, so the user
        # must be told the temp value to complete first logon. The success
        # message must interpolate the plaintext $Password.
        $script:source | Should -Match 'New temporary password:\s*\$Password'
    }
    It 'widens Out-String so the DistinguishedName is not clipped in the dialog' {
        $script:source | Should -Match 'Out-String\s+-Width\s+\d+'
    }
    It 'shows the exception message in a MsgBox on failure' {
        $script:source | Should -Match 'Failed to reset password.*?\$\(\$_\.Exception\.Message\)'
    }
    It 'does NOT leak the password on the failure path' {
        # The failure MsgBox message must not interpolate $Password - a failed
        # reset has no temp value worth (or safe) surfacing.
        $failMatch = [regex]::Match(
            $script:source,
            '(?s)catch\s*\{[^}]*Failed to reset password[^}]*\}'
        )
        $failMatch.Success | Should -BeTrue
        $failMatch.Value | Should -Not -Match '\$Password\b'
    }
}

Describe 'ChangePassword.ps1 SecureString and plaintext hygiene' {
    BeforeAll {
        $cpPath = Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'
        $script:source = Get-Content $cpPath -Raw
        $script:ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $cpPath, [ref]$null, [ref]$null
        )
        $script:finallies = $script:ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Where-Object { $_.Finally }
    }
    It 'converts to SecureString with -AsPlainText -Force' {
        $script:source | Should -Match 'ConvertTo-SecureString\s+-String\s+\$Password\s+-AsPlainText\s+-Force'
    }
    It 'retains $Password until the success dialog (does not null it at conversion time)' {
        # Premature nulling is exactly the regression that hid the temp
        # password. There must be exactly one '$Password = $null' in the whole
        # script, and (next test) it must live in a finally block.
        ([regex]::Matches($script:source, '\$Password\s*=\s*\$null')).Count |
            Should -Be 1
    }
    It 'clears the plaintext $Password in a finally block (every exit path)' {
        $clearing = $script:finallies | Where-Object {
            $_.Finally.Extent.Text -match '\$Password\s*=\s*\$null'
        }
        $clearing | Should -Not -BeNullOrEmpty
    }
    It 'clears the SecureString copy $NewPassword in the same finally block' {
        $clearing = $script:finallies | Where-Object {
            $_.Finally.Extent.Text -match '\$NewPassword\s*=\s*\$null'
        }
        $clearing | Should -Not -BeNullOrEmpty
    }
}

Describe 'ChangePassword.ps1 logging' {
    BeforeAll {
        $script:source = Get-Content (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1') -Raw
    }
    It 'writes to a date-stamped log file under LogFiles' {
        $script:source | Should -Match 'password_change_log-\$\(Get-Date\s+-Format\s+''yyyyMMdd''\)'
    }
    It 'avoids overwriting same-day log files via $LogIndex' {
        $script:source | Should -Match '\$LogIndex\s*\+\+'
    }
}

Describe 'ChangePassword.ps1 non-interactive operability' {
    BeforeAll {
        $script:source = Get-Content (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1') -Raw
    }
    It 'exposes a -NonInteractive switch' {
        $script:source | Should -Match '\[switch\]\$NonInteractive'
    }
    It 'derives interactivity from the environment, overridable by the switch' {
        $script:source | Should -Match '\$Interactive\s*=\s*\[Environment\]::UserInteractive\s*-and\s*-not\s*\$NonInteractive'
    }
    It 'gates the modal dialog behind the interactivity flag, with a console fallback' {
        # Msgbox must only be reachable inside the $Interactive branch...
        $script:source | Should -Match '(?s)function Show-OperatorMessage[\s\S]*?if\s*\(\s*\$Interactive\s*\)[\s\S]*?Msgbox'
        # ...and the non-interactive path writes to the console instead.
        $script:source | Should -Match '(?s)function Show-OperatorMessage[\s\S]*?else[\s\S]*?Write-Host'
    }
    It 'requires -UserName when non-interactive (no reliance on Out-GridView)' {
        $script:source | Should -Match '(?s)if\s*\(!\$UserName\)[\s\S]*?if\s*\(\s*-not\s*\$Interactive\s*\)[\s\S]*?throw'
    }
    It 'never writes the temporary password to the log or audit' {
        # The temp value may be shown to the operator but must not persist:
        # no Write-LogFile of the success message, and no Password field in the audit row.
        $script:source | Should -Not -Match 'Write-LogFile[^\r\n]*New temporary password'
        $script:source | Should -Not -Match 'Password\s*=\s*\$Password[\s\S]*?Export-Csv'
    }
}

# =============================================================================
# Audit CSV write
# =============================================================================
# ChangePassword.ps1 writes a transient .log file under LogFiles\ for
# human inspection AND a durable, append-only audit trail at
# LogFiles\password-resets.csv, equivalent to ElevateUser.ps1's
# on-prem-elevations.csv. The audit row is written from a finally
# block so it lands even on failure. Columns:
#   Timestamp, Operator, Account, UserType, Outcome
# =============================================================================
Describe 'ChangePassword.ps1 audit CSV write' {
    BeforeAll {
        $script:source = Get-Content (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1') -Raw
        $script:ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'),
            [ref]$null, [ref]$null
        )
    }
    It 'writes to password-resets.csv' {
        $script:source | Should -Match 'password-resets\.csv'
    }
    It 'audit row write lives in a finally block (so it lands even on failure)' {
        $tryStmts = $script:ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $auditFinally = $tryStmts | Where-Object {
            $_.Finally -and $_.Finally.Extent.Text -match 'Export-Csv'
        }
        $auditFinally | Should -Not -BeNullOrEmpty
    }
    It 'audit row includes the security-relevant fields' {
        foreach ($field in @('Timestamp', 'Operator', 'Account', 'UserType', 'Outcome')) {
            $script:source | Should -Match ([regex]::Escape($field))
        }
    }
    It 'creates the CSV header on first write and appends thereafter' {
        $script:source | Should -Match 'Export-Csv\s+-Path\s+.+\s+-NoTypeInformation\s+-Append'
    }
    It 'guards the audit write and falls back to the log file on failure' {
        # A locked / inaccessible CSV must not throw out of the finally and
        # mask the reset outcome - the Export-Csv sits in its own try/catch
        # whose catch records the row via Write-LogFile instead.
        $auditFinally = $script:ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Where-Object {
            $_.Finally -and $_.Finally.Extent.Text -match 'Export-Csv'
        } | Select-Object -First 1
        $auditFinally | Should -Not -BeNullOrEmpty
        $innerTry = $auditFinally.Finally.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Where-Object {
            $_.Body.Extent.Text -match 'Export-Csv'
        } | Select-Object -First 1
        $innerTry | Should -Not -BeNullOrEmpty
        ($innerTry.CatchClauses.Extent.Text -join "`n") | Should -Match 'Write-LogFile'
    }
}

# =============================================================================
# Target-user resolution and OU-scope guard
# =============================================================================
# -UserName is declared [string] (a SamAccountName). The resolved AD object MUST
# live in a separate variable ($UserObject); dereferencing $UserName as an object
# ($User.SamAccountName) coerces/throws under Set-StrictMode Latest and silently
# breaks BOTH the interactive and the parameter path. These tests pin the fixed
# shape so that regression cannot return, and pin the OU-scope refusal so the
# -UserName path can't reset an account outside the run's Staff / HiPriv subtree.
# =============================================================================
Describe 'ChangePassword.ps1 target-user resolution and OU-scope guard' {
    BeforeAll {
        $cpPath = Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'
        $script:source = Get-Content $cpPath -Raw
        $script:ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $cpPath, [ref]$null, [ref]$null
        )
        # All Get-ADUser calls, then the subset that are *lookups* (carry -filter)
        # as opposed to the success-pipeline projection (Get-ADUser -Properties).
        $script:getAdUserCalls = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Get-ADUser'
        }, $true)
        $script:lookupCalls = @($script:getAdUserCalls | Where-Object {
            $_.CommandElements | Where-Object {
                $_ -is [System.Management.Automation.Language.CommandParameterAst] -and
                $_.ParameterName -eq 'filter'
            }
        })
    }
    It 'never dereferences the [string] -UserName parameter as an object ($UserName.SamAccountName)' {
        # The exact regression this guards: $UserName is [string]; $UserName.<prop> is a
        # property reference on a string -> throws under StrictMode Latest.
        $script:source | Should -Not -Match '\$UserName\.SamAccountName'
    }
    It 'resolves into a separate $UserObject and takes the reset Identity from it' {
        $script:source | Should -Match '\$UserObject\.SamAccountName'
        $script:source | Should -Match 'Identity\s*=\s*\$UserObject\.SamAccountName'
    }
    It 'gates the reset body on the resolved object, not the raw parameter' {
        # if ($UserObject) - not if ($UserName) - guards the New-Password / reset body.
        $script:source | Should -Match '(?s)if\s*\(\s*\$UserObject\s*\)\s*\{.*?New-Password'
    }
    It 'has an interactive (Out-GridView) path and a parameter (-UserName) path' {
        $script:source | Should -Match '(?s)if\s*\(\s*!\s*\$UserName\s*\).*?Out-GridView'
        $script:source | Should -Match '(?s)else\s*\{.*?Get-ADUser'
    }
    It 'scopes BOTH Get-ADUser lookups (the -filter calls) to -SearchBase $OUPath -SearchScope Subtree' {
        # The parameter path must be scoped too, or -UserName could reach an account
        # outside the Staff / HiPriv subtree this run is gated to. NB: the
        # success-pipeline Get-ADUser -Properties projection is intentionally
        # excluded - it isn't a lookup and carries no -filter.
        $script:lookupCalls.Count | Should -BeGreaterOrEqual 2
        foreach ($call in $script:lookupCalls) {
            $call.Extent.Text | Should -Match '-SearchBase\s+\$OUPath'
            $call.Extent.Text | Should -Match '-SearchScope\s+Subtree'
        }
    }
    It 'binds $UserName as a value in the parameter-path filter (single-quoted, injection-safe)' {
        # Single quotes => the AD filter engine binds $UserName as a value rather than
        # interpolating it. A double-quoted filter here would be an AD-filter
        # injection foothold.
        $script:source | Should -Match '-filter\s+''SamAccountName\s+-eq\s+\$UserName'''
    }
    It 'refuses (throws) when a supplied -User is not found within the OU scope' {
        $elseThrow = [regex]::Match(
            $script:source,
            '(?s)else\s*\{.*?if\s*\(\s*!\s*\$UserObject\s*\).*?throw'
        )
        $elseThrow.Success | Should -BeTrue
    }
    It 'logs the refusal before throwing (durable trace of a denied attempt)' {
        $script:source | Should -Match 'Write-LogFile[^\r\n]*REFUSED'
    }
}
