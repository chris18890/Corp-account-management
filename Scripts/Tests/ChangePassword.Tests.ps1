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
#  - GUI feedback via VisualBasic MsgBox on both success and failure
#  - SecureString conversion zeroed after use
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
    It 'shows a MsgBox and returns when the password fails policy' {
        # On policy failure, the script should NOT call Set-ADAccountPassword.
        $script:source | Should -Match '(?s)try\s*\{\s*Test-Password.*?\}\s*catch\s*\{.*?Msgbox.*?return'
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
    It 'shows the resulting user state in a MsgBox on success' {
        $script:source | Should -Match '(?s)Set-ADAccountPassword[\s\S]*?Msgbox'
    }
    It 'shows the exception message in a MsgBox on failure' {
        $script:source | Should -Match 'Failed to reset password.*?\$\(\$_\.Exception\.Message\)'
    }
}

Describe 'ChangePassword.ps1 SecureString hygiene' {
    BeforeAll {
        $script:source = Get-Content (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1') -Raw
    }
    It 'converts to SecureString with -AsPlainText -Force' {
        $script:source | Should -Match 'ConvertTo-SecureString\s+-String\s+\$Password\s+-AsPlainText\s+-Force'
    }
    It 'zeroes the plain-text Password variable after conversion' {
        # Limits how long the cleartext lives in process memory.
        $script:source | Should -Match '\$Password\s*=\s*\$null'
    }
    It 'zeroes the SecureString variable after the reset' {
        $script:source | Should -Match '\$NewPassword\s*=\s*\$null'
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
}
