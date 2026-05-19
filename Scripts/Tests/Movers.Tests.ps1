#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Body-structure tests for Movers.ps1. The parameter contract is already
# pinned in ScriptParameters.Tests.ps1; this file pins the load-bearing
# safety logic that lives in the body and would otherwise be untested:
#
#  - Authorisation gate (Standard_Account_Admins / Standard_Group_Admins)
#  - Refuses to move Hi-Priv accounts (security-relevant: Hi-Priv accounts
#    have their own lifecycle scripts and must never be moved by Movers)
#  - Refuses to move disabled accounts
#  - Updates Department attribute only when current value differs
#  - Removes from old dept group only when OldDept differs from NewDept
#  - Tolerates "not a member" without throwing
#  - Updates manager only when NewMgrSam is non-empty
#
# SEQUENCE (load-bearing, see the ordering Describe below):
#  The department change must run as: verify new dept group exists
#  (skip the row with continue if not) -> join new group -> write
#  Department attribute -> remove from old group. This add-before-write
#  and add-before-remove ordering means a mid-row failure leaves the
#  user over-entitled (in both groups) or fully untouched, never with a
#  Department attribute pointing at a group they were never added to.
#
# AST-based rather than behavioural because every code path in Movers.ps1
# mutates AD; mocking the full Get-ADUser / Set-ADUser / Add-GroupMember /
# Remove-ADGroupMember surface end-to-end would be more brittle than
# asserting the source structure directly.
#
# Convention: single-quoted regex patterns to sidestep PowerShell's
# variable interpolation and backslash handling.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $script:scriptPath  = Join-Path $script:scriptsRoot 'Users\Movers.ps1'
    $script:source      = Get-Content $script:scriptPath -Raw
    $script:lines       = Get-Content $script:scriptPath
    $tokens = $null; $errors = $null
    $script:ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:scriptPath, [ref]$tokens, [ref]$errors
    )
    # Helper: 1-based line number of the first source line matching a regex.
    function Script:Get-FirstLine {
        param([string]$Pattern)
        ($script:lines | Select-String -Pattern $Pattern | Select-Object -First 1).LineNumber
    }
}

Describe 'Movers.ps1 required modules and hardening' {
    It 'requires ActiveDirectory'    { $script:source | Should -Match '#Requires\s+-Modules\s+ActiveDirectory' }
    It 'requires RunAsAdministrator' { $script:source | Should -Match '#Requires\s+-RunAsAdministrator' }
    It 'imports CorpAdmin from Modules\CorpAdmin\CorpAdmin.psd1' {
        $script:source | Should -Match 'Modules.CorpAdmin.CorpAdmin\.psd1'
    }
    It 'declares Set-StrictMode -Version Latest' {
        $script:source | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
}

Describe 'Movers.ps1 authorisation gate' {
    It 'gates on Standard_Account_Admins' {
        $script:source | Should -Match 'Standard_Account_Admins'
    }
    It 'gates on Standard_Group_Admins' {
        $script:source | Should -Match 'Standard_Group_Admins'
    }
    It 'throws when invoker lacks required group membership' {
        $script:source | Should -Match '(?ms)not\s+authorised.*?throw'
    }
}

Describe 'Movers.ps1 Hi-Priv refusal' {
    # Hi-Priv accounts must never be processed by Movers - they belong in
    # the Administration OU and have separate lifecycle scripts. Losing
    # this check is a privilege-escalation surface (a department move
    # that touched a Hi-Priv account could fold it into a Staff dept
    # group, granting Staff-level access to a Hi-Priv account).
    It 'computes the HiPrivOU path from environment.psd1' {
        $script:source | Should -Match '\$HiPrivOU\s*=.*HiPrivAccounts'
    }
    It 'compares user DistinguishedName against the HiPriv OU path' {
        $script:source | Should -Match '\$User\.DistinguishedName\s+-like\s+"\*\$HiPrivOU\*"'
    }
    It 'skips Hi-Priv accounts with continue (not throw)' {
        # The Hi-Priv detection block should `continue` to the next CSV row,
        # not throw - throwing would abort the whole batch.
        $script:source | Should -Match '(?s)is Hi-Priv.*?continue'
    }
    It 'logs the refusal at WARNING level' {
        $script:source | Should -Match 'Hi-Priv.*?-ForegroundColor\s+Yellow'
    }
}

Describe 'Movers.ps1 disabled-account refusal' {
    It 'checks -not $User.Enabled' {
        $script:source | Should -Match '-not\s+\$User\.Enabled'
    }
    It 'skips disabled accounts with continue (not throw)' {
        $script:source | Should -Match '(?s)is\s+disabled.*?continue'
    }
}

Describe 'Movers.ps1 department update' {
    BeforeAll {
        $script:setUserCalls = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Set-ADUser'
        }, $true)
    }
    It 'calls Set-ADUser at least once' {
        $script:setUserCalls.Count | Should -BeGreaterThan 0
    }
    It 'sets the -Department parameter on the user' {
        $deptCalls = @($script:setUserCalls | Where-Object {
            $_.Extent.Text -match '-Department\s+\$NewDept'
        })
        $deptCalls.Count | Should -BeGreaterOrEqual 1
    }
    It 'only updates Department when current value differs from new' {
        # Avoids no-op writes against the DC and keeps the audit log clean.
        $script:source | Should -Match '\$User\.Department\s+-ne\s+\$NewDept'
    }
}

Describe 'Movers.ps1 new-department group membership' {
    It 'verifies the new department group exists before joining' {
        # Get-ADGroup gate before Add-GroupMember; a missing group must
        # short-circuit the row before any mutation. Otherwise a typo in
        # NEW_DEPT would silently leave the user without a dept group.
        $script:source | Should -Match '(?s)Get-ADGroup\s+-Filter\s+"Name\s+-eq\s+''\$NewDept''".*?Add-GroupMember'
    }
    It 'skips the row cleanly when the new department group does not exist (no throw, no mutation)' {
        # REGRESSION GUARD for the reorder fix. The previous version wrote
        # the Department attribute first and then threw on a missing group,
        # leaving the attribute set to a dept the user was never added to;
        # the idempotent "already set" guard then masked the half-completed
        # row on re-run. The fix gates ALL mutation behind the group-exists
        # check and aborts the row via `continue` (not throw) so a re-run
        # after the group is created starts from a clean state.
        #
        # Line-anchored (not (?s)) on purpose: the missing-group message must
        # be followed immediately by `continue`. A greedy (?s) here would
        # reach the legitimate `throw` in the old-dept removal catch further
        # down the same row and give a false reading either way.
        $script:source | Should -Match '(?m)Department group \$NewDept does not exist[^\n]*\r?\n\s*continue'
    }
    It 'does NOT throw on a missing new department group' {
        # Explicit negative: the line after the missing-new-group message
        # must not be a throw. (A throw here would be caught by the per-row
        # catch and logged as a generic "ERROR processing" - misleading for
        # what is a data-entry problem - and under the old ordering it only
        # fired after the Department attribute had already been written.)
        # Line-anchored so it inspects only the missing-group branch, not the
        # unrelated old-dept removal throw later in the same row.
        $script:source | Should -Not -Match '(?m)Department group \$NewDept does not exist[^\n]*\r?\n\s*throw'
    }
}

Describe 'Movers.ps1 department-change ordering (safe sequence)' {
    # The whole point of the reorder is sequence. These pins lock it in:
    #   join new group  ->  write Department attribute  ->  remove old group
    # add-before-write : a failure can't leave Department pointing at a
    #                    group the user was never added to.
    # add-before-remove: a failure leaves the user over-entitled (in both
    #                    groups) rather than under-entitled (in neither).
    BeforeAll {
        $script:addNewLine  = Get-FirstLine 'Add-GroupMember.*-Group\s+\$NewDept'
        $script:setDeptLine = Get-FirstLine 'Set-ADUser\s+-Identity\s+\$User\s+-Department\s+\$NewDept'
        $script:removeLine  = Get-FirstLine 'Remove-ADGroupMember\s+-Identity\s+\$OldDept'
        $script:gateLine    = Get-FirstLine 'Get-ADGroup\s+-Filter\s+"Name\s+-eq\s+''\$NewDept''"'
    }
    It 'all four anchor lines are present' {
        $script:gateLine    | Should -Not -BeNullOrEmpty
        $script:addNewLine  | Should -Not -BeNullOrEmpty
        $script:setDeptLine | Should -Not -BeNullOrEmpty
        $script:removeLine  | Should -Not -BeNullOrEmpty
    }
    It 'verifies the new group exists before joining it' {
        $script:gateLine | Should -BeLessThan $script:addNewLine
    }
    It 'joins the new department group before writing the Department attribute' {
        $script:addNewLine | Should -BeLessThan $script:setDeptLine
    }
    It 'removes from the old department group only after joining the new one (add-before-remove)' {
        $script:addNewLine | Should -BeLessThan $script:removeLine
    }
}

Describe 'Movers.ps1 old-department group cleanup' {
    It 'only attempts removal when OldDept differs from NewDept' {
        $script:source | Should -Match '\$OldDept\s+-and\s+\$OldDept\s+-ne\s+\$NewDept'
    }
    It 'verifies the old group exists before Remove-ADGroupMember' {
        $script:source | Should -Match '(?s)Get-ADGroup\s+-Filter\s+"Name\s+-eq\s+''\$OldDept''".*?Remove-ADGroupMember'
    }
    It 'tolerates "not a member" without rethrowing' {
        # If the user was never in the old dept group (data drift), the
        # mover should log and continue rather than abort the batch.
        $script:source | Should -Match 'not\s+a\s+member'
    }
}

Describe 'Movers.ps1 manager update' {
    It 'only attempts manager update when NewMgrSam is non-empty' {
        $script:source | Should -Match 'if\s*\(\s*\$NewMgrSam\s*\)'
    }
    It 'looks up the manager via Get-ADUser before setting' {
        $script:source | Should -Match '(?s)Get-ADUser\s+-Filter\s+"sAMAccountName\s+-eq\s+''\$NewMgrSam''".*?Set-ADUser.*?-Manager'
    }
    It 'logs a warning when the new manager is not found (does not throw)' {
        # A bad NEW_MANAGER value shouldn't abort the batch.
        $script:source | Should -Match 'Manager\s+\$NewMgrSam\s+not\s+found'
    }
}

Describe 'Movers.ps1 CSV contract' {
    It 'reads from movers.csv next to the script' {
        $script:source | Should -Match '\$ScriptPath\\movers\.csv'
    }
    It 'validates all four required CSV headers' {
        foreach ($header in @('USERNAME', 'OLD_DEPT', 'NEW_DEPT', 'NEW_MANAGER')) {
            $script:source | Should -Match ([regex]::Escape($header))
        }
    }
    It 'throws when a required header is missing' {
        $script:source | Should -Match 'missing required column'
    }
}

Describe 'Movers.ps1 batch-error tolerance' {
    It 'wraps per-row processing in try/catch with continue on error' {
        # One bad row in the CSV shouldn't abort the rest of the batch.
        $perRowCatches = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $perRowCatches.Count | Should -BeGreaterThan 0
    }
    It 'continues to the next user on error' {
        $script:source | Should -Match '(?s)ERROR processing.*?continue'
    }
}
