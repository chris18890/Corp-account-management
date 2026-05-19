#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Body-structure tests for Enable-CloudAdmin.ps1 and Disable-CloudAdmin.ps1.
#
# These scripts implement the JIT cloud-admin model that stands in for PIM
# (the tenant has no Entra ID P2). The most security-relevant property is
# the *rollback* path on Enable-CloudAdmin: if Update-MgUser succeeds but
# Register-ScheduledTask then fails, the account would otherwise stay
# enabled indefinitely with no auto-disable. The catch must call
# Update-MgUser -AccountEnabled:$false before rethrowing.
#
# AST + source-pattern based rather than behavioural. Full mocks of the
# AD + Graph + Scheduled-Task surface would be brittle for the value
# gained; the structural shape of the try/catch/finally is what matters
# and the AST captures that precisely.
#
# What's pinned:
#  - Tier -> Prefix derivation (Cloud -> ca., Global -> ga.)
#  - Authorisation gate (HiPriv_Account_Admins / Domain Admins)
#  - Reason required (refuses on empty)
#  - Duration default and cap against Env.Security.MaxElevationMinutes
#  - Enable: Update-MgUser -AccountEnabled:$true followed by Register-ScheduledTask
#  - Rollback: catch block calls Update-MgUser -AccountEnabled:$false then throws
#  - Audit CSV row written in a finally block (lands even on failure)
#  - Disable: replaces any pending auto-disable task
#  - Disable: handles "already disabled" path with AlreadyDisabled outcome
#
# Convention: single-quoted regex patterns throughout.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    function Script:Read-Ps1 {
        param([string]$Path)
        $tokens = $null; $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    }
    function Script:Get-CommandsByName {
        param($Node, [string]$Name)
        $Node.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq $Name
        }, $true)
    }
}

# =============================================================================
# Enable-CloudAdmin.ps1
# =============================================================================
Describe 'Enable-CloudAdmin.ps1 parameter contract (cross-check with ScriptParameters)' {
    BeforeAll {
        $script:enablePath = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
        $script:enableCmd  = Get-Command $script:enablePath
    }
    It 'DurationMinutes has a default of 0 (meaning use the cap)' {
        # The script applies the cap when DurationMinutes -eq 0, so the
        # default value is semantically load-bearing.
        $defaults = $script:enableCmd.Parameters['DurationMinutes'].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] }
        # Default is at the AST level; check via source
        $src = Get-Content $script:enablePath -Raw
        $src | Should -Match '\[int\]\$DurationMinutes\s*=\s*0'
    }
}

Describe 'Enable-CloudAdmin.ps1 prefix derivation' {
    BeforeAll {
        $script:enableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1') -Raw
    }
    It "maps Tier 'Cloud' to prefix 'ca.'" {
        $script:enableSource | Should -Match '"Cloud"\s*\{\s*"ca\."\s*\}'
    }
    It "maps Tier 'Global' to prefix 'ga.'" {
        $script:enableSource | Should -Match '"Global"\s*\{\s*"ga\."\s*\}'
    }
    It 'constructs UPN as <prefix><UserName>@<EmailSuffix>' {
        $script:enableSource | Should -Match '\$AccountName\s*=\s*"\$Prefix\$UserName"'
        $script:enableSource | Should -Match '\$UPN\s*=\s*"\$AccountName@\$EmailSuffix"'
    }
}

Describe 'Enable-CloudAdmin.ps1 authorisation gate' {
    BeforeAll {
        $script:enableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1') -Raw
    }
    It 'gates on HiPriv_Account_Admins' {
        $script:enableSource | Should -Match 'HiPriv_Account_Admins'
    }
    It 'gates on Domain Admins' {
        $script:enableSource | Should -Match "'Domain Admins'"
    }
    It 'throws when invoker lacks required group membership' {
        $script:enableSource | Should -Match '(?ms)not\s+authorised.*?throw'
    }
}

Describe 'Enable-CloudAdmin.ps1 reason requirement' {
    BeforeAll {
        $script:enableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1') -Raw
    }
    It 'prompts for Reason via Read-Host when not supplied' {
        $script:enableSource | Should -Match 'Reason\s*=\s*Read-Host'
    }
    It 'throws when Reason is null or whitespace' {
        $script:enableSource | Should -Match '(?ms)IsNullOrWhiteSpace\(\$Reason\).*?throw'
    }
}

Describe 'Enable-CloudAdmin.ps1 duration handling' {
    BeforeAll {
        $script:enableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1') -Raw
    }
    It 'reads MaxMinutes from environment.psd1' {
        $script:enableSource | Should -Match '\$MaxMinutes\s*=\s*\$Env\.Security\.MaxElevationMinutes'
    }
    It 'applies the cap default when DurationMinutes is 0' {
        $script:enableSource | Should -Match '\$DurationMinutes\s+-eq\s+0.*?\$DurationMinutes\s*=\s*\$MaxMinutes'
    }
    It 'caps any request exceeding MaxMinutes' {
        # (?s) so . matches newlines - source spans the if-block across
        # multiple lines with a Write-LogFile between the condition
        # and the assignment.
        $script:enableSource | Should -Match '(?s)\$DurationMinutes\s+-gt\s+\$MaxMinutes.*?\$DurationMinutes\s*=\s*\$MaxMinutes'
    }
    It 'computes DisableTime as Now + DurationMinutes' {
        $script:enableSource | Should -Match '\$DisableTime\s*=\s*\$EnableTime\.AddMinutes\(\$DurationMinutes\)'
    }
}

Describe 'Enable-CloudAdmin.ps1 enable path' {
    BeforeAll {
        $script:enablePath   = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
        $script:enableSource = Get-Content $script:enablePath -Raw
        $script:enableAst    = Read-Ps1 $script:enablePath
        $script:enableCalls  = Get-CommandsByName $script:enableAst 'Update-MgUser'
    }
    It 'calls Update-MgUser at least twice (enable + rollback paths)' {
        $script:enableCalls.Count | Should -BeGreaterOrEqual 2
    }
    It 'enables with -AccountEnabled:$true at least once' {
        $enableCalls = @($script:enableCalls | Where-Object {
            $_.Extent.Text -match '-AccountEnabled:\$true'
        })
        $enableCalls.Count | Should -BeGreaterOrEqual 1
    }
    It 'registers the auto-disable scheduled task' {
        Get-CommandsByName $script:enableAst 'Register-ScheduledTask' | Should -Not -BeNullOrEmpty
    }
    It 'replaces any existing auto-disable task before registering' {
        # Avoids stale tasks lingering when re-elevating an already-enabled account.
        $script:enableSource | Should -Match '(?s)Get-ScheduledTask\s+-TaskName\s+\$TaskName.*?Unregister-ScheduledTask'
    }
    It 'verifies Disable-CloudAdmin.ps1 exists before registering the task' {
        # The scheduled task action points at Disable-CloudAdmin.ps1; if
        # that file is missing, the task is non-functional. Better to
        # fail fast at enable time.
        $script:enableSource | Should -Match 'Disable-CloudAdmin\.ps1'
        $script:enableSource | Should -Match 'Test-Path\s+\$DisableScript'
    }
}

Describe 'Enable-CloudAdmin.ps1 rollback path (security-critical)' {
    BeforeAll {
        $script:enablePath   = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
        $script:enableSource = Get-Content $script:enablePath -Raw
        $script:enableAst    = Read-Ps1 $script:enablePath
    }
    # If the script enables the account but then fails to register the
    # auto-disable task, the account would otherwise stay enabled with
    # no automatic cleanup. The rollback in the catch is what prevents
    # that. This is the single most important property of the script.
    It 'has a try block containing both Update-MgUser AND Register-ScheduledTask' {
        $tryStmts = $script:enableAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $enableTry = $tryStmts | Where-Object {
            $_.Body.Extent.Text -match 'Update-MgUser.*-AccountEnabled:\$true' -and
            $_.Body.Extent.Text -match 'Register-ScheduledTask'
        }
        $enableTry | Should -Not -BeNullOrEmpty
    }
    It 'catch block calls Update-MgUser -AccountEnabled:$false (rollback)' {
        # Find the try that contains the enable, then inspect its catch.
        $tryStmts = $script:enableAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $enableTry = $tryStmts | Where-Object {
            $_.Body.Extent.Text -match 'Update-MgUser.*-AccountEnabled:\$true' -and
            $_.Body.Extent.Text -match 'Register-ScheduledTask'
        } | Select-Object -First 1
        $enableTry | Should -Not -BeNullOrEmpty
        $catchBody = ($enableTry.CatchClauses | Select-Object -First 1).Body.Extent.Text
        $catchBody | Should -Match 'Update-MgUser.*-AccountEnabled:\$false'
    }
    It 'catch block rethrows after attempting rollback' {
        $tryStmts = $script:enableAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $enableTry = $tryStmts | Where-Object {
            $_.Body.Extent.Text -match 'Update-MgUser.*-AccountEnabled:\$true' -and
            $_.Body.Extent.Text -match 'Register-ScheduledTask'
        } | Select-Object -First 1
        $catchBody = ($enableTry.CatchClauses | Select-Object -First 1).Body.Extent.Text
        $catchBody | Should -Match '\bthrow\b'
    }
    It 'sets Outcome = RollbackFailed if the rollback itself fails' {
        # Tells the audit trail when the catastrophic case happened.
        # Pin the assignment form rather than the bare quoted string -
        # less brittle and matches the source's double-quoted style.
        $script:enableSource | Should -Match '\$Outcome\s*=\s*"RollbackFailed"'
    }
}

Describe 'Enable-CloudAdmin.ps1 audit CSV write' {
    BeforeAll {
        $script:enablePath = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
        $script:enableAst  = Read-Ps1 $script:enablePath
        $script:enableSrc  = Get-Content $script:enablePath -Raw
    }
    It 'writes to cloud-admin-elevations.csv' {
        $script:enableSrc | Should -Match 'cloud-admin-elevations\.csv'
    }
    It 'audit row write lives in a finally block (lands even on failure)' {
        $tryStmts = $script:enableAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $auditFinally = $tryStmts | Where-Object {
            $_.Finally -and $_.Finally.Extent.Text -match 'Export-Csv'
        }
        $auditFinally | Should -Not -BeNullOrEmpty
    }
    It 'audit row includes the security-relevant fields' {
        foreach ($field in @('Timestamp', 'Action', 'Operator', 'Account', 'Tier',
                             'DurationMinutes', 'DisableAt', 'Reason', 'Outcome')) {
            $script:enableSrc | Should -Match ([regex]::Escape($field))
        }
    }
    It 'creates the CSV header on first write and appends thereafter' {
        $script:enableSrc | Should -Match 'Test-Path\s+\$AuditFile'
        $script:enableSrc | Should -Match 'Export-Csv\s+-Path\s+\$AuditFile\s+-NoTypeInformation\s+-Append'
    }
}

# =============================================================================
# Disable-CloudAdmin.ps1
# =============================================================================
Describe 'Disable-CloudAdmin.ps1 prefix derivation' {
    BeforeAll {
        $script:disableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1') -Raw
    }
    # Mirror of Enable-CloudAdmin's derivation - Disable must produce the
    # exact same UPN for the same input or it can't find the account to
    # disable.
    It "maps Tier 'Cloud' to prefix 'ca.'" {
        $script:disableSource | Should -Match '"Cloud"\s*\{\s*"ca\."\s*\}'
    }
    It "maps Tier 'Global' to prefix 'ga.'" {
        $script:disableSource | Should -Match '"Global"\s*\{\s*"ga\."\s*\}'
    }
    It 'derives the same UPN form as Enable-CloudAdmin (round-trip property)' {
        # If Enable and Disable disagree on UPN form, scheduled auto-disable
        # tasks call a script that can't find what was enabled. Pin both.
        $script:disableSource | Should -Match '\$UPN\s*=\s*"\$AccountName@\$EmailSuffix"'
    }
}

Describe 'Disable-CloudAdmin.ps1 authorisation gate' {
    BeforeAll {
        $script:disableSource = Get-Content (Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1') -Raw
    }
    It 'gates on HiPriv_Account_Admins' {
        $script:disableSource | Should -Match 'HiPriv_Account_Admins'
    }
    It 'gates on Domain Admins' {
        $script:disableSource | Should -Match "'Domain Admins'"
    }
    It 'throws when invoker lacks required group membership' {
        $script:disableSource | Should -Match '(?ms)not\s+authorised.*?throw'
    }
}

Describe 'Disable-CloudAdmin.ps1 disable path' {
    BeforeAll {
        $script:disablePath = Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1'
        $script:disableSrc  = Get-Content $script:disablePath -Raw
        $script:disableAst  = Read-Ps1 $script:disablePath
    }
    It 'calls Update-MgUser with -AccountEnabled:$false' {
        $calls = Get-CommandsByName $script:disableAst 'Update-MgUser'
        $disableCalls = @($calls | Where-Object {
            $_.Extent.Text -match '-AccountEnabled:\$false'
        })
        $disableCalls.Count | Should -BeGreaterOrEqual 1
    }
    It 'unregisters the auto-disable scheduled task if one exists' {
        $script:disableSrc | Should -Match 'Get-ScheduledTask\s+-TaskName\s+\$TaskName'
        $script:disableSrc | Should -Match 'Unregister-ScheduledTask\s+-TaskName\s+\$TaskName'
    }
}

Describe 'Disable-CloudAdmin.ps1 already-disabled idempotency' {
    BeforeAll {
        $script:disableSrc = Get-Content (Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1') -Raw
    }
    # Running Disable twice should be safe - the second run records
    # AlreadyDisabled in the audit trail and does not call Update-MgUser.
    It 'checks $MgUser.AccountEnabled before calling Update-MgUser' {
        $script:disableSrc | Should -Match '!\$MgUser\.AccountEnabled'
    }
    It "sets Outcome = AlreadyDisabled on the idempotent path" {
        $script:disableSrc | Should -Match '\$Outcome\s*=\s*"AlreadyDisabled"'
    }
}

Describe 'Disable-CloudAdmin.ps1 audit CSV write' {
    BeforeAll {
        $script:disablePath = Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1'
        $script:disableSrc  = Get-Content $script:disablePath -Raw
        $script:disableAst  = Read-Ps1 $script:disablePath
    }
    It 'writes to the same cloud-admin-elevations.csv as Enable-CloudAdmin' {
        $script:disableSrc | Should -Match 'cloud-admin-elevations\.csv'
    }
    It 'audit row write lives in a finally block' {
        $tryStmts = $script:disableAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $auditFinally = $tryStmts | Where-Object {
            $_.Finally -and $_.Finally.Extent.Text -match 'Export-Csv'
        }
        $auditFinally | Should -Not -BeNullOrEmpty
    }
    It "Action column is 'Disable'" {
        $script:disableSrc | Should -Match 'Action\s*=\s*"Disable"'
    }
}

# =============================================================================
# Cross-script consistency
# =============================================================================
Describe 'Enable/Disable-CloudAdmin audit CSV column consistency' {
    BeforeAll {
        $script:enableSrc  = Get-Content (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1')  -Raw
        $script:disableSrc = Get-Content (Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1') -Raw
    }
    # Both scripts append to the same CSV; column drift between them
    # would produce a malformed audit trail.
    It 'both scripts reference the same audit file path' {
        $script:enableSrc  | Should -Match 'LogFiles\\cloud-admin-elevations\.csv'
        $script:disableSrc | Should -Match 'LogFiles\\cloud-admin-elevations\.csv'
    }
    It 'both write the same column set' {
        $columns = @('Timestamp','Action','Operator','Account','Tier','DurationMinutes','DisableAt','Reason','Outcome')
        foreach ($col in $columns) {
            $script:enableSrc  | Should -Match ([regex]::Escape($col))
            $script:disableSrc | Should -Match ([regex]::Escape($col))
        }
    }
}
