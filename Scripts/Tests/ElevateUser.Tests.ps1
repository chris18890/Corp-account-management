#====================================================================
# ElevateUser.ps1 - security-contract tests
#
# ElevateUser.ps1 performs live AD operations and prompts interactively,
# so it is pinned by source/AST assertions rather than executed. These lock
# the security-critical decisions: tiered authorisation, the protected-group
# gate, the Tier-0 target rule, the elevation cap, Remove idempotency, the
# forced reason, and the guaranteed audit row.
#====================================================================

BeforeAll {
    $script:scriptsRoot = Split-Path $PSScriptRoot -Parent
    $script:path = Join-Path $script:scriptsRoot 'Users\ElevateUser.ps1'
    $script:src  = Get-Content $script:path -Raw
    $tokens = $null
    $errors = $null
    $script:ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:path,
        [ref]$tokens,
        [ref]$errors
    )
    $errors | Should -BeNullOrEmpty
}

Describe 'ElevateUser.ps1 parameter contract' {
    It 'restricts UserAction to E or R' {
        $script:src | Should -Match 'ValidateSet\("E","R"\)'
    }
    It 'restricts TempOrPerm to P or T' {
        $script:src | Should -Match 'ValidateSet\("P","T"\)'
    }
    It 'requires TimeSpan to be a positive int' {
        $script:src | Should -Match 'ValidateRange\(1,\s*\[int\]::MaxValue\)'
    }
    It 'makes GroupName, UserName, UserAction mandatory' {
        $script:src | Should -Match 'Mandatory\)\]\[string\]\$GroupName'
        $script:src | Should -Match 'Mandatory\)\]\[string\]\$UserName'
        $script:src | Should -Match 'Mandatory\)\]\[ValidateSet\("E","R"\)\]\[string\]\$UserAction'
    }
}

Describe 'ElevateUser.ps1 tiered authorisation' {
    It 'defines the protected (forest + Tier-0) group set' {
        $script:src | Should -Match "'Domain Admins'"
        $script:src | Should -Match "'Enterprise Admins'"
        $script:src | Should -Match "'Schema Admins'"
        $script:src | Should -Match 'RolePrefix\)Tier0_Level_3_Admins'
    }
    It 'classifies the invoker as high-tier and standard via Test-IsMemberOf' {
        $script:src | Should -Match '\$HighTierAdminGroups[\s\S]*?HiPriv_Group_Admins'
        $script:src | Should -Match '\$HighTierAdminGroups[\s\S]*?Domain Admins'
        $script:src | Should -Match '\$InvokerIsHighTier\s*=\s*Test-IsMemberOf'
        $script:src | Should -Match '\$InvokerIsStandard\s*=\s*Test-IsMemberOf'
    }
    It 'admits the level-1 access roles into the STANDARD tier' {
        $script:src | Should -Match '\$StandardAdminGroups[\s\S]*?Standard_Group_Admins'
        $script:src | Should -Match '\$StandardAdminGroups[\s\S]*?SER_Access_Admins'
        $script:src | Should -Match '\$StandardAdminGroups[\s\S]*?Local_Admin_Group_Admins'
        $script:src | Should -Match '\$InvokerIsStandard\s*=\s*Test-IsMemberOf[^\r\n]*\$StandardAdminGroups'
    }
    It 'keeps protected groups high-tier-only (access roles must NOT be in the high tier)' {
        $hiTier = ($script:src -split "\r?\n") | Where-Object { $_ -match '\$HighTierAdminGroups\s*=\s*@\(' }
        $hiTier | Should -Not -BeNullOrEmpty
        $hiTier | Should -Not -Match 'SER_Access_Admins'
        $hiTier | Should -Not -Match 'Local_Admin_Group_Admins'
    }
    It 'gates protected groups on high-tier alone, so a multi-tier member is not shadowed by the standard tier' {
        # The tier flags are additive (-or), so someone in both tiers - e.g. a level-3
        # tech in SER_Access_Admins + Standard_Group_Admins + HiPriv_Group_Admins -
        # keeps high-tier reach. The protected-group gate must read $InvokerIsHighTier
        # directly, never a "standard and NOT high-tier" composite that would shadow it.
        $script:src | Should -Match '-not\s+\$InvokerIsHighTier'
        $script:src | Should -Not -Match '\$InvokerIsStandard\s+-and\s+-not\s+\$InvokerIsHighTier'
    }
    It 'aborts when the invoker is in neither tier' {
        $script:src | Should -Match 'if\s*\(-not\s*\(\$InvokerIsHighTier\s*-or\s*\$InvokerIsStandard\)\)'
        $script:src | Should -Match 'throw "Invoker is not authorised'
    }
    It 'audits unauthorised invoker attempts before throwing' {
        $idxAuth  = $script:src.IndexOf('if (-not ($InvokerIsHighTier -or $InvokerIsStandard))')
        $idxAudit = $script:src.IndexOf('Write-ElevationAuditRow', $idxAuth)
        $idxDeny  = $script:src.IndexOf('-Outcome "Denied"', $idxAudit)
        $idxThrow = $script:src.IndexOf('throw "Invoker is not authorised', $idxDeny)
        $idxAuth  | Should -BeGreaterThan -1
        $idxAudit | Should -BeGreaterThan $idxAuth
        $idxDeny  | Should -BeGreaterThan $idxAudit
        $idxThrow | Should -BeGreaterThan $idxDeny
    }
    It 'uses a specific audit reason when an unauthorised invoker has not supplied one' {
        $script:src | Should -Match '\$UnauthorisedAuditReason\s*=\s*if\s*\(\[string\]::IsNullOrWhiteSpace\(\$Reason\)\)'
        $script:src | Should -Match 'Unauthorised execution attempt before reason capture'
        $script:src | Should -Match '-Reason\s+\$UnauthorisedAuditReason'
    }
}

Describe 'ElevateUser.ps1 audit helper' {
    It 'defines a reusable audit writer' {
        $script:src | Should -Match 'function\s+Write-ElevationAuditRow'
    }
    It 'writes the expected audit fields' {
        foreach ($f in 'Timestamp','Action','Operator','Account','Group','DurationMinutes','Reason','Outcome') {
            $script:src | Should -Match $f
        }
    }
    It 'supports append and first-write CSV creation' {
        $script:src | Should -Match 'Export-Csv'
        $script:src | Should -Match '-Append'
        $script:src | Should -Match 'Test-Path\s+\$AuditFile'
    }
}

Describe 'ElevateUser.ps1 protected-group gate (symmetric across E and R)' {
    It 'gates protected groups on high-tier BEFORE the E/R action switch' {
        $idxResolve   = $script:src.IndexOf('$TargetGroupObj = Get-ADGroup -Identity $GroupName -Server $DCHostName -ErrorAction Stop')
        $idxProtected = $script:src.IndexOf('$TargetIsProtected = $ProtectedGroups -contains $TargetGroupObj.Name')
        $idxGate      = $script:src.IndexOf('only high-tier admins may modify it')
        $actionSwitch = $script:ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.SwitchStatementAst] -and
            $n.Condition.Extent.Text -match '\$UserAction\.ToUpper\(\)' -and
            $n.Extent.Text -match 'Add-GroupMember' -and
            $n.Extent.Text -match 'Remove-GroupMember'
        }, $true) | Select-Object -First 1
        $actionSwitch | Should -Not -BeNullOrEmpty
        $idxSwitch = $actionSwitch.Extent.StartOffset
        $idxResolve   | Should -BeGreaterThan -1
        $idxProtected | Should -BeGreaterThan $idxResolve
        $idxGate      | Should -BeGreaterThan $idxProtected
        $idxSwitch    | Should -BeGreaterThan $idxGate
    }
    It 'denies-and-audits rather than throwing on the protected-group rejection' {
        $script:src | Should -Match 'DENIED:[^\r\n]*protected/Tier-0 group'
        $script:src | Should -Match '\$Outcome\s*=\s*"Denied"'
    }
}

Describe 'ElevateUser.ps1 elevate rules' {
    It 'only allows a Tier-0 / Level-3 account into a protected group' {
        $script:src | Should -Match 'Test-IsMemberOf -Sam \$UserName -GroupNames @\("\$\(\$Env\.Groups\.RolePrefix\)Tier0_Level_3_Admins"\)'
        $script:src | Should -Match 'non-Tier-0 / non-Level-3 account'
    }
    It 'caps a temporary elevation at MaxElevationMinutes' {
        $script:src | Should -Match 'if\s*\(\$TimeSpan\s*-gt\s*\$Env\.Security\.MaxElevationMinutes\)'
        $script:src | Should -Match '\$Outcome\s*=\s*"Rejected"'
    }
    It 'defaults the temporary window to 60 minutes' {
        $script:src | Should -Match '\$TimeSpan\s*=\s*60'
    }
    It 'uses a TTL add for temporary and a plain add for permanent' {
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Member \$UserName -TimeSpan \$TimeSpan'
        $script:src | Should -Match 'Add-GroupMember -LogFile \$LogFile -DCHostName \$DCHostName -Group \$GroupName -Member \$UserName\r?\n'
    }
}

Describe 'ElevateUser.ps1 protected group canonicalisation' {
    BeforeAll {
        $script:scriptsRoot = Split-Path $PSScriptRoot -Parent
        $script:path = Join-Path $script:scriptsRoot 'Users\ElevateUser.ps1'
        $script:src  = Get-Content $script:path -Raw
    }
    It 'resolves the target group before protected-group comparison' {
        $script:src | Should -Match '\$TargetGroupObj\s*=\s*Get-ADGroup\s+-Identity\s+\$GroupName\s+-Server\s+\$DCHostName\s+-ErrorAction\s+Stop'
    }
    It 'compares protected groups using the resolved group Name' {
        $script:src | Should -Match '\$TargetIsProtected\s*=\s*\$ProtectedGroups\s+-contains\s+\$TargetGroupObj\.Name'
    }
}

Describe 'ElevateUser.ps1 remove delegates to the verified module helper' {
    # The removal ALGORITHM - class-agnostic resolution, direct-membership
    # idempotency (NoChange), ambiguity rejection (Rejected), DN-keyed removal
    # and post-remove verification - now lives in Remove-GroupMember
    # (CorpAdmin.psm1) and is exercised behaviourally in CorpAdmin.Tests.ps1's
    # 'Remove-GroupMember' / 'Resolve-GroupMemberObject' describes.
    # ElevateUser.ps1's only job is to DELEGATE to it and fold the returned
    # outcome into the audited $Outcome, so these tests pin the delegation
    # contract rather than re-scraping the (now-moved) algorithm.
    BeforeAll {
        # Locate the action switch on $UserAction (NOT the audit-label switch in
        # the finally; disambiguated by the presence of Add-GroupMember), then
        # pull the text of its "R" clause so assertions are scoped to that branch.
        $script:actionSwitch = $script:ast.FindAll(
            { param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] }, $true
        ) | Where-Object { $_.Condition.Extent.Text -match '\$UserAction' -and $_.Extent.Text -match 'Add-GroupMember' } |
            Select-Object -First 1
        $script:rBody = if ($script:actionSwitch) {
            (
                $script:actionSwitch.Clauses | Where-Object {
                    $clauseLabels = @($_.Item1 | ForEach-Object { $_.Extent.Text.Trim('"''') })
                    $clauseLabels -contains 'R'
                } | Select-Object -First 1
            ).Item2.Extent.Text
        } else {
            ''
        }
    }
    It 'exposes a resolvable R (remove) clause in the UserAction switch' {
        $script:rBody | Should -Not -BeNullOrEmpty
    }
    It 'delegates the removal to the Remove-GroupMember module helper' {
        $script:rBody | Should -Match '\$Outcome\s*=\s*Remove-GroupMember\b'
    }
    It 'passes the script-level GroupName and UserName straight through' {
        $script:rBody | Should -Match 'Remove-GroupMember[\s\S]*?-Group \$GroupName'
        $script:rBody | Should -Match 'Remove-GroupMember[\s\S]*?-Member \$UserName'
    }
    It 'threads the PDC host and log file through to the helper' {
        $script:rBody | Should -Match 'Remove-GroupMember[\s\S]*?-DCHostName \$DCHostName'
        $script:rBody | Should -Match 'Remove-GroupMember[\s\S]*?-LogFile \$LogFile'
    }
    It 'leaves resolution and AD mutation to the helper (no inline AD calls in the branch)' {
        $script:rBody | Should -Not -Match 'Remove-ADGroupMember'
        $script:rBody | Should -Not -Match 'Get-ADGroupMember'
        $script:rBody | Should -Not -Match 'Get-ADObject'
    }
    It 'maps a thrown removal error to a Failed outcome' {
        $script:rBody | Should -Match 'catch'
        $script:rBody | Should -Match '\$Outcome\s*=\s*"Failed"'
    }
}

Describe 'ElevateUser.ps1 forced reason' {
    It 'prompts for a reason when not supplied' {
        $script:src | Should -Match 'Read-Host "Reason for elevation/removal'
    }
    It 'aborts when the reason is blank' {
        $script:src | Should -Match '\[string\]::IsNullOrWhiteSpace\(\$Reason\)'
        $script:src | Should -Match 'throw "No reason supplied'
    }
    It 'audits blank reason rejection before throwing' {
        $script:src | Should -Match 'No reason supplied\. Aborting'
        $script:src | Should -Match 'Write-ElevationAuditRow[\s\S]*?-Reason\s+"No reason supplied"[\s\S]*?-Outcome\s+"Rejected"'
        $script:src | Should -Match 'throw "No reason supplied'
    }
}

Describe 'ElevateUser.ps1 guaranteed audit (try/finally)' {
    It 'wraps the action in a try with a finally' {
        $script:ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $true) | Where-Object { $_.Finally } | Should -Not -BeNullOrEmpty
    }
    It 'writes an audit row with all required fields in the finally' {
        $fin = ($script:ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $true) | Where-Object { $_.Finally } | Select-Object -Last 1).Finally.Extent.Text
        foreach ($f in 'Timestamp','Action','Operator','Account','Group','DurationMinutes','Reason','Outcome') { $fin | Should -Match $f }
        $fin | Should -Match 'Write-ElevationAuditRow'
    }
    It 'records every terminal outcome the script sets directly' {
        foreach ($o in 'Failed','Denied','Rejected','Success') {
            $script:src | Should -Match "`"$o`""
        }
    }
    It 'folds the remove helper''s outcome into the audited $Outcome' {
        $script:src | Should -Match '\$Outcome\s*=\s*Remove-GroupMember\b'
    }
}
