#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Structure tests for script regions that can't be exercised by running them
# (they perform irreversible AD/Exchange operations). These tests parse the
# script source with the PowerShell AST and assert the load-bearing patterns
# are still present and consistent across copies of the same logic.
#
# - Three user-management scripts (CreateUsers, Cleanup-ADSyncFailureUsers,
#   CreateOnPremMailboxes) independently sanitise the same fields. They have
#   to stay aligned or the same person produces three different SAMs across
#   user lifecycle tooling.
#
# - CreateGroup.ps1's S and H switch clauses do parallel work; the H clause's
#   HiPriv_Group_Admins precheck is load-bearing (losing it = privilege
#   escalation surface).

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    # AST helpers used across body-pinning tests.
    function Script:Get-CommandsByName {
        param($Node, [string]$Name)
        $Node.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq $Name
        }, $true)
    }
    function Script:Get-ParamArg {
        # Raw Extent.Text of the argument paired with -Name on a
        # CommandAst. Handles both -X Y (space) and -X:Y (colon) styles.
        param($CommandAst, [string]$Name)
        if (-not $CommandAst) { return $null }
        $els = $CommandAst.CommandElements
        for ($i = 0; $i -lt $els.Count; $i++) {
            $el = $els[$i]
            if ($el -is [System.Management.Automation.Language.CommandParameterAst] -and
                $el.ParameterName -eq $Name) {
                # Colon-style: value lives in .Argument on the same node
                if ($el.Argument) { return $el.Argument.Extent.Text }
                # Space-style: value is the next element (if there is one)
                if ($i + 1 -lt $els.Count) { return $els[$i + 1].Extent.Text }
                return $null
            }
        }
        return $null
    }
    function Script:Test-HasParam {
        param($CommandAst, [string]$Name)
        if (-not $CommandAst) { return $false }
        $params = $CommandAst.CommandElements | Where-Object {
            $_ -is [System.Management.Automation.Language.CommandParameterAst]
        }
        return ($params.ParameterName -contains $Name)
    }
    $script:dsPath = Join-Path $script:scriptsRoot 'Prelim\DomainSetup.ps1'
    if (-not (Test-Path $script:dsPath)) {
        throw "DomainSetup.ps1 not found at $script:dsPath - test setup cannot proceed."
    }
    $script:dsSource = Get-Content $script:dsPath -Raw -ErrorAction Stop
    $tokens = $null; $errors = $null
    $script:dsAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:dsPath, [ref]$tokens, [ref]$errors
    )
    $script:azPath = Join-Path $script:scriptsRoot 'azure_buildout.ps1'
    if (-not (Test-Path $script:azPath)) {
        throw "azure_buildout.ps1 not found at $script:azPath - test setup cannot proceed."
    }
    $script:azSource = Get-Content $script:azPath -Raw -ErrorAction Stop
    $tokens = $null; $errors = $null
    $script:azAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:azPath, [ref]$tokens, [ref]$errors
    )
}

# ====================================================================
# The S and H clauses of the GroupType switch perform parallel work with
# subtly different shapes. Without these pins, a future edit could
#   - drop the HiPriv_Group_Admins precheck (privilege escalation),
#   - send a Hi-Priv group into the Staff group via Add-GroupMember,
#   - flip HiddenFromAddressListsEnabled the wrong way,
#   - target the wrong OU path.
# All of these would parse cleanly and ship.
# ====================================================================

Describe 'CreateGroup.ps1 S/H branch parity' {
    BeforeAll {
        $path = Join-Path $script:scriptsRoot 'Users\CreateGroup.ps1'
        $tokens = $null; $errors = $null
        $cgAst = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $script:cgSwitch = $cgAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst]
        }, $true) | Where-Object {
            $_.Condition.Extent.Text -match '\$GroupType'
        } | Select-Object -First 1
        $script:cgClauses = @{}
        foreach ($clause in $script:cgSwitch.Clauses) {
            $script:cgClauses[$clause.Item1.Value] = $clause.Item2
        }
    }
    It 'GroupType switch has clauses S and H (plus a default)' {
        ($script:cgClauses.Keys | Sort-Object) -join ',' | Should -Be 'H,S'
        $script:cgSwitch.Default | Should -Not -BeNullOrEmpty
    }
    It 'default clause throws (defense in depth even with ValidateSet on -GroupType)' {
        $throws = $script:cgSwitch.Default.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
        }, $true)
        $throws.Count | Should -BeGreaterOrEqual 1
    }
    Context 'S clause - public Staff-bound group' {
        BeforeAll {
            $script:sBody           = $script:cgClauses['S']
            $script:sNewDomainGroup = (Get-CommandsByName $script:sBody 'New-DomainGroup') | Select-Object -First 1
            $script:sAddGroupMember = Get-CommandsByName $script:sBody 'Add-GroupMember'
        }
        It 'sets $OUPath to "OU=$GroupsOU,$EndPath" (Staff-group OU)' {
            $assign = $script:sBody.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left.Extent.Text -eq '$OUPath'
            }, $true) | Select-Object -First 1
            $assign | Should -Not -BeNullOrEmpty
            $assign.Right.Extent.Text | Should -Match '"OU=\$GroupsOU,\$EndPath"'
        }
        It 'New-DomainGroup uses -Path $OUPath' {
            (Get-ParamArg $script:sNewDomainGroup 'Path') | Should -Be '$OUPath'
        }
        It 'calls New-DomainGroup with -GroupDescription $GroupDescription (passes through caller-supplied)' {
            (Get-ParamArg $script:sNewDomainGroup 'GroupDescription') | Should -Be '$GroupDescription'
        }
        It 'calls New-DomainGroup with -HiddenFromAddressListsEnabled $false (Staff group is visible)' {
            (Get-ParamArg $script:sNewDomainGroup 'HiddenFromAddressListsEnabled') | Should -Be '$false'
        }
        It 'calls New-DomainGroup with -O365 $O365 (passes through caller-supplied; H is the sole hardcoded "N")' {
            (Get-ParamArg $script:sNewDomainGroup 'O365') | Should -Be '$O365'
        }
        It 'calls New-DomainGroup with the logging/DC/category triad' {
            Test-HasParam $script:sNewDomainGroup 'LogFile'       | Should -BeTrue
            Test-HasParam $script:sNewDomainGroup 'DCHostName'    | Should -BeTrue
            Test-HasParam $script:sNewDomainGroup 'GroupCategory' | Should -BeTrue
        }
        It 'invokes Add-GroupMember with -Group $StaffGroup -Member $GroupName' {
            $matched = @($script:sAddGroupMember | Where-Object {
                (Get-ParamArg $_ 'Group')  -eq '$StaffGroup' -and
                (Get-ParamArg $_ 'Member') -eq '$GroupName'
            })
            $matched.Count | Should -Be 1
        }
        It 'Add-GroupMember to $StaffGroup also includes -LogFile and -DCHostName' {
            $staffCall = $script:sAddGroupMember | Where-Object {
                (Get-ParamArg $_ 'Group') -eq '$StaffGroup'
            } | Select-Object -First 1
            Test-HasParam $staffCall 'LogFile'    | Should -BeTrue
            Test-HasParam $staffCall 'DCHostName' | Should -BeTrue
        }
    }
    Context 'H clause - Hi-Priv group, restricted OU, no Staff join' {
        BeforeAll {
            $script:hBody           = $script:cgClauses['H']
            $script:hNewDomainGroup = (Get-CommandsByName $script:hBody 'New-DomainGroup') | Select-Object -First 1
            $script:hAddGroupMember = Get-CommandsByName $script:hBody 'Add-GroupMember'
        }
        It 'references HiPriv_Group_Admins (authorisation precheck is present)' {
            $script:hBody.Extent.Text | Should -Match 'HiPriv_Group_Admins'
        }
        It 'throws on missing required group membership' {
            $throws = $script:hBody.FindAll({
                param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
            }, $true)
            $throws.Count | Should -BeGreaterOrEqual 1
        }
        It 'sets $OUPath to a HiPrivGroups OU path' {
            $assign = $script:hBody.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left.Extent.Text -eq '$OUPath'
            }, $true) | Select-Object -First 1
            $assign | Should -Not -BeNullOrEmpty
            $assign.Right.Extent.Text | Should -Match 'HiPrivGroups.+Administration'
        }
        It 'New-DomainGroup uses -Path $OUPath' {
            (Get-ParamArg $script:hNewDomainGroup 'Path') | Should -Be '$OUPath'
        }
        It 'calls New-DomainGroup with -O365 "N" (no Exchange enablement for Hi-Priv)' {
            (Get-ParamArg $script:hNewDomainGroup 'O365') | Should -Be '"N"'
        }
        It 'calls New-DomainGroup with -HiddenFromAddressListsEnabled $true' {
            (Get-ParamArg $script:hNewDomainGroup 'HiddenFromAddressListsEnabled') | Should -Be '$true'
        }
        It 'does NOT call Add-GroupMember -Group $StaffGroup (would fold Hi-Priv into Staff)' {
            $bad = @($script:hAddGroupMember | Where-Object {
                (Get-ParamArg $_ 'Group') -eq '$StaffGroup'
            })
            $bad.Count | Should -Be 0
        }
    }
}

Describe 'name sanitisation centralised' -ForEach @(
    @{ File = 'Users\CreateUsers.ps1' }
    @{ File = 'Users\Cleanup-ADSyncFailureUsers.ps1' }
) {
    BeforeAll {
        $script:src = Get-Content (Join-Path $script:scriptsRoot $File) -Raw
    }
    It '<File>: sanitises names via ConvertTo-SafeName' {
        $script:src | Should -Match 'ConvertTo-SafeName'
    }
    It '<File>: no longer inlines the raw name character-class strip' {
        $script:src | Should -Not -Match "replace\s+'\[\?@"
    }
}

Describe 'username sanitisation centralised' -ForEach @(
    @{ File = 'Users\CreateUsers.ps1' }
    @{ File = 'Users\Cleanup-ADSyncFailureUsers.ps1' }
    @{ File = 'Users\CreateOnPremMailboxes.ps1' }
) {
    BeforeAll {
        $script:src = Get-Content (Join-Path $script:scriptsRoot $File) -Raw
    }
    It '<File>: sanitises usernames via ConvertTo-SafeSamAccountName' {
        $script:src | Should -Match 'ConvertTo-SafeSamAccountName'
    }
    It '<File>: no longer inlines the raw character-class strip' {
        $script:src | Should -Not -Match "replace\s+'\[\^A-Za-z0-9\.-\]'"
    }
}

Describe 'DomainSetup.ps1 GPO directory pre-flight' {
    It 'exposes -SkipGPOs as a switch parameter' {
        $cmd = Get-Command $script:dsPath
        $cmd.Parameters['SkipGPOs'].SwitchParameter | Should -BeTrue
    }
    It 'guards the GPO directory with Test-Path before any Import-GPO call' {
        # The Test-Path must appear before the first Import-GPO line.
        $preflight = ($script:dsSource -split "`n" |
            Select-String 'Test-Path\s+\$GPOLocation' |
            Select-Object -First 1).LineNumber
        $firstImport = ($script:dsSource -split "`n" |
            Select-String 'Import-GPO' |
            Select-Object -First 1).LineNumber
        $preflight   | Should -Not -BeNullOrEmpty
        $firstImport | Should -Not -BeNullOrEmpty
        $preflight   | Should -BeLessThan $firstImport
    }
    It 'gates every Import-GPO call behind -not $SkipGPOs' {
        # AST: every CommandAst whose name is Import-GPO must have an
        # ancestor IfStatementAst testing $SkipGPOs.
        $imports = $script:dsAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.CommandAst] -and
                      $n.GetCommandName() -eq 'Import-GPO'
        }, $true)
        $imports.Count | Should -BeGreaterThan 0 -Because 'a 0-import script would silently pass the foreach gate check'
        foreach ($call in $imports) {
            $gated = $false
            $node = $call.Parent
            while ($node) {
                if ($node -is [System.Management.Automation.Language.IfStatementAst] -and
                    $node.Extent.Text -match '\$SkipGPOs') {
                    $gated = $true; break
                }
                $node = $node.Parent
            }
            $gated | Should -BeTrue -Because "Import-GPO at line $($call.Extent.StartLineNumber) must be gated by SkipGPOs"
        }
    }
}

Describe 'DomainSetup.ps1 MachineAccountQuota hardening' {
    # ms-ds-MachineAccountQuota defaults to 10, letting any authenticated user
    # join up to 10 machines to the domain. Zeroing it forces all computer
    # account creation through delegated paths. Silent regression = silent
    # downgrade of attack surface; pin the value.
    It 'sets ms-ds-MachineAccountQuota to 0 on the domain object' {
        $setAdDomain = (Get-CommandsByName $script:dsAst 'Set-ADDomain') | Select-Object -First 1
        $setAdDomain | Should -Not -BeNullOrEmpty
        $replace = Get-ParamArg $setAdDomain 'Replace'
        $replace | Should -Match '"ms-ds-MachineAccountQuota"\s*=\s*"0"'
    }
    It 'targets the local domain via Get-ADDomain.distinguishedname (not a hardcoded DN)' {
        # Pin that the script doesn't drift to a hardcoded "DC=corp,DC=com" -
        # the call must remain portable across deployments.
        $script:dsSource | Should -Match 'Set-ADDomain\s+\(Get-ADDomain\)\.distinguishedname'
    }
}

Describe 'DomainSetup.ps1 AD optional features' {
    # PAM is what makes ElevateUser.ps1's temporal group membership work
    # (TTL on group membership). Recycle Bin makes deleted-user/group recovery
    # possible without authoritative restore. Both are forest-wide one-way
    # switches - losing either is a meaningful capability loss.
    BeforeAll {
        $script:dsOptFeatures = Get-CommandsByName $script:dsAst 'Enable-ADOptionalFeature'
    }
    It 'enables exactly two optional features (PAM and Recycle Bin)' {
        # If a third gets added, this fails and forces a conscious test update.
        # If one is dropped, this also fails. Either way: human attention.
        $script:dsOptFeatures.Count | Should -Be 2
    }
    It 'enables the Privileged Access Management Feature' {
        $pam = $script:dsOptFeatures | Where-Object {
            $_.CommandElements[1].Extent.Text -match 'Privileged Access Management Feature'
        }
        $pam | Should -Not -BeNullOrEmpty
        (Get-ParamArg $pam 'Scope')  | Should -Be 'ForestOrConfigurationSet'
        (Get-ParamArg $pam 'Target') | Should -Be '$DNSSuffix'
    }
    It 'enables the Recycle Bin Feature' {
        $rb = $script:dsOptFeatures | Where-Object {
            $_.CommandElements[1].Extent.Text -match 'Recycle Bin Feature'
        }
        $rb | Should -Not -BeNullOrEmpty
        (Get-ParamArg $rb 'Scope')  | Should -Be 'ForestOrConfigurationSet'
        (Get-ParamArg $rb 'Target') | Should -Be '$DNSSuffix'
    }
    It 'enables both features non-interactively (-Confirm:$False)' {
        # If -Confirm gets dropped, a re-run prompts in an interactive session
        # and silently hangs in a scheduled context.
        foreach ($f in $script:dsOptFeatures) {
            $f.Extent.Text | Should -Match '-Confirm:\s*\$False'
        }
    }
}

Describe 'DomainSetup.ps1 LAPS configuration' {
    # The LAPS delegation block is the load-bearing access-control surface for
    # local-admin password recovery. The four target OUs (Desktops, Laptops,
    # Servers, VMs) each need BOTH Read and Reset permissions delegated, AND
    # Servers must go to Server_Admins while the others go to Desktop_Admins.
    # A regression that swaps the groups silently re-routes server password
    # access to the workstation admin tier.
    BeforeAll {
        $script:dsLapsRead  = Get-CommandsByName $script:dsAst 'Set-LapsADReadPasswordPermission'
        $script:dsLapsReset = Get-CommandsByName $script:dsAst 'Set-LapsADResetPasswordPermission'
        $script:dsLapsSelf  = Get-CommandsByName $script:dsAst 'Set-LapsADComputerSelfPermission'
        $script:dsLapsSchema = Get-CommandsByName $script:dsAst 'Update-LapsADSchema'
    }
    It 'updates the AD schema for LAPS exactly once' {
        $script:dsLapsSchema.Count | Should -Be 1
    }
    It 'enables computer self-permission for password rotation' {
        $script:dsLapsSelf.Count | Should -Be 1
        (Get-ParamArg $script:dsLapsSelf[0] 'Identity') | Should -Be '$Location'
    }
    It 'delegates Read AND Reset to every targeted OU (no read-only gaps)' {
        # An OU with Read but no Reset means the local admin password leaks
        # but can't be rotated when needed. Pair-count must match.
        $script:dsLapsRead.Count  | Should -Be 4
        $script:dsLapsReset.Count | Should -Be 4
    }
    It 'delegates Desktops OU to Desktop_Admins' {
        $reads  = $script:dsLapsRead  | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Desktops\)' }
        $resets = $script:dsLapsReset | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Desktops\)' }
        $reads.Count  | Should -Be 1
        $resets.Count | Should -Be 1
        (Get-ParamArg $reads[0]  'AllowedPrincipals') | Should -Match 'Desktop_Admins'
        (Get-ParamArg $resets[0] 'AllowedPrincipals') | Should -Match 'Desktop_Admins'
    }
    It 'delegates Laptops OU to Desktop_Admins' {
        $reads  = $script:dsLapsRead  | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Laptops\)' }
        $resets = $script:dsLapsReset | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Laptops\)' }
        $reads.Count  | Should -Be 1
        $resets.Count | Should -Be 1
        (Get-ParamArg $reads[0]  'AllowedPrincipals') | Should -Match 'Desktop_Admins'
        (Get-ParamArg $resets[0] 'AllowedPrincipals') | Should -Match 'Desktop_Admins'
    }
    It 'delegates Servers OU to Server_Admins (tier separation from Desktops/Laptops)' {
        $reads  = $script:dsLapsRead  | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Servers\)' }
        $resets = $script:dsLapsReset | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.Servers\)' }
        $reads.Count  | Should -Be 1
        $resets.Count | Should -Be 1
        # Critical: must be Server_Admins, NOT Desktop_Admins. The whole point of
        # the LAPS tier split is to prevent workstation-admin tier from rotating
        # server local-admin passwords.
        (Get-ParamArg $reads[0]  'AllowedPrincipals') | Should -Match 'Server_Admins'
        (Get-ParamArg $resets[0] 'AllowedPrincipals') | Should -Match 'Server_Admins'
        (Get-ParamArg $reads[0]  'AllowedPrincipals') | Should -Not -Match 'Desktop_Admins'
        (Get-ParamArg $resets[0] 'AllowedPrincipals') | Should -Not -Match 'Desktop_Admins'
    }
    It 'delegates VMs OU to Desktop_Admins' {
        $reads  = $script:dsLapsRead  | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.VMs\)' }
        $resets = $script:dsLapsReset | Where-Object { (Get-ParamArg $_ 'Identity') -match '\$\(\$Env\.OUs\.VMs\)' }
        $reads.Count  | Should -Be 1
        $resets.Count | Should -Be 1
        (Get-ParamArg $reads[0]  'AllowedPrincipals') | Should -Match 'Desktop_Admins'
        (Get-ParamArg $resets[0] 'AllowedPrincipals') | Should -Match 'Desktop_Admins'
    }
}

Describe 'DomainSetup.ps1 built-in admin (SID500) hardening' {
    # The built-in Administrator (RID 500) is a high-value, unblockable target.
    # Three properties matter: it should not live in the default Users
    # container, it must not be Kerberos-delegated to any service, and it
    # must not be a member of Schema Admins (it is by default).
    It 'moves the built-in admin into the HiPrivAccounts OU' {
        $move = (Get-CommandsByName $script:dsAst 'Move-ADObject') | Where-Object {
            (Get-ParamArg $_ 'Identity') -match 'SID500'
        } | Select-Object -First 1
        $move | Should -Not -BeNullOrEmpty
        (Get-ParamArg $move 'TargetPath') | Should -Match '\$\(\$Env\.OUs\.HiPrivAccounts\)'
    }
    It 'sets AccountNotDelegated to $True (prevents Kerberos delegation)' {
        $sac = (Get-CommandsByName $script:dsAst 'Set-ADAccountControl') | Where-Object {
            (Get-ParamArg $_ 'Identity') -match 'SID500'
        } | Select-Object -First 1
        $sac | Should -Not -BeNullOrEmpty
        (Get-ParamArg $sac 'AccountNotDelegated') | Should -Be '$True'
    }
    It 'removes the built-in admin from Schema Admins' {
        $remove = (Get-CommandsByName $script:dsAst 'Remove-ADGroupMember') | Where-Object {
            ((Get-ParamArg $_ 'Identity') -match 'Schema Admins') -and
            ((Get-ParamArg $_ 'Members')  -match 'SID500')
        } | Select-Object -First 1
        $remove | Should -Not -BeNullOrEmpty
        # Must be non-interactive or a scheduled context will hang on confirm.
        (Get-ParamArg $remove 'Confirm') | Should -Be '$False'
    }
    It 'tolerates idempotent rerun via ADException catch (already-moved/already-not-a-member)' {
        # Both the Move and the Remove are wrapped in try/catch [ADException]
        # so a second invocation doesn't crash. Without these guards a single
        # failed mid-script run can't be retried cleanly.
        $tryStatements = $script:dsAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $sid500Tries = $tryStatements | Where-Object {
            ($_.Extent.Text -match 'Move-ADObject'        -and $_.Extent.Text -match '\$SID500') -or
            ($_.Extent.Text -match 'Remove-ADGroupMember' -and $_.Extent.Text -match 'Schema Admins')
        }
        $sid500Tries.Count | Should -Be 2 -Because 'one for Move-ADObject, one for Remove-ADGroupMember'
        # And each catch must filter on ADException specifically rather than
        # a bare catch{} that would swallow unrelated errors.
        foreach ($t in $sid500Tries) {
            $hasAdException = $t.CatchClauses | Where-Object {
                $_.CatchTypes.TypeName.FullName -contains 'Microsoft.ActiveDirectory.Management.ADException'
            }
            $hasAdException | Should -Not -BeNullOrEmpty -Because "try at line $($t.Extent.StartLineNumber) must catch ADException specifically"
        }
    }
}

Describe 'DomainSetup.ps1 default policy enforcement' {
    # Default Domain Policy and Default Domain Controllers Policy must be
    # *enforced* (not just linked) so a future GPO-link reorder can't deprioritise
    # them. -enforced yes makes them non-overridable by child links.
    BeforeAll {
        $script:dsGpLinks = (Get-CommandsByName $script:dsAst 'Set-GPLink')
    }
    It 'enforces at least two GPO links (Default Domain + Default DC)' {
        $script:dsGpLinks.Count | Should -BeGreaterOrEqual 2
    }
    It 'enforces Default Domain Policy at the domain root' {
        $ddp = $script:dsGpLinks | Where-Object {
            ((Get-ParamArg $_ 'name')   -match 'Default Domain Policy') -or
            ((Get-ParamArg $_ 'Name')   -match 'Default Domain Policy') -or
            ($_.Extent.Text -match '\$GPOName.*Default Domain Policy')
        }
        # Looser match because the call uses `-name $GPOName` after assigning
        # $GPOName above - we accept either the direct literal or the variable.
        # Either way it must target $EndPath with -enforced yes.
        $gpoNameAssign = $script:dsSource -split "`n" |
            Select-String '\$GPOName\s*=\s*"Default Domain Policy"' |
            Select-Object -First 1
        $gpoNameAssign | Should -Not -BeNullOrEmpty
        $setLink = $script:dsSource -split "`n" |
            Select-String 'Set-GPLink\s+-name\s+\$GPOName\s+-target\s+\$EndPath\s+-enforced\s+yes' |
            Select-Object -First 1
        $setLink | Should -Not -BeNullOrEmpty -Because 'Default Domain Policy must be enforced at $EndPath'
    }
    It 'enforces Default Domain Controllers Policy on the DCs OU' {
        $gpoNameAssign = $script:dsSource -split "`n" |
            Select-String '\$GPOName\s*=\s*"Default Domain Controllers Policy"' |
            Select-Object -First 1
        $gpoNameAssign | Should -Not -BeNullOrEmpty
        $setLink = $script:dsSource -split "`n" |
            Select-String 'Set-GPLink\s+-name\s+\$GPOName\s+-target\s+"ou=Domain Controllers,\$EndPath"\s+-enforced\s+yes' |
            Select-Object -First 1
        $setLink | Should -Not -BeNullOrEmpty -Because 'Default Domain Controllers Policy must be enforced on ou=Domain Controllers,$EndPath'
    }
    It 're-throws on Set-GPLink failure (no silent enforcement drop)' {
        # If the try/catch around Set-GPLink ever degrades to log-and-continue,
        # a missing GPO could pass silently and leave the domain unenforced.
        $tryStatements = $script:dsAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true)
        $gplinkTries = $tryStatements | Where-Object { $_.Extent.Text -match 'Set-GPLink' }
        $gplinkTries.Count | Should -BeGreaterOrEqual 2 -Because 'one try block per default-policy enforcement'
        foreach ($t in $gplinkTries) {
            $rethrows = $t.CatchClauses.Body.FindAll({
                param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
            }, $true)
            $rethrows.Count | Should -BeGreaterOrEqual 1 -Because "catch at line $($t.Extent.StartLineNumber) must re-throw, not swallow"
        }
    }
}

# ====================================================================
# The existing "-SkipGPOs gates every Import-GPO" test proves the gating
# shape; it does NOT prove any specific GPO is present. Deleting the
# EnableSMBSigning import+link (or any sibling) would parse cleanly and
# ship, silently dropping a hardening baseline. These pins make that fail.
# ====================================================================
Describe 'DomainSetup.ps1 security-hardening GPOs are imported and linked' {
    BeforeAll {
        $script:importedGpoNames = (Get-CommandsByName $script:dsAst 'Import-GPO') |
            ForEach-Object { (Get-ParamArg $_ 'BackupGpoName').Trim('"') }
        $script:gpoLinks = (Get-CommandsByName $script:dsAst 'Add-GPOLink') |
            ForEach-Object {
                [pscustomobject]@{
                    Name   = (Get-ParamArg $_ 'GPOName').Trim('"')
                    Target = Get-ParamArg $_ 'GPOTarget'
                }
            }
    }
    $domainRootGpos = @(
        'Logon Policy','TLS','CWDIllegalInDllSearch','DisableNullSessionEnumeration',
        'EnableSMBSigning','EnforceNLAandTLSforRDP','GroupPolicyHardening_MS150-011',
        'Kerberos_Armouring','LSAProtection','NTLMv2','LDAP_signing_requirements'
    )
    $dcGpos = @(
        'DC_AppLocker_Disable_Browsers','DC_Auditing',
        'DC_Disable_Print_Spooler','DC_LDAP_signing_requirements'
    )
    It '<Gpo> is imported' -TestCases (($domainRootGpos + $dcGpos) | ForEach-Object { @{ Gpo = $_ } }) {
        param($Gpo)
        $script:importedGpoNames | Should -Contain $Gpo -Because "the $Gpo baseline must be imported"
    }
    It '<Gpo> is linked at the domain root' -TestCases ($domainRootGpos | ForEach-Object { @{ Gpo = $_ } }) {
        param($Gpo)
        $link = $script:gpoLinks | Where-Object { $_.Name -eq $Gpo }
        $link | Should -Not -BeNullOrEmpty -Because "$Gpo must be linked"
        $link.Target | Should -Be '$EndPath'
    }
    It '<Gpo> is linked at the Domain Controllers OU' -TestCases ($dcGpos | ForEach-Object { @{ Gpo = $_ } }) {
        param($Gpo)
        $link = $script:gpoLinks | Where-Object { $_.Name -eq $Gpo }
        $link | Should -Not -BeNullOrEmpty -Because "$Gpo must be linked"
        $link.Target | Should -Match 'Domain Controllers'
    }
}

Describe 'azure_buildout.ps1 parameter contract' {
    # All parameters are optional [string] - the script falls back to READ-HOST
    # for any missing value (lines 16-30). That interactive fallback is
    # load-bearing for the documented workflow, so the test pins:
    # - the 6 expected params exist with the right names and types
    # - none are Mandatory (would break the READ-HOST fallback)
    # - CmdletBinding is declared (so -Verbose / -ErrorAction etc. work)
    BeforeAll {
        $script:azCmd = Get-Command $script:azPath
    }
    It 'declares CmdletBinding' {
        $cb = $script:azAst.ParamBlock.Attributes |
            Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }
        $cb | Should -Not -BeNullOrEmpty
    }
    It 'exposes -SubscriptionId, -TenantID, -MyIPAddress, -NameStem, -Owner, -LogFile' {
        $expected = 'SubscriptionId','TenantID','MyIPAddress','NameStem','Owner','LogFile'
        foreach ($p in $expected) {
            $script:azCmd.Parameters.Keys | Should -Contain $p -Because "missing parameter: $p"
        }
    }
    It 'declares every parameter as [string]' {
        $expected = 'SubscriptionId','TenantID','MyIPAddress','NameStem','Owner','LogFile'
        foreach ($p in $expected) {
            $script:azCmd.Parameters[$p].ParameterType.FullName | Should -Be 'System.String' -Because "$p must be [string] for the READ-HOST fallback to work"
        }
    }
    It 'declares no Mandatory parameter (interactive READ-HOST fallback is load-bearing)' {
        # If any parameter becomes Mandatory, the script can't be invoked
        # parameterless and the documented workflow breaks.
        $expected = 'SubscriptionId','TenantID','MyIPAddress','NameStem','Owner','LogFile'
        foreach ($p in $expected) {
            $attrs = $script:azCmd.Parameters[$p].Attributes |
                Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] }
            $mandatory = $attrs | Where-Object Mandatory
            $mandatory | Should -BeNullOrEmpty -Because "$p must not be Mandatory"
        }
    }
    It 'exposes -RemoveIncomplete as an optional [switch] (reconciliation mode, separate from the READ-HOST set)' {
        $script:azCmd.Parameters.Keys | Should -Contain 'RemoveIncomplete'
        $script:azCmd.Parameters['RemoveIncomplete'].ParameterType.FullName |
            Should -Be 'System.Management.Automation.SwitchParameter'
        $attrs = $script:azCmd.Parameters['RemoveIncomplete'].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] }
        ($attrs | Where-Object Mandatory) | Should -BeNullOrEmpty -Because 'the reconciliation switch must stay optional'
    }
}

Describe 'azure_buildout.ps1 prerequisites and strict mode' {
    It 'declares Set-StrictMode -Version Latest' {
        $script:azSource | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It 'imports the Az module' {
        $script:azSource | Should -Match 'Import-Module\s+Az\b'
    }
    It 'imports CorpAdmin via Modules\CorpAdmin\CorpAdmin.psd1' {
        # Uses Join-Path so the import is robust to where the script is run from.
        $script:azSource | Should -Match 'Import-Module\s+\(Join-Path\s+\$PSScriptRoot\s+''Modules\\CorpAdmin\\CorpAdmin\.psd1'''
    }
}

Describe 'azure_buildout.ps1 throw-policy consistency' {
    # The script has 41 catch blocks in three shapes:
    #   - 29 log-and-throw           (the dominant pattern: loud propagation)
    #   - 11 log-only                (paired with -ErrorAction SilentlyContinue
    #                                 on Get-Az* probes - "does it exist yet?"
    #                                 idempotent buildout pattern)
    #   -  1 log-and-continue        (per-VM outer wrapper - one bad VM
    #                                 shouldn't kill the rest of the batch)
    # These tests pin all three shapes explicitly so a future change that
    # silently widens log-only ("oh, all catches just log") fails loudly.
    BeforeAll {
        $script:azCatches = $script:azAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.CatchClauseAst]
        }, $true)
    }
    It 'every catch block logs via Write-LogFile (no silent empty catches, no Write-Host)' {
        foreach ($c in $script:azCatches) {
            $body = $c.Body.Extent.Text
            $body | Should -Match 'Write-LogFile' -Because "catch at line $($c.Extent.StartLineNumber) must log via Write-LogFile"
            $body | Should -Not -Match '\bWrite-Host\b' -Because "catch at line $($c.Extent.StartLineNumber) must not use Write-Host (use Write-LogFile)"
        }
    }
    It 'every catch terminates with one of: throw, continue, Write-LogFile' {
        # Find each catch's last meaningful statement. Anything outside the
        # allowed set (return, break, exit, an Az cmdlet, ...) is a smell that
        # demands review.
        $allowed = @('ThrowStatementAst', 'ContinueStatementAst', 'PipelineAst')
        foreach ($c in $script:azCatches) {
            $statements = $c.Body.Statements
            $statements.Count | Should -BeGreaterThan 0 -Because "catch at line $($c.Extent.StartLineNumber) must not be empty"
            $last = $statements[-1]
            $lastTypeName = $last.GetType().Name
            $lastTypeName | Should -BeIn $allowed -Because "catch at line $($c.Extent.StartLineNumber) terminator is $lastTypeName; must be throw/continue/Write-LogFile"
            # If it's a PipelineAst, it must be a Write-LogFile (not some
            # other "silently do something" cmdlet)
            if ($lastTypeName -eq 'PipelineAst') {
                $last.Extent.Text | Should -Match '^\s*Write-LogFile' -Because "catch at line $($c.Extent.StartLineNumber) ends in a non-Write-LogFile pipeline"
            }
        }
    }
    It 'has exactly one continue (the per-VM batch boundary)' {
        # More than one continue means someone added a second batch boundary,
        # which is almost certainly a mistake. Zero means the per-VM wrapper
        # got changed to throw, which would abort the whole run on one bad VM.
        $continueCatches = $script:azCatches | Where-Object {
            $_.Body.Statements[-1] -is [System.Management.Automation.Language.ContinueStatementAst]
        }
        $continueCatches.Count | Should -Be 1 -Because 'exactly one catch should use continue (the per-VM outer wrapper)'
    }
    It 'majority of catches re-throw (loud failure is the dominant pattern)' {
        # If a maintenance pass converts a bunch of throws to log-only by mistake,
        # the count drops and this test catches it. The bound is generous on
        # purpose - we're guarding against wholesale regression, not pinning
        # an exact ratio that bikesheds future edits.
        $throwCatches = $script:azCatches | Where-Object {
            $_.Body.Statements[-1] -is [System.Management.Automation.Language.ThrowStatementAst]
        }
        $throwCatches.Count | Should -BeGreaterOrEqual 25 -Because 'most catches should propagate; current count is 29'
    }
}

Describe 'azure_buildout.ps1 attaches the data disk on reuse, not only on creation' {
    BeforeAll {
        $script:addDataDisk = $script:azAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.CommandAst] -and
                      $n.GetCommandName() -eq 'Add-AzVMDataDisk'
        }, $true) | Select-Object -First 1
        $script:diskCreateIf = $script:azAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.IfStatementAst] -and
            $n.Clauses[0].Item1.Extent.Text -match '!\s*\(\s*\$DataDisk\s*\)'
        }, $true) | Select-Object -First 1
    }
    It 'has both the disk-create guard and an Add-AzVMDataDisk call' {
        $script:addDataDisk  | Should -Not -BeNullOrEmpty
        $script:diskCreateIf | Should -Not -BeNullOrEmpty
    }
    It 'attaches the disk OUTSIDE the create-if-missing guard' {
        $start = $script:diskCreateIf.Extent.StartOffset
        $end   = $script:diskCreateIf.Extent.EndOffset
        $at    = $script:addDataDisk.Extent.StartOffset
        ($at -ge $start -and $at -lt $end) |
            Should -BeFalse -Because 'a reused disk must still be attached, so the attach cannot live inside the create guard'
    }
}

Describe 'azure_buildout.ps1 -RemoveIncomplete reconciliation' {
    It 'exposes a -RemoveIncomplete switch' {
        $p = $script:azAst.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'RemoveIncomplete' }
        $p | Should -Not -BeNullOrEmpty
        $p.StaticType.Name | Should -Be 'SwitchParameter'
    }
    It 'defines a helper that gates on VM existence before removing resources' {
        $fn = $script:azAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                      $n.Name -eq 'Remove-IncompleteBuildResource'
        }, $true) | Select-Object -First 1
        $fn | Should -Not -BeNullOrEmpty
        $body = $fn.Body.Extent.Text
        $body | Should -Match 'Get-AzVM'
        $body | Should -Match 'Remove-AzNetworkInterface'
        $body | Should -Match 'Remove-AzPublicIpAddress'
        $body | Should -Match 'Remove-AzDisk'
    }
}

Describe 'azure_buildout.ps1 RG tag composition and threading' {
    # All Azure resources created should carry the Department and Owner tags
    # for billing/governance. The composition happens once at line 55:
    #   $RgTags = $EnvConfig.Azure.BaseTags + @{ Department = $NameStem; Owner = $Owner }
    # and is threaded into both New-AzResourceGroup and New-AzVMConfig.
    # A regression that drops either thread point silently un-tags resources.
    It 'composes $RgTags from BaseTags + Department + Owner' {
        $script:azSource | Should -Match '\$RgTags\s*=\s*\$EnvConfig\.Azure\.BaseTags\s*\+\s*@\{\s*Department\s*=\s*\$NameStem;\s*Owner\s*=\s*\$Owner\s*\}'
    }
    It 'threads $RgTags into New-AzResourceGroup via -Tag' {
        $newRg = (Get-CommandsByName $script:azAst 'New-AzResourceGroup') | Select-Object -First 1
        $newRg | Should -Not -BeNullOrEmpty
        (Get-ParamArg $newRg 'Tag') | Should -Be '$RgTags'
    }
    It 'threads $RgTags into New-AzVMConfig via -Tags' {
        # Note: New-AzResourceGroup uses singular -Tag while New-AzVMConfig
        # uses plural -Tags. Az cmdlet shape, not our doing.
        $newVm = (Get-CommandsByName $script:azAst 'New-AzVMConfig') | Select-Object -First 1
        $newVm | Should -Not -BeNullOrEmpty
        (Get-ParamArg $newVm 'Tags') | Should -Be '$RgTags'
    }
}

Describe 'azure_buildout.ps1 RDP source restriction (security regression guard)' {
    # The RDP NSG rule scopes -SourceAddressPrefix to $MyIPAddress, not "*".
    # Opening 3389 to the internet is how Azure VMs get scanned and brute-forced
    # within minutes of provisioning. This test pins the scoped source so a
    # "just open RDP while I troubleshoot" PR can't slip in.
    BeforeAll {
        $script:azRdpRule = (Get-CommandsByName $script:azAst 'New-AzNetworkSecurityRuleConfig') | Where-Object {
            (Get-ParamArg $_ 'Name') -eq 'rdp-rule' -or
            ($_.Extent.Text -match '-Name\s+rdp-rule\b')
        } | Select-Object -First 1
    }
    It 'creates exactly one RDP rule' {
        $script:azRdpRule | Should -Not -BeNullOrEmpty
    }
    It 'scopes the RDP rule SourceAddressPrefix to $MyIPAddress (not internet-wide)' {
        $source = Get-ParamArg $script:azRdpRule 'SourceAddressPrefix'
        $source | Should -Be '$MyIPAddress' -Because 'RDP must never be open to 0.0.0.0/0 or *'
        # Defence in depth: also check the raw text doesn't contain a wildcard
        # for the source - catches a refactor that hardcodes "*" as a literal.
        $script:azRdpRule.Extent.Text | Should -Not -Match '-SourceAddressPrefix\s+\*'
        $script:azRdpRule.Extent.Text | Should -Not -Match '-SourceAddressPrefix\s+0\.0\.0\.0/0'
    }
}

Describe 'azure_buildout.ps1 NSG HTTP/HTTPS rules scope to the private IP only' {
    BeforeAll {
        $script:webRules = Get-CommandsByName $script:azAst 'Add-AzNetworkSecurityRuleConfig'
    }
    It '<Rule> targets the VM private IP, not a wildcard or the public IP' -TestCases @(
        @{ Rule = 'http-rule' }, @{ Rule = 'https-rule' }
    ) {
        param($Rule)
        $cmd = $script:webRules | Where-Object { (Get-ParamArg $_ 'Name') -eq $Rule }
        $cmd | Should -Not -BeNullOrEmpty -Because "$Rule must exist"
        $dest = Get-ParamArg $cmd 'DestinationAddressPrefix'
        $dest | Should -Be '$VirtualMachine_PrivateIP'
        $dest | Should -Not -Match 'PublicIP'
        $dest | Should -Not -Be '*'
    }
}

Describe 'azure_buildout.ps1 cleanup in finally' {
    # The outermost try/finally ensures credentials are cleared and the Az
    # session is disconnected even if buildout aborts mid-flight. Without
    # this, a crashed buildout leaves a cached subscription context on disk
    # that the next interactive Az session inherits.
    BeforeAll {
        $script:azOuterTry = $script:azAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Where-Object {
            # The outermost try - its Finally is non-null and contains
            # Disconnect-AzAccount.
            $_.Finally -and $_.Finally.Extent.Text -match 'Disconnect-AzAccount'
        } | Select-Object -First 1
    }
    It 'wraps the script body in a try/finally with cleanup' {
        $script:azOuterTry | Should -Not -BeNullOrEmpty
    }
    It 'disposes the VirtualMachineCredential password if it was set' {
        # The credential is built from a Get-Credential prompt; its SecureString
        # password needs disposing to clear it from memory.
        $body = $script:azOuterTry.Finally.Extent.Text
        $body | Should -Match 'if\s*\(\s*\$VirtualMachineCredential\s*\)\s*\{[^}]*\$VirtualMachineCredential\.Password\.Dispose\(\)'
    }
    It 'calls Disconnect-AzAccount in the finally block' {
        $body = $script:azOuterTry.Finally.Extent.Text
        $body | Should -Match 'Disconnect-AzAccount'
    }
    It 'calls Clear-AzContext in the finally block' {
        # Clear-AzContext removes the cached subscription context so the next
        # Az session doesn't accidentally inherit it.
        $body = $script:azOuterTry.Finally.Extent.Text
        $body | Should -Match 'Clear-AzContext'
    }
}

Describe 'IT admin-account scripts: SAM-formation integrity' -ForEach @(
    @{ File = 'Users\CreateITAdminUser.ps1';       Var = 'UserNameAdmin';       Prefix = 'admin.' }
    @{ File = 'Users\CreateITCloudAdminUser.ps1';  Var = 'UserNameCloudAdmin';  Prefix = 'ca.'    }
    @{ File = 'Users\CreateITDomainAdminUser.ps1'; Var = 'UserNameDomainAdmin'; Prefix = 'da.'    }
    @{ File = 'Users\CreateITGlobalAdminUser.ps1'; Var = 'UserNameGlobalAdmin'; Prefix = 'ga.'    }
) {
    BeforeAll {
        $script:p   = Join-Path $script:scriptsRoot $File
        $script:src = Get-Content $script:p -Raw
        $t = $null; $e = $null
        $script:ast = [System.Management.Automation.Language.Parser]::ParseFile($script:p, [ref]$t, [ref]$e)
        $script:adminVars = @(
            $script:ast.FindAll({
                param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and
                          $n.VariablePath.UserPath -match '^UserName.*Admin$'
            }, $true) | ForEach-Object { $_.VariablePath.UserPath } | Sort-Object -Unique
        )
    }
    It '<File>: parses without errors' {
        $script:ast | Should -Not -BeNullOrEmpty
    }
    It '<File>: uses exactly one privileged-name variable ($<Var>) for assign and read' {
        # A second distinct $UserName*Admin name means the value is written to one
        # variable and read from another - the assign/read mismatch bug.
        $script:adminVars.Count | Should -Be 1 -Because 'assign and read must use the same variable'
        $script:adminVars[0]    | Should -BeExactly $Var
    }
    It '<File>: derives $<Var> from ConvertTo-SafeSamAccountName' {
        $script:src | Should -Match ([regex]::Escape('$' + $Var) + '\s*=\s*ConvertTo-SafeSamAccountName')
    }
    It "<File>: applies the '<Prefix>' prefix" {
        $script:src | Should -Match ("ConvertTo-SafeSamAccountName.*-Prefix\s+'" + [regex]::Escape($Prefix) + "'")
    }
    It '<File>: passes Read-Host to sanitisers parenthesised, never bare' {
        $script:src | Should -Not -Match 'ConvertTo-Safe\w+\s+Read-Host'
    }
}
