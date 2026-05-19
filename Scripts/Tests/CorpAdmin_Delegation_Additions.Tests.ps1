#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# =============================================================================
# CorpAdmin.Delegation.Tests.ps1
# =============================================================================
# AST-level tests pinning the contract of each Grant-*Delegation helper in
# CorpAdmin.psm1. These tests don't hit a real domain - they parse the module
# source and verify that each helper:
#   1. Has the expected parameter contract (5 mandatory params, names match)
#   2. Resolves the admin group's SID via Get-ADGroup + SecurityIdentifier
#   3. Reads the ACL via Get-Acl with the expected path shape
#   4. Calls AddAccessRule the expected number of times with expected shapes
#   5. Writes the ACL back via $Acl | Set-Acl
#
# This is the same pattern used for the LAPS/MAQ/SID500 Describes in
# ScriptBody.Tests.ps1: pin the source-pattern contract rather than
# mock-execute the function against a real domain.
#
# The test file lives next to ScriptBody.Tests.ps1 in Scripts/Tests/.
# =============================================================================

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $script:modPath = Join-Path $script:scriptsRoot 'Modules\CorpAdmin\CorpAdmin.psm1'
    if (-not (Test-Path $script:modPath)) {
        throw "CorpAdmin.psm1 not found at $script:modPath - test setup cannot proceed."
    }
    $script:modSource = Get-Content $script:modPath -Raw -ErrorAction Stop
    $tokens = $null; $errors = $null
    $script:modAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:modPath, [ref]$tokens, [ref]$errors
    )
    # Helper: find a named function in the module AST.
    function Script:Get-ModuleFunction {
        param([string]$Name)
        $script:modAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                      $n.Name -eq $Name
        }, $true) | Select-Object -First 1
    }
    # Helper: count AddAccessRule (or RemoveAccessRule) calls in a function body.
    function Script:Get-AclRuleCall {
        param(
            $FunctionAst,
            [string]$MethodName = 'AddAccessRule'
        )
        $FunctionAst.Body.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $n.Member.Value -eq $MethodName
        }, $true)
    }
}

# Common assertions every Grant-*Delegation helper must satisfy. These run as
# a single It block per helper rather than separate Describes per assertion -
# tighter output and the assertions are conceptually one contract.
Describe 'CorpAdmin.psm1 delegation helpers - common contract' {
    $expectedParams = @{
        'Grant-ComputerJoinDelegation'         = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-GroupDelegation'                = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-GroupMembershipEditDelegation'  = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-PasswordResetDelegation'        = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-UserDelegation'                 = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-OUDelegation'                   = @('AdminGroupName','TargetOU','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-DNSOperatorsPermissionDelegation'  = @('AdminGroupName','TargetDN','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-DNSReadOnlyPermissionDelegation'   = @('AdminGroupName','TargetDN','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-ADObjectPermissionDelegation'   = @('AdminGroupName','TargetDN','BaseDN','GuidMap','ExtendedRightsMap')
        'Grant-GPOPermissionDelegation'        = @('AdminGroupName','BaseDN','GuidMap','ExtendedRightsMap')  # no TargetDN - operates on the domain root
        'Grant-GPOCreationDelegation'          = @('AdminGroupName','TargetDN','BaseDN','GuidMap','ExtendedRightsMap')
    }
    It '<Name> declares the expected mandatory parameters' -TestCases (
        $expectedParams.GetEnumerator() | ForEach-Object {
            @{ Name = $_.Key; Expected = $_.Value }
        }
    ) {
        param($Name, $Expected)
        $fn = Get-ModuleFunction $Name
        $fn | Should -Not -BeNullOrEmpty -Because "function $Name must exist in CorpAdmin.psm1"
        $actual = $fn.Body.ParamBlock.Parameters.Name.VariablePath.UserPath
        foreach ($p in $Expected) {
            $actual | Should -Contain $p -Because "$Name must declare -$p"
        }
        # Every parameter must be Mandatory.
        foreach ($pAst in $fn.Body.ParamBlock.Parameters) {
            $mandatoryAttr = $pAst.Attributes | Where-Object {
                $_.TypeName.Name -eq 'Parameter' -and
                $_.NamedArguments.ArgumentName -contains 'Mandatory'
            }
            $mandatoryAttr | Should -Not -BeNullOrEmpty -Because "$Name -$($pAst.Name.VariablePath.UserPath) must be Mandatory"
        }
    }
    It '<Name> resolves the admin group SID via Get-ADGroup + SecurityIdentifier' -TestCases (
        $expectedParams.Keys | ForEach-Object { @{ Name = $_ } }
    ) {
        param($Name)
        $fn = Get-ModuleFunction $Name
        $body = $fn.Body.Extent.Text
        $body | Should -Match 'New-Object System\.Security\.Principal\.SecurityIdentifier \(Get-ADGroup \$AdminGroupName\)\.SID'
    }
    It '<Name> reads the ACL via Get-Acl then writes back via Set-Acl' -TestCases (
        $expectedParams.Keys | ForEach-Object { @{ Name = $_ } }
    ) {
        param($Name)
        $fn = Get-ModuleFunction $Name
        $body = $fn.Body.Extent.Text
        $body | Should -Match 'Get-Acl\s+"AD:\\'   -Because 'each helper must read the target ACL via Get-Acl "AD:\..."'
        $body | Should -Match '\$Acl\s*\|\s*Set-Acl' -Because 'each helper must persist via $Acl | Set-Acl'
    }
}

Describe 'CorpAdmin.psm1 Grant-ComputerJoinDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-ComputerJoinDelegation' }
    It 'adds exactly 8 AccessRules (CreateChild/DeleteChild/WriteProperty/WriteDacl + 2 validated writes + Reset/Change Password)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 8
    }
    It 'grants CreateChild and DeleteChild on computer objects' {
        $script:fn.Body.Extent.Text | Should -Match "'CreateChild',\s*\`$AccessControlTypeAllow,\s*\`$GuidMap\['computer'\]"
        $script:fn.Body.Extent.Text | Should -Match "'DeleteChild',\s*\`$AccessControlTypeAllow,\s*\`$GuidMap\['computer'\]"
    }
    It 'grants validated writes for DNS host name and SPN (so domain joins can update those)' {
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Validated write to DNS host name'\]"
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Validated write to service principal name'\]"
    }
    It 'grants Reset Password and Change Password extended rights' {
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Reset Password'\]"
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Change Password'\]"
    }
}

Describe 'CorpAdmin.psm1 Grant-GroupDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-GroupDelegation' }
    It 'adds exactly 3 AccessRules (CreateChild, DeleteChild, WriteProperty on group objects)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 3
    }
    It 'scopes all rules to group objects via $GuidMap[''group'']' {
        ([regex]::Matches($script:fn.Body.Extent.Text, "\`$GuidMap\['group'\]")).Count | Should -BeGreaterOrEqual 3
    }
}

Describe 'CorpAdmin.psm1 Grant-GroupMembershipEditDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-GroupMembershipEditDelegation' }
    It 'adds exactly 1 AccessRule (WriteProperty on the member attribute)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 1
    }
    It 'targets the member attribute specifically (not arbitrary group properties)' {
        $script:fn.Body.Extent.Text | Should -Match "'WriteProperty',\s*\`$AccessControlTypeAllow,\s*\`$GuidMap\['member'\],'Descendents',\s*\`$GuidMap\['group'\]"
    }
}

Describe 'CorpAdmin.psm1 Grant-PasswordResetDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-PasswordResetDelegation' }
    It 'adds exactly 3 AccessRules (pwdLastSet, lockoutTime, Reset Password)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 3
    }
    It 'grants WriteProperty on pwdLastSet and lockoutTime' {
        $script:fn.Body.Extent.Text | Should -Match "\`$GuidMap\['pwdLastSet'\]"
        $script:fn.Body.Extent.Text | Should -Match "\`$GuidMap\['lockoutTime'\]"
    }
    It 'grants the Reset Password extended right (the actual reset operation)' {
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Reset Password'\]"
    }
}

Describe 'CorpAdmin.psm1 Grant-UserDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-UserDelegation' }
    It 'adds exactly 4 AccessRules (CreateChild/DeleteChild/WriteProperty/Reset Password)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 4
    }
    It 'includes Reset Password extended right (user-creation delegation implies the ability to reset newly-set passwords)' {
        $script:fn.Body.Extent.Text | Should -Match "ExtendedRightsMap\['Reset Password'\]"
    }
}

Describe 'CorpAdmin.psm1 Grant-OUDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-OUDelegation' }
    It 'adds exactly 3 AccessRules on organizationalUnit objects' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 3
    }
    It 'targets organizationalUnit specifically (not group, user, etc.)' {
        ([regex]::Matches($script:fn.Body.Extent.Text, "\`$GuidMap\['organizationalUnit'\]")).Count | Should -BeGreaterOrEqual 3
    }
}

Describe 'CorpAdmin.psm1 Grant-DNSOperatorsPermissionDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-DNSOperatorsPermissionDelegation' }
    It 'adds exactly 5 Allow AccessRules (no Deny rules)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 5
        $script:fn.Body.Extent.Text | Should -Not -Match 'AccessControlTypeDeny'
    }
    It 'grants the five generic operator rights' {
        $body = $script:fn.Body.Extent.Text
        $body | Should -Match 'GenericRead'
        $body | Should -Match 'GenericExecute'
        $body | Should -Match 'GenericWrite'
        $body | Should -Match 'CreateChild'
        $body | Should -Match 'DeleteChild'
    }
}

Describe 'CorpAdmin.psm1 Grant-DNSReadOnlyPermissionDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-DNSReadOnlyPermissionDelegation' }
    It 'adds 9 AccessRules and removes 1 (net 8 rules)' {
        # 2 Allow (GR, GE) + 7 Deny (GW, CC, DC, WO, WD, DT, DEL) = 9 Add
        # Then 1 Remove of the implicit ReadControl Deny that GenericWrite-Deny drags in
        (Get-AclRuleCall $script:fn 'AddAccessRule').Count | Should -Be 9
        (Get-AclRuleCall $script:fn 'RemoveAccessRule').Count | Should -Be 1
    }
    It 'denies all the dangerous mutating rights (GW, CC, DC, WO, WD, DT, DEL)' {
        $body = $script:fn.Body.Extent.Text
        foreach ($right in @('adRightsGW','adRightsCC','adRightsDC','adRightsWO','adRightsWD','adRightsDT','adRightsDEL')) {
            $body | Should -Match "\`$$right,\`$AccessControlTypeDeny" -Because "$right must be Denied for read-only delegation"
        }
    }
    It 'removes the implicit ReadControl Deny (so the group can still inspect the ACL)' {
        # The comment explaining this in the source must persist - it documents
        # the only non-trivial line in the function.
        $script:fn.Body.Extent.Text | Should -Match 'ReadControl bit'
        $script:fn.Body.Extent.Text | Should -Match 'RemoveAccessRule.*adRightsRC.*AccessControlTypeDeny'
    }
}

Describe 'CorpAdmin.psm1 Grant-ADObjectPermissionDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-ADObjectPermissionDelegation' }
    It 'adds exactly 3 AccessRules (GenericAll, CreateChild, DeleteChild)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 3
    }
    It 'grants GenericAll (full control on the target object class)' {
        $script:fn.Body.Extent.Text | Should -Match 'GenericAll'
    }
}

Describe 'CorpAdmin.psm1 Grant-GPOPermissionDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-GPOPermissionDelegation' }
    It 'does not require a TargetDN parameter (operates on the domain root)' {
        $params = $script:fn.Body.ParamBlock.Parameters.Name.VariablePath.UserPath
        $params | Should -Not -Contain 'TargetDN'
        $params | Should -Not -Contain 'TargetOU'
    }
    It 'reads the ACL at the BaseDN (domain root), not a sub-container' {
        $script:fn.Body.Extent.Text | Should -Match 'Get-Acl\s+"AD:\\\$BaseDN"'
    }
    It 'adds exactly 6 AccessRules (4 for gPLink/gPOptions read+write + 2 RSoP extended rights)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 6
    }
    It 'grants ReadProperty and WriteProperty on gPLink (the link attribute) and gPOptions (the inheritance flag)' {
        $body = $script:fn.Body.Extent.Text
        $body | Should -Match "\`$GuidMap\['gPLink'\]"
        $body | Should -Match "\`$GuidMap\['gPOptions'\]"
    }
    It 'grants both RSoP extended rights (Logging and Planning)' {
        $body = $script:fn.Body.Extent.Text
        $body | Should -Match "ExtendedRightsMap\['Generate Resultant Set of Policy \(Logging\)'\]"
        $body | Should -Match "ExtendedRightsMap\['Generate Resultant Set of Policy \(Planning\)'\]"
    }
}

Describe 'CorpAdmin.psm1 Grant-GPOCreationDelegation' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Grant-GPOCreationDelegation' }
    It 'adds exactly 1 AccessRule (CreateChild only)' {
        (Get-AclRuleCall $script:fn).Count | Should -Be 1
    }
    It 'uses InheritanceType.None (the right applies only to the target container, not children)' {
        # This is the key distinction from other CreateChild delegations.
        # GPOs live in CN=Policies,CN=System and creating a GPO must not
        # inherit "create more containers" down the tree.
        $script:fn.Body.Extent.Text | Should -Match "ActiveDirectorySecurityInheritance\]\s*'None'"
    }
}

Describe 'CorpAdmin.psm1 Get-ADSchemaGuidMap' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Get-ADSchemaGuidMap' }
    It 'declares Server as a mandatory parameter' {
        $params = $script:fn.Body.ParamBlock.Parameters
        $serverParam = $params | Where-Object { $_.Name.VariablePath.UserPath -eq 'Server' }
        $serverParam | Should -Not -BeNullOrEmpty
    }
    It 'queries the schema naming context (not the configuration NC or domain NC)' {
        $script:fn.Body.Extent.Text | Should -Match 'SchemaNamingContext'
        $script:fn.Body.Extent.Text | Should -Not -Match 'ConfigurationNamingContext'
    }
    It 'filters on schemaidguid presence and returns lDAPDisplayName -> GUID' {
        $body = $script:fn.Body.Extent.Text
        $body | Should -Match "LDAPFilter\s*=\s*'\(schemaidguid=\*\)'"
        $body | Should -Match "\`$map\[\`$_\.lDAPDisplayName\]\s*=\s*\[System\.GUID\]\`$_\.schemaIDGUID"
    }
}

Describe 'CorpAdmin.psm1 Get-ADExtendedRightsMap' {
    BeforeAll { $script:fn = Get-ModuleFunction 'Get-ADExtendedRightsMap' }
    It 'declares Server as a mandatory parameter' {
        $params = $script:fn.Body.ParamBlock.Parameters
        $serverParam = $params | Where-Object { $_.Name.VariablePath.UserPath -eq 'Server' }
        $serverParam | Should -Not -BeNullOrEmpty
    }
    It 'queries the configuration naming context (where extended rights live)' {
        $script:fn.Body.Extent.Text | Should -Match 'ConfigurationNamingContext'
        $script:fn.Body.Extent.Text | Should -Not -Match 'SchemaNamingContext'
    }
    It 'filters on controlAccessRight + rightsguid and returns displayName -> GUID' {
        $body = $script:fn.Body.Extent.Text
        $body | Should -Match "LDAPFilter\s*=\s*'\(&\(objectclass=controlAccessRight\)\(rightsguid=\*\)\)'"
        $body | Should -Match "\`$map\[\`$_\.displayName\]\s*=\s*\[System\.GUID\]\`$_\.rightsGuid"
    }
}
