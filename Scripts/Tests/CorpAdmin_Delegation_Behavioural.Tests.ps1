#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# =============================================================================
# CorpAdmin.Delegation.Behavioural.Tests.ps1
# =============================================================================
# BEHAVIOURAL cover for the Grant-*Delegation helpers in CorpAdmin.psm1.
# Complements the AST contract suite by EXECUTING each helper against an
# in-memory descriptor and asserting on the ACEs it builds.
#
# TWO THINGS THAT SHAPE THE ASSERTIONS
#  1. Set-Acl pipeline binding: a real Get-Acl on AD:\ returns a descriptor
#     decorated with provider note properties (PSPath); the helper's trailing
#     `$Acl | Set-Acl` binds its path from PSPath. A bare New-Object descriptor
#     lacks it and null-refs during binding, so the Get-Acl mock decorates the
#     descriptor like a real result.
#  2. DACL canonicalisation: AddAccessRule MERGES rules that share identity,
#     type, objectType, inheritedObjectType and inheritance flags into a single
#     ACE with the rights OR'd together. So the ACE COUNT is the number of
#     distinct such tuples (not the number of AddAccessRule calls), and a right
#     is matched by testing that an ACE CONTAINS it as a flag - not by string-
#     equality against the whole ActiveDirectoryRights value.
#
# PATTERN
#   Helpers are exported, so called directly; Get-ADGroup/Get-Acl/Set-Acl are
#   mocked -ModuleName in BeforeEach; the built descriptor is captured via a
#   $global the Get-Acl mock writes (the one scope a module-scope mock body and
#   the test both see). Fixtures use DISTINCT GUIDs per key so a swap is caught.
#
# BOUNDARY
#   Verifies the descriptor the helper BUILDS - not that Set-Acl persists it to
#   AD:\, nor effective access once deny-precedence/inheritance resolve live.
# =============================================================================

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1'
    if (-not (Test-Path $modulePath)) {
        throw "CorpAdmin module not found at: $modulePath"
    }
    Import-Module $modulePath -Force
    $script:empty = [guid]::Empty
    $script:g = @{
        computer           = [guid]'c0000000-0000-0000-0000-000000000001'
        group              = [guid]'c0000000-0000-0000-0000-000000000002'
        member             = [guid]'c0000000-0000-0000-0000-000000000003'
        user               = [guid]'c0000000-0000-0000-0000-000000000004'
        organizationalUnit = [guid]'c0000000-0000-0000-0000-000000000005'
        pwdLastSet         = [guid]'c0000000-0000-0000-0000-000000000006'
        lockoutTime        = [guid]'c0000000-0000-0000-0000-000000000007'
        gPLink             = [guid]'c0000000-0000-0000-0000-000000000008'
        gPOptions          = [guid]'c0000000-0000-0000-0000-000000000009'
    }
    $script:er = @{
        'Reset Password'                              = [guid]'e0000000-0000-0000-0000-000000000001'
        'Change Password'                             = [guid]'e0000000-0000-0000-0000-000000000002'
        'Validated write to DNS host name'            = [guid]'e0000000-0000-0000-0000-000000000003'
        'Validated write to service principal name'   = [guid]'e0000000-0000-0000-0000-000000000004'
        'Generate Resultant Set of Policy (Logging)'  = [guid]'e0000000-0000-0000-0000-000000000005'
        'Generate Resultant Set of Policy (Planning)' = [guid]'e0000000-0000-0000-0000-000000000006'
    }
    # Count ACEs matching the supplied criteria. Rights = CONTAINED-flag test
    # (see canonicalisation note); other criteria exact; omitted = any.
    function script:Select-Ace {
        param(
            [Parameter(Mandatory)]$Aces,
            [string]$Rights,
            [string]$Type,
            [guid]$ObjectType,
            [guid]$InheritedObjectType,
            [string]$Inheritance
        )
        $checkOT  = $PSBoundParameters.ContainsKey('ObjectType')
        $checkIOT = $PSBoundParameters.ContainsKey('InheritedObjectType')
        $want = if ($Rights) {
            [System.DirectoryServices.ActiveDirectoryRights]$Rights
        } else { $null }
        @($Aces | Where-Object {
            (-not $Rights -or (($_.ActiveDirectoryRights -band $want) -eq $want)) -and
            (-not $Type -or $_.AccessControlType.ToString() -eq $Type) -and
            (-not $checkOT -or $_.ObjectType -eq $ObjectType) -and
            (-not $checkIOT -or $_.InheritedObjectType -eq $InheritedObjectType) -and
            (-not $Inheritance -or $_.InheritanceType.ToString() -eq $Inheritance)
        })
    }
    function script:Get-CapturedAce {
        @($script:__delegAcl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]) | Where-Object { $_.IdentityReference.Value -eq $script:__delegSid })
    }
}

Describe 'Grant-*Delegation helpers build correct ACLs' {
    BeforeEach {
        $script:__delegSid  = 'S-1-5-21-1004336348-1177238915-682003330-1104'
        $script:__delegAcl  = $null
        $script:__delegPath = $null
        Mock Get-ADGroup -ModuleName CorpAdmin {
            [pscustomobject]@{
                SID = (New-Object System.Security.Principal.SecurityIdentifier $script:__delegSid)
            }
        }
        Mock Get-Acl -ModuleName CorpAdmin {
            $script:__delegPath = @($Path)[0] # Get-Acl -Path is [string[]]; unwrap
            $acl = New-Object System.DirectoryServices.ActiveDirectorySecurity
            # Decorate like a real Get-Acl result so `$Acl | Set-Acl` binds (PSPath).
            Add-Member -InputObject $acl -Force -NotePropertyName PSPath      -NotePropertyValue $script:__delegPath
            Add-Member -InputObject $acl -Force -NotePropertyName PSChildName -NotePropertyValue 'delegated'
            Add-Member -InputObject $acl -Force -NotePropertyName PSDrive     -NotePropertyValue 'AD'
            Add-Member -InputObject $acl -Force -NotePropertyName PSProvider  -NotePropertyValue 'ActiveDirectory'
            $script:__delegAcl = $acl
            $acl
        }
        Mock Set-Acl -ModuleName CorpAdmin { }
    }
    
    Describe 'Grant-ComputerJoinDelegation' {
        It 'builds the correct computer-join delegation ACL' {
            Grant-ComputerJoinDelegation -AdminGroupName 'Corp_Computer_Admins' -TargetOU 'Servers' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl     -ModuleName CorpAdmin -Times 1 -Exactly
            Should -Invoke Get-ADGroup -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Servers,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 6 # merged: CC+DC, WP+WD, then 2 Self + 2 ExtendedRight
            # CC/DC -> (objectType, inheritance): ObjectType=computer, All (one merged ACE).
            (Select-Ace $aces -Rights CreateChild -Type Allow -ObjectType $script:g.computer -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights DeleteChild -Type Allow -ObjectType $script:g.computer -Inheritance All).Count | Should -Be 1
            # WP/WD -> (inheritance, inheritedObjectType): ObjectType empty, inherited computer.
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:empty -InheritedObjectType $script:g.computer -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteDacl     -Type Allow -InheritedObjectType $script:g.computer -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights Self -Type Allow -ObjectType $script:er.'Validated write to DNS host name').Count | Should -Be 1
            (Select-Ace $aces -Rights Self -Type Allow -ObjectType $script:er.'Validated write to service principal name').Count | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Reset Password'  -InheritedObjectType $script:g.computer).Count | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Change Password' -InheritedObjectType $script:g.computer).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-GroupDelegation' {
        It 'builds the correct group delegation ACL' {
            Grant-GroupDelegation -AdminGroupName 'Corp_Group_Admins' -TargetOU 'Groups' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Groups,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 2 # merged: CreateChild+DeleteChild
            (Select-Ace $aces -Rights CreateChild   -Type Allow -ObjectType $script:g.group -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights DeleteChild   -Type Allow -ObjectType $script:g.group -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -InheritedObjectType $script:g.group -Inheritance Descendents).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-GroupMembershipEditDelegation' {
        It 'grants WriteProperty on member only, scoped to group descendents' {
            Grant-GroupMembershipEditDelegation -AdminGroupName 'Corp_GroupMembership_Admins' -TargetOU 'Groups' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Groups,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.member -InheritedObjectType $script:g.group -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.user).Count | Should -Be 0
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:empty).Count  | Should -Be 0
        }
    }
    
    Describe 'Grant-PasswordResetDelegation' {
        It 'builds the correct password-reset delegation ACL' {
            Grant-PasswordResetDelegation -AdminGroupName 'Corp_PwReset_Admins' -TargetOU 'Users' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Users,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 3 # distinct objectTypes (pwdLastSet, lockoutTime) + ExtendedRight: no merge
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.pwdLastSet  -InheritedObjectType $script:g.user -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.lockoutTime -InheritedObjectType $script:g.user -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Reset Password' -InheritedObjectType $script:g.user -Inheritance Descendents).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-UserDelegation' {
        It 'builds the correct user delegation ACL' {
            Grant-UserDelegation -AdminGroupName 'Corp_User_Admins' -TargetOU 'Users' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Users,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 3 # merged: CreateChild+DeleteChild; + WriteProperty + ExtendedRight
            (Select-Ace $aces -Rights CreateChild   -Type Allow -ObjectType $script:g.user -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights DeleteChild   -Type Allow -ObjectType $script:g.user -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -InheritedObjectType $script:g.user -Inheritance Descendents).Count | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Reset Password' -InheritedObjectType $script:g.user).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-OUDelegation' {
        It 'builds the correct OU delegation ACL' {
            Grant-OUDelegation -AdminGroupName 'Corp_OU_Admins' -TargetOU 'Corp' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\OU=Corp,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 2 # merged: CreateChild+DeleteChild
            (Select-Ace $aces -Rights CreateChild   -Type Allow -ObjectType $script:g.organizationalUnit -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights DeleteChild   -Type Allow -ObjectType $script:g.organizationalUnit -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -InheritedObjectType $script:g.organizationalUnit -Inheritance Descendents).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-DNSOperatorsPermissionDelegation' {
        It 'grants the five DNS-operator rights as Allow at the DN path' {
            Grant-DNSOperatorsPermissionDelegation -AdminGroupName 'Corp_DNS_Operators' -TargetDN 'CN=MicrosoftDNS,CN=System' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\CN=MicrosoftDNS,CN=System,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 1 # all five Allow/All/no-GUID rules merge into one ACE
            foreach ($rt in 'GenericRead','GenericExecute','GenericWrite','CreateChild','DeleteChild') {
                (Select-Ace $aces -Rights $rt -Type Allow -Inheritance All).Count | Should -Be 1 -Because "$rt should be granted"
            }
            (Select-Ace $aces -Type Deny).Count | Should -Be 0
        }
    }
    
    Describe 'Grant-DNSReadOnlyPermissionDelegation' {
        It 'allows read/execute, denies the write family, and leaves ReadControl NOT denied' {
            Grant-DNSReadOnlyPermissionDelegation -AdminGroupName 'Corp_DNS_ReadOnly' -TargetDN 'CN=MicrosoftDNS,CN=System' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\CN=MicrosoftDNS,CN=System,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            (Select-Ace $aces -Rights GenericRead    -Type Allow).Count | Should -Be 1
            (Select-Ace $aces -Rights GenericExecute -Type Allow).Count | Should -Be 1
            # OR the denied rights into one mask (robust to ACE coalescing).
            $denyMask = 0
            foreach ($ace in ($aces | Where-Object AccessControlType -eq 'Deny')) {
                $denyMask = $denyMask -bor [int]$ace.ActiveDirectoryRights
            }
            # The write/delete family must be denied. Asserted as a mask rather than
            # per-right so it survives coalescing. GenericWrite is deliberately NOT
            # in this set: it carries the ReadControl bit the helper strips via
            # RemoveAccessRule, so a whole-GenericWrite match would fail post-strip -
            # that path is covered by the ReadControl-absent assertion below instead.
            $writeDelete =
                [int][System.DirectoryServices.ActiveDirectoryRights]::CreateChild -bor
                [int][System.DirectoryServices.ActiveDirectoryRights]::DeleteChild -bor
                [int][System.DirectoryServices.ActiveDirectoryRights]::WriteDacl   -bor
                [int][System.DirectoryServices.ActiveDirectoryRights]::WriteOwner  -bor
                [int][System.DirectoryServices.ActiveDirectoryRights]::Delete
            ($denyMask -band $writeDelete) | Should -Be $writeDelete -Because 'the write/delete family must be denied'
            # Load-bearing: after RemoveAccessRule strips the ReadControl bit that the
            # GenericWrite Deny drags in, the SID's effective denied rights must NOT
            # include READ_CONTROL, so the group can still read the ACL.
            $readControl = [int][System.DirectoryServices.ActiveDirectoryRights]::ReadControl
            ($denyMask -band $readControl) | Should -Be 0 -Because 'ReadControl deny must be stripped so the group can inspect the ACL'
        }
    }
    
    Describe 'Grant-ADObjectPermissionDelegation' {
        It 'grants GenericAll/CreateChild/DeleteChild Allow at the DN path' {
            Grant-ADObjectPermissionDelegation -AdminGroupName 'Corp_Sites_Admins' -TargetDN 'CN=Sites,CN=Configuration' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\CN=Sites,CN=Configuration,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 1 # GenericAll subsumes/merges with CreateChild+DeleteChild
            (Select-Ace $aces -Rights GenericAll  -Type Allow -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights CreateChild -Type Allow -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights DeleteChild -Type Allow -Inheritance All).Count | Should -Be 1
        }
    }
    
    Describe 'Grant-GPOPermissionDelegation' {
        It 'reads the domain-root path and grants the GPO management ACEs' {
            Grant-GPOPermissionDelegation -AdminGroupName 'Corp_GPO_Admins' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 4 # merged: Read+Write per attribute (gPLink, gPOptions) + 2 ExtendedRight
            (Select-Ace $aces -Rights ReadProperty  -Type Allow -ObjectType $script:g.gPLink    -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.gPLink    -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights ReadProperty  -Type Allow -ObjectType $script:g.gPOptions -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights WriteProperty -Type Allow -ObjectType $script:g.gPOptions -Inheritance All).Count | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Generate Resultant Set of Policy (Logging)').Count  | Should -Be 1
            (Select-Ace $aces -Rights ExtendedRight -Type Allow -ObjectType $script:er.'Generate Resultant Set of Policy (Planning)').Count | Should -Be 1
        }
    }
    
    Describe 'Grant-GPOCreationDelegation' {
        It 'grants a single CreateChild Allow ACE with NO inheritance' {
            Grant-GPOCreationDelegation -AdminGroupName 'Corp_GPO_Creators' -TargetDN 'CN=Policies,CN=System' -BaseDN 'DC=corp,DC=example,DC=com' -GuidMap $script:g -ExtendedRightsMap $script:er
            Should -Invoke Set-Acl -ModuleName CorpAdmin -Times 1 -Exactly
            $script:__delegPath | Should -Be 'AD:\CN=Policies,CN=System,DC=corp,DC=example,DC=com'
            $aces = Get-CapturedAce
            $aces.Count | Should -Be 1
            (Select-Ace $aces -Rights CreateChild -Type Allow -Inheritance None).Count | Should -Be 1
        }
    }
}
