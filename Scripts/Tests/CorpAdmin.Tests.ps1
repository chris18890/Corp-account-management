#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

BeforeAll {
    # Runs once before any test in the file.
    # Import the thing under test here.
    $modulePath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1'
    if (-not (Test-Path $modulePath)) {
        throw "CorpAdmin module not found at: $modulePath"
    }
    Import-Module $modulePath -Force
}

Describe 'Get-EnvironmentConfig' {
    BeforeAll {
        $testEnvPath = Join-Path $TestDrive 'environment.psd1'
        "@{ Network = @{}; OUs = @{}; Groups = @{}; Shares = @{}; Locale = @{};
           Security = @{ PasswordLength = 20; MaxElevationMinutes = 480 };
           Azure = @{}; Exchange = @{}; EntraRoles = @{}; WSUS = @{} }" |
            Set-Content $testEnvPath
    }
    It 'loads from an explicit path' {
        $cfg = Get-EnvironmentConfig -Path $testEnvPath -Force
        $cfg | Should -Not -BeNullOrEmpty
    }
    It 'returns the same object on a second call (cache hit)' {
        $env:CORPADMIN_ENV_PSD1 = $testEnvPath
        try {
            $first  = Get-EnvironmentConfig -Force  # no -Path, so populates cache
            $second = Get-EnvironmentConfig         # returns from cache
            [object]::ReferenceEquals($first, $second) | Should -BeTrue
        } finally {
            $env:CORPADMIN_ENV_PSD1 = $null
        }
    }
    It '-Force bypasses the cache' {
        $env:CORPADMIN_ENV_PSD1 = $testEnvPath
        try {
            $first  = Get-EnvironmentConfig -Force  # populates cache, returns $config1
            $second = Get-EnvironmentConfig -Force  # Force set, skips cache, returns $config2
            [object]::ReferenceEquals($first, $second) | Should -BeFalse
        } finally {
            $env:CORPADMIN_ENV_PSD1 = $null
        }
    }
    It 'invalidates the cache when CORPADMIN_ENV_PSD1 changes between calls' {
        $envA = Join-Path $TestDrive 'envA.psd1'
        $envB = Join-Path $TestDrive 'envB.psd1'
        "@{ Network=@{}; OUs=@{}; Groups=@{}; Shares=@{}; Locale=@{};
           Security=@{ PasswordLength=20; MaxElevationMinutes=480 };
           Azure=@{}; Exchange=@{}; EntraRoles=@{}; WSUS=@{};
           Marker='A' }" | Set-Content $envA
        "@{ Network=@{}; OUs=@{}; Groups=@{}; Shares=@{}; Locale=@{};
           Security=@{ PasswordLength=20; MaxElevationMinutes=480 };
           Azure=@{}; Exchange=@{}; EntraRoles=@{}; WSUS=@{};
           Marker='B' }" | Set-Content $envB
        try {
            $env:CORPADMIN_ENV_PSD1 = $envA
            (Get-EnvironmentConfig).Marker | Should -Be 'A' # no -Force: populates cache keyed on envA
            $env:CORPADMIN_ENV_PSD1 = $envB
            # No -Force. Old behaviour returned the stale 'A'; fixed behaviour
            # re-resolves because the path key no longer matches.
            (Get-EnvironmentConfig).Marker | Should -Be 'B'
        } finally {
            $env:CORPADMIN_ENV_PSD1 = $null
        }
    }
    It 'falls back to the default path when CORPADMIN_ENV_PSD1 is unset after being cached' {
        $envA = Join-Path $TestDrive 'envA.psd1'
        "@{ Network=@{}; OUs=@{}; Groups=@{}; Shares=@{}; Locale=@{};
           Security=@{ PasswordLength=20; MaxElevationMinutes=480 };
           Azure=@{}; Exchange=@{}; EntraRoles=@{}; WSUS=@{};
           Marker='A' }" | Set-Content $envA
        try {
            $env:CORPADMIN_ENV_PSD1 = $envA
            (Get-EnvironmentConfig).Marker | Should -Be 'A'
            $env:CORPADMIN_ENV_PSD1 = $null
            # Now resolves to the repo default environment.psd1 (no Marker key).
            (Get-EnvironmentConfig).Marker | Should -BeNullOrEmpty
        } finally {
            $env:CORPADMIN_ENV_PSD1 = $null
        }
    }
    It 'throws when a required section is missing' {
        $bad = Join-Path $TestDrive 'env-missing.psd1'
        "@{ Network = @{}; OUs = @{}; Groups = @{}; Shares = @{}; Locale = @{};
           Security = @{ PasswordLength = 20; MaxElevationMinutes = 480 };
           Azure = @{}; Exchange = @{}; EntraRoles = @{} }" | Set-Content $bad
        { Get-EnvironmentConfig -Path $bad -Force } | Should -Throw -ExpectedMessage '*missing required section*WSUS*'
    }
    It 'names every missing section, not just the first' {
        $bad = Join-Path $TestDrive 'env-missing-two.psd1'
        "@{ Network = @{}; OUs = @{}; Groups = @{}; Shares = @{}; Locale = @{};
           Security = @{}; Azure = @{}; Exchange = @{} }" | Set-Content $bad
        { Get-EnvironmentConfig -Path $bad -Force } | Should -Throw -ExpectedMessage '*EntraRoles, WSUS*'
    }
    It 'accepts a config that has every required section' {
        { Get-EnvironmentConfig -Path $testEnvPath -Force } | Should -Not -Throw
    }
}

Describe 'CorpAdmin.psd1 export manifest integrity' {
    BeforeAll {
        $manifestPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1'
        $modPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $script:exportedFns = @((Import-PowerShellDataFile -LiteralPath $manifestPath).FunctionsToExport)
        $tokens = $null
        $script:parseErrors = $null
        $script:moduleAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $modPath,
            [ref]$tokens,
            [ref]$script:parseErrors
        )
        $script:definedFns = @(
            $script:moduleAst.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true).Name
        )
    }
    It 'CorpAdmin.psm1 parses cleanly' {
        $script:parseErrors | Should -BeNullOrEmpty -Because 'a syntax error in CorpAdmin.psm1 breaks module import for every consumer'
    }
    It 'every exported function is actually defined (no stale or misspelt entry)' {
        foreach ($name in $script:exportedFns) {
            $script:definedFns | Should -Contain $name -Because "$name is exported but not defined in CorpAdmin.psm1"
        }
    }
    It 'FunctionsToExport has no duplicate entries' {
        ($script:exportedFns | Group-Object | Where-Object Count -gt 1) | Should -BeNullOrEmpty
    }
    It '<Function> is part of the public surface and must be exported' -TestCases (
        @(
            'Get-EnvironmentConfig','Write-LogFile','New-Password','Test-Password',
            'ConvertTo-IntOrDefault','ConvertTo-SafeName','ConvertTo-SafeSamAccountName',
            'Resolve-GroupMemberObject','Test-IsMemberOf','Get-ADGroupMemberTTLState','Add-GroupMember','Remove-GroupMember',
            'New-ADOU','New-DomainGroup','Add-GPOLink','Invoke-ADSync',
            'New-UserMailbox','Update-UserMailbox','New-UserOnPremMailbox','Update-UserOnPremMailbox',
            'Send-NotificationEmail','Get-ADSchemaGuidMap','Get-ADExtendedRightsMap',
            'Grant-ComputerJoinDelegation','Grant-GroupDelegation','Grant-GroupMembershipEditDelegation',
            'Grant-PasswordResetDelegation','Grant-UserDelegation','Grant-OUDelegation',
            'Grant-DNSOperatorsPermissionDelegation','Grant-DNSReadOnlyPermissionDelegation',
            'Grant-ADObjectPermissionDelegation','Grant-GPOPermissionDelegation','Grant-GPOCreationDelegation'
        ) | ForEach-Object { @{ Function = $_ } }
    ) {
        param($Function)
        $script:definedFns  | Should -Contain $Function -Because "$Function should be implemented in CorpAdmin.psm1"
        $script:exportedFns | Should -Contain $Function -Because "$Function is part of the public CorpAdmin module surface"
    }
    It 'keeps mailbox helper functions private' {
        foreach ($helper in 'Confirm-MailboxType','Grant-MailboxAccess','Grant-MailboxAccessOnPrem') {
            $script:definedFns  | Should -Contain $helper -Because "$helper should exist in CorpAdmin.psm1"
            $script:exportedFns | Should -Not -Contain $helper -Because "$helper is an internal implementation helper"
        }
    }
}

Describe 'ConvertTo-SafeName' {
    It 'strips ? @ \ and + characters' {
        ConvertTo-SafeName 'Jo?h@n\+' | Should -BeExactly 'John'
    }
    It 'trims surrounding whitespace' {
        ConvertTo-SafeName '  Smith  ' | Should -BeExactly 'Smith'
    }
    It 'leaves a clean name unchanged' {
        ConvertTo-SafeName "O'Brien" | Should -BeExactly "O'Brien"
    }
    It 'returns empty string for empty input' {
        ConvertTo-SafeName '' | Should -BeExactly ''
    }
}

Describe 'ConvertTo-SafeSamAccountName' {
    It 'strips characters outside [A-Za-z0-9.-]' {
        ConvertTo-SafeSamAccountName 'jo hn_smith!' | Should -BeExactly 'johnsmith'
    }
    It 'preserves dots and hyphens' {
        ConvertTo-SafeSamAccountName 'john.s-mith' | Should -BeExactly 'john.s-mith'
    }
    It 'truncates to 20 characters by default' {
        (ConvertTo-SafeSamAccountName ('a' * 30)).Length | Should -Be 20
    }
    It 'applies the prefix before truncating' {
        ConvertTo-SafeSamAccountName 'johnsmith' -Prefix 'admin.' | Should -BeExactly 'admin.johnsmith'
    }
    It 'caps the prefixed result at 20 chars' {
        (ConvertTo-SafeSamAccountName ('a' * 30) -Prefix 'admin.').Length | Should -Be 20
    }
    It 'strips the value but trusts the prefix' {
        ConvertTo-SafeSamAccountName 'jo hn' -Prefix 'admin.' | Should -BeExactly 'admin.john'
    }
}

Describe 'New-Password' {
    Context 'Default parameters' {
        It 'returns a string of the default length' {
            $pw = New-Password
            $pw | Should -BeOfType [string]
            $pw.Length | Should -Be 20
        }
        It 'contains at least one uppercase letter' {
            (New-Password) -cmatch '[A-Z]' | Should -BeTrue
        }
        It 'contains at least one lowercase letter' {
            (New-Password) -cmatch '[a-z]' | Should -BeTrue
        }
        It 'contains at least one digit' {
            (New-Password) -cmatch '[0-9]' | Should -BeTrue
        }
        It 'contains at least four special characters' {
            $pw = New-Password
            ([regex]::Matches($pw, '[^A-Za-z0-9]')).Count | Should -BeGreaterOrEqual 4
        }
        It 'never contains visually ambiguous characters (100 samples)' {
            1..100 | ForEach-Object {
                New-Password | Should -Not -CMatch '[IOl01]'
            }
        }
    }
    Context 'Custom length' {
        It 'honours the -Length parameter' {
            (New-Password -Length 32).Length | Should -Be 32
        }
        It 'rejects lengths below 12' {
            { New-Password -Length 8 } | Should -Throw -ExpectedMessage '*minimum allowed range of 12*'
        }
    }
    Context 'Custom minimums' {
        It 'throws when minimums exceed length' {
            { New-Password -Length 12 -MinSpecial 20 } | Should -Throw '*exceed total password length*'
        }
    }
}

Describe 'Test-Password' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'test.log'
    }
    It 'passes for a compliant password' {
        { Test-Password -LogFile $logFile -Password 'GoodP@ssw0rd!ThatsLong#X' -PasswordLength 20 } | Should -Not -Throw
    }
    It 'throws when too short' {
        { Test-Password -LogFile $logFile -Password 'Short1!' -PasswordLength 20 } | Should -Throw -ExpectedMessage '*does not comply*'
    }
    It 'throws when missing an uppercase letter' {
        { Test-Password -LogFile $logFile -Password 'nouppercase1234!@#$' -PasswordLength 20 } | Should -Throw -ExpectedMessage '*does not comply*'
    }
    It 'throws when missing a lowercase letter' {
        { Test-Password -LogFile $logFile -Password 'NOLOWERCASE1234!@#$' -PasswordLength 20 } | Should -Throw -ExpectedMessage '*does not comply*'
    }
    It 'throws when missing a digit' {
        { Test-Password -LogFile $logFile -Password 'NoDigitsHereAtAll!@#$' -PasswordLength 20 } | Should -Throw -ExpectedMessage '*does not comply*'
    }
    It 'throws when missing a special character' {
        { Test-Password -LogFile $logFile -Password 'NoSpecialsHere1234567' -PasswordLength 20 } | Should -Throw -ExpectedMessage '*does not comply*'
    }
}

Describe 'Resolve-GroupMemberObject' {
    It 'returns Found with the object for a single match' {
        Mock Get-ADObject -ModuleName CorpAdmin { [pscustomobject]@{ ObjectClass='user'; DistinguishedName='CN=alice,OU=Users,DC=corp,DC=local' } }
        $r = Resolve-GroupMemberObject -Member 'alice' -DCHostName 'DC1'
        $r.Status | Should -Be 'Found'
        $r.Count  | Should -Be 1
        $r.Object.DistinguishedName | Should -Be 'CN=alice,OU=Users,DC=corp,DC=local'
    }
    It 'returns NotFound for zero matches' {
        Mock Get-ADObject -ModuleName CorpAdmin { $null }
        $r = Resolve-GroupMemberObject -Member 'ghost' -DCHostName 'DC1'
        $r.Status | Should -Be 'NotFound'; $r.Count | Should -Be 0; $r.Object | Should -BeNullOrEmpty
    }
    It 'returns Ambiguous with the count for multiple matches' {
        Mock Get-ADObject -ModuleName CorpAdmin {
            @(
                [pscustomobject]@{ ObjectClass='user';  DistinguishedName='CN=dup,OU=Users,DC=corp,DC=local' }
                [pscustomobject]@{ ObjectClass='group'; DistinguishedName='CN=dup,OU=Groups,DC=corp,DC=local' }
            )
        }
        $r = Resolve-GroupMemberObject -Member 'dup' -DCHostName 'DC1'
        $r.Status | Should -Be 'Ambiguous'; $r.Count | Should -Be 2; $r.Object | Should -BeNullOrEmpty
    }
    It 'queries with the class-agnostic user-or-group filter' {
        Mock Get-ADObject -ModuleName CorpAdmin { [pscustomobject]@{ DistinguishedName='CN=x' } }
        Resolve-GroupMemberObject -Member 'x' -DCHostName 'DC1' | Out-Null
        Should -Invoke Get-ADObject -ModuleName CorpAdmin -ParameterFilter {
            $LDAPFilter -eq "(|(sAMAccountName=x)(&(objectClass=group)(cn=x)))"
        }
    }
    It 'LDAP-escapes the member name' {
        Mock Get-ADObject -ModuleName CorpAdmin { [pscustomobject]@{ DistinguishedName='CN=x' } }
        Resolve-GroupMemberObject -Member 'a(b)c' -DCHostName 'DC1' | Out-Null
        Should -Invoke Get-ADObject -ModuleName CorpAdmin -ParameterFilter { $LDAPFilter -like '*a\28b\29c*' }
    }
    It 'resolves distinguishedName input via -Identity' {
        $dn = 'CN=Role_Reception,OU=Groups,DC=corp,DC=local'
        $script:getADObjectCalls = @()
        Mock Get-ADObject -ModuleName CorpAdmin {
            param(
                $Identity,
                $LDAPFilter,
                $Server,
                $ErrorAction
            )
            $script:getADObjectCalls += [pscustomobject]@{
                Identity   = $Identity
                LDAPFilter = $LDAPFilter
                Server     = $Server
            }
            [pscustomobject]@{
                ObjectClass       = 'group'
                DistinguishedName = $dn
            }
        }
        $r = Resolve-GroupMemberObject -Member $dn -DCHostName 'DC1'
        $r.Status | Should -Be 'Found'
        $r.Count  | Should -Be 1
        $r.Object.DistinguishedName | Should -Be $dn
        $identityCalls = @(
            $script:getADObjectCalls | Where-Object {
                [string]$_.Identity -eq $dn -and $_.Server -eq 'DC1' -and [string]::IsNullOrWhiteSpace([string]$_.LDAPFilter)
            }
        )
        $identityCalls.Count | Should -Be 1
    }
    It 'returns NotFound when distinguishedName input does not resolve' {
        $dn = 'CN=Missing,OU=Groups,DC=corp,DC=local'
        Mock Get-ADObject -ModuleName CorpAdmin {
            throw [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]::new("not found")
        }
        $r = Resolve-GroupMemberObject -Member $dn -DCHostName 'DC1'
        $r.Status | Should -Be 'NotFound'
        $r.Count  | Should -Be 0
        $r.Object | Should -BeNullOrEmpty
    }
    It 'queries the requested domain controller' {
        Mock Get-ADObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                DistinguishedName = 'CN=x,DC=corp,DC=local'
            }
        }
        Resolve-GroupMemberObject -Member 'x' -DCHostName 'DC1' | Out-Null
        Should -Invoke Get-ADObject -ModuleName CorpAdmin -ParameterFilter {
            $Server -eq 'DC1'
        }
    }
    It 'LDAP-escapes wildcard and backslash characters' {
        Mock Get-ADObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                DistinguishedName = 'CN=x,DC=corp,DC=local'
            }
        }
        Resolve-GroupMemberObject -Member 'a*b\c' -DCHostName 'DC1' | Out-Null
        Should -Invoke Get-ADObject -ModuleName CorpAdmin -ParameterFilter {
            $LDAPFilter -like '*a\2ab\5cc*'
        }
    }
}

Describe 'Test-IsMemberOf' {
    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1') -Force
    }
    It 'returns $true when the user is a recursive member of the group' {
        Mock -ModuleName CorpAdmin Get-ADGroup       { [pscustomobject]@{ Name = 'GroupA'; DistinguishedName = 'CN=GroupA,DC=corp,DC=local' } }
        Mock -ModuleName CorpAdmin Get-ADGroupMember { @([pscustomobject]@{ SamAccountName = 'alice' }) }
        Test-IsMemberOf -Sam 'alice' -GroupNames @('GroupA') -DCHostName 'dc01' | Should -BeTrue
    }
    It 'returns $false when the user is in none of the groups' {
        Mock -ModuleName CorpAdmin Get-ADGroup       { [pscustomobject]@{ Name = 'GroupA'; DistinguishedName = 'CN=GroupA,DC=corp,DC=local' } }
        Mock -ModuleName CorpAdmin Get-ADGroupMember { @([pscustomobject]@{ SamAccountName = 'bob' }) }
        Test-IsMemberOf -Sam 'alice' -GroupNames @('GroupA','GroupB') -DCHostName 'dc01' | Should -BeFalse
    }
    It 'returns $true if the user is in ANY one of several groups' {
        Mock -ModuleName CorpAdmin Get-ADGroup { [pscustomobject]@{ Name = 'g'; DistinguishedName = 'CN=g,DC=corp,DC=local' } }
        $script:call = 0
        Mock -ModuleName CorpAdmin Get-ADGroupMember {
            $script:call++
            if ($script:call -ge 2) { @([pscustomobject]@{ SamAccountName = 'alice' }) } else { @() }
        }
        Test-IsMemberOf -Sam 'alice' -GroupNames @('GroupA','GroupB') -DCHostName 'dc01' | Should -BeTrue
    }
    It 'skips a group that does not resolve, without querying its membership' {
        Mock -ModuleName CorpAdmin Get-ADGroup       { $null }
        Mock -ModuleName CorpAdmin Get-ADGroupMember { throw 'should not be called' }
        Test-IsMemberOf -Sam 'alice' -GroupNames @('Nonexistent') -DCHostName 'dc01' | Should -BeFalse
    }
    It 'queries membership recursively' {
        Mock -ModuleName CorpAdmin Get-ADGroup       { [pscustomobject]@{ Name = 'g'; DistinguishedName = 'CN=g,DC=corp,DC=local' } }
        Mock -ModuleName CorpAdmin Get-ADGroupMember { @([pscustomobject]@{ SamAccountName = 'alice' }) }
        Test-IsMemberOf -Sam 'alice' -GroupNames @('g') -DCHostName 'dc01' | Out-Null
        Should -Invoke -ModuleName CorpAdmin Get-ADGroupMember -ParameterFilter { $Recursive -eq $true }
    }
    It 'escapes a single quote in the group name so the -Filter does not break' {
        Mock -ModuleName CorpAdmin Get-ADGroup       { [pscustomobject]@{ Name = "O'Brien Admins"; DistinguishedName = "CN=O'Brien Admins,DC=corp,DC=local" } }
        Mock -ModuleName CorpAdmin Get-ADGroupMember { @([pscustomobject]@{ SamAccountName = 'alice' }) }
        Test-IsMemberOf -Sam 'alice' -GroupNames @("O'Brien Admins") -DCHostName 'dc01' | Out-Null
        Should -Invoke -ModuleName CorpAdmin Get-ADGroup -ParameterFilter { $Filter -match "O''Brien Admins" }
    }
}

Describe 'Add-GroupMember' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'test.log'
    }
    Context 'When Resolve-GroupMemberObject returns an unexpected status' {
        BeforeEach {
            Mock Write-LogFile -ModuleName CorpAdmin { }
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Status = 'WeirdState'
                    Object = $null
                    Count  = 0
                }
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin { throw 'should not be called' }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'throws a clear unexpected-status error before membership checks or add, and logs END' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            } | Should -Throw "*Unexpected Resolve-GroupMemberObject status 'WeirdState'*"
            Should -Invoke Get-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq 'Add-GroupMember Failed'
            }
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq '=== Add-GroupMember END ==='
            }
        }
    }
    Context "When Resolve-GroupMemberObject returns Found without a DistinguishedName" {
        BeforeEach {
            Mock Write-LogFile -ModuleName CorpAdmin { }
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Status = 'Found'
                    Object = [pscustomobject]@{
                        SamAccountName = 'alice'
                    }
                    Count = 1
                }
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin { throw 'should not be called' }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'throws a clear malformed-found-object error before membership checks or add, and logs END' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            } | Should -Throw "*Status='Found' but did not return an object with DistinguishedName*"
            Should -Invoke Get-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq 'Add-GroupMember Failed'
            }
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq '=== Add-GroupMember END ==='
            }
        }
    }
    Context 'When group exists and member exists' {
        BeforeEach {
            $script:gmCalls = 0
            # Normal group resolution.
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            } -ParameterFilter {
                -not $ShowMemberTimeToLive
            }
            # TTL readback path. This is the new post-write validation route.
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                    member            = @(
                        '<TTL=3600>,CN=alice,OU=Users,DC=corp,DC=local'
                    )
                }
            } -ParameterFilter {
                $ShowMemberTimeToLive -eq $true -and
                $Properties -contains 'member'
            }
            # Current Add-GroupMember implementation resolves members via Get-ADObject,
            # not Get-ADUser.
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            } -ParameterFilter {
                $LDAPFilter -like '*(sAMAccountName=alice)*'
            }
            # PAM / TTL feature gate.
            Mock Get-ADOptionalFeature -ModuleName CorpAdmin {
                [pscustomobject]@{
                    EnabledScopes = @('DC=corp,DC=local')
                }
            }
            # First call is the pre-check: alice is not yet a member.
            # Second call is the post-add check: alice is now present.
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                $script:gmCalls = ([int]$script:gmCalls) + 1
                if ($script:gmCalls -le 1) {
                    @()
                } else {
                    [pscustomobject]@{
                        SamAccountName    = 'alice'
                        DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                    }
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
            Mock Start-Sleep -ModuleName CorpAdmin { }
        }
        It 'calls Add-ADGroupMember once' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'passes exact -MemberTimeToLive when -TimeSpan is supplied' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $MemberTimeToLive -is [timespan] -and
                $MemberTimeToLive.TotalMinutes -eq 60
            }
        }
        It 'validates TTL membership using Get-ADGroup -ShowMemberTimeToLive' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            Should -Invoke Get-ADGroup -ModuleName CorpAdmin -Times 1 -ParameterFilter {
                $ShowMemberTimeToLive -eq $true -and
                $Properties -contains 'member'
            }
        }
        It 'does not roll back when TTL membership is confirmed' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'When TTL post-write validation cannot confirm an expiring link' {
        BeforeEach {
            $script:gmCalls = 0
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            } -ParameterFilter {
                -not $ShowMemberTimeToLive
            }
            # Deliberately returns alice as a plain member DN, without <TTL=...>.
            # This simulates AD showing membership but not proving that the link is TTL-bound.
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                    member            = @(
                        'CN=alice,OU=Users,DC=corp,DC=local'
                    )
                }
            } -ParameterFilter {
                $ShowMemberTimeToLive -eq $true -and
                $Properties -contains 'member'
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            } -ParameterFilter {
                $LDAPFilter -like '*(sAMAccountName=alice)*'
            }
            Mock Get-ADOptionalFeature -ModuleName CorpAdmin {
                [pscustomobject]@{
                    EnabledScopes = @('DC=corp,DC=local')
                }
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                $script:gmCalls = ([int]$script:gmCalls) + 1
                if ($script:gmCalls -le 1) {
                    @()
                } else {
                    [pscustomobject]@{
                        SamAccountName    = 'alice'
                        DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                    }
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
            Mock Start-Sleep -ModuleName CorpAdmin { }
        }
        It 'throws when TTL membership cannot be confirmed after add' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            } | Should -Throw '*TTL post-add verification failed*'
        }
        It 'rolls back the membership when TTL validation fails' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            } | Should -Throw '*TTL post-add verification failed*'
            Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly
        }
    }
    Context 'When PAM / TTL capability is not enabled' {
        BeforeEach {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            } -ParameterFilter {
                $LDAPFilter -like '*(sAMAccountName=alice)*'
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                @()
            }
            Mock Get-ADOptionalFeature -ModuleName CorpAdmin {
                [pscustomobject]@{
                    EnabledScopes = @()
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'throws before adding when PAM feature is not enabled' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            } | Should -Throw '*PAM feature is not enabled*'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'When group does not exist' {
        BeforeEach {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                throw "Cannot find an object with identity: 'NoSuchGroup' under: 'DC=corp,DC=local'."
            }
        }
        It 'rethrows the identity-not-found error' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'NoSuchGroup' -Member 'alice'
            } | Should -Throw '*Cannot find an object with identity*'
        }
    }
    Context 'When member does not exist' {
        BeforeEach {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            # Current implementation resolves the member using Get-ADObject -LDAPFilter.
            # No object returned means the member does not exist.
            Mock Get-ADObject -ModuleName CorpAdmin {
                $null
            } -ParameterFilter {
                $LDAPFilter -like '*(sAMAccountName=ghostMember)*'
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'throws member-not-found when the member LDAP lookup returns no object' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'ghostMember'
            } | Should -Throw "*Member 'ghostMember' not found*"
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'When member is already in the group' {
        BeforeEach {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            } -ParameterFilter {
                $LDAPFilter -like '*(sAMAccountName=alice)*'
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'is idempotent: no re-add and no throw when already a member' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            } | Should -Not -Throw
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
        It 'rejects TTL add when the member is already directly present' {
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            } | Should -Throw '*already a member*'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
        It 'logs END even when already a member and no TTL is requested' {
            Mock Write-LogFile -ModuleName CorpAdmin { }
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'Domain Admins'
                    DistinguishedName = 'CN=Domain Admins,OU=Groups,DC=corp,DC=local'
                }
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            }
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                [pscustomobject]@{
                    SamAccountName    = 'alice'
                    DistinguishedName = 'CN=alice,OU=Users,DC=corp,DC=local'
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq 'Add-GroupMember NoChange'
            }
            Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                $LogString -eq '=== Add-GroupMember END ==='
            }
        }
    }
    Context 'When the member is a GROUP (group-in-group nesting)' {
        BeforeEach {
            $script:gmCalls = 0
            # Target group resolution (no TTL readback).
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'SER_Mailbox_Access'
                    DistinguishedName = 'CN=SER_Mailbox_Access,OU=Groups,DC=corp,DC=local'
                }
            } -ParameterFilter {
                -not $ShowMemberTimeToLive
            }
            # TTL readback: the nested group is present with a TTL-bound link.
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Name              = 'SER_Mailbox_Access'
                    DistinguishedName = 'CN=SER_Mailbox_Access,OU=Groups,DC=corp,DC=local'
                    member            = @('<TTL=3600>,CN=Role_Reception,OU=Groups,DC=corp,DC=local')
                }
            } -ParameterFilter {
                $ShowMemberTimeToLive -eq $true -and
                $Properties -contains 'member'
            }
            # The member resolves as a GROUP via the broadened filter (matched here by cn).
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{
                    ObjectClass       = 'group'
                    DistinguishedName = 'CN=Role_Reception,OU=Groups,DC=corp,DC=local'
                }
            } -ParameterFilter {
                $LDAPFilter -like '*Role_Reception*'
            }
            Mock Get-ADOptionalFeature -ModuleName CorpAdmin {
                [pscustomobject]@{
                    EnabledScopes = @('DC=corp,DC=local')
                }
            }
            # Pre-check: not yet a member. Post-add: present as a DIRECT member (by DN).
            Mock Get-ADGroupMember -ModuleName CorpAdmin {
                $script:gmCalls = ([int]$script:gmCalls) + 1
                if ($script:gmCalls -le 1) {
                    @()
                } else {
                    [pscustomobject]@{
                        objectClass       = 'group'
                        DistinguishedName = 'CN=Role_Reception,OU=Groups,DC=corp,DC=local'
                    }
                }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
            Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
            Mock Start-Sleep -ModuleName CorpAdmin { }
        }
        It 'resolves the group member by name and adds it with a TTL (no rollback)' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'SER_Mailbox_Access' -Member 'Role_Reception' -TimeSpan 60
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
                [string]$Members -eq 'CN=Role_Reception,OU=Groups,DC=corp,DC=local' -and
                $MemberTimeToLive.TotalMinutes -eq 60
            }
            Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
        It 'rejects an ambiguous member match before adding' {
            Mock Get-ADObject -ModuleName CorpAdmin {
                @(
                    [pscustomobject]@{
                        ObjectClass       = 'user'
                        DistinguishedName = 'CN=Role_Reception,OU=Users,DC=corp,DC=local'
                    }
                    [pscustomobject]@{
                        ObjectClass       = 'group'
                        DistinguishedName = 'CN=Role_Reception,OU=Groups,DC=corp,DC=local'
                    }
                )
            } -ParameterFilter {
                $LDAPFilter -like '*Role_Reception*'
            }
            {
                Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'SER_Mailbox_Access' -Member 'Role_Reception' -TimeSpan 60
            } | Should -Throw '*ambiguous*'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
}

Describe 'Remove-GroupMember' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'rgm.log'
    }
    BeforeEach {
        Mock Get-ADGroup -ModuleName CorpAdmin { [pscustomobject]@{ Name=$Identity; DistinguishedName="CN=$Identity,OU=Groups,DC=corp,DC=local" } }
        Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
    }
    It "returns 'NoChange' and does not remove when the member is not found" {
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='NotFound'; Object=$null; Count=0 } }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'ghost' | Should -Be 'NoChange'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It "returns 'Rejected' and does not remove when the member is ambiguous" {
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='Ambiguous'; Object=$null; Count=2 } }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'dup' | Should -Be 'Rejected'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It "returns 'NoChange' when the resolved object is not a DIRECT member" {
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='Found'; Object=[pscustomobject]@{ DistinguishedName='CN=alice,OU=Users,DC=corp,DC=local' }; Count=1 } }
        Mock Get-ADGroupMember -ModuleName CorpAdmin { @() }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice' | Should -Be 'NoChange'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It "removes a direct USER member by DN and returns 'Success'" {
        $dn = 'CN=alice,OU=Users,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='Found'; Object=[pscustomobject]@{ DistinguishedName=$dn }; Count=1 } }
        $script:gmCalls = 0
        Mock Get-ADGroupMember -ModuleName CorpAdmin {
            $script:gmCalls = ([int]$script:gmCalls) + 1
            if ($script:gmCalls -le 1) { [pscustomobject]@{ DistinguishedName=$dn } } else { @() }
        }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice' | Should -Be 'Success'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter { [string]$Members -eq $dn }
    }
    It "removes a nested GROUP member by DN and returns 'Success'" {
        $dn = 'CN=Role_Reception,OU=Groups,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='Found'; Object=[pscustomobject]@{ ObjectClass='group'; DistinguishedName=$dn }; Count=1 } }
        $script:gmCalls = 0
        Mock Get-ADGroupMember -ModuleName CorpAdmin {
            $script:gmCalls = ([int]$script:gmCalls) + 1
            if ($script:gmCalls -le 1) { [pscustomobject]@{ objectClass='group'; DistinguishedName=$dn } } else { @() }
        }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'SER_Mailbox_Access' -Member 'Role_Reception' | Should -Be 'Success'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter { [string]$Members -eq $dn }
    }
    It 'throws when post-remove verification still finds the member' {
        $dn = 'CN=alice,OU=Users,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin { [pscustomobject]@{ Status='Found'; Object=[pscustomobject]@{ DistinguishedName=$dn }; Count=1 } }
        Mock Get-ADGroupMember -ModuleName CorpAdmin { [pscustomobject]@{ DistinguishedName=$dn } }
        {
            Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice'
        } | Should -Throw '*Post-remove verification failed*'
    }
    It 'throws when Remove-ADGroupMember fails' {
        $dn = 'CN=alice,OU=Users,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                Status = 'Found'
                Object = [pscustomobject]@{
                    DistinguishedName = $dn
                }
                Count = 1
            }
        }
        Mock Get-ADGroupMember -ModuleName CorpAdmin {
            [pscustomobject]@{
                DistinguishedName = $dn
            }
        }
        Mock Remove-ADGroupMember -ModuleName CorpAdmin {
            throw 'remove failed'
        }
        {
            Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice'
        } | Should -Throw '*remove failed*'
    }
    It "logs END when the member is not found and returns 'NoChange'" {
        Mock Write-LogFile -ModuleName CorpAdmin { }
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                Status = 'NotFound'
                Object = $null
                Count  = 0
            }
        }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'ghost' | Should -Be 'NoChange'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq 'Remove-GroupMember NoChange'
        }
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq '=== Remove-GroupMember END ==='
        }
    }
    It "logs END when the member is ambiguous and returns 'Rejected'" {
        Mock Write-LogFile -ModuleName CorpAdmin { }
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                Status = 'Ambiguous'
                Object = $null
                Count  = 2
            }
        }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'dup' | Should -Be 'Rejected'
        Should -Invoke Remove-ADGroupMember -ModuleName CorpAdmin -Times 0 -Exactly
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq 'Remove-GroupMember Rejected'
        }
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq '=== Remove-GroupMember END ==='
        }
    }
    It 'logs END even when Remove-ADGroupMember fails' {
        Mock Write-LogFile -ModuleName CorpAdmin { }
        $dn = 'CN=alice,OU=Users,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                Status = 'Found'
                Object = [pscustomobject]@{
                    DistinguishedName = $dn
                }
                Count = 1
            }
        }
        Mock Get-ADGroupMember -ModuleName CorpAdmin {
            [pscustomobject]@{
                DistinguishedName = $dn
            }
        }
        Mock Remove-ADGroupMember -ModuleName CorpAdmin {
            throw 'remove failed'
        }
        {
            Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice'
        } | Should -Throw '*remove failed*'
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq 'Remove-GroupMember Failed'
        }
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq '=== Remove-GroupMember END ==='
        }
    }
    It "logs END when removal succeeds and returns 'Success'" {
        Mock Write-LogFile -ModuleName CorpAdmin { }
        $dn = 'CN=alice,OU=Users,DC=corp,DC=local'
        Mock Resolve-GroupMemberObject -ModuleName CorpAdmin {
            [pscustomobject]@{
                Status = 'Found'
                Object = [pscustomobject]@{
                    DistinguishedName = $dn
                }
                Count = 1
            }
        }
        $script:gmCalls = 0
        Mock Get-ADGroupMember -ModuleName CorpAdmin {
            $script:gmCalls = ([int]$script:gmCalls) + 1
            if ($script:gmCalls -le 1) {
                [pscustomobject]@{
                    DistinguishedName = $dn
                }
            } else {
                @()
            }
        }
        Mock Remove-ADGroupMember -ModuleName CorpAdmin { }
        Remove-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'G' -Member 'alice' | Should -Be 'Success'
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq 'Remove-GroupMember Success'
        }
        Should -Invoke Write-LogFile -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $LogString -eq '=== Remove-GroupMember END ==='
        }
    }
}

Describe 'New-ADOU' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'test.log'
    }
    Context 'Happy path' {
        BeforeAll {
            Mock New-ADOrganizationalUnit -ModuleName CorpAdmin { }
        }
        It 'calls New-ADOrganizationalUnit once' {
            New-ADOU -LogFile $logFile -DCHostName 'DC1' -OUName 'TestOU' -Path 'DC=example,DC=com'
            Should -Invoke New-ADOrganizationalUnit -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'sets ProtectedFromAccidentalDeletion to true' {
            New-ADOU -LogFile $logFile -DCHostName 'DC1' -OUName 'TestOU' -Path 'DC=example,DC=com'
            Should -Invoke New-ADOrganizationalUnit -ModuleName CorpAdmin -ParameterFilter {
                $ProtectedFromAccidentalDeletion -eq $true
            }
        }
    }
    Context 'OU already exists' {
        BeforeAll {
            Mock New-ADOrganizationalUnit -ModuleName CorpAdmin {
                throw [System.Exception]::new('The name reference is already in use.')
            }
        }
        It 'does not rethrow' {
            { New-ADOU -LogFile $logFile -DCHostName 'DC1' -OUName 'TestOU' -Path 'DC=example,DC=com' } | Should -Not -Throw
        }
    }
    Context 'Other AD errors' {
        BeforeAll {
            Mock New-ADOrganizationalUnit -ModuleName CorpAdmin {
                throw [System.Exception]::new('Access denied')
            }
        }
        It 'rethrows non-already-in-use errors' {
            { New-ADOU -LogFile $logFile -DCHostName 'DC1' -OUName 'TestOU' -Path 'DC=example,DC=com' } | Should -Throw -ExpectedMessage '*Access denied*'
        }
    }
}

Describe 'New-DomainGroup' {
    BeforeAll {
        # Stub Exchange cmdlets so Pester can mock them inside the module scope.
        # The real Enable-DistributionGroup / Set-DistributionGroup come from a
        # remote Exchange PSSession, which isn't available in unit-test context.
        function global:Enable-DistributionGroup {
            param($Identity, $DomainController)
        }
        function global:Set-DistributionGroup {
            param(
                $Identity,$HiddenFromAddressListsEnabled,$RequireSenderAuthenticationEnabled,$DomainController
            )
        }
        $logFile = Join-Path $TestDrive 'test.log'
    }
    AfterAll {
        Remove-Item function:global:Enable-DistributionGroup -ErrorAction SilentlyContinue
        Remove-Item function:global:Set-DistributionGroup -ErrorAction SilentlyContinue
    }
    Context 'Happy path, O365 = N' {
        BeforeAll {
            Mock New-ADGroup -ModuleName CorpAdmin { }
            Mock Set-ADObject -ModuleName CorpAdmin { }
            Mock Enable-DistributionGroup -ModuleName CorpAdmin { }
            Mock Set-DistributionGroup -ModuleName CorpAdmin { }
        }
        It 'creates the group and protects it from deletion' {
            New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'N' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com'
            Should -Invoke New-ADGroup  -ModuleName CorpAdmin -Times 1 -Exactly
            Should -Invoke Set-ADObject -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'does not call Exchange cmdlets when O365 = N' {
            New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'N' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com'
            Should -Invoke Enable-DistributionGroup -ModuleName CorpAdmin -Times 0 -Exactly
            Should -Invoke Set-DistributionGroup    -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'Group already exists, O365 = N' {
        BeforeAll {
            Mock New-ADGroup -ModuleName CorpAdmin {
                throw [System.Exception]::new('The specified group already exists.')
            }
            Mock Set-ADObject -ModuleName CorpAdmin { }
        }
        It 'does not rethrow' {
            { New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'N' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com' } | Should -Not -Throw
        }
    }
    Context 'O365 = E enables and configures distribution group' {
        BeforeAll {
            Mock New-ADGroup    -ModuleName CorpAdmin { }
            Mock Set-ADObject   -ModuleName CorpAdmin { }
            Mock Enable-DistributionGroup -ModuleName CorpAdmin { }
            Mock Set-DistributionGroup    -ModuleName CorpAdmin { }
        }
        It 'calls Enable-DistributionGroup once' {
            New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'E' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com'
            Should -Invoke Enable-DistributionGroup -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'calls Set-DistributionGroup once with -HiddenFromAddressListsEnabled $true' {
            New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'E' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com'
            Should -Invoke Set-DistributionGroup -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter { $HiddenFromAddressListsEnabled -eq $true }
        }
    }
    Context 'O365 = E but Enable-DistributionGroup throws' {
        BeforeAll {
            Mock New-ADGroup  -ModuleName CorpAdmin { }
            Mock Set-ADObject -ModuleName CorpAdmin { }
            Mock Enable-DistributionGroup -ModuleName CorpAdmin {
                throw [System.Exception]::new('Exchange not available')
            }
            Mock Set-DistributionGroup -ModuleName CorpAdmin { }
        }
        It 'swallows the Exchange error (logs WARNING, does not rethrow)' {
            { New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'E' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com' } | Should -Not -Throw
        }
    }
    Context 'Other AD errors during creation' {
        BeforeAll {
            Mock New-ADGroup -ModuleName CorpAdmin {
                throw [System.Exception]::new('Access denied')
            }
            Mock Set-ADObject -ModuleName CorpAdmin { }
        }
        It 'rethrows non-already-exists errors' {
            { New-DomainGroup -LogFile $logFile -DCHostName 'DC1' -GroupName 'TestGroup' -GroupCategory 'Security' -GroupScope 'Universal' -O365 'N' -HiddenFromAddressListsEnabled $true -Path 'OU=Groups,DC=example,DC=com' } | Should -Throw -ExpectedMessage '*Access denied*'
        }
    }
}

Describe 'Add-GPOLink' {
    BeforeAll {
        # Stub New-GPLink (GroupPolicy module isn't in CorpAdmin's RequiredModules
        # so it's not guaranteed to be loaded in the test session).
        function global:New-GPLink {
            param($Name, $Target, $LinkEnabled, $Enforced, $Order, $Server)
        }
        $logFile = Join-Path $TestDrive 'test.log'
    }
    AfterAll {
        Remove-Item function:global:New-GPLink -ErrorAction SilentlyContinue
    }
    Context 'Happy path' {
        BeforeAll {
            Mock New-GPLink -ModuleName CorpAdmin { }
        }
        It 'calls New-GPLink once' {
            Add-GPOLink -LogFile $logFile -DCHostName 'DC1' -GPOName 'TestPolicy' -GPOTarget 'OU=Computers,DC=example,DC=com'
            Should -Invoke New-GPLink -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'passes -LinkEnabled Yes and -Enforced yes' {
            Add-GPOLink -LogFile $logFile -DCHostName 'DC1' -GPOName 'TestPolicy' -GPOTarget 'OU=Computers,DC=example,DC=com'
            Should -Invoke New-GPLink -ModuleName CorpAdmin -ParameterFilter {
                $LinkEnabled -eq 'Yes' -and $Enforced -eq 'yes'
            }
        }
    }
    Context 'GPO already linked' {
        BeforeAll {
            Mock New-GPLink -ModuleName CorpAdmin {
                throw [System.Exception]::new('The GPO is already linked to this target.')
            }
        }
        It 'does not rethrow' {
            { Add-GPOLink -LogFile $logFile -DCHostName 'DC1' -GPOName 'TestPolicy' -GPOTarget 'OU=Computers,DC=example,DC=com' } | Should -Not -Throw
        }
    }
    Context 'Other errors' {
        BeforeAll {
            Mock New-GPLink -ModuleName CorpAdmin {
                throw [System.Exception]::new('Access denied')
            }
        }
        It 'rethrows non-already-linked errors' {
            { Add-GPOLink -LogFile $logFile -DCHostName 'DC1' -GPOName 'TestPolicy' -GPOTarget 'OU=Computers,DC=example,DC=com' } | Should -Throw -ExpectedMessage '*Access denied*'
        }
    }
}

Describe 'ConvertTo-IntOrDefault' {
    Context 'Blank and whitespace input' {
        It 'returns 0 for an empty string' {
            ConvertTo-IntOrDefault '' | Should -Be 0
        }
        It 'returns 0 for $null' {
            ConvertTo-IntOrDefault $null | Should -Be 0
        }
        It 'returns 0 for whitespace only' {
            ConvertTo-IntOrDefault '   ' | Should -Be 0
        }
    }
    Context 'Valid numeric input' {
        It 'parses a plain integer string' {
            ConvertTo-IntOrDefault '30' | Should -Be 30
        }
        It 'trims surrounding whitespace before parsing' {
            ConvertTo-IntOrDefault ' 30 ' | Should -Be 30
        }
        It 'parses a negative integer' {
            ConvertTo-IntOrDefault '-5' | Should -Be -5
        }
        It 'parses zero as zero (not the default-by-coincidence)' {
            # Pin that an explicit "0" is parsed, not just falling through
            # to the default - they happen to coincide, but a custom default
            # proves it is the parse path, not the fallback.
            ConvertTo-IntOrDefault '0' -Default 99 | Should -Be 0
        }
    }
    Context 'Non-numeric junk' {
        It 'returns the default for non-numeric text' {
            ConvertTo-IntOrDefault 'thirty' | Should -Be 0
        }
        It 'returns the default for a number with trailing junk' {
            ConvertTo-IntOrDefault '30abc' | Should -Be 0
        }
        It 'returns the default for a decimal (TryParse [int] rejects it)' {
            # [int]::TryParse does not accept "4.5" - it is not a whole number.
            ConvertTo-IntOrDefault '4.5' | Should -Be 0
        }
    }
    Context 'Custom default' {
        It 'honours a custom -Default on blank input' {
            ConvertTo-IntOrDefault '' -Default 7 | Should -Be 7
        }
        It 'honours a custom -Default on junk input' {
            ConvertTo-IntOrDefault 'nope' -Default 7 | Should -Be 7
        }
        It 'ignores the default when input is valid' {
            ConvertTo-IntOrDefault '12' -Default 7 | Should -Be 12
        }
    }
    Context 'Return type' {
        It 'returns an [int], not a string' {
            (ConvertTo-IntOrDefault '30') | Should -BeOfType [int]
        }
    }
}

# ====================================================================
# Invoke-ADSync runs a remote AAD Connect sync over a PSSession using
# ADSync cmdlets that exist only on the Connect server, so it can't be
# behaviourally exercised on a runner. Pin its contract at the source/AST
# level: mandatory params, a Delta sync, and a finally that always tears
# the session down so an error can't leak a remote session.
# ====================================================================
Describe 'Invoke-ADSync source contract' {
    BeforeAll {
        $modPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($modPath, [ref]$t, [ref]$e)
        $script:fn = $ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-ADSync'
        }, $true) | Select-Object -First 1
        if (-not $script:fn) { throw 'Invoke-ADSync not found in CorpAdmin.psm1' }
        $script:try = $script:fn.Body.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Select-Object -First 1
    }
    It 'declares LogFile, Cred, AzureADConnect and O365EmailSuffix as mandatory' {
        foreach ($p in 'LogFile','Cred','AzureADConnect','O365EmailSuffix') {
            $pAst = $script:fn.Body.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $p }
            $pAst | Should -Not -BeNullOrEmpty -Because "-$p must be declared"
            ($pAst.Attributes | Where-Object {
                $_.TypeName.Name -eq 'Parameter' -and $_.NamedArguments.ArgumentName -contains 'Mandatory'
            }) | Should -Not -BeNullOrEmpty -Because "-$p must be Mandatory"
        }
    }
    It 'initialises ADConnectSession before try so finally is StrictMode-safe' {
        $script:fn.Body.Extent.Text | Should -Match '\$ADConnectSession\s*=\s*\$null'
    }
    It 'runs a Delta sync cycle' {
        $script:fn.Body.Extent.Text | Should -Match 'Start-ADSyncSyncCycle\s+-PolicyType\s+Delta'
    }
    It 'wraps the work in try/catch/finally' {
        $script:try | Should -Not -BeNullOrEmpty
        $script:try.CatchClauses.Count | Should -BeGreaterThan 0
        $script:try.Finally | Should -Not -BeNullOrEmpty -Because 'the session must be torn down in finally'
    }
    It 'removes the PSSession in finally (no leaked remote session on error)' {
        $script:try.Finally.Extent.Text | Should -Match 'Remove-PSSession'
    }
    It 'logs sync failures rather than rethrowing' {
        foreach ($catch in $script:try.CatchClauses) {
            ($catch.Body.FindAll({
                param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
            }, $true)) | Should -BeNullOrEmpty -Because 'a sync failure is logged, not fatal'
        }
    }
}

# ====================================================================
# Send-NotificationEmail is a thin transport over Send-MailKitMessage, which
# (with MimeKit) isn't present on a CI runner - so, like Invoke-ADSync, its
# contract is pinned at the source/AST level rather than mock-executed.
# Module parse-clean coverage lives in the manifest integrity tests above.
# ====================================================================
Describe 'Send-NotificationEmail source contract' {
    BeforeAll {
        $modPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $modPath,
            [ref]$tokens,
            [ref]$errors
        )
        $script:fn = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Send-NotificationEmail'
        }, $true) | Select-Object -First 1
        if (-not $script:fn) {
            throw 'Send-NotificationEmail not found in CorpAdmin.psm1'
        }
        $script:body = $script:fn.Body.Extent.Text
        $script:try = $script:fn.Body.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Select-Object -First 1
    }
    It 'declares LogFile, SMTPServer, EmailTo, EmailFrom, EmailSubject and EmailBody, all mandatory' {
        foreach ($p in 'LogFile','SMTPServer','EmailTo','EmailFrom','EmailSubject','EmailBody') {
            $pAst = $script:fn.Body.ParamBlock.Parameters | Where-Object {
                $_.Name.VariablePath.UserPath -eq $p
            }
            $pAst | Should -Not -BeNullOrEmpty -Because "-$p must be declared"
            ($pAst.Attributes | Where-Object {
                $_.TypeName.Name -eq 'Parameter' -and $_.NamedArguments.ArgumentName -contains 'Mandatory'
            }) | Should -Not -BeNullOrEmpty -Because "-$p must be Mandatory"
        }
    }
    It 'types EmailTo as [string] (callers pass an address; the function builds the list)' {
        $emailTo = $script:fn.Body.ParamBlock.Parameters | Where-Object {
            $_.Name.VariablePath.UserPath -eq 'EmailTo'
        }
        $emailTo.StaticType.Name | Should -Be 'String'
    }
    It 'builds its own recipient list from EmailTo' {
        $script:body | Should -Match '\[MimeKit\.InternetAddressList\]::new\(\)'
        $script:body | Should -Match '\[MimeKit\.InternetAddress\]\$EmailTo'
    }
    It 'splats the expected fields to Send-MailKitMessage' {
        foreach ($key in 'RecipientList','From','Body','Subject','SmtpServer','UseSecureConnectionIfAvailable') {
            $script:body | Should -Match "$key\s*="
        }
        $script:body | Should -Match 'Send-MailKitMessage\s+@Splat'
    }
    It 'guards the send in try/catch and logs on both paths' {
        $script:try | Should -Not -BeNullOrEmpty
        $script:try.Body.Extent.Text | Should -Match 'Send-MailKitMessage'
        $script:try.Body.Extent.Text | Should -Match 'Write-LogFile'
        $script:try.CatchClauses[0].Body.Extent.Text | Should -Match 'Write-LogFile'
    }
    It 'does not rethrow a send failure (a failed notification must not abort user creation)' {
        ($script:try.CatchClauses[0].Body.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.ThrowStatementAst]
        }, $true)) | Should -BeNullOrEmpty
    }
}
