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
            $second = Get-EnvironmentConfig          # returns from cache
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
            (Get-EnvironmentConfig).Marker | Should -Be 'A'   # no -Force: populates cache keyed on envA
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
           Azure = @{}; Exchange = @{}; EntraRoles = @{} }" |   # no WSUS
            Set-Content $bad
        { Get-EnvironmentConfig -Path $bad -Force } |
            Should -Throw -ExpectedMessage '*missing required section*WSUS*'
    }
    It 'names every missing section, not just the first' {
        $bad = Join-Path $TestDrive 'env-missing-two.psd1'
        "@{ Network = @{}; OUs = @{}; Groups = @{}; Shares = @{}; Locale = @{};
           Security = @{}; Azure = @{}; Exchange = @{} }" |     # no EntraRoles, no WSUS
            Set-Content $bad
        { Get-EnvironmentConfig -Path $bad -Force } |
            Should -Throw -ExpectedMessage '*EntraRoles, WSUS*'
    }
    It 'accepts a config that has every required section' {
        { Get-EnvironmentConfig -Path $testEnvPath -Force } | Should -Not -Throw
    }
}

Describe 'CorpAdmin.psd1 export manifest integrity' {
    BeforeAll {
        $manifestPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1'
        $modPath      = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $script:exportedFns = @((Import-PowerShellDataFile -LiteralPath $manifestPath).FunctionsToExport)
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($modPath, [ref]$t, [ref]$e)
        $script:definedFns = @($ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true).Name)
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
            'Add-GroupMember','New-ADOU','New-DomainGroup','Add-GPOLink','Invoke-ADSync',
            'New-UserMailbox','Update-UserMailbox','New-UserOnPremMailbox','Update-UserOnPremMailbox',
            'Send-NotificationEmail','Get-ADSchemaGuidMap','Get-ADExtendedRightsMap',
            'Grant-ComputerJoinDelegation','Grant-GroupDelegation','Grant-GroupMembershipEditDelegation',
            'Grant-PasswordResetDelegation','Grant-UserDelegation','Grant-OUDelegation',
            'Grant-DNSOperatorsPermissionDelegation','Grant-DNSReadOnlyPermissionDelegation',
            'Grant-ADObjectPermissionDelegation','Grant-GPOPermissionDelegation','Grant-GPOCreationDelegation'
        ) | ForEach-Object { @{ Function = $_ } }
    ) {
        param($Function)
        $script:exportedFns | Should -Contain $Function
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

Describe 'Add-GroupMember' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'test.log'
    }
    Context 'When group exists and member exists' {
        BeforeAll {
            # Mocks live in the CorpAdmin module's scope, not the test scope.
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{ Name = 'Domain Admins' }
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{ SamAccountName = 'alice' }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin { }
        }
        It 'calls Add-ADGroupMember once' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice'
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'passes -MemberTimeToLive when -TimeSpan is supplied' {
            Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' -TimeSpan 60
            Should -Invoke Add-ADGroupMember -ModuleName CorpAdmin -Times 1 -ParameterFilter {$null -ne  $MemberTimeToLive}
        }
    }
    Context 'When group does not exist' {
        BeforeAll {
            Mock Get-ADGroup -ModuleName CorpAdmin { $null }
        }
        It 'throws' {
            { Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'NoSuchGroup' -Member 'alice' } | Should -Throw "*does not exist*"
        }
    }
    Context 'When member does not exist' {
        BeforeAll {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{ Name = 'Domain Admins' }
            }
            Mock Get-ADObject -ModuleName CorpAdmin { $null }
        }
        It 'throws' {
            { Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'ghostMember' } | Should -Throw "*does not exist*"
        }
    }
    Context 'When member is already in the group' {
        BeforeAll {
            Mock Get-ADGroup -ModuleName CorpAdmin {
                [pscustomobject]@{ Name = 'Domain Admins' }
            }
            Mock Get-ADObject -ModuleName CorpAdmin {
                [pscustomobject]@{ SamAccountName = 'alice' }
            }
            Mock Add-ADGroupMember -ModuleName CorpAdmin {
                throw [System.Exception]::new('The specified account name is already a member of the group.')
            }
        }
        It 'does not rethrow' {
            { Add-GroupMember -LogFile $logFile -DCHostName 'DC1' -Group 'Domain Admins' -Member 'alice' } | Should -Not -Throw
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
            param($Identity, $HiddenFromAddressListsEnabled,
                  $RequireSenderAuthenticationEnabled, $DomainController)
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
# contract is pinned at the source/AST level rather than mock-executed. The
# parse-clean assertion also guards the whole module against a stray syntax
# error.
# ====================================================================
Describe 'Send-NotificationEmail source contract' {
    BeforeAll {
        $modPath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $t = $null; $script:parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($modPath, [ref]$t, [ref]$script:parseErrors)
        $script:fn = $ast.FindAll({
            param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                      $n.Name -eq 'Send-NotificationEmail'
        }, $true) | Select-Object -First 1
        if (-not $script:fn) { throw 'Send-NotificationEmail not found in CorpAdmin.psm1' }
        $script:body = $script:fn.Body.Extent.Text
        $script:try  = $script:fn.Body.FindAll({
            param($n) $n -is [System.Management.Automation.Language.TryStatementAst]
        }, $true) | Select-Object -First 1
    }
    It 'CorpAdmin.psm1 parses with no syntax errors' {
        $script:parseErrors | Should -BeNullOrEmpty -Because 'a malformed attribute breaks Import-Module for every consumer'
    }
    It 'declares LogFile, SMTPServer, EmailTo, EmailFrom, EmailSubject and EmailBody, all mandatory' {
        foreach ($p in 'LogFile','SMTPServer','EmailTo','EmailFrom','EmailSubject','EmailBody') {
            $pAst = $script:fn.Body.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $p }
            $pAst | Should -Not -BeNullOrEmpty -Because "-$p must be declared"
            ($pAst.Attributes | Where-Object {
                $_.TypeName.Name -eq 'Parameter' -and $_.NamedArguments.ArgumentName -contains 'Mandatory'
            }) | Should -Not -BeNullOrEmpty -Because "-$p must be Mandatory"
        }
    }
    It 'types EmailTo as [string] (callers pass an address; the function builds the list)' {
        $emailTo = $script:fn.Body.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'EmailTo' }
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
        $script:try.Body.Extent.Text                 | Should -Match 'Send-MailKitMessage'
        $script:try.Body.Extent.Text                 | Should -Match 'Write-LogFile'
        $script:try.CatchClauses[0].Body.Extent.Text | Should -Match 'Write-LogFile'
    }
    It 'does not rethrow a send failure (a failed notification must not abort user creation)' {
        ($script:try.CatchClauses[0].Body.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ThrowStatementAst]
        }, $true)) | Should -BeNullOrEmpty
    }
}
