#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Coverage for the four mailbox functions in CorpAdmin.psm1 - collectively
# ~40% of the module's lines and previously untested.
#
# Function shape:
#   New-UserMailbox          calls Enable-RemoteMailbox (hybrid stub)
#   New-UserOnPremMailbox    calls Enable-Mailbox       (on-prem)
#   Update-UserMailbox       calls Update-MgUser + Set-Mailbox* cloud cmdlets
#   Update-UserOnPremMailbox calls on-prem Set-Mailbox* with -DomainController
#
# The mode parameter on all four is -SharedEquipmentRoom (S/E/R/empty).
# O365 mode E/H/N lives in the caller (CreateUsers.ps1) and dispatches
# between New-UserMailbox (H) and New-UserOnPremMailbox (E).
#
# Stubs for Exchange and Graph cmdlets live in BeforeAll/AfterAll so the
# module imports cleanly on a runner without ExchangeOnlineManagement or
# Microsoft.Graph installed. Per-Context mocks are applied with
# -ModuleName CorpAdmin, matching the existing CorpAdmin_Tests.ps1 style.
#
# Private helpers covered via InModuleScope:
#   Confirm-MailboxType
#   Grant-MailboxAccess
#   Grant-MailboxAccessOnPrem

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psd1'
    if (-not (Test-Path $modulePath)) {
        throw "CorpAdmin module not found at: $modulePath"
    }
    Import-Module $modulePath -Force
    # Stubs for Exchange / Graph cmdlets so Pester can mock them in module scope.
    function global:Enable-RemoteMailbox {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $PrimarySmtpAddress, $alias, $DomainController, $remoteroutingaddress, [switch]$shared, [switch]$equipment, [switch]$room
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Enable remote mailbox')) {
            # Test stub only.
        }
    }
    function global:Enable-Mailbox {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $PrimarySmtpAddress, $alias, $DomainController, [switch]$shared, [switch]$equipment, [switch]$room
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Enable mailbox')) {
            # Test stub only.
        }
    }
    function global:Get-Mailbox {
        param(
            $Identity, $DomainController, $ErrorAction
        )
    }
    function global:Set-Mailbox {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $Type, $DomainController, $ResourceCapacity
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set mailbox')) {
            # Test stub only.
        }
    }
    function global:Set-MailboxSpellingConfiguration {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $DictionaryLanguage, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set mailbox spelling configuration')) {
            # Test stub only.
        }
    }
    function global:Set-MailboxRegionalConfiguration {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $Language, $DateFormat, $TimeFormat, $TimeZone, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set mailbox regional configuration')) {
            # Test stub only.
        }
    }
    function global:Set-MailboxFolderPermission {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $User, $AccessRights, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set mailbox folder permission')) {
            # Test stub only.
        }
    }
    function global:Get-MailboxPermission {
        param(
            $Identity, $User, $DomainController, $ErrorAction
        )
    }
    function global:Add-MailboxPermission {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $User, $AccessRights, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Add mailbox permission')) {
            # Test stub only.
        }
    }
    function global:Get-RecipientPermission {
        param(
            $Identity, $Trustee, $ErrorAction
        )
    }
    function global:Add-RecipientPermission {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $Trustee, $AccessRights
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Add recipient permission')) {
            # Test stub only.
        }
    }
    function global:Get-ADPermission {
        param(
            $Identity, $User, $DomainController, $ErrorAction
        )
    }
    function global:Add-ADPermission {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $User, $AccessRights, $ExtendedRights, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Add AD permission')) {
            # Test stub only.
        }
    }
    function global:Set-CalendarProcessing {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $AutomateProcessing, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set calendar processing')) {
            # Test stub only.
        }
    }
    function global:Enable-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Enable distribution group')) {
            # Test stub only.
        }
    }
    function global:Set-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            $Identity, $HiddenFromAddressListsEnabled, $RequireSenderAuthenticationEnabled, $DomainController
        )
        if ($PSCmdlet.ShouldProcess($Identity, 'Set distribution group')) {
            # Test stub only.
        }
    }
    # Stub for Microsoft Graph cmdlet
    function global:Update-MgUser {
        param(
            $UserId, $UsageLocation, $PreferredLanguage
        )
    }
    # Stub environment.psd1 with the locale defaults the Update-* functions
    # consume. CorpAdmin's Get-EnvironmentConfig caches per-process, so we
    # point at a temp psd1 via the env var override.
    $script:logFile = Join-Path $TestDrive 'test.log'
    $logFile = $script:logFile
    $testEnvPath = Join-Path $TestDrive 'environment.psd1'
    @'
@{
    Network  = @{}
    OUs      = @{}
    Groups   = @{
        SharedAccessPrefix    = 'sh_'
        EquipmentAccessPrefix = 'eq_'
        RoomAccessPrefix      = 'ro_'
    }
    Shares   = @{}
    Locale   = @{
        Language      = 'en-GB'
        TimeZone      = 'GMT Standard Time'
        DateFormat    = 'dd/MM/yyyy'
        TimeFormat    = 'HH:mm'
        UsageLocation = 'GB'
        Dictionary    = 'EnglishUnitedKingdom'
    }
    Security = @{ PasswordLength = 20; MaxElevationMinutes = 480 }
    Azure    = @{}
    Exchange = @{}
    EntraRoles = @{}
    WSUS     = @{}
}
'@ | Set-Content $testEnvPath
    $env:CORPADMIN_ENV_PSD1 = $testEnvPath
    Get-EnvironmentConfig -Force | Out-Null
    # Silence module-internal logging so tests don't write files
    Mock Write-LogFile -ModuleName CorpAdmin { }
}

AfterAll {
    @(
        'Enable-RemoteMailbox',
        'Enable-Mailbox',
        'Get-Mailbox',
        'Set-Mailbox',
        'Set-MailboxSpellingConfiguration',
        'Set-MailboxRegionalConfiguration',
        'Set-MailboxFolderPermission',
        'Get-MailboxPermission',
        'Add-MailboxPermission',
        'Get-RecipientPermission',
        'Add-RecipientPermission',
        'Get-ADPermission',
        'Add-ADPermission',
        'Set-CalendarProcessing',
        'Enable-DistributionGroup',
        'Set-DistributionGroup',
        'Update-MgUser'
    ) | ForEach-Object {
        Remove-Item "function:global:$_" -ErrorAction SilentlyContinue
    }
    $env:CORPADMIN_ENV_PSD1 = $null
    Remove-Module CorpAdmin -ErrorAction SilentlyContinue
}

# =============================================================================
# New-UserMailbox  -  Enable-RemoteMailbox (hybrid)
# =============================================================================
Describe 'New-UserMailbox' {
    Context 'happy path - regular user (no SharedEquipmentRoom)' {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'jdoe' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'calls Enable-RemoteMailbox once' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -Times 1 -Exactly
        }
        It 'sets remoteroutingaddress to <UserName>@<O365EmailSuffix>' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $remoteroutingaddress -eq 'jdoe@tenant.onmicrosoft.com'
            }
        }
        It 'uses uppercase alias derived from username' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $alias -ceq 'JDOE'
            }
        }
        It 'passes -DomainController matching -DCHostName' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $DomainController -eq 'DC1'
            }
        }
    }
    Context 'realname provided' {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'jdoe' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'sets PrimarySmtpAddress = <realname>@<EmailSuffix>' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com' -realname 'John.Doe'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $PrimarySmtpAddress -eq 'John.Doe@example.com'
            }
        }
    }
    Context "SharedEquipmentRoom 'S' (shared)" {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'shared1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -shared' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'shared1' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com' -SharedEquipmentRoom 'S'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $shared -eq $true
            }
        }
    }
    Context "SharedEquipmentRoom 'E' (equipment)" {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'projector1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -equipment' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'projector1' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com' -SharedEquipmentRoom 'E'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $equipment -eq $true
            }
        }
    }
    Context "SharedEquipmentRoom 'R' (room)" {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'lab1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -room' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'lab1' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com' -SharedEquipmentRoom 'R'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -ParameterFilter {
                $room -eq $true
            }
        }
    }
    Context 'idempotency - mailbox already exists' {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'jdoe' } }
        }
        It 'does not call Enable-RemoteMailbox' {
            New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'error swallow' {
        BeforeAll {
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { throw 'simulated Exchange failure' }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null }
        }
        It 'does not propagate Exchange exceptions' {
            {
                New-UserMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com' -O365EmailSuffix 'tenant.onmicrosoft.com'
            } | Should -Not -Throw
        }
    }
}

# =============================================================================
# New-UserOnPremMailbox  -  Enable-Mailbox (on-prem)
# =============================================================================
Describe 'New-UserOnPremMailbox' {
    Context 'happy path - regular user' {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = $Identity } }
            Mock Enable-RemoteMailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'jdoe' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'calls Enable-Mailbox (not Enable-RemoteMailbox)' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -Times 1 -Exactly
            Should -Invoke Enable-RemoteMailbox -ModuleName CorpAdmin -Times 0 -Exactly
        }
        It 'passes -DomainController matching -DCHostName' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $DomainController -eq 'DC1'
            }
        }
        It 'uses uppercase alias derived from username' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $alias -ceq 'JDOE'
            }
        }
    }
    Context "SharedEquipmentRoom 'S' (shared)" {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'shared1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -shared' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'shared1' -EmailSuffix 'example.com' -SharedEquipmentRoom 'S'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $shared -eq $true
            }
        }
    }
    Context "SharedEquipmentRoom 'E' (equipment)" {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'projector1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -equipment' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'projector1' -EmailSuffix 'example.com' -SharedEquipmentRoom 'E'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $equipment -eq $true
            }
        }
    }
    Context "SharedEquipmentRoom 'R' (room)" {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null } -ParameterFilter { -not $DomainController }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'lab1' } } -ParameterFilter { $DomainController -eq 'DC1' }
        }
        It 'passes -room' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'lab1' -EmailSuffix 'example.com' -SharedEquipmentRoom 'R'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $room -eq $true
            }
        }
    }
    Context 'idempotency - mailbox already exists' {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { }
            Mock Get-Mailbox -ModuleName CorpAdmin { [pscustomobject]@{ Identity = 'jdoe' } }
        }
        It 'does not call Enable-Mailbox' {
            New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com'
            Should -Invoke Enable-Mailbox -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'error swallow' {
        BeforeAll {
            Mock Enable-Mailbox -ModuleName CorpAdmin { throw 'simulated Exchange failure' }
            Mock Get-Mailbox -ModuleName CorpAdmin { $null }
        }
        It 'does not propagate Exchange exceptions' {
            {
                New-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe' -EmailSuffix 'example.com'
            } | Should -Not -Throw
        }
    }
}

# =============================================================================
# Private mailbox helpers
# =============================================================================
Describe 'Confirm-MailboxType' {
    BeforeEach {
        Mock Start-Sleep -ModuleName CorpAdmin { }
    }
    It 'does not fail on null Get-Mailbox responses before eventual success' {
        $script:getMailboxCalls = 0
        Mock Get-Mailbox -ModuleName CorpAdmin {
            $script:getMailboxCalls++
            if ($script:getMailboxCalls -lt 2) {
                return $null
            }
            [pscustomobject]@{
                RecipientTypeDetails = 'RoomMailbox'
            }
        }
        InModuleScope CorpAdmin {
            {
                Confirm-MailboxType -Upn 'room1' -ExpectedType 'RoomMailbox'
            } | Should -Not -Throw
        }
        $script:getMailboxCalls | Should -BeGreaterOrEqual 2
        Should -Invoke Start-Sleep -ModuleName CorpAdmin -Times 1
    }
    It 'throws a clear error when the expected mailbox type is never reached' {
        Mock Get-Mailbox -ModuleName CorpAdmin {
            [pscustomobject]@{
                RecipientTypeDetails = 'UserMailbox'
            }
        }
        InModuleScope CorpAdmin {
            {
                Confirm-MailboxType -Upn 'room1' -ExpectedType 'RoomMailbox'
            } | Should -Throw '*did not convert to RoomMailbox*'
        }
    }
    It 'passes DomainController through for on-prem checks' {
        Mock Get-Mailbox -ModuleName CorpAdmin {
            [pscustomobject]@{
                RecipientTypeDetails = 'SharedMailbox'
            }
        }
        InModuleScope CorpAdmin {
            Confirm-MailboxType -Upn 'shared1' -ExpectedType 'SharedMailbox' -DomainController 'DC1'
        }
        Should -Invoke Get-Mailbox -ModuleName CorpAdmin -ParameterFilter {
            $Identity -eq 'shared1' -and $DomainController -eq 'DC1'
        }
    }
}

Describe 'Grant-MailboxAccess' {
    BeforeEach {
        Mock Write-LogFile -ModuleName CorpAdmin { }
    }
    It 'does not duplicate existing FullAccess or SendAs permissions' {
        Mock Get-MailboxPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                AccessRights = @('FullAccess')
                IsInherited = $false
                Deny  = $false
            }
        }
        Mock Get-RecipientPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                AccessRights = @('SendAs')
                Deny = $false
            }
        }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccess -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1'
        }
        Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -Times 0 -Exactly
        Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It 'adds FullAccess and SendAs when missing' {
        Mock Get-MailboxPermission -ModuleName CorpAdmin { $null }
        Mock Get-RecipientPermission -ModuleName CorpAdmin { $null }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccess -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1'
        }
        Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $Identity -eq 'shared1' -and $User -eq 'sh_shared1' -and $AccessRights -eq 'FullAccess'
        }
        Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $Identity -eq 'shared1' -and $Trustee -eq 'sh_shared1' -and $AccessRights -eq 'SendAs'
        }
    }
    It 'ignores inherited or deny FullAccess entries and adds a direct allow' {
        Mock Get-MailboxPermission -ModuleName CorpAdmin {
            @(
                [pscustomobject]@{
                    AccessRights = @('FullAccess')
                    IsInherited = $true
                    Deny = $false
                },
                [pscustomobject]@{
                    AccessRights = @('FullAccess')
                    IsInherited = $false
                    Deny = $true
                }
            )
        }
        Mock Get-RecipientPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                AccessRights = @('SendAs')
                Deny = $false
            }
        }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccess -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1'
        }
        Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -Times 1 -Exactly
        Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -Times 0 -Exactly
    }
}

Describe 'Grant-MailboxAccessOnPrem' {
    BeforeEach {
        Mock Write-LogFile -ModuleName CorpAdmin { }
        Mock Enable-DistributionGroup -ModuleName CorpAdmin { }
        Mock Set-DistributionGroup -ModuleName CorpAdmin { }
    }
    It 'does not duplicate existing FullAccess or Send-As permissions' {
        Mock Get-MailboxPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                AccessRights = @('FullAccess')
                IsInherited = $false
                Deny = $false
            }
        }
        Mock Get-ADPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                ExtendedRights = @('Send-As')
                Deny = $false
            }
        }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-ADPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccessOnPrem -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1' -DomainController 'DC1'
        }
        Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -Times 0 -Exactly
        Should -Invoke Add-ADPermission -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It 'treats Send As, Send-As and SendAs as existing Send-As permissions' -TestCases @(
        @{ Right = 'Send As' }
        @{ Right = 'Send-As' }
        @{ Right = 'SendAs' }
    ) {
        param($Right)
        Mock Get-MailboxPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                AccessRights = @('FullAccess')
                IsInherited = $false
                Deny = $false
            }
        }
        Mock Get-ADPermission -ModuleName CorpAdmin {
            [pscustomobject]@{
                ExtendedRights = @($Right)
                Deny = $false
            }
        }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-ADPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccessOnPrem -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1' -DomainController 'DC1'
        }
        Should -Invoke Add-ADPermission -ModuleName CorpAdmin -Times 0 -Exactly
    }
    It 'adds FullAccess and Send-As when missing' {
        Mock Get-MailboxPermission -ModuleName CorpAdmin { $null }
        Mock Get-ADPermission -ModuleName CorpAdmin { $null }
        Mock Add-MailboxPermission -ModuleName CorpAdmin { }
        Mock Add-ADPermission -ModuleName CorpAdmin { }
        InModuleScope CorpAdmin -Parameters @{ logFile = $script:logFile } {
            param($logFile)
            Grant-MailboxAccessOnPrem -LogFile $logFile -Upn 'shared1' -GroupName 'sh_shared1' -DomainController 'DC1'
        }
        Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $Identity -eq 'shared1' -and $User -eq 'sh_shared1' -and $DomainController -eq 'DC1'
        }
        Should -Invoke Add-ADPermission -ModuleName CorpAdmin -Times 1 -Exactly -ParameterFilter {
            $Identity -eq 'shared1' -and $User -eq 'sh_shared1' -and $ExtendedRights -eq 'Send As' -and $DomainController -eq 'DC1'
        }
    }
}

# =============================================================================
# Update-UserMailbox  -  cloud locale + per-type configuration
# =============================================================================
Describe 'Update-UserMailbox' {
    BeforeEach {
        Mock Start-Sleep -ModuleName CorpAdmin { }
        Mock Get-MailboxPermission -ModuleName CorpAdmin { $null }
        Mock Get-RecipientPermission -ModuleName CorpAdmin { $null }
    }
    Context 'locale defaults from environment.psd1' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'jdoe@example.com'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            Mock Update-MgUser -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
        }
        It 'sets UsageLocation = GB on the Graph user' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'jdoe@example.com'
            Should -Invoke Update-MgUser -ModuleName CorpAdmin -ParameterFilter {
                $UsageLocation -eq 'GB'
            }
        }
        It 'applies Language en-GB and TimeZone GMT Standard Time' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'jdoe@example.com'
            Should -Invoke Set-MailboxRegionalConfiguration -ModuleName CorpAdmin -ParameterFilter {
                $Language -eq 'en-GB' -and $TimeZone -eq 'GMT Standard Time'
            }
        }
        It 'applies the EnglishUnitedKingdom dictionary' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'jdoe@example.com'
            Should -Invoke Set-MailboxSpellingConfiguration -ModuleName CorpAdmin -ParameterFilter {
                $DictionaryLanguage -eq 'EnglishUnitedKingdom'
            }
        }
        It 'sets Calendar folder permission to Reviewer for Default user' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'jdoe@example.com'
            Should -Invoke Set-MailboxFolderPermission -ModuleName CorpAdmin -ParameterFilter {
                $Identity -eq 'jdoe@example.com:\Calendar' -and $AccessRights -eq 'Reviewer'
            }
        }
    }
    Context "SharedEquipmentRoom 'S' - convert to SharedMailbox" {
        BeforeAll {
            Mock Update-MgUser -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Set-Mailbox -ModuleName CorpAdmin { }
            Mock Add-MailboxPermission -ModuleName CorpAdmin { }
            Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        }
        It 'converts the mailbox type when not already shared' {
            $script:getMailboxCalls = 0
            Mock Get-Mailbox -ModuleName CorpAdmin {
                $script:getMailboxCalls++
                if ($script:getMailboxCalls -eq 1) {
                    [pscustomobject]@{
                        Identity = 'shared1'
                        RecipientTypeDetails = 'UserMailbox'
                    }
                } else {
                    [pscustomobject]@{
                        Identity = 'shared1'
                        RecipientTypeDetails = 'SharedMailbox'
                    }
                }
            }
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Set-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $Type -eq 'Shared'
            }
        }
        It 'skips type conversion when already SharedMailbox' {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'shared1'
                    RecipientTypeDetails = 'SharedMailbox'
                }
            }
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Set-Mailbox -ModuleName CorpAdmin -Times 0 -Exactly -ParameterFilter {
                $Type -eq 'Shared'
            }
        }
        It 'grants FullAccess to the sh_<user> delegation group' {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'shared1'
                    RecipientTypeDetails = 'SharedMailbox'
                }
            }
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -ParameterFilter {
                $User -eq 'sh_shared1' -and $AccessRights -eq 'FullAccess'
            }
        }
        It 'grants SendAs to the sh_<user> delegation group via Add-RecipientPermission' {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'shared1'
                    RecipientTypeDetails = 'SharedMailbox'
                }
            }
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -ParameterFilter {
                $Trustee -eq 'sh_shared1' -and $AccessRights -eq 'SendAs'
            }
        }
    }
    Context "SharedEquipmentRoom 'E' - Equipment with capacity" {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'projector1'
                    RecipientTypeDetails = 'EquipmentMailbox'
                }
            }
            Mock Update-MgUser -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Set-Mailbox -ModuleName CorpAdmin { }
            Mock Set-CalendarProcessing -ModuleName CorpAdmin { }
            Mock Add-MailboxPermission -ModuleName CorpAdmin { }
            Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        }
        It 'sets ResourceCapacity when Capacity is supplied' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'projector1' -SharedEquipmentRoom 'E' -Capacity 4
            Should -Invoke Set-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $ResourceCapacity -eq 4
            }
        }
        It 'enables AutoAccept calendar processing' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'projector1' -SharedEquipmentRoom 'E' -Capacity 4
            Should -Invoke Set-CalendarProcessing -ModuleName CorpAdmin -ParameterFilter {
                $AutomateProcessing -eq 'AutoAccept'
            }
        }
        It 'uses exact equipment calendar identity without extra spaces' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'projector1' -SharedEquipmentRoom 'E' -Capacity 4
            Should -Invoke Set-MailboxFolderPermission -ModuleName CorpAdmin -ParameterFilter {
                $Identity -eq 'projector1:\Calendar' -and $AccessRights -eq 'Author'
            }
        }
        It 'grants FullAccess and SendAs to the eq_<user> delegation group' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'projector1' -SharedEquipmentRoom 'E' -Capacity 4
            Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -ParameterFilter {
                $User -eq 'eq_projector1'
            }
            Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -ParameterFilter {
                $Trustee -eq 'eq_projector1'
            }
        }
    }
    Context "SharedEquipmentRoom 'R' - Room" {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity  = 'lab1'
                    RecipientTypeDetails = 'RoomMailbox'
                }
            }
            Mock Update-MgUser -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Set-Mailbox -ModuleName CorpAdmin { }
            Mock Set-CalendarProcessing -ModuleName CorpAdmin { }
            Mock Add-MailboxPermission -ModuleName CorpAdmin { }
            Mock Add-RecipientPermission -ModuleName CorpAdmin { }
        }
        It 'uses exact room calendar identity without extra spaces' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'lab1' -SharedEquipmentRoom 'R'
            Should -Invoke Set-MailboxFolderPermission -ModuleName CorpAdmin -ParameterFilter {
                $Identity -eq 'lab1:\Calendar' -and $AccessRights -eq 'Author'
            }
        }
        It 'grants FullAccess and SendAs to the ro_<user> delegation group' {
            Update-UserMailbox -LogFile $logFile -UserPrincipalName 'lab1' -SharedEquipmentRoom 'R'
            Should -Invoke Add-MailboxPermission -ModuleName CorpAdmin -ParameterFilter {
                $User -eq 'ro_lab1'
            }
            Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -ParameterFilter {
                $Trustee -eq 'ro_lab1'
            }
        }
    }
    Context 'mailbox not found after retries' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin { $null }
            Mock Start-Sleep -ModuleName CorpAdmin { }
            Mock Update-MgUser -ModuleName CorpAdmin { }
        }
        It 'does not throw and does not update the user' {
            {
                Update-UserMailbox -LogFile $logFile -UserPrincipalName 'ghost@example.com'
            } | Should -Not -Throw
            Should -Invoke Update-MgUser -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'error swallow' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'jdoe@example.com'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            Mock Update-MgUser -ModuleName CorpAdmin { throw 'simulated Graph failure' }
        }
        It 'does not propagate Graph exceptions' {
            {
                Update-UserMailbox -LogFile $logFile -UserPrincipalName 'jdoe@example.com'
            } | Should -Not -Throw
        }
    }
}

# =============================================================================
# Update-UserOnPremMailbox  -  on-prem locale + per-type configuration
# =============================================================================
Describe 'Update-UserOnPremMailbox' {
    BeforeEach {
        Mock Start-Sleep -ModuleName CorpAdmin { }
        Mock Get-MailboxPermission -ModuleName CorpAdmin { $null }
        Mock Get-ADPermission -ModuleName CorpAdmin { $null }
    }
    Context 'locale defaults with -DomainController' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'jdoe'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Update-MgUser -ModuleName CorpAdmin { }
        }
        It 'applies regional config with -DomainController' {
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe'
            Should -Invoke Set-MailboxRegionalConfiguration -ModuleName CorpAdmin -ParameterFilter {
                $Language -eq 'en-GB' -and $DomainController -eq 'DC1'
            }
        }
        It 'applies spelling config with -DomainController' {
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe'
            Should -Invoke Set-MailboxSpellingConfiguration -ModuleName CorpAdmin -ParameterFilter {
                $DictionaryLanguage -eq 'EnglishUnitedKingdom' -and $DomainController -eq 'DC1'
            }
        }
        It 'does NOT call Update-MgUser (on-prem-only path)' {
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe'
            Should -Invoke Update-MgUser -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context "SharedEquipmentRoom 'S' - convert and enable distribution group" {
        BeforeAll {
            Mock Set-Mailbox -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Add-MailboxPermission -ModuleName CorpAdmin { }
            Mock Add-ADPermission -ModuleName CorpAdmin { }
            Mock Add-RecipientPermission -ModuleName CorpAdmin { }
            Mock Enable-DistributionGroup -ModuleName CorpAdmin { }
            Mock Set-DistributionGroup -ModuleName CorpAdmin { }
        }
        It 'converts the mailbox type when not already shared' {
            $script:getMailboxCalls = 0
            Mock Get-Mailbox -ModuleName CorpAdmin {
                $script:getMailboxCalls++
                if ($script:getMailboxCalls -eq 1) {
                    [pscustomobject]@{
                        Identity = 'shared1'
                        RecipientTypeDetails = 'UserMailbox'
                    }
                } else {
                    [pscustomobject]@{
                        Identity = 'shared1'
                        RecipientTypeDetails = 'SharedMailbox'
                    }
                }
            }
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Set-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $Type -eq 'Shared' -and $DomainController -eq 'DC1'
            }
        }
        It 'enables and hides the sh_<user> distribution group' {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'shared1'
                    RecipientTypeDetails = 'SharedMailbox'
                }
            }
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Enable-DistributionGroup -ModuleName CorpAdmin -ParameterFilter {
                $Identity -eq 'sh_shared1' -and $DomainController -eq 'DC1'
            }
            Should -Invoke Set-DistributionGroup -ModuleName CorpAdmin -ParameterFilter {
                $HiddenFromAddressListsEnabled -eq $true -and $DomainController -eq 'DC1'
            }
        }
        It "grants 'Send As' via Add-ADPermission, not Add-RecipientPermission" {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'shared1'
                    RecipientTypeDetails = 'SharedMailbox'
                }
            }
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'shared1' -SharedEquipmentRoom 'S'
            Should -Invoke Add-ADPermission -ModuleName CorpAdmin -ParameterFilter {
                $ExtendedRights -eq 'Send As' -and $DomainController -eq 'DC1'
            }
            Should -Invoke Add-RecipientPermission -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context "SharedEquipmentRoom 'R' - Room with capacity" {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity  = 'lab1'
                    RecipientTypeDetails = 'RoomMailbox'
                }
            }
            Mock Set-Mailbox -ModuleName CorpAdmin { }
            Mock Set-MailboxSpellingConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
            Mock Set-MailboxFolderPermission -ModuleName CorpAdmin { }
            Mock Set-CalendarProcessing -ModuleName CorpAdmin { }
            Mock Add-MailboxPermission -ModuleName CorpAdmin { }
            Mock Add-ADPermission -ModuleName CorpAdmin { }
            Mock Enable-DistributionGroup -ModuleName CorpAdmin { }
            Mock Set-DistributionGroup -ModuleName CorpAdmin { }
        }
        It 'sets ResourceCapacity and enables AutoAccept' {
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'lab1' -SharedEquipmentRoom 'R' -Capacity 30
            Should -Invoke Set-Mailbox -ModuleName CorpAdmin -ParameterFilter {
                $ResourceCapacity -eq 30 -and $DomainController -eq 'DC1'
            }
            Should -Invoke Set-CalendarProcessing -ModuleName CorpAdmin -ParameterFilter {
                $AutomateProcessing -eq 'AutoAccept' -and $DomainController -eq 'DC1'
            }
        }
        It 'uses exact room calendar identity without extra spaces' {
            Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'lab1' -SharedEquipmentRoom 'R'
            Should -Invoke Set-MailboxFolderPermission -ModuleName CorpAdmin -ParameterFilter {
                $Identity -eq 'lab1:\Calendar' -and $AccessRights -eq 'Author' -and $DomainController -eq 'DC1'
            }
        }
    }
    Context 'mailbox not found after retries' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin { $null }
            Mock Start-Sleep -ModuleName CorpAdmin { }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { }
        }
        It 'does not throw and does not apply regional config' {
            {
                Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'ghost'
            } | Should -Not -Throw
            Should -Invoke Set-MailboxRegionalConfiguration -ModuleName CorpAdmin -Times 0 -Exactly
        }
    }
    Context 'error swallow' {
        BeforeAll {
            Mock Get-Mailbox -ModuleName CorpAdmin {
                [pscustomobject]@{
                    Identity = 'jdoe'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            Mock Set-MailboxRegionalConfiguration -ModuleName CorpAdmin { throw 'simulated Exchange failure' }
        }
        It 'does not propagate Exchange exceptions' {
            {
                Update-UserOnPremMailbox -LogFile $logFile -DCHostName 'DC1' -UserName 'jdoe'
            } | Should -Not -Throw
        }
    }
}

# =============================================================================
# Cross-function module-source guards
# =============================================================================
Describe 'CorpAdmin.psm1 mailbox functions: source-level guards' {
    BeforeAll {
        $psm = Join-Path $PSScriptRoot '..\Modules\CorpAdmin\CorpAdmin.psm1'
        $script:moduleSource = Get-Content $psm -Raw
    }
    It 'does not hardcode en-US anywhere in the module' {
        $script:moduleSource | Should -Not -Match "'en-US'"
    }
    It 'does not hardcode Pacific Standard Time' {
        $script:moduleSource | Should -Not -Match "'Pacific Standard Time'"
    }
    It "does not hardcode UsageLocation = 'US'" {
        $script:moduleSource | Should -Not -Match "UsageLocation\s*=\s*'US'"
    }
    It 'does not hardcode EnglishUnitedStates' {
        $script:moduleSource | Should -Not -Match "'EnglishUnitedStates'"
    }
}
