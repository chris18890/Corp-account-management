# AST/body coverage for the optional-role bootstrap scripts (dcpromo, the
# Additional\ server scripts, and WDS), in the same style as Setup.Tests.ps1.

BeforeAll {
    $script:scriptsRoot = Split-Path $PSScriptRoot -Parent   # Tests -> Scripts
    function Get-Src { param($Rel) Get-Content (Join-Path $script:scriptsRoot $Rel) -Raw }
    function Get-Ast {
        param($Rel)
        $t = $null; $e = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $script:scriptsRoot $Rel), [ref]$t, [ref]$e)
    }
}

Describe 'Tier-0 standalone infra script: <Rel>' -ForEach @(
    @{ Rel = 'Machine1\dcpromo.ps1' }
    @{ Rel = 'Additional\additional-dcpromo.ps1' }
    @{ Rel = 'Additional\setup-RootCA.ps1' }
    @{ Rel = 'WDS\wds.ps1' }
) {
    BeforeAll { $script:src = Get-Src $Rel; $script:ast = Get-Ast $Rel }
    It '<Rel>: parses' { $script:ast | Should -Not -BeNullOrEmpty }
    It '<Rel>: is RunAsAdministrator + StrictMode' {
        $script:src | Should -Match '#Requires\s+-RunAsAdministrator'
        $script:src | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It '<Rel>: imports no shared modules (runs pre-domain / pre-module)' {
        $script:src | Should -Not -Match 'Import-Module.+CorpAdmin'
        $script:src | Should -Not -Match 'Get-EnvironmentConfig'
    }
}

Describe 'Module-backed infra script: <Rel>' -ForEach @(
    @{ Rel = 'Additional\setup-InterCA.ps1' }
    @{ Rel = 'Additional\setup-wsus.ps1' }
    @{ Rel = 'Additional\setup-exch-PreReqs.ps1' }
    @{ Rel = 'Additional\Setup-Exch-PostInstall.ps1' }
) {
    BeforeAll { $script:src = Get-Src $Rel; $script:ast = Get-Ast $Rel }
    It '<Rel>: parses' { $script:ast | Should -Not -BeNullOrEmpty }
    It '<Rel>: is RunAsAdministrator + StrictMode' {
        $script:src | Should -Match '#Requires\s+-RunAsAdministrator'
        $script:src | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It '<Rel>: loads config via the CorpAdmin module' {
        $script:src | Should -Match 'Import-Module.+CorpAdmin'
        $script:src | Should -Match 'Get-EnvironmentConfig'
    }
}

Describe 'dcpromo.ps1 promotes the FOREST ROOT' {
    BeforeAll { $script:src = Get-Src 'Machine1\dcpromo.ps1' }
    It 'uses Install-ADDSForest'                  { $script:src | Should -Match 'Install-ADDSForest' }
    It 'does NOT use Install-ADDSDomainController' { $script:src | Should -Not -Match 'Install-ADDSDomainController' }
    It 'installs the AD DS role'                  { $script:src | Should -Match 'Install-WindowsFeature[^\r\n]*AD-Domain-Services' }
    It 'sets WinThreshold forest mode'            { $script:src | Should -Match 'ForestMode\s+"WinThreshold"' }
}

Describe 'additional-dcpromo.ps1 promotes a REPLICA DC' {
    BeforeAll { $script:src = Get-Src 'Additional\additional-dcpromo.ps1' }
    It 'uses Install-ADDSDomainController' { $script:src | Should -Match 'Install-ADDSDomainController' }
    It 'does NOT use Install-ADDSForest'   { $script:src | Should -Not -Match 'Install-ADDSForest' }
    It 'joins the domain first if not a member' { $script:src | Should -Match 'Add-Computer\s+-DomainName' }
    It 'gates on Domain Admins and throws if unauthorised' {
        $script:src | Should -Match "'Domain Admins'"
        $script:src | Should -Match 'Where-Object\s+SamAccountName\s+-eq\s+\$env:USERNAME'
        $script:src | Should -Match 'throw'
    }
    It 'targets the PDC emulator' { $script:src | Should -Match 'PDCEmulator' }
}

Describe 'setup-RootCA.ps1 installs a STANDALONE ROOT CA' {
    BeforeAll { $script:src = Get-Src 'Additional\setup-RootCA.ps1' }
    It 'installs CAType StandaloneRootCa' { $script:src | Should -Match 'Install-AdcsCertificationAuthority[^\r\n]*-CAType\s+StandaloneRootCa' }
    It 'is NOT a subordinate/enterprise CA' { $script:src | Should -Not -Match 'EnterpriseSubordinateCa' }
    It 'requires a DomainSuffix' { $script:src | Should -Match 'throw\s+"DomainSuffix' }
    It 'guards every certutil call with an exit-code check' {
        $certutil = ([regex]::Matches($script:src, 'certutil\.exe')).Count
        $guards   = ([regex]::Matches($script:src, 'LASTEXITCODE\s+-ne\s+0')).Count
        $guards | Should -BeGreaterOrEqual $certutil
    }
}

Describe 'setup-InterCA.ps1 installs an ENTERPRISE SUBORDINATE CA' {
    BeforeAll { $script:src = Get-Src 'Additional\setup-InterCA.ps1' }
    It 'installs CAType EnterpriseSubordinateCa' { $script:src | Should -Match 'Install-AdcsCertificationAuthority[^\r\n]*-CAType\s+EnterpriseSubordinateCa' }
    It 'is NOT a standalone root CA' { $script:src | Should -Not -Match 'StandaloneRootCa' }
    It 'gates on Enterprise Admins and throws if unauthorised' {
        $script:src | Should -Match "'Enterprise Admins'"
        $script:src | Should -Match 'Test-IsMemberOf[^\r\n]*-Sam\s+\$env:USERNAME'
        $script:src | Should -Match 'throw'
    }
    It 'guards every certutil call with an exit-code check' {
        $certutil = ([regex]::Matches($script:src, 'certutil\.exe')).Count
        $guards   = ([regex]::Matches($script:src, 'LASTEXITCODE\s+-ne\s+0')).Count
        $guards | Should -BeGreaterOrEqual $certutil
    }
}

Describe 'setup-wsus.ps1' {
    BeforeAll { $script:src = Get-Src 'Additional\setup-wsus.ps1' }
    It 'loads the WSUS admin assembly' { $script:src | Should -Match 'Add-Type\s+-AssemblyName\s+Microsoft\.UpdateServices\.Administration' }
    It 'installs the UpdateServices feature' { $script:src | Should -Match '"UpdateServices"' }
    It 'takes the share name from config' { $script:src | Should -Match '\$Env\.Shares\.WSUS' }
    It 'installs features idempotently' { $script:src | Should -Match 'Get-WindowsFeature[^\r\n]*InstallState\s+-Eq\s+Installed' }
    It 'domain-joins before configuring (two-phase)' {
        $script:src | Should -Match 'partofdomain\s+-eq\s+\$false'
        $script:src | Should -Match 'Add-Computer\s+-DomainName'
    }
}

Describe 'wds.ps1' {
    BeforeAll { $script:src = Get-Src 'WDS\wds.ps1' }
    It 'installs the WDS feature' { $script:src | Should -Match 'Install-WindowsFeature\s+-name\s+WDS' }
    It 'initialises and authorises the server via WDSUTIL' {
        $script:src | Should -Match 'WDSUTIL\s+/initialize-server'
        $script:src | Should -Match '/Authorize'
    }
    It 'guards every WDSUTIL call with an exit-code check' {
        $wdsutil = ([regex]::Matches($script:src, 'WDSUTIL\s+/')).Count
        $guards  = ([regex]::Matches($script:src, 'LASTEXITCODE\s+-ne\s+0')).Count
        $guards | Should -BeGreaterOrEqual $wdsutil
    }
    It 'creates the "Servers" install-image group' { $script:src | Should -Match 'New-WdsInstallImageGroup\s+-Name\s+"Servers"' }
}

Describe 'WDS image setup is consistent with wds.ps1' {
    It 'imports boot + install images into the "Servers" group wds.ps1 creates' {
        $img = Get-Src 'WDS\Server2022ImageSetup.ps1'
        $img | Should -Match 'Import-WdsBootImage'
        $img | Should -Match 'Import-WdsInstallImage'
        $img | Should -Match '-ImageGroup\s+"Servers"'
    }
}

Describe 'setup-exch-PreReqs.ps1' {
    BeforeAll { $script:src = Get-Src 'Additional\setup-exch-PreReqs.ps1' }
    It 'gates on Domain/Enterprise/Schema Admins and throws if unauthorised' {
        $script:src | Should -Match "@\('Domain Admins',\s*'Enterprise Admins',\s*'Schema Admins'\)"
        $script:src | Should -Match 'Test-IsMemberOf[^\r\n]*-Sam\s+\$env:USERNAME'
        $script:src | Should -Match 'throw'
    }
    It 'time-boxes the temporary Enterprise + Schema Admins elevation via MaxElevationMinutes' {
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Group\s+"Enterprise Admins"[^\r\n]*-TimeSpan\s+\$Env\.Security\.MaxElevationMinutes'
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Group\s+"Schema Admins"[^\r\n]*-TimeSpan\s+\$Env\.Security\.MaxElevationMinutes'
    }
    It 'runs Exchange /PrepareAD guarded by an exit-code check' {
        $script:src | Should -Match 'Setup\.exe[^\r\n]*/PrepareAD'
        $script:src | Should -Match 'IAcceptExchangeServerLicenseTerms'
        $script:src | Should -Match 'LASTEXITCODE\s+-ne\s+0'
    }
    It 'guarantees de-elevation: removes Enterprise + Schema Admins inside a finally' {
        # The most important property - a leaked Schema/Enterprise Admin is severe.
        $script:src | Should -Match 'Remove-ADGroupMember\s+-Identity\s+"Enterprise Admins"'
        $script:src | Should -Match 'Remove-ADGroupMember\s+-Identity\s+"Schema Admins"'
        $finallyIdx = $script:src.IndexOf('finally')
        $entIdx     = $script:src.IndexOf('Remove-ADGroupMember -Identity "Enterprise Admins"')
        $schIdx     = $script:src.IndexOf('Remove-ADGroupMember -Identity "Schema Admins"')
        $finallyIdx | Should -BeGreaterThan -1
        $entIdx     | Should -BeGreaterThan $finallyIdx
        $schIdx     | Should -BeGreaterThan $finallyIdx
    }
    It 'tolerates re-runs (handles the "not a member" ADException on removal)' {
        $script:src | Should -Match 'catch\s+\[Microsoft\.ActiveDirectory\.Management\.ADException\]'
        $script:src | Should -Match 'not a member'
    }
}

Describe 'Setup-Exch-PostInstall.ps1' {
    BeforeAll { $script:src = Get-Src 'Additional\Setup-Exch-PostInstall.ps1' }
    It 'requires the ActiveDirectory module' {
        $script:src | Should -Match '#Requires\s+-Modules\s+ActiveDirectory'
    }
    It 'gates on Domain Admins and throws if unauthorised' {
        $script:src | Should -Match "@\('Domain Admins'\)"
        $script:src | Should -Match 'Test-IsMemberOf[^\r\n]*-Sam\s+\$env:USERNAME'
        $script:src | Should -Match 'throw'
    }
    It 'maps Exchange RBAC role groups to the tiered admin model' {
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Group\s+"Organization Management"'
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Group\s+"Server Management"'
        $script:src | Should -Match 'Add-GroupMember[^\r\n]*-Group\s+"Recipient Management"'
    }
    It 'connects to remote Exchange over a Kerberos PSSession' {
        $script:src | Should -Match 'New-PSSession\s+-ConfigurationName\s+Microsoft\.Exchange'
        $script:src | Should -Match '-authentication\s+Kerberos'
    }
    It 'aborts if the session fails to connect' {
        $script:src | Should -Match 'if\s*\(!\$ExSession\)'
        $script:src | Should -Match 'Exit'
    }
    It 'tears down the session and disposes the credential at the end' {
        $script:src | Should -Match 'Remove-PsSession\s+\$ExSession'
        $script:src | Should -Match '\$Cred\.Password\.Dispose'
    }
    It 'configures <VDir> with both external and internal URLs' -ForEach @(
        @{ VDir = 'Ecp' }
        @{ VDir = 'WebServices' }
        @{ VDir = 'Mapi' }
        @{ VDir = 'ActiveSync' }
        @{ VDir = 'Oab' }
        @{ VDir = 'Owa' }
        @{ VDir = 'PowerShell' }
    ) {
        $script:src | Should -Match "Set-${VDir}VirtualDirectory[^\r\n]*-ExternalUrl"
        $script:src | Should -Match "Set-${VDir}VirtualDirectory[^\r\n]*-InternalUrl"
    }
    It 'creates the primary and secondary email address policies' {
        $script:src | Should -Match 'New-EmailAddressPolicy[^\r\n]*-Priority 1'
        $script:src | Should -Match 'New-EmailAddressPolicy[^\r\n]*-Priority 2'
        $script:src | Should -Match 'SMTP:%g\.%s@'
        $script:src | Should -Match 'SMTP:%m@'
    }
}
