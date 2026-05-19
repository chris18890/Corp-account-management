# AST/body coverage for the Tier-0 bootstrap scripts.
# Complements Environment.Tests.ps1, which pins the environment.psd1 values these read.

BeforeAll {
    $script:scriptsRoot = Split-Path $PSScriptRoot -Parent   # Tests -> Scripts
    $script:machineRoot = Join-Path $script:scriptsRoot 'Machine1'
}

Describe 'setup.ps1 (Tier-0 bootstrap)' {
    BeforeAll {
        $script:setupPath = Join-Path $script:machineRoot 'setup.ps1'
        $script:setupSrc  = Get-Content $script:setupPath -Raw
        $t = $null; $e = $null
        $script:setupAst  = [System.Management.Automation.Language.Parser]::ParseFile($script:setupPath, [ref]$t, [ref]$e)
        $roleParam = $script:setupAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ParameterAst] -and
                      $n.Name.VariablePath.UserPath -eq 'Role'
        }, $true) | Select-Object -First 1
        $vs = $roleParam.Attributes | Where-Object { $_.TypeName.Name -match 'ValidateSet' } | Select-Object -First 1
        $script:roleSet = @($vs.PositionalArguments.Value | Sort-Object)
        $sw = $script:setupAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst]
        }, $true) | Select-Object -First 1
        $script:switchRoles = @($sw.Clauses | ForEach-Object { $_.Item1.Value } | Sort-Object)
    }
    It 'parses, and is RunAsAdministrator + StrictMode' {
        $script:setupAst | Should -Not -BeNullOrEmpty
        $script:setupSrc | Should -Match '#Requires\s+-RunAsAdministrator'
        $script:setupSrc | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It 'aborts when environment.psd1 is missing' {
        $script:setupSrc | Should -Match 'Test-Path.+EnvironmentConfig'
        $script:setupSrc | Should -Match 'throw'
    }
    It 'every accepted -Role has a matching switch case (and no orphan case)' {
        # An accepted role with no case leaves $IPAddress unset -> StrictMode fault later.
        $script:switchRoles | Should -Be $script:roleSet
    }
    It '-Role set is exactly the eight infra roles' {
        $script:roleSet | Should -Be @('DC1','DC2','DC3','EXCH','InterCA','RootCA','RTR','WSUS')
    }
    It 'renames the machine to "$Domain-$Role"' {
        $script:setupSrc | Should -Match '\$Machine\s*=\s*"\$Domain-\$Role"'
    }
}

Describe 'setup.ps1 per-role static IP' -ForEach @(
    @{ Role = 'DC1';     Net = 'IPNet1'; Offset = 'DC1'     }
    @{ Role = 'DC2';     Net = 'IPNet1'; Offset = 'DC2'     }
    @{ Role = 'RTR';     Net = 'IPNet1'; Offset = 'RTR'     }
    @{ Role = 'RootCA';  Net = 'IPNet1'; Offset = 'RootCA'  }
    @{ Role = 'InterCA'; Net = 'IPNet1'; Offset = 'InterCA' }
    @{ Role = 'WSUS';    Net = 'IPNet1'; Offset = 'WSUS'    }
    @{ Role = 'EXCH';    Net = 'IPNet1'; Offset = 'EXCH'    }
    # DC3 is the second-site DC: it takes DC1's offset (.2) on the second
    # subnet ($IPNet2), so it is intentionally NOT $Hosts.DC3 (no such key) -
    # it mirrors DC1 one subnet over.
    @{ Role = 'DC3';     Net = 'IPNet2'; Offset = 'DC1'     }
) {
    BeforeAll { $script:setupSrc = Get-Content (Join-Path $script:machineRoot 'setup.ps1') -Raw }
    It '<Role>: assigns IP as $<Net>.$($Hosts.<Offset>)' {
        # Pins each role's subnet + offset; still catches a real copy/paste
        # (e.g. DC2 -> $Hosts.DC1, or DC3 drifting onto $IPNet1).
        $script:setupSrc | Should -Match ('\$IPAddress\s*=\s*"\$' + $Net + '\.\$\(\$Hosts\.' + $Offset + '\)"')
    }
}

Describe 'dhcp.ps1 (Tier-0 network setup)' {
    BeforeAll {
        $script:dhcpPath = Join-Path $script:machineRoot 'dhcp.ps1'
        $script:dhcpSrc  = Get-Content $script:dhcpPath -Raw
        $t = $null; $e = $null
        $script:dhcpAst  = [System.Management.Automation.Language.Parser]::ParseFile($script:dhcpPath, [ref]$t, [ref]$e)
    }
    It 'parses, and is RunAsAdministrator + StrictMode' {
        $script:dhcpAst | Should -Not -BeNullOrEmpty
        $script:dhcpSrc | Should -Match '#Requires\s+-RunAsAdministrator'
        $script:dhcpSrc | Should -Match 'Set-StrictMode\s+-Version\s+Latest'
    }
    It 'loads config through the CorpAdmin module' {
        $script:dhcpSrc | Should -Match 'Get-EnvironmentConfig'
    }
    It '-Platform is constrained to Azure or Local' {
        $pp = $script:dhcpAst.FindAll({
            param($n) $n -is [System.Management.Automation.Language.ParameterAst] -and
                      $n.Name.VariablePath.UserPath -eq 'Platform'
        }, $true) | Select-Object -First 1
        $vs = $pp.Attributes | Where-Object { $_.TypeName.Name -match 'ValidateSet' } | Select-Object -First 1
        @($vs.PositionalArguments.Value | Sort-Object) | Should -Be @('Azure','Local')
    }
    It 'derives the DHCP scope bounds from config (not hardcoded)' {
        $script:dhcpSrc | Should -Match '\$StartRange\s*=\s*\$Env\.Network\.DhcpStart'
        $script:dhcpSrc | Should -Match '\$EndRange\s*=\s*\$Env\.Network\.DhcpEnd'
    }
    It 'derives router + DNS from HostOffsetsLocal' {
        $script:dhcpSrc | Should -Match '\$Router\s*=\s*\$Env\.Network\.HostOffsetsLocal\.RTR'
        $script:dhcpSrc | Should -Match '\$DNSServer2\s*=\s*\$Env\.Network\.HostOffsetsLocal\.DC1'
        $script:dhcpSrc | Should -Match '\$DNSServer3\s*=\s*\$Env\.Network\.HostOffsetsLocal\.DC2'
    }
    It 'Local platform adds the routing + dhcp features and renames NICs from config' {
        $script:dhcpSrc | Should -Match '\$FeatureName\s*\+=\s*@\("routing"'
        $script:dhcpSrc | Should -Match '\$Env\.Network\.Interfaces\.External'
        $script:dhcpSrc | Should -Match '\$Env\.Network\.Interfaces\.Internal'
    }
}
