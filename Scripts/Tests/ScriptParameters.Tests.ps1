#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Parameter-contract tests for scripts that take parameters with declared
# constraints. These tests don't execute the script bodies - they just
# inspect the parameter metadata via Get-Command.
#
# Catches: dropped Mandatory flags, drifted ValidateSet values, drifted
# ValidateRange bounds. These regressions are silent at script load time
# but break callers downstream.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    function Get-ParamAttr {
        param($Cmd, [string]$Name, [Type]$Type)
        $Cmd.Parameters[$Name].Attributes | Where-Object { $_ -is $Type }
    }
    
    function Test-IsMandatory {
        param($Cmd, [string]$Name)
        $attrs = Get-ParamAttr $Cmd $Name ([System.Management.Automation.ParameterAttribute])
        $null -ne ($attrs | Where-Object Mandatory)
    }
}

# ====================================================================
# Tier 0 / Tier 1 infrastructure
# ====================================================================

Describe 'setup.ps1 (Machine1) parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Machine1\setup.ps1')
    }
    It 'requires -Domain'   { Test-IsMandatory $script:cmd 'Domain'     | Should -BeTrue }
    It 'requires -Platform' { Test-IsMandatory $script:cmd 'Platform'   | Should -BeTrue }
    It 'requires -Role'     { Test-IsMandatory $script:cmd 'Role'       | Should -BeTrue }
    It 'Platform ValidateSet is Azure,Local' {
        $vs = Get-ParamAttr $script:cmd 'Platform' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('Azure','Local') | Sort-Object)
    }
    It 'Role ValidateSet covers all infrastructure roles' {
        $vs = Get-ParamAttr $script:cmd 'Role' ([System.Management.Automation.ValidateSetAttribute])
        $expected = 'DC1','RTR','DC2','DC3','RootCA','InterCA','WSUS','EXCH'
        ($vs.ValidValues | Sort-Object) | Should -Be ($expected | Sort-Object)
    }
}

Describe 'DomainSetup.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Prelim\DomainSetup.ps1')
    }
    It 'requires -EmailSuffix' { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -Drive'       { Test-IsMandatory $script:cmd 'Drive'       | Should -BeTrue }
}

Describe 'dhcp.ps1 (Machine1) parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Machine1\dhcp.ps1')
    }
    It 'requires -Domain'    { Test-IsMandatory $script:cmd 'Domain'   | Should -BeTrue }
    It 'requires -Platform'  { Test-IsMandatory $script:cmd 'Platform' | Should -BeTrue }
    It 'Platform ValidateSet is Azure,Local' {
        $vs = Get-ParamAttr $script:cmd 'Platform' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('Azure','Local') | Sort-Object)
    }
}

Describe 'dcpromo.ps1 (Machine1) parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Machine1\dcpromo.ps1')
    }
    It 'requires -Domain'       { Test-IsMandatory $script:cmd 'Domain'       | Should -BeTrue }
    It 'requires -DomainSuffix' { Test-IsMandatory $script:cmd 'DomainSuffix' | Should -BeTrue }
}

Describe 'wds.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'WDS\wds.ps1')
    }
    It 'requires -Drive' { Test-IsMandatory $script:cmd 'Drive' | Should -BeTrue }
}

Describe 'additional-dcpromo.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\additional-dcpromo.ps1')
    }
    It 'requires -Domain'       { Test-IsMandatory $script:cmd 'Domain'       | Should -BeTrue }
    It 'requires -DomainSuffix' { Test-IsMandatory $script:cmd 'DomainSuffix' | Should -BeTrue }
}

Describe 'additional-DFS-Setup.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\additional-DFS-Setup.ps1')
    }
    It 'requires -Drive' { Test-IsMandatory $script:cmd 'Drive' | Should -BeTrue }
}

Describe 'Setup-Exch-PostInstall.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\Setup-Exch-PostInstall.ps1')
    }
    It 'requires -EmailSuffix' { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
}

Describe 'setup-exch-PreReqs.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\setup-exch-PreReqs.ps1')
    }
    It 'requires -Domain' { Test-IsMandatory $script:cmd 'Domain' | Should -BeTrue }
}

Describe 'setup-InterCA.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\setup-InterCA.ps1')
    }
    It 'requires -Drive'  { Test-IsMandatory $script:cmd 'Drive'  | Should -BeTrue }
    It 'requires -Domain' { Test-IsMandatory $script:cmd 'Domain' | Should -BeTrue }
}

Describe 'setup-wsus.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Additional\setup-wsus.ps1')
    }
    It 'requires -Drive'  { Test-IsMandatory $script:cmd 'Drive'  | Should -BeTrue }
    It 'requires -Domain' { Test-IsMandatory $script:cmd 'Domain' | Should -BeTrue }
}

# ====================================================================
# Security-sensitive scripts
# ====================================================================

Describe 'ElevateUser.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\ElevateUser.ps1')
    }
    It 'requires -GroupName'    { Test-IsMandatory $script:cmd 'GroupName'  | Should -BeTrue }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'   | Should -BeTrue }
    It 'requires -UserAction'   { Test-IsMandatory $script:cmd 'UserAction' | Should -BeTrue }
    It 'UserAction ValidateSet is E,R' {
        $vs = Get-ParamAttr $script:cmd 'UserAction' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('E','R') | Sort-Object)
    }
    It 'TempOrPerm ValidateSet is P,T' {
        $vs = Get-ParamAttr $script:cmd 'TempOrPerm' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('P','T') | Sort-Object)
    }
    It 'TimeSpan ValidateRange has minimum 1' {
        $vr = Get-ParamAttr $script:cmd 'TimeSpan' ([System.Management.Automation.ValidateRangeAttribute])
        $vr.MinRange | Should -Be 1
    }
    It 'TimeSpan ValidateRange has no hard upper bound' {
        $vr = Get-ParamAttr $script:cmd 'TimeSpan' ([System.Management.Automation.ValidateRangeAttribute])
        $vr.MaxRange | Should -Be ([int]::MaxValue)
    }
}

Describe 'ElevateUserLocal.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\ElevateUserLocal.ps1')
        $script:src = Get-Content $script:cmd.Path -Raw
        $tokens = $null; $errors = $null
        $script:elevateAst = [System.Management.Automation.Language.Parser]::ParseFile($script:cmd.Path, [ref]$tokens, [ref]$errors)
    }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'   | Should -BeTrue }
    It 'requires -UserAction'   { Test-IsMandatory $script:cmd 'UserAction' | Should -BeTrue }
    It 'UserAction ValidateSet is E,R' {
        $vs = Get-ParamAttr $script:cmd 'UserAction' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('E','R') | Sort-Object)
    }
    # Negative contract - bootstrap-friendly, no shared module.
    # AST-based so header comments referencing these names can't false-positive.
    It 'has no -Reason parameter' {
        $script:cmd.Parameters.Keys | Should -Not -Contain 'Reason'
    }
    It 'does not import CorpAdmin' {
        $imports = $script:elevateAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Import-Module' -and
            $n.Extent.Text -match 'CorpAdmin'
        }, $true)
        $imports | Should -BeNullOrEmpty
    }
    It 'does not call Write-LogFile' {
        $calls = $script:elevateAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Write-LogFile'
        }, $true)
        $calls | Should -BeNullOrEmpty
    }
    It 'does not write an audit CSV' {
        $exports = $script:elevateAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Export-Csv'
        }, $true)
        $exports | Should -BeNullOrEmpty
    }
}

Describe 'Enable-CloudAdmin.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1')
    }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -Tier'         { Test-IsMandatory $script:cmd 'Tier'        | Should -BeTrue }
    It 'Tier ValidateSet is Cloud,Global' {
        $vs = Get-ParamAttr $script:cmd 'Tier' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('Cloud','Global') | Sort-Object)
    }
    It 'DurationMinutes ValidateRange is 0..480' {
        $vr = Get-ParamAttr $script:cmd 'DurationMinutes' ([System.Management.Automation.ValidateRangeAttribute])
        $vr.MinRange | Should -Be 0
        $vr.MaxRange | Should -Be 480
    }
}

Describe 'Disable-CloudAdmin.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\Disable-CloudAdmin.ps1')
    }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -Tier'         { Test-IsMandatory $script:cmd 'Tier'        | Should -BeTrue }
    It 'Tier ValidateSet is Cloud,Global' {
        $vs = Get-ParamAttr $script:cmd 'Tier' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('Cloud','Global') | Sort-Object)
    }
}

# ====================================================================
# User-management scripts
# ====================================================================

Describe 'CreateUsers.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateUsers.ps1')
    }
    It 'requires -O365'         { Test-IsMandatory $script:cmd 'O365'        | Should -BeTrue }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'O365 ValidateSet is E,H,N' {
        $vs = Get-ParamAttr $script:cmd 'O365' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('E','H','N') | Sort-Object)
    }
    It 'validates all required CSV headers are checked' {
        $src = Get-Content $script:cmd.Path -Raw
        foreach ($header in @('USERNAME','FIRSTNAME','LASTNAME','DEPT','COMPANY','MANAGER',
                              'Requester','S/E/R','AdminID','Managed','Cap','REALNAME',
                              'PHONE','HIPRIV','PrivLevel','Description')) {
            $src | Should -Match ([regex]::Escape($header))
        }
    }
}

Describe 'Cleanup-ADSyncFailureUsers.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\Cleanup-ADSyncFailureUsers.ps1')
    }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'validates all required CSV headers are checked' {
        $src = Get-Content $script:cmd.Path -Raw
        foreach ($header in @('USERNAME','FIRSTNAME','LASTNAME','DEPT','COMPANY',
                              'S/E/R','CAP','HIPRIV','PrivLevel')) {
            $src | Should -Match ([regex]::Escape($header))
        }
    }
}

Describe 'CreateITAdminUser.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateITAdminUser.ps1')
    }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
    It 'PrivLevel ValidateSet is 1,2,3' {
        $vs = Get-ParamAttr $script:cmd 'PrivLevel' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('1','2','3') | Sort-Object)
    }
}

Describe 'CreateITDomainAdminUser.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateITDomainAdminUser.ps1')
    }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
    # Domain Admin accounts cannot be Level 1 (desktop admins) - only 2 or 3.
    It 'PrivLevel ValidateSet is 2,3 (NOT 1)' {
        $vs = Get-ParamAttr $script:cmd 'PrivLevel' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('2','3') | Sort-Object)
    }
}

Describe 'CreateITCloudAdminUser.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateITCloudAdminUser.ps1')
    }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
    It 'PrivLevel ValidateSet is 1,2,3' {
        $vs = Get-ParamAttr $script:cmd 'PrivLevel' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('1','2','3') | Sort-Object)
    }
}

Describe 'CreateITGlobalAdminUser.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateITGlobalAdminUser.ps1')
    }
    It 'requires -EmailSuffix'  { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'requires -UserName'     { Test-IsMandatory $script:cmd 'UserName'    | Should -BeTrue }
}

Describe 'CreateOnPremMailboxes.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateOnPremMailboxes.ps1')
    }
    It 'requires -EmailSuffix' { Test-IsMandatory $script:cmd 'EmailSuffix' | Should -BeTrue }
    It 'validates all required CSV headers are checked' {
        $src = Get-Content $script:cmd.Path -Raw
        foreach ($header in @('USERNAME','S/E/R','CAP','REALNAME')) {
            $src | Should -Match ([regex]::Escape($header))
        }
    }
}

Describe 'Movers.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\Movers.ps1')
        $script:src = Get-Content $script:cmd.Path -Raw
    }
    It 'validates all four required CSV headers are checked' {
        foreach ($header in @('USERNAME', 'OLD_DEPT', 'NEW_DEPT', 'NEW_MANAGER')) {
            $script:src | Should -Match ([regex]::Escape($header))
        }
    }
}

Describe 'CreateGroup.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\CreateGroup.ps1')
    }
    It 'requires -GroupName'    { Test-IsMandatory $script:cmd 'GroupName'   | Should -BeTrue }
    It 'requires -GroupType'    { Test-IsMandatory $script:cmd 'GroupType'   | Should -BeTrue }
    It 'requires -O365'         { Test-IsMandatory $script:cmd 'O365'        | Should -BeTrue }
    It 'GroupType ValidateSet is S,H' {
        $vs = Get-ParamAttr $script:cmd 'GroupType' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('S','H') | Sort-Object)
    }
    It 'O365 ValidateSet is E,H,N' {
        $vs = Get-ParamAttr $script:cmd 'O365' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('E','H','N') | Sort-Object)
    }
}

Describe 'ChangePassword.ps1 parameter contract' {
    BeforeAll {
        $script:cmd = Get-Command (Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1')
    }
    It 'requires -UserType' { Test-IsMandatory $script:cmd 'UserType' | Should -BeTrue }
    It 'UserType ValidateSet is S,H' {
        $vs = Get-ParamAttr $script:cmd 'UserType' ([System.Management.Automation.ValidateSetAttribute])
        ($vs.ValidValues | Sort-Object) | Should -Be (@('S','H') | Sort-Object)
    }
    # -UserName is the optional non-interactive target (SamAccountName). It must NOT
    # be mandatory - omitting it is what triggers the interactive Out-GridView
    # picker. -LogFile is the optional caller/test-supplied log path.
    It 'has an optional -UserName parameter (interactive fallback when omitted)' {
        $script:cmd.Parameters.Keys           | Should -Contain 'UserName'
        Test-IsMandatory $script:cmd 'UserName'   | Should -BeFalse
    }
    It 'has an optional -LogFile parameter' {
        $script:cmd.Parameters.Keys             | Should -Contain 'LogFile'
        Test-IsMandatory $script:cmd 'LogFile'  | Should -BeFalse
    }
}

Describe 'ChangePassword.ps1 uses explicit SearchScope Subtree' {
    BeforeAll {
        $path = Join-Path $script:scriptsRoot 'Users\ChangePassword.ps1'
        $tokens = $null; $errors = $null
        $script:cpAst = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'Get-ADUser call specifies -SearchScope Subtree explicitly' {
        $call = $script:cpAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Get-ADUser'
        }, $true) | Select-Object -First 1
        $call | Should -Not -BeNullOrEmpty
        $call.Extent.Text | Should -Match '-SearchScope\s+Subtree'
    }
}
