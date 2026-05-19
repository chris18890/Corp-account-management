#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Parses every .ps1 and .psm1 under Scripts/ and asserts no parse errors.
# Catches the [[ValidateRange / transposed ]) class of regressions that
# bit scripts in the past - they manifest as parse errors and the
# script won't even load.

BeforeDiscovery {
    $scriptsRoot = Join-Path $PSScriptRoot '..'
    $scriptFiles = Get-ChildItem -Path $scriptsRoot -Include '*.ps1','*.psm1' -Recurse -File |
        Where-Object {
            # Skip the Tests folder (this file lives there) and any LogFiles output.
            $_.FullName -notmatch '[\\/]Tests[\\/]'    -and
            $_.FullName -notmatch '[\\/]LogFiles[\\/]'
        }
}

Describe 'Script and module syntax' {
    It 'parses cleanly: <Name>' -ForEach $scriptFiles {
        $errors = $null
        $tokens = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            $_.FullName, [ref]$tokens, [ref]$errors) | Out-Null
        $errors | Should -BeNullOrEmpty -Because "PowerShell parser reported errors in $($_.FullName)"
    }
}
