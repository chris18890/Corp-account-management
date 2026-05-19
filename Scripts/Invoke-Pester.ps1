Import-Module Pester -MinimumVersion 5.5.0
$config = New-PesterConfiguration
$config.Output.Verbosity = 'Detailed'
$config.Run.Path = Join-Path $PSScriptRoot 'Tests'
$config.TestResult.Enabled = $true
$config.Run.Exit = $true
$config.TestResult.OutputPath = 'TestResults.xml'
$config.TestResult.OutputFormat = 'JUnitXml'
$config.CodeCoverage.Enabled = $true
# =============================================================================
# Code coverage scope
# =============================================================================
# Recursively enumerates every .ps1 and .psm1 under Scripts\, with three
# explicit exclusions. The exclusion list is deliberate, not defensive bloat:
#
#   \Tests\        - test files themselves. A test file is "covered" by the
#                    act of being run, which is meaningless signal. Including
#                    them inflates the line-rate denominator with code that
#                    can't have a regression hidden in it.
#
#   *.Tests.ps1    - test files that end up outside the \Tests\ directory
#                    (sibling files, future restructuring). Defence in depth
#                    against the first exclusion's path-shape assumption.
#
#   Invoke-Pester  - the runner itself. It can't measure its own execution
#                    while running tests, and including it makes the report
#                    structurally confusing.
#
# When adding a new subdirectory under Scripts\ that contains vendored or
# third-party code outside your maintenance, EXTEND THIS LIST rather than
# letting the coverage number drop. A 0% on vendored code looks like a gap
# but isn't actionable - it's noise.
# =============================================================================
$config.CodeCoverage.Path = @(
    Get-ChildItem -Path $PSScriptRoot -Recurse -Include '*.ps1','*.psm1' -File |
        Where-Object {
            $_.FullName -notmatch '\\Tests\\'   -and
            $_.Name     -notlike '*.Tests.ps1' -and
            $_.Name     -ne     'Invoke-Pester.ps1'
        } |
        ForEach-Object FullName
)
Invoke-Pester -Configuration $config
