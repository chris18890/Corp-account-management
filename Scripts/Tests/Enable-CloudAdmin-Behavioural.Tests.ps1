#====================================================================
# Enable-CloudAdmin.ps1 - BEHAVIOURAL rollback test.
# CloudAdmin_Tests.ps1 pins the rollback STRUCTURE via AST; this EXECUTES the
# script and proves that when scheduling fails after the account is enabled,
# the catch actually re-disables it.
#====================================================================
BeforeAll {
    $script:scriptsRoot = Split-Path $PSScriptRoot -Parent
    $script:enablePath  = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
    # Stub Microsoft.Graph so #Requires is satisfied and the cmdlets are mockable.
    $script:stubRoot = Join-Path $TestDrive 'StubModules'
    $gph = Join-Path $script:stubRoot 'Microsoft.Graph'
    New-Item -ItemType Directory -Path $gph -Force | Out-Null
    Set-Content (Join-Path $gph 'Microsoft.Graph.psd1') "@{ ModuleVersion='1.0.0'; RootModule='Microsoft.Graph.psm1'; FunctionsToExport=@('Get-MgContext','Connect-MgGraph','Get-MgUser','Update-MgUser','Disconnect-MgGraph') }"
    Set-Content (Join-Path $gph 'Microsoft.Graph.psm1') @'
function Get-MgContext { [pscustomobject]@{ Account = "stub" } }
function Connect-MgGraph { param([switch]$NoWelcome, $Scopes) }
function Get-MgUser { param($Filter, $UserId, $ErrorAction) [pscustomobject]@{ Id = "1111"; AccountEnabled = $false } }
function Update-MgUser { param($UserId, [bool]$AccountEnabled, $ErrorAction) }
function Disconnect-MgGraph { }
'@
    Get-Module Microsoft.Graph | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module (Join-Path $gph 'Microsoft.Graph.psd1') -Force
    $st = Join-Path $script:stubRoot 'ScheduledTasks'
    New-Item -ItemType Directory -Path $st -Force | Out-Null
    Set-Content (Join-Path $st 'ScheduledTasks.psd1') "@{ ModuleVersion='1.0.0'; RootModule='ScheduledTasks.psm1'; FunctionsToExport=@('Get-ScheduledTask','Unregister-ScheduledTask','New-ScheduledTaskAction','New-ScheduledTaskTrigger','New-ScheduledTaskPrincipal','New-ScheduledTaskSettingsSet','Register-ScheduledTask') }"
    Set-Content (Join-Path $st 'ScheduledTasks.psm1') @'
function Get-ScheduledTask        { param($TaskName, $ErrorAction) }
function Unregister-ScheduledTask { param($Confirm) }
function New-ScheduledTaskAction      { param([string]$Execute, [string]$Argument)                                                              [pscustomobject]@{} }
function New-ScheduledTaskTrigger     { param([switch]$Once, [datetime]$At)                                                                     [pscustomobject]@{} }
function New-ScheduledTaskPrincipal   { param([string]$UserId, [string]$RunLevel, [string]$LogonType)                                           [pscustomobject]@{} }
function New-ScheduledTaskSettingsSet { param([switch]$AllowStartIfOnBatteries, [switch]$DontStopIfGoingOnBatteries, [switch]$StartWhenAvailable) [pscustomobject]@{} }
function Register-ScheduledTask       { param($TaskName, $Action, $Trigger, $Principal, $Settings, $Description, $ErrorAction) }
'@
    # Make sure the real CDXML module isn't already loaded, then force the stub in
    Get-Module ScheduledTasks | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module (Join-Path $st 'ScheduledTasks.psd1') -Force
    $realEnable = Join-Path $script:scriptsRoot 'Users\Enable-CloudAdmin.ps1'
    Set-Content (Join-Path $TestDrive 'Disable-CloudAdmin.ps1') '# stub'
    $testEnable = Join-Path $TestDrive 'Enable-CloudAdmin.ps1'
    $content = (Get-Content $realEnable -Raw) -replace '(?m)^#Requires -RunAsAdministrator\s*$','' -replace 'Import-Module \(Join-Path \(Split-Path \$PSScriptRoot -Parent\) ''Modules\\CorpAdmin\\CorpAdmin\.psd1''\)',"Import-Module '$(Join-Path $script:scriptsRoot 'Modules\CorpAdmin\CorpAdmin.psd1')'"
    $content | Set-Content $testEnable
    $script:enablePath = $testEnable
    if ($env:PSModulePath -notlike "*$script:stubRoot*") {
        $env:PSModulePath = "$script:stubRoot$([IO.Path]::PathSeparator)$env:PSModulePath"
    }
}

Describe 'Enable-CloudAdmin.ps1 rollback (behavioural)' {
    BeforeEach {
        # Auth gate: invoker treated as authorised (script-scope call -> no -ModuleName).
        Mock Get-ADDomain   { [pscustomobject]@{ PDCEmulator = 'dc.example.com' } }
        Mock Test-IsMemberOf { $true }
        Mock Get-ADGroup    { [pscustomobject]@{ Name = 'x' } }
        Mock Get-ADGroupMember { [pscustomobject]@{ SamAccountName = $env:USERNAME } }
        # Graph
        Mock Get-MgContext { [pscustomobject]@{ Account = 'stub' } }
        Mock Connect-MgGraph { }
        Mock Get-MgUser  { [pscustomobject]@{ Id = '1111'; AccountEnabled = $false } }
        Mock Update-MgUser { }
        # Scheduled-task surface
        Mock Get-ScheduledTask        { }
        Mock Unregister-ScheduledTask { }
        Mock New-ScheduledTaskAction      { [pscustomobject]@{} }
        Mock New-ScheduledTaskTrigger     { [pscustomobject]@{} }
        Mock New-ScheduledTaskPrincipal   { [pscustomobject]@{} }
        Mock New-ScheduledTaskSettingsSet { [pscustomobject]@{} }
        Mock Register-ScheduledTask { }
    }
    It 're-disables the account when scheduling fails after enable' {
        Mock Register-ScheduledTask { throw 'simulated scheduling failure' }
        { & $script:enablePath -UserName foo -EmailSuffix example.com -Tier Cloud -DurationMinutes 60 -Reason 'behavioural test' } | Should -Throw
        Should -Invoke Update-MgUser -ParameterFilter { $AccountEnabled -eq $true } -Times 1
        Should -Invoke Update-MgUser -ParameterFilter { $AccountEnabled -eq $false } -Times 1 # rollback fired
    }
    It 'does NOT disable the account when scheduling succeeds' {
        Mock Register-ScheduledTask { }
        & $script:enablePath -UserName foo -EmailSuffix example.com -Tier Cloud -DurationMinutes 60 -Reason 'behavioural test'
        Should -Invoke Update-MgUser -ParameterFilter { $AccountEnabled -eq $true } -Times 1
        Should -Invoke Update-MgUser -ParameterFilter { $AccountEnabled -eq $false } -Times 0
    }
}
