@{
    RootModule        = 'CorpAdmin.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a6a3dbe5-7251-4670-a14e-4c6638b3dc06'
    Author            = 'Chris Murray'
    CompanyName       = 'Corp-Account-Management'
    Copyright         = '(c) Chris Murray. Licensed under MIT.'
    Description       = 'Shared helper functions for the Corp-Account-Management scripts: environment config loading, AD group/mailbox helpers, logging, password generation/validation.'
    PowerShellVersion = '5.1'
    
    # Hard requirement - every consumer of this module uses AD cmdlets.
    RequiredModules   = @('ActiveDirectory')
    
    # ExchangeOnlineManagement, Microsoft.Graph, and Send-MailKitMessage are deliberately
    # NOT listed here. Only the mailbox helpers need them, and only when called against a
    # tenant that has them. Listing them would force-load on every Import-Module call,
    # which would break Hi-Priv setup scripts that run before the modules are installed.
    
    FunctionsToExport = @(
        'Get-EnvironmentConfig'
        ,'Write-LogFile'
        ,'Resolve-GroupMemberObject'
        ,'Test-IsMemberOf'
        ,'Get-ADGroupMemberTTLState'
        ,'Add-GroupMember'
        ,'Remove-GroupMember'
        ,'New-ADOU'
        ,'New-DomainGroup'
        ,'New-UserMailbox'
        ,'Update-UserMailbox'
        ,'New-UserOnPremMailbox'
        ,'Update-UserOnPremMailbox'
        ,'New-Password'
        ,'Test-Password'
        ,'Add-GPOLink'
        ,'Invoke-ADSync'
        ,'Get-ADSchemaGuidMap'
        ,'Get-ADExtendedRightsMap'
        ,'Grant-ComputerJoinDelegation'
        ,'Grant-GroupDelegation'
        ,'Grant-GroupMembershipEditDelegation'
        ,'Grant-PasswordResetDelegation'
        ,'Grant-UserDelegation'
        ,'Grant-OUDelegation'
        ,'Grant-DNSOperatorsPermissionDelegation'
        ,'Grant-DNSReadOnlyPermissionDelegation'
        ,'Grant-ADObjectPermissionDelegation'
        ,'Grant-GPOPermissionDelegation'
        ,'Grant-GPOCreationDelegation'
        ,'ConvertTo-IntOrDefault'
        ,'Send-NotificationEmail'
        ,'ConvertTo-SafeSamAccountName'
        ,'ConvertTo-SafeName'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    
    PrivateData = @{
        PSData = @{
            Tags         = @('ActiveDirectory','Exchange','UserManagement','Internal')
            ProjectUri   = 'https://github.com/chris18890/Corp-Account-Management'
            ReleaseNotes = 'Initial module-form release; functionally identical to the dot-sourced helpers.ps1 it replaces.'
        }
    }
}
