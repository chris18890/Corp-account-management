@{
    Severity = @('Error', 'Warning')
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'           # CorpAdmin uses Write-Host in Write-LogFile
        ,'PSUseShouldProcessForStateChangingFunctions'  # noisy for setup scripts
        ,'PSAvoidUsingConvertToSecureStringWithPlainText'
        ,'PSAvoidUsingPlainTextForPassword'
        ,'PSReviewUnusedParameter'
        ,'PSUseDeclaredVarsMoreThanAssignments'
    )
    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable = $true
            TargetVersions = @('5.1', '7.4')
        }
    }
}
