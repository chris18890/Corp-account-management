[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][string]$UserAction,
    [string]$GroupName,
    [string]$Domain
)
if (!$Domain) {
    $Domain = "$env:userdomain"
}
if (!$GroupName) {
    $GroupName = "Administrators"
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter a username to be elevated/removed - '
}
if (!$UserAction) {
    $UserAction = READ-HOST 'Enter an action;  E to elevate, R to remove - '
}
$UserAction=$UserAction.ToUpper()
switch ($UserAction) {
    "E" {
        try {
            Add-LocalGroupMember -Group $GroupName -Member "$Domain\$UserName" -ErrorAction Stop
        } catch [Microsoft.ActiveDirectory.Management.ADException] {
            switch ($Error[0].Exception.ErrorCode) {
                1378 { # 'The specified object is already a member of the group'
                    Write-Host "'$UserName' is already a member of group '$GroupName'" -ForegroundColor Green
                }
                default {
                    Write-Host "ERROR: An unexpected error occurred while attempting to add user '$UserName' to a group:`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                }
            }
        }
    }
    "R" {
        Remove-LocalGroupMember -Group $GroupName -Member "$Domain\$UserName" -ErrorAction Stop
    }
}
