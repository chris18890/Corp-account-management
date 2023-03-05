[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$GroupName,
    [Parameter(Mandatory)][string]$UserName,
    [Parameter(Mandatory)][string]$UserAction,
    [string]$TempOrPerm,
    [string]$TimeSpan
)

#====================================================================
#group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Group,
        [Parameter(Mandatory)][string]$Member,
        [string]$TimeSpan
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          User will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    #                   $TimeSpan - number of minutes to add temporal memebership for
    # Returns:
    # Notes:
    #================================================================
    $Error.Clear()
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)"
    if ($checkGroup -ne $null) {
        Write-Host "Adding $Member to $Group"
        try {
            if ($TimeSpan) {
                Add-ADGroupMember -Identity $Group -Members $Member -MemberTimeToLive (New-TimeSpan -Minutes $TimeSpan)
            } else {
                Add-ADGroupMember -Identity $Group -Members $Member
            }
            Write-Host "Added $Member to $Group"
        } catch [Microsoft.ActiveDirectory.Management.ADException] {
            switch ($Error[0].Exception.ErrorCode) {
                1378 { # 'The specified object is already a member of the group'
                    Write-Host "'$Member' is already a member of group '$Group'" -ForegroundColor Green
                }
                default {
                    Write-Host "ERROR: An unexpected error occurred while attempting to add user '$Member' to a group:`n$($Error[0].InvocationInfo.InvocationName) : $($Error[0].Exception.message)`n$($Error[0].InvocationInfo.PositionMessage)`n+ CategoryInfo : $($Error[0].CategoryInfo)`n+ FullyQualifiedErrorId : $($Error[0].FullyQualifiedErrorId)" -ForegroundColor Red
                }
            }
        }
    } else {
        Write-Host "$Group does not exist" -ForegroundColor Red
    }
}
#====================================================================

if (!$GroupName) {
    $GroupName = READ-HOST 'Enter a group name for elevation - '
}
if (!$UserName) {
    $UserName = READ-HOST 'Enter a username to be elevated - '
}
if (!$UserAction) {
    $UserAction = READ-HOST 'Enter an action;  E to elevate, R to remove - '
}
$UserAction=$UserAction.ToUpper()
switch ($UserAction) {
    "E" {
        if (!$TempOrPerm) {
            $TempOrPerm = READ-HOST 'Enter a duration;  P for permanent, T for temporary - '
        }
        $TempOrPerm=$TempOrPerm.ToUpper()
        switch ($TempOrPerm) {
            "P" {
                Add-ADGroupMember -Group $GroupName -Member $UserName
            }
            "T" {
                if (!$TimeSpan) {
                    Write-Host "No timespan specified, using default of 60 minutes"
                    $TimeSpan = "60"
                }
                Add-GroupMember -Group $GroupName -Member $UserName -TimeSpan $TimeSpan
            }
        }
    }
    "R" {
        Remove-ADGroupMember -Identity $GroupName -Members $UserName
    }
}
#====================================================================
