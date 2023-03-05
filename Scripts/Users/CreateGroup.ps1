[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EmailSuffix
    , [Parameter(Mandatory)][string]$O365
    , [Parameter(Mandatory)][string]$UserType
)

#====================================================================
#Domain Names in ADS & DNS format, and main OU name
#====================================================================
$Domain = "$env:userdomain"
$EndPath = (Get-ADDomain -Identity $Domain).DistinguishedName
$DNSSuffix = (Get-ADDomain -Identity $Domain).DNSRoot
$DCHostName = (Get-ADDomainController).HostName # Use this DC for all create/update operations, otherwise aspects may fail due to replication/timing issues
Write-Host "DC being used is '$DCHostName'"
$ExServer = "$Domain-EXCH.$DNSSuffix" #Remote Exchange PS session
#====================================================================

#====================================================================
#Drive where all the folders will be created
#====================================================================
$RootShare = "Store"
#====================================================================

#====================================================================
#Group Variables
#====================================================================
$GroupsOU = "Groups"
$GroupCategory = "Security"
$GroupScope = "Universal"
$StaffGroup = "Staff"
#====================================================================

#====================================================================
#group creation function
#====================================================================
function Create-ADGroup {
    [CmdletBinding()]
    param(
        [string]$GroupName,[String]$Path,[String]$GroupDescription
    )
    $Error.Clear()
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -OtherAttributes:@{mail="$GroupName@$EmailSuffix"} -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
        Write-Host "'$GroupName' already exists" -ForegroundColor Green
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
        Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
    }
}
#====================================================================

#====================================================================
#group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Group
        , [Parameter(Mandatory)][string]$Member
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          User will be added to the group
    # Inputs:           $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    # Calls:            Write-Log function
    # Returns:
    # Notes:
    #================================================================
    $Error.Clear()
    $checkGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$Group)"
    if ($checkGroup -ne $null) {
        Write-Host "Adding $Member to $Group"
        try {
            Add-ADGroupMember -Identity $Group -Members $Member -Server $DCHostName
            Write-Host "Added $Member to $Group"
        }
        catch [Microsoft.ActiveDirectory.Management.ADException] {
            switch ($Error[0].Exception.ErrorCode) {
                1378 { # 'The specified object is already a member of the group'
                    Write-Host "'$Member' is already a member of group '$Group'" -ForegroundColor Yellow
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

#====================================================================
function Test-Cred {
    [CmdletBinding()]
    [OutputType([String])]
    Param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias(
            'PSCredential'
        )]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credentials
    )
    #================================================================
    # Purpose:          Test Provided Credentials Prior to Use
    # Assumptions:      
    # Effects:
    # Inputs:           $Cred
    # $LogString:       
    # Calls:
    # Returns:         Authenticated or Unauthenticated
    # Notes:
    #================================================================
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
    if ($Credentials -eq $null) {
        try {
            $Credentials = Get-Credential "domain\$env:username" -ErrorAction Stop
        } catch {
            $ErrorMsg = $_.Exception.Message
            Write-Warning "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
    }
    # Checking module
    try {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    } catch {
        $_.Exception.Message
        Continue
    }
    if (!$domain) {
        Write-Warning "Something went wrong"
    } else {
        if ($domain.name -ne $null) {
            return "Authenticated"
        } else {
            return "Not authenticated"
        }
    }
}
#====================================================================

#====================================================================
if ($O365 -eq "E" -or $O365 -eq "H") {
    # Get user credentials for server connectivity (Non-MFA)
    try {
        $Cred = Get-Credential -ErrorAction Stop -Message "Admin credentials for remote sessions:"
    } catch {
        $ErrorMsg = $_.Exception.Message
        Write-Host "Failed to validate credentials: $ErrorMsg "
        Pause
        Break
    }
    $CredCheck = $Cred | Test-Cred
    if ($CredCheck -ne "Authenticated") {
        Write-Host "Credential validation failed - Script Terminating"
        pause
        Exit
    }
    #Connect to remote Exchange PowerShell
    Write-Host "Connecting to remote Exchange PowerShell session... "
    try {
        $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExServer/PowerShell" -Name ExSession -Credential $cred -ErrorAction Stop -authentication Kerberos
        $ExConnected = $true
        Write-Host "connected."
        Write-Host "Importing Exchange session... "
        Import-PSSession -Session $ExSession -ErrorAction Stop -AllowClobber > $null
        Write-Host "done."
    } catch {
        $e = $_.Exception
        Write-Host $e
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Host $line
        $msg = $e.Message
        Write-Host $msg
        $Action = "Error Importing Exchange Session"
        Write-Host $Action
        Write-Host "failed."
        Write-Host "ERROR: $_" -ForegroundColor Red
    }
    if (!$ExSession) {
        Write-log "Exchange session not connected Stopping Script"
        Exit
    }
}
#====================================================================

#====================================================================
#group creation
#====================================================================
$Membership = READ-HOST 'Enter group name - '
$Description = READ-HOST 'Enter group description - '
switch ($UserType) {
    "S" {
        $OU = $GroupsOU
        $OUPath = "OU=$OU,$EndPath"
        Create-ADGroup -GroupName "$Membership" -Path $OUPath -GroupDescription "$Description"
        Add-GroupMember -Group $StaffGroup -Member $Membership
    }
    "H" {
        $OU = "Hi_Priv_Groups"
        $OUPath = "OU=$OU,OU=IT,$EndPath"
        Create-ADGroup -GroupName "$Membership" -Path $OUPath -GroupDescription "$Description"
    }
}
#====================================================================

#====================================================================
#Create Group Share
#====================================================================
switch ($UserType) {
    "S" {
        $ShareName = $Membership
        if (!(TEST-PATH "\\$Domain\$RootShare\$ShareName")) { 
            New-Item "\\$Domain\$RootShare\$ShareName" -type directory -force
            $Acl = Get-Acl "\\$Domain\$RootShare\$ShareName"
            $isProtected = $true 
            $preserveInheritance = $false
            $Acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule($ShareName,"Modify","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("RG_DFS_Admins","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            $Ar = New-Object system.security.accesscontrol.filesystemaccessrule("System","FullControl","ContainerInherit, ObjectInherit", "None", "Allow")
            $Acl.SetAccessRule($Ar)
            Set-Acl "\\$Domain\$RootShare\$ShareName" $Acl
        } else {
            Write-Host "\\$Domain\$RootShare\$ShareName already exists" -ForegroundColor Green
        }
    }
}
#====================================================================

#====================================================================
if ($O365 -eq "E" -or $O365 -eq "H") {
    if (Get-PSSession -Name ExSession -ErrorAction SilentlyContinue) {
        Remove-PsSession $ExSession
        Write-Host "Closed Exchange session."
    }
}
#====================================================================
