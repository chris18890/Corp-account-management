Set-StrictMode -Version Latest

# Capture the directory containing corpadmin.psm1 at dot-source time so the
# default environment.psd1 path resolves regardless of the caller's location.
$Script:HelpersDir = $PSScriptRoot
$Script:CachedEnv = $null
$Script:CachedPath = $null

#====================================================================
# Load the centralised environment configuration
#====================================================================
function Get-EnvironmentConfig {
    [CmdletBinding()]
    param(
        [string]$Path
        ,[switch]$Force
    )
    # Resolve the effective path FIRST, so the cache can be keyed on it.
    if (-not $Path) {
        # Allow ops override: set $env:CORPADMIN_ENV_PSD1 to point at an
        # alternate config file (useful for multi-tenant deployments).
        if ($env:CORPADMIN_ENV_PSD1 -and (Test-Path -LiteralPath $env:CORPADMIN_ENV_PSD1)) {
            $Path = $env:CORPADMIN_ENV_PSD1
        }
        else {
            # Module lives at Scripts/Modules/CorpAdmin/; environment.psd1 lives at Scripts/.
            $Path = Join-Path (Split-Path (Split-Path $Script:HelpersDir -Parent) -Parent) 'environment.psd1'
        }
    }
    # Cache hit only when not forced AND the resolved path matches what we cached.
    # This means an env-var change (or unset) between calls correctly invalidates.
    if (-not $Force -and $Script:CachedEnv -and $Script:CachedPath -eq $Path) {
        return $Script:CachedEnv
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "environment.psd1 not found at '$Path'. Pass -Path explicitly."
    }
    $config = Import-PowerShellDataFile -LiteralPath $Path
    # Structural validation: fail fast and loudly if a required top-level
    # section is absent, rather than letting a downstream script dereference
    # a null section (e.g. $Env.WSUS.Products) and fail far from the cause.
    $requiredSections = @(
        'Network','OUs','Groups','Shares','Locale',
        'Security','Azure','Exchange','EntraRoles','WSUS'
    )
    $missing = $requiredSections | Where-Object {
        -not $config.ContainsKey($_) -or $null -eq $config[$_]
    }
    if ($missing) {
        throw "environment.psd1 at '$Path' is missing required section(s): $($missing -join ', ')."
    }
    # Cache every resolution, keyed on the path it came from. An explicit
    # -Path call now also seeds the cache (keyed on that path) rather than
    # deliberately not populating it - which is fine, because the key
    # guarantees a later default-path call with a different resolved path
    # won't get this entry.
    $Script:CachedEnv  = $config
    $Script:CachedPath = $Path
    $config
}
#====================================================================

#====================================================================
# Set up logging
#====================================================================
function Write-LogFile {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$LogString
        ,[String]$ForegroundColor
    )
    #================================================================
    # Purpose:          To write a string with a date and time stamp to a log file
    # Assumptions:      $LogFile set with path to log file to write to
    # Effects:
    # Inputs:
    # $LogString:       String to write to log file
    # Calls:
    # Returns:
    # Notes:
    #================================================================
    $line = "$(Get-Date -Format 'G') $LogString"
    $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    if ($ForegroundColor) {
        Write-Host $LogString -ForegroundColor $ForegroundColor
    } else {
        Write-Host $LogString
    }
}
#====================================================================

#====================================================================
# Group resolver function
#====================================================================
function Resolve-GroupMemberObject {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Member
        ,[Parameter(Mandatory)][string]$DCHostName
    )
    #================================================================
    # Purpose:  Resolve $Member to a single AD object, class-agnostic -
    #           a user OR a group. The sAMAccountName arm covers users
    #           and security groups; the group-by-cn arm covers groups
    #           without a usable sAMAccountName. Returns a STATUS object
    #           rather than throwing: a missing member is a hard error
    #           for an add but a no-op for a remove, so that policy
    #           belongs to the caller, not to resolution.
    # Returns:  [pscustomobject] @{ Status; Object; Count }
    #             Status : 'Found' | 'NotFound' | 'Ambiguous'
    #             Object : the resolved AD object when Found, else $null
    #             Count  : number of objects matched
    # Notes:    $Member is LDAP-escaped so parentheses, backslashes,
    #           stars and nulls cannot break or inject the filter.
    #================================================================
    # DistinguishedName input path.
    # This makes the existing "supply a unique identifier or a distinguished name"
    # messages truthful and gives callers a deterministic way to bypass ambiguity.
    if ($Member -match '^(?i)(CN|OU)=[^,]+,.*DC=') {
        try {
            $resolvedByDn = Get-ADObject -Identity $Member -Server $DCHostName -ErrorAction Stop
            return [pscustomobject]@{
                Status = 'Found'
                Object = $resolvedByDn
                Count  = 1
            }
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            return [pscustomobject]@{
                Status = 'NotFound'
                Object = $null
                Count  = 0
            }
        }
    }
    $SafeMember = $Member -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00'
    $resolved = @(
        Get-ADObject -LDAPFilter "(|(sAMAccountName=$SafeMember)(&(objectClass=group)(cn=$SafeMember)))" -Server $DCHostName -ErrorAction Stop | Where-Object { $_ }
    )
    if ($resolved.Count -eq 0) {
        return [pscustomobject]@{
            Status = 'NotFound'
            Object = $null
            Count = 0
        }
    }
    if ($resolved.Count -gt 1) {
        return [pscustomobject]@{
            Status = 'Ambiguous'
            Object = $null
            Count = $resolved.Count
        }
    }
    return [pscustomobject]@{
        Status = 'Found'
        Object = $resolved[0]
        Count = 1
    }
}
#====================================================================

#====================================================================
# Group membership test function
#====================================================================
function Test-IsMemberOf {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Sam
        ,[Parameter(Mandatory)][string[]]$GroupNames
        ,[Parameter(Mandatory)][string]$DCHostName
    )
    #================================================================
    # Purpose:      Return $true if $Sam is a recursive member of ANY
    #               of the named groups, else $false. Backs the
    #               per-script authorisation gates.
    # Inputs:       $Sam        - sAMAccountName to test (e.g. $env:USERNAME)
    #               $GroupNames - one or more group Names to check
    #               $DCHostName - DC to target (PDC emulator by convention)
    # Returns:      [bool]
    # Notes:        Group names are escaped for the AD -Filter, so a name
    #               containing a single quote does not break the lookup.
    #================================================================
    foreach ($GroupName in $GroupNames) {
        $Safe  = $GroupName -replace "'", "''"
        $Group = Get-ADGroup -Filter "Name -eq '$Safe'" -Server $DCHostName -ErrorAction SilentlyContinue
        if (-not $Group) { continue }
        $Match = Get-ADGroupMember -Identity $Group.DistinguishedName -Recursive -Server $DCHostName -ErrorAction SilentlyContinue | Where-Object SamAccountName -eq $Sam
        if ($Match) { return $true }
    }
    return $false
}
#====================================================================

#====================================================================
# Group TTL status function
#====================================================================
function Get-ADGroupMemberTTLState {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$GroupDistinguishedName,
        [Parameter(Mandatory)][string]$MemberDistinguishedName,
        [Parameter(Mandatory)][string]$Server
    )
    $groupWithTtl = Get-ADGroup -Identity $GroupDistinguishedName -Properties member -ShowMemberTimeToLive -Server $Server -ErrorAction Stop
    foreach ($rawMember in @($groupWithTtl.member)) {
        $rawText = [string]$rawMember
        $ttlSeconds = $null
        $memberDn = $rawText
        if ($rawText -match '^<TTL=(\d+)>,(.+)$') {
            $ttlSeconds = [int]$Matches[1]
            $memberDn = $Matches[2]
        }
        if ($memberDn -ieq $MemberDistinguishedName) {
            return [pscustomobject]@{
                IsMember   = $true
                HasTTL     = ($null -ne $ttlSeconds)
                TTLSeconds = $ttlSeconds
                RawValue   = $rawText
            }
        }
    }
    [pscustomobject]@{
        IsMember   = $false
        HasTTL     = $false
        TTLSeconds = $null
        RawValue   = $null
    }
}
#====================================================================

#====================================================================
# Group addition function
#====================================================================
function Add-GroupMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogFile
        ,[Parameter(Mandatory)][string]$DCHostName
        ,[Parameter(Mandatory)][string]$Group
        ,[Parameter(Mandatory)][string]$Member
        ,[ValidateRange(1,[int]::MaxValue)][int]$TimeSpan
    )
    #================================================================
    # Purpose:          To add a user account or group to a group
    # Assumptions:      Parameters have been set correctly
    # Effects:          Member will be added to the group
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $Group - Group name as set before calling the function
    #                   $Member - Object to be added
    #                   $TimeSpan - number of minutes to add temporal memebership for
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:
    #================================================================
    $AddGroupMemberOutcome = 'Failed'
    try {
        Write-LogFile -LogFile $LogFile -LogString "=== Add-GroupMember START ==="
        Write-LogFile -LogFile $LogFile -LogString "Target Group: $Group"
        Write-LogFile -LogFile $LogFile -LogString "Target Member: $Member"
        # ============================================================
        # Resolve objects (fail fast if invalid)
        # ============================================================
        $GroupObj = Get-ADGroup -Identity $Group -Server $DCHostName -ErrorAction Stop
        $resolution = Resolve-GroupMemberObject -Member $Member -DCHostName $DCHostName
        switch ($resolution.Status) {
            'NotFound' {
                throw "Member '$Member' not found"
            }
            'Ambiguous' {
                throw "Member '$Member' is ambiguous ($($resolution.Count) objects matched); supply a unique identifier or a distinguished name."
            }
            'Found' {
                # Continue below.
            }
            default {
                throw "Unexpected Resolve-GroupMemberObject status '$($resolution.Status)' for member '$Member'."
            }
        }
        $MemberObj = $resolution.Object
        $MemberDN  = $null
        if ($MemberObj -and $MemberObj.PSObject.Properties.Match('DistinguishedName').Count -gt 0) {
            $MemberDN = [string]$MemberObj.DistinguishedName
        }
        if ([string]::IsNullOrWhiteSpace($MemberDN)) {
            throw "Resolve-GroupMemberObject returned Status='Found' but did not return an object with DistinguishedName for '$Member'."
        }
        # ============================================================
        # Check existing direct membership
        # ============================================================
        $ExistingMember = Get-ADGroupMember -Identity $GroupObj.DistinguishedName -Server $DCHostName | Where-Object DistinguishedName -eq $MemberDN
        if ($ExistingMember) {
            Write-LogFile -LogFile $LogFile -LogString "$Member is already a member of $Group"
            if (-not $TimeSpan) {
                Write-LogFile -LogFile $LogFile -LogString "No TTL requested > no change needed"
                $AddGroupMemberOutcome = 'NoChange'
                return
            }
            throw "$Member is already a member of $Group. Remove existing membership before applying TTL."
        }
        # ============================================================
        # Apply membership
        # ============================================================
        if ($TimeSpan) {
            # ========================================================
            # Validate PAM feature (TTL capability)
            # ========================================================
            $PAMFeature = Get-ADOptionalFeature -Filter "Name -eq 'Privileged Access Management Feature'" -Server $DCHostName -ErrorAction Stop
            if (-not $PAMFeature.EnabledScopes) {
                throw "AD PAM feature is not enabled on this domain"
            }
            Write-LogFile -LogFile $LogFile -LogString "Applying TTL membership: $TimeSpan minutes"
            $TTL = New-TimeSpan -Minutes $TimeSpan
            Add-ADGroupMember -Identity $GroupObj.DistinguishedName -Members $MemberDN -MemberTimeToLive $TTL -Server $DCHostName -ErrorAction Stop
            # ========================================================
            # Basic post-write membership validation
            # ========================================================
            $MemberCheck = Get-ADGroupMember -Identity $GroupObj.DistinguishedName -Server $DCHostName | Where-Object DistinguishedName -eq $MemberDN
            if (-not $MemberCheck) {
                throw "Post-add verification failed: membership not present"
            }
            # ========================================================
            # TTL post-write validation
            # ========================================================
            $TTLVerified = $false
            $TTLState = $null
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $TTLState = Get-ADGroupMemberTTLState -GroupDistinguishedName $GroupObj.DistinguishedName -MemberDistinguishedName $MemberDN -Server $DCHostName
                    if ($TTLState.IsMember -and $TTLState.HasTTL -and $TTLState.TTLSeconds -gt 0) {
                        $TTLVerified = $true
                        break
                    }
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "WARNING: TTL validation attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                Start-Sleep -Seconds 2
            }
            if (-not $TTLVerified) {
                Write-LogFile -LogFile $LogFile -LogString "ERROR: TTL membership could not be confirmed for $Member in $Group." -ForegroundColor Red
                try {
                    Remove-ADGroupMember -Identity $GroupObj.DistinguishedName -Members $MemberDN -Server $DCHostName -Confirm:$false -ErrorAction Stop
                    Write-LogFile -LogFile $LogFile -LogString "Rollback completed: removed $Member from $Group after failed TTL validation." -ForegroundColor Yellow
                } catch {
                    Write-LogFile -LogFile $LogFile -LogString "CRITICAL: TTL validation failed and rollback also failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                throw "TTL post-add verification failed: '$Member' is present or may be present, but an expiring TTL membership could not be confirmed for '$Group'."
            }
            if ($TTLState.TTLSeconds -gt ($TTL.TotalSeconds + 60)) {
                Write-LogFile -LogFile $LogFile -LogString "WARNING: TTL returned by AD ($($TTLState.TTLSeconds)s) is greater than requested TTL ($([int]$TTL.TotalSeconds)s)." -ForegroundColor Yellow
            }
            Write-LogFile -LogFile $LogFile -LogString "TTL membership verified for $Member in $Group. Remaining TTL: $($TTLState.TTLSeconds) seconds."
        } else {
            # ========================================================
            # Permanent membership (use VERY carefully)
            # ========================================================
            Write-LogFile -LogFile $LogFile -LogString "Applying PERMANENT membership" -ForegroundColor Yellow
            Add-ADGroupMember -Identity $GroupObj.DistinguishedName -Members $MemberDN -Server $DCHostName -ErrorAction Stop
            # Validation
            $MemberCheck = Get-ADGroupMember -Identity $GroupObj.DistinguishedName -Server $DCHostName | Where-Object DistinguishedName -eq $MemberDN
            if (-not $MemberCheck) {
                throw "Post-add verification failed: membership not present"
            }
        }
        $AddGroupMemberOutcome = 'Success'
        Write-LogFile -LogFile $LogFile -LogString "Add-GroupMember SUCCESS"
    } catch {
        $AddGroupMemberOutcome = 'Failed'
        Write-LogFile -LogFile $LogFile -LogString "ERROR in Add-GroupMember: $($_.Exception.Message)" -ForegroundColor Red
        throw
    } finally {
        # Do not allow final logging failure to mask the original result/exception.
        try {
            Write-LogFile -LogFile $LogFile -LogString "Add-GroupMember $AddGroupMemberOutcome"
            Write-LogFile -LogFile $LogFile -LogString "=== Add-GroupMember END ==="
        }
        catch {
            Write-Warning "Failed to write Add-GroupMember END log entry: $($_.Exception.Message)"
        }
    }
}
#====================================================================

#====================================================================
# Group removal function
#====================================================================
function Remove-GroupMember {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$LogFile
        ,[Parameter(Mandatory)][string]$DCHostName
        ,[Parameter(Mandatory)][string]$Group
        ,[Parameter(Mandatory)][string]$Member
    )
    #================================================================
    # Purpose:  Remove a DIRECT member (user OR group) from $Group.
    #           Symmetric with Add-GroupMember and sharing its resolver.
    #           Nested (transitive) members are intentionally not touched
    #           - Remove-ADGroupMember only removes direct members, and
    #           TTL-based nests expire server-side regardless.
    # Returns:  [string] outcome for the caller's audit:
    #             'Success'  - direct member removed and verified gone
    #             'NoChange' - member not found, or not a direct member
    #             'Rejected' - the name resolved to more than one object
    #           THROWS on a hard failure (the AD remove itself, or
    #           post-remove verification still finding the member), so
    #           the caller maps the exception to 'Failed'.
    #================================================================
    $RemoveGroupMemberOutcome = 'Failed'
    try {
        Write-LogFile -LogFile $LogFile -LogString "=== Remove-GroupMember START ==="
        Write-LogFile -LogFile $LogFile -LogString "Target Group: $Group"
        Write-LogFile -LogFile $LogFile -LogString "Target Member: $Member"
        $GroupObj = Get-ADGroup -Identity $Group -Server $DCHostName -ErrorAction Stop
        $resolution = Resolve-GroupMemberObject -Member $Member -DCHostName $DCHostName
        switch ($resolution.Status) {
            'NotFound' {
                Write-LogFile -LogFile $LogFile -LogString "Member '$Member' not found; nothing to remove."
                $RemoveGroupMemberOutcome = 'NoChange'
                return 'NoChange'
            }
            'Ambiguous' {
                Write-LogFile -LogFile $LogFile -LogString "Member '$Member' is ambiguous ($($resolution.Count) objects matched); supply a unique identifier or a distinguished name." -ForegroundColor Red
                $RemoveGroupMemberOutcome = 'Rejected'
                return 'Rejected'
            }
            'Found' {
                # Continue below.
            }
            default {
                throw "Unexpected Resolve-GroupMemberObject status '$($resolution.Status)' for member '$Member'."
            }
        }
        $MemberObj = $resolution.Object
        $MemberDN  = $null
        if ($MemberObj -and $MemberObj.PSObject.Properties.Match('DistinguishedName').Count -gt 0) {
            $MemberDN = [string]$MemberObj.DistinguishedName
        }
        if ([string]::IsNullOrWhiteSpace($MemberDN)) {
            throw "Resolve-GroupMemberObject returned Status='Found' but did not return an object with DistinguishedName for '$Member'."
        }
        # Direct membership only. Remove-ADGroupMember cannot remove a nested member.
        $DirectMember = Get-ADGroupMember -Identity $GroupObj.DistinguishedName -Server $DCHostName -ErrorAction SilentlyContinue | Where-Object DistinguishedName -eq $MemberDN
        if (-not $DirectMember) {
            Write-LogFile -LogFile $LogFile -LogString "$Member is not a direct member of $Group; nothing to remove."
            $RemoveGroupMemberOutcome = 'NoChange'
            return 'NoChange'
        }
        Remove-ADGroupMember -Identity $GroupObj.DistinguishedName -Members $MemberDN -Server $DCHostName -Confirm:$false -ErrorAction Stop
        # Post-remove verification: direct membership, DN-exact.
        $Still = Get-ADGroupMember -Identity $GroupObj.DistinguishedName -Server $DCHostName | Where-Object DistinguishedName -eq $MemberDN
        if ($Still) {
            throw "Post-remove verification failed: '$Member' still present in '$Group'"
        }
        Write-LogFile -LogFile $LogFile -LogString "Removed $Member from $Group."
        $RemoveGroupMemberOutcome = 'Success'
        return 'Success'
    } catch {
        $RemoveGroupMemberOutcome = 'Failed'
        Write-LogFile -LogFile $LogFile -LogString "ERROR in Remove-GroupMember: $($_.Exception.Message)" -ForegroundColor Red
        throw
    } finally {
        # Do not allow final logging failure to mask the original result/exception.
        try {
            Write-LogFile -LogFile $LogFile -LogString "Remove-GroupMember $RemoveGroupMemberOutcome"
            Write-LogFile -LogFile $LogFile -LogString "=== Remove-GroupMember END ==="
        }
        catch {
            Write-Warning "Failed to write Remove-GroupMember END log entry: $($_.Exception.Message)"
        }
    }
}
#====================================================================

#====================================================================
# OU creation function
#====================================================================
function New-ADOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][string]$OUName
        ,[Parameter(Mandatory)][String]$Path
        ,[String]$OUDescription
    )
    Write-LogFile -LogFile $LogFile -LogString "Creating OU $OUName"
    try {
        New-ADOrganizationalUnit -Name $OUName -Path $Path -ProtectedFromAccidentalDeletion:$true -Description $OUDescription -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Created OU $OUName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already in use") {
            Write-LogFile -LogFile $LogFile -LogString "'$OUName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
}
#====================================================================

#====================================================================
# Group creation function
#====================================================================
function New-DomainGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$GroupName
        ,[Parameter(Mandatory)][String]$GroupCategory
        ,[Parameter(Mandatory)][String]$GroupScope
        ,[Parameter(Mandatory)][ValidateSet("E","H","N")][String]$O365
        ,[Parameter(Mandatory)][Boolean]$HiddenFromAddressListsEnabled
        ,[Parameter(Mandatory)][String]$Path
        ,[String]$GroupDescription
    )
    Write-LogFile -LogFile $LogFile -LogString "Creating Group $GroupName"
    try {
        New-ADGroup -GroupCategory $GroupCategory -GroupScope $GroupScope -Name $GroupName -Path $Path -SamAccountName $GroupName -Server $DCHostName -Description $GroupDescription
        Set-ADObject -Identity "CN=$GroupName,$Path" -Server $DCHostName -ProtectedFromAccidentalDeletion $true
        Write-LogFile -LogFile $LogFile -LogString "Created Group $GroupName"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already exists") {
            Write-LogFile -LogFile $LogFile -LogString "'$GroupName' already exists" -ForegroundColor Green
        } else {
            throw
        }
    }
    if ($O365 -eq "E" -or $O365 -eq "H") {
        try {
            Enable-DistributionGroup -Identity $GroupName -DomainController $DCHostName
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        try {
            Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $HiddenFromAddressListsEnabled -RequireSenderAuthenticationEnabled $true -DomainController $DCHostName
        } catch {
            Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not configure $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
#====================================================================

#====================================================================
# Mailbox helper: verify a mailbox reached the expected recipient type
# within the wait window, or throw. Shared by the cloud and on-prem
# mailbox-update functions. -DomainController is threaded only for the
# on-prem path (EXO's Get-Mailbox has no such parameter).
#====================================================================
function Confirm-MailboxType {
    param(
        [Parameter(Mandatory)][string]$Upn
        ,[Parameter(Mandatory)][string]$ExpectedType
        ,[string]$DomainController
    )
    $gm = @{
        Identity = $Upn
        ErrorAction = 'SilentlyContinue'
    }
    if ($DomainController) {
        $gm['DomainController'] = $DomainController
    }
    $x = 0
    $mailbox = Get-Mailbox @gm
    $type = if ($mailbox) { $mailbox.RecipientTypeDetails } else { $null }
    while ($type -ne $ExpectedType -and $x -lt 6) {
        Start-Sleep -Seconds 10
        $mailbox = Get-Mailbox @gm
        $type = if ($mailbox) { $mailbox.RecipientTypeDetails } else { $null }
        $x++
    }
    if ($type -ne $ExpectedType) {
        throw "Mailbox '$Upn' did not convert to $ExpectedType (still '$type') after the wait window."
    }
}

#====================================================================
# Mailbox helper (CLOUD / EXO): idempotent FullAccess + SendAs grant.
# Pre-checks so a re-run is a logged no-op.
#====================================================================
function Grant-MailboxAccess {
    param(
        [Parameter(Mandatory)][string]$LogFile
        ,[Parameter(Mandatory)][string]$Upn
        ,[Parameter(Mandatory)][string]$GroupName
    )
    # FullAccess (idempotent)
    # Wrap in @() so a "no permissions found" / $null result becomes an empty array
    # rather than flowing a null object into the Where-Object script block.
    $mailboxPermissions = @(
        Get-MailboxPermission -Identity $Upn -User $GroupName -ErrorAction SilentlyContinue
    )
    $fullAccessExisting = $mailboxPermissions | Where-Object {
        if ($null -eq $_) {
            return $false
        }
        $hasAccessRights = $_.PSObject.Properties.Match('AccessRights').Count -gt 0
        if (-not $hasAccessRights) {
            return $false
        }
        $isInherited = if ($_.PSObject.Properties.Match('IsInherited').Count -gt 0) {
            [bool]$_.IsInherited
        } else {
            $false
        }
        $isDeny = if ($_.PSObject.Properties.Match('Deny').Count -gt 0) {
            [bool]$_.Deny
        } else {
            $false
        }
        $_.AccessRights -contains 'FullAccess' -and -not $isInherited -and -not $isDeny
    }
    if (-not $fullAccessExisting) {
        Add-MailboxPermission -Identity $Upn -User $GroupName -AccessRights FullAccess -Confirm:$false | Out-Null
        Write-LogFile -LogFile $LogFile -LogString "Granted FullAccess for $GroupName on $Upn"
    } else {
        Write-LogFile -LogFile $LogFile -LogString "FullAccess for $GroupName on $Upn already present; no change."
    }
    # SendAs (idempotent)
    # Same null-safe pattern as above.
    $recipientPermissions = @(
        Get-RecipientPermission -Identity $Upn -Trustee $GroupName -ErrorAction SilentlyContinue
    )
    $sendAsExisting = $recipientPermissions | Where-Object {
        if ($null -eq $_) {
            return $false
        }
        $hasAccessRights = $_.PSObject.Properties.Match('AccessRights').Count -gt 0
        if (-not $hasAccessRights) {
            return $false
        }
        $isDeny = if ($_.PSObject.Properties.Match('Deny').Count -gt 0) {
            [bool]$_.Deny
        } else {
            $false
        }
        $_.AccessRights -contains 'SendAs' -and -not $isDeny
    }
    if (-not $sendAsExisting) {
        Add-RecipientPermission -Identity $Upn -Trustee $GroupName -AccessRights SendAs -Confirm:$false | Out-Null
        Write-LogFile -LogFile $LogFile -LogString "Granted SendAs for $GroupName on $Upn"
    } else {
        Write-LogFile -LogFile $LogFile -LogString "SendAs for $GroupName on $Upn already present; no change."
    }
}

#====================================================================
# Mailbox helper (ON-PREM Exchange): mail-enable the access group, then
# idempotently grant FullAccess (Add-MailboxPermission) and Send-As.
# Send-As on-prem is an AD EXTENDED RIGHT, not a recipient permission -
# hence Add-ADPermission / Get-ADPermission, not Add/Get-RecipientPermission.
#====================================================================
function Grant-MailboxAccessOnPrem {
    param(
        [Parameter(Mandatory)][string]$LogFile
        ,[Parameter(Mandatory)][string]$Upn
        ,[Parameter(Mandatory)][string]$GroupName
        ,[Parameter(Mandatory)][string]$DomainController
    )
    # Mail-enable the access group (idempotent-tolerant: an already-enabled
    # group throws, which we downgrade to a warning - preserves prior behaviour).
    try {
        Enable-DistributionGroup -Identity $GroupName -DomainController $DomainController
        Set-DistributionGroup -Identity $GroupName -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $true -DomainController $DomainController
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "WARNING: Could not enable $GroupName - $($_.Exception.Message)" -ForegroundColor Yellow
    }
    # FullAccess (idempotent)
    # Wrap in @() so a "no permissions found" / $null result becomes an empty array
    # rather than flowing a null object into the Where-Object script block.
    $mailboxPermissions = @(
        Get-MailboxPermission -Identity $Upn -User $GroupName -DomainController $DomainController -ErrorAction SilentlyContinue
    )
    $fullAccessExisting = @(
        $mailboxPermissions | Where-Object {
            if ($null -eq $_) {
                $false
            } else {
                $hasAccessRights = $_.PSObject.Properties.Match('AccessRights').Count -gt 0
                if (-not $hasAccessRights) {
                    $false
                } else {
                    $isInherited = if ($_.PSObject.Properties.Match('IsInherited').Count -gt 0) {
                        [bool]$_.IsInherited
                    } else {
                        $false
                    }
                    $isDeny = if ($_.PSObject.Properties.Match('Deny').Count -gt 0) {
                        [bool]$_.Deny
                    } else {
                        $false
                    }
                    $_.AccessRights -contains 'FullAccess' -and -not $isInherited -and -not $isDeny
                }
            }
        }
    )
    if ($fullAccessExisting.Count -eq 0) {
        Add-MailboxPermission -Identity $Upn -User $GroupName -AccessRights FullAccess -Confirm:$false -DomainController $DomainController | Out-Null
        Write-LogFile -LogFile $LogFile -LogString "Granted FullAccess for $GroupName on $Upn"
    } else {
        Write-LogFile -LogFile $LogFile -LogString "FullAccess for $GroupName on $Upn already present; no change."
    }
    # Send-As (idempotent). NOTE: Get-ADPermission can surface this right as
    # 'Send-As', 'Send As', or 'SendAs'. Normalise whitespace/hyphen so all
    # forms compare as 'SendAs'.
    $adPermissions = @(
        Get-ADPermission -Identity $Upn -User $GroupName -DomainController $DomainController -ErrorAction SilentlyContinue
    )
    $sendAsExisting = @(
        $adPermissions | Where-Object {
            if ($null -eq $_) {
                $false
            } else {
                $hasExtendedRights = $_.PSObject.Properties.Match('ExtendedRights').Count -gt 0
                if (-not $hasExtendedRights) {
                    $false
                } else {
                    $isDeny = if ($_.PSObject.Properties.Match('Deny').Count -gt 0) {
                        [bool]$_.Deny
                    } else {
                        $false
                    }
                    $normalisedRights = @($_.ExtendedRights) | ForEach-Object {
                        ([string]$_) -replace '[\s-]', ''
                    }
                    $normalisedRights -contains 'SendAs' -and -not $isDeny
                }
            }
        }
    )
    if ($sendAsExisting.Count -eq 0) {
        Add-ADPermission -Identity $Upn -User $GroupName -AccessRights ExtendedRight -ExtendedRights "Send As" -Confirm:$false -DomainController $DomainController | Out-Null
        Write-LogFile -LogFile $LogFile -LogString "Granted Send-As for $GroupName on $Upn"
    } else {
        Write-LogFile -LogFile $LogFile -LogString "Send-As for $GroupName on $Upn already present; no change."
    }
}

#====================================================================
# Create mailbox function
#====================================================================
function New-UserMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[Parameter(Mandatory)][String]$EmailSuffix
        ,[Parameter(Mandatory)][String]$O365EmailSuffix
        ,[String]$realname,[String]$SharedEquipmentRoom,[Int]$Capacity
    )
    #================================================================
    # Purpose:      To create an Exchange Online Mailbox for a user account
    # Assumptions:  Parameters have been set correctly
    # Effects:      Mailbox should be created for user
    # Inputs:       $LogFile - String of log location passed to Write-LogFile
    #               $UserName - SAM account name of user
    #               $realname - Real Name to set as Primary SMTP address, read from CSV
    #               $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #               $Capacity - If mailbox is a room account, set the capacity
    # Calls:        Write-LogFile function
    # Returns:      $EnabledMailbox
    # Notes:        Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        $mbxParams = @{
            Identity             = $UserName
            alias                = $alias
            DomainController     = $DCHostName
            remoteroutingaddress = "$UserName@$O365EmailSuffix"
        }
        if ($realname) {
            $mbxParams['PrimarySmtpAddress'] = "$realname@$EmailSuffix"
        }
        switch ($SharedEquipmentRoom) {
            "S" { $mbxParams['shared']    = $true }
            "E" { $mbxParams['equipment'] = $true }
            "R" { $mbxParams['room']      = $true }
        }
        $smtp = if ($realname) { " -PrimarySmtpAddress $realname@$EmailSuffix" } else { "" }
        $flag = switch ($SharedEquipmentRoom) { "S" { " -shared" } "E" { " -equipment" } "R" { " -room" } default { "" } }
        $action = "Enable-RemoteMailbox -Identity $UserName$smtp -alias $alias -DomainController $DCHostName -remoteroutingaddress $UserName@$O365EmailSuffix$flag"
        Write-LogFile -LogFile $LogFile -LogString $Action
        if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
            Enable-RemoteMailbox @mbxParams
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-LogFile -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Creation Function"
}
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$UserPrincipalName
        ,[String]$SharedEquipmentRoom = "",[Int]$Capacity = 0
    )
    #================================================================
    # Purpose:      Update Mailbox parameters which need to be configured in O365
    # Assumptions:  Parameters have been set correctly
    # Effects:      Mailbox defaults should be assigned to the new mailbox
    # Inputs:       $LogFile - String of log location passed to Write-LogFile
    #               $UserPrincipalName - UPN of user
    #               $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #               $Capacity - If mailbox is a room account, set the capacity
    # Calls:        Write-LogFile function
    # Returns:
    # Notes:        Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Updating Mailbox"
    $MBX = $null
    try {
        $Env = Get-EnvironmentConfig
        $MBX = $null
        $i = 0
        while (-not $MBX -and $i -le 6) {
            $MBX = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
            if (-not $MBX) {
                Start-Sleep -Seconds 10
            }
            $i++
        }
        if ($MBX) {
            Write-LogFile -LogFile $LogFile -LogString " "
            Write-LogFile -LogFile $LogFile -LogString "Assigning region for $UserPrincipalName"
            Update-MgUser -UserId $UserPrincipalName -UsageLocation $Env.Locale.UsageLocation
            Set-MailboxSpellingConfiguration -Identity $UserPrincipalName -DictionaryLanguage $Env.Locale.Dictionary
            Set-MailboxRegionalConfiguration -Identity $UserPrincipalName -Language $Env.Locale.Language -DateFormat $Env.Locale.DateFormat -TimeFormat $Env.Locale.TimeFormat -TimeZone $Env.Locale.TimeZone
            $identityStr = $UserPrincipalName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to shared type"
                        Set-Mailbox -Identity $UserPrincipalName -Type Shared
                        Confirm-MailboxType -Upn $UserPrincipalName -ExpectedType 'SharedMailbox'
                    }
                    Grant-MailboxAccess -LogFile $LogFile -Upn $UserPrincipalName -GroupName "$($Env.Groups.SharedAccessPrefix)$UserPrincipalName"
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserPrincipalName -Type Equipment
                        Confirm-MailboxType -Upn $UserPrincipalName -ExpectedType 'EquipmentMailbox'
                    }
                    $identityStr = $UserPrincipalName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserPrincipalName : Updating Calendar Processing"
                    #Set calendar resource attendant to auto-accept
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Grant-MailboxAccess -LogFile $LogFile -Upn $UserPrincipalName -GroupName "$($Env.Groups.EquipmentAccessPrefix)$UserPrincipalName"
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserPrincipalName Mailbox to room type"
                        Set-Mailbox -Identity $UserPrincipalName -Type Room
                        Confirm-MailboxType -Upn $UserPrincipalName -ExpectedType 'RoomMailbox'
                    }
                    $identityStr = $UserPrincipalName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -confirm:$false
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserPrincipalName : Updating Calendar Processing"
                    #Set calendar resource attendant to auto-accept
                    Set-CalendarProcessing -Identity $UserPrincipalName -AutomateProcessing AutoAccept -confirm:$false
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserPrincipalName -ResourceCapacity $Capacity
                    }
                    Grant-MailboxAccess -LogFile $LogFile -Upn $UserPrincipalName -GroupName "$($Env.Groups.RoomAccessPrefix)$UserPrincipalName"
                }
            }
        } else {
            $logmsg = "Mailbox: " + $UserPrincipalName +" not found in AzureAD"
            Write-LogFile -LogFile $LogFile -LogString $LogMsg
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Create mailbox function
#====================================================================
function New-UserOnPremMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[Parameter(Mandatory)][String]$EmailSuffix
        ,[String]$realname,[String]$SharedEquipmentRoom,[Int]$Capacity
    )
    #================================================================
    # Purpose:      To create an Exchange On-Prem Mailbox for a user account
    # Assumptions:  Parameters have been set correctly
    # Effects:      Mailbox should be created for user
    # Inputs:       $LogFile - String of log location passed to Write-LogFile
    #               $UserName - SAM account name of user
    #               $realname - Real Name to set as Primary SMTP address, read from CSV
    #               $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #               $Capacity - If mailbox is a room account, set the capacity
    # Calls:        Write-LogFile function
    # Returns:      $EnabledMailbox
    # Notes:        Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Creating mailbox"
    $alias = $UserName.ToUpper()    #Alias = uppercase UserName
    try {
        $mbxParams = @{
            Identity         = $UserName
            alias            = $alias
            DomainController = $DCHostName
        }
        if ($realname) {
            $mbxParams['PrimarySmtpAddress'] = "$realname@$EmailSuffix"
        }
        switch ($SharedEquipmentRoom) {
            "S" { $mbxParams['shared']    = $true }
            "E" { $mbxParams['equipment'] = $true }
            "R" { $mbxParams['room']      = $true }
        }
        $smtp = if ($realname) { " -PrimarySmtpAddress $realname@$EmailSuffix" } else { "" }
        $flag = switch ($SharedEquipmentRoom) { "S" { " -shared" } "E" { " -equipment" } "R" { " -room" } default { "" } }
        $action = "Enable-Mailbox -Identity $UserName$smtp -alias $alias -DomainController $DCHostName$flag"
        Write-LogFile -LogFile $LogFile -LogString $Action
        if (-not (Get-Mailbox $UserName -ErrorAction SilentlyContinue)) {
            Enable-Mailbox @mbxParams
        }
        $mailboxExists = Get-Mailbox $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
        if (-not $mailboxExists) {
            Write-LogFile -LogFile $LogFile -LogString "ERROR: Mailbox for $UserName could not be created" -ForegroundColor Red
            throw
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Mailbox created for $UserName successfully"
            $EnabledMailbox = New-Object -Property @{"Alias" = ""; "SharedEquipmentRoom" = ""; "Capacity" = ""} -TypeName PSObject
            $EnabledMailbox.alias = $alias
            $EnabledMailbox.SharedEquipmentRoom = $SharedEquipmentRoom
            $EnabledMailbox.Capacity = $Capacity
            Return $EnabledMailbox
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to enable Mailbox or update settings"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Creation Function"
}
#====================================================================

#====================================================================
# Update mailbox Default Settings
#====================================================================
function Update-UserOnPremMailbox {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][String]$UserName
        ,[String]$SharedEquipmentRoom = "",[Int]$Capacity = 0
    )
    #================================================================
    # Purpose:      Update Mailbox parameters which need to be configured On-Prem
    # Assumptions:  Parameters have been set correctly
    # Effects:      Mailbox defaults should be assigned to the new mailbox
    # Inputs:       $LogFile - String of log location passed to Write-LogFile
    #               $UserName - SAM account name of user
    #               $SharedEquipmentRoom - Flag to set if mailbox is a human ID or a resouce or shared mailbox
    #               $Capacity - If mailbox is a room account, set the capacity
    # Calls:        Write-LogFile function
    # Returns:
    # Notes:        Uses Remote Powershell & Exchange Management shell snapin
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Updating Mailbox"
    $MBX = $null
    try {
        $Env = Get-EnvironmentConfig
        $MBX = $null
        $i = 0
        while (-not $MBX -and $i -le 6) {
            $MBX = Get-Mailbox -Identity $UserName -DomainController $DCHostName -ErrorAction SilentlyContinue
            if (-not $MBX) {
                Start-Sleep -Seconds 10
            }
            $i++
        }
        if ($MBX) {
            Set-MailboxSpellingConfiguration -Identity $UserName -DictionaryLanguage $Env.Locale.Dictionary -DomainController $DCHostName
            Set-MailboxRegionalConfiguration -Identity $UserName -Language $Env.Locale.Language -DateFormat $Env.Locale.DateFormat -TimeFormat $Env.Locale.TimeFormat -TimeZone $Env.Locale.TimeZone -DomainController $DCHostName
            $identityStr = $UserName + ":\Calendar"
            Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Reviewer -DomainController $DCHostName
            switch ($SharedEquipmentRoom) {
                "S" {
                    if ($MBX.RecipientTypeDetails -ne "SharedMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to shared type"
                        Set-Mailbox -Identity $UserName -Type Shared -DomainController $DCHostName
                        Confirm-MailboxType -Upn $UserName -ExpectedType 'SharedMailbox' -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Shared Mailbox $UserName : Adding Permissions"
                    Grant-MailboxAccessOnPrem -LogFile $LogFile -Upn $UserName -GroupName "$($Env.Groups.SharedAccessPrefix)$UserName" -DomainController $DCHostName
                }
                "E" {
                    if ($MBX.RecipientTypeDetails -ne "EquipmentMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to equipment type"
                        Set-Mailbox -Identity $UserName -Type Equipment -DomainController $DCHostName
                        Confirm-MailboxType -Upn $UserName -ExpectedType 'EquipmentMailbox' -DomainController $DCHostName
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -Confirm:$false -DomainController $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Updating Calendar Processing"
                    #Set calendar resource attendant to auto-accept
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -Confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Equipment Mailbox $UserName : Adding Permissions"
                    Grant-MailboxAccessOnPrem -LogFile $LogFile -Upn $UserName -GroupName "$($Env.Groups.EquipmentAccessPrefix)$UserName" -DomainController $DCHostName
                }
                "R" {
                    if ($MBX.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-LogFile -LogFile $LogFile -LogString "Converting $UserName Mailbox to room type"
                        Set-Mailbox -Identity $UserName -Type Room -DomainController $DCHostName
                        Confirm-MailboxType -Upn $UserName -ExpectedType 'RoomMailbox' -DomainController $DCHostName
                    }
                    $identityStr = $UserName + ":\Calendar"
                    Set-MailboxFolderPermission -Identity $identityStr -User Default -AccessRights Author -Confirm:$false -DomainController $DCHostName
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Updating Calendar Processing"
                    #Set calendar resource attendant to auto-accept
                    Set-CalendarProcessing -Identity $UserName -AutomateProcessing AutoAccept -Confirm:$false -DomainController $DCHostName
                    if ($Capacity) {
                        Set-Mailbox -Identity $UserName -ResourceCapacity $Capacity -DomainController $DCHostName
                    }
                    Write-LogFile -LogFile $LogFile -LogString "Updating Room Mailbox $UserName : Adding Permissions"
                    Grant-MailboxAccessOnPrem -LogFile $LogFile -Upn $UserName -GroupName "$($Env.Groups.RoomAccessPrefix)$UserName" -DomainController $DCHostName
                }
            }
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Failed to Complete Mailbox Update"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    }
    Write-LogFile -LogFile $LogFile -LogString "End of Mailbox Update Function"
}
#====================================================================

#====================================================================
# Bias-free random index in [0, $max) via rejection sampling.
# uint64 throughout so the threshold doesn't overflow when $max divides 2^32.
#====================================================================
function Get-CryptoRandIndex {
    param([int]$max)
    $rng = $null
    if ($max -le 1) { return 0 }
    try {
        $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        $bytes = [byte[]]::new(4)
        $rangeSize = [uint64]4294967296   # 2^32
        $threshold = $rangeSize - ($rangeSize % [uint64]$max)
        do {
            $rng.GetBytes($bytes)
            $r = [uint64][BitConverter]::ToUInt32($bytes, 0)
        } while ($r -ge $threshold)
        [int]($r % [uint64]$max)
    } finally {
        if ($rng) {
            $rng.Dispose()
        }
    }
}
#====================================================================

#====================================================================
# Generate a cryptographically random password 
#====================================================================
function New-Password {
    <#
    .SYNOPSIS
        Generate a cryptographically random password.
    .DESCRIPTION
        Drop-in replacement for [Web.Security.Membership]::GeneratePassword.
        Uses System.Security.Cryptography.RandomNumberGenerator with rejection
        sampling so each character draw is unbiased.
        
        Character sets exclude visually ambiguous characters (I, O, l, 0, 1)
        to reduce transcription errors when passwords are handed off to users
        verbally or in writing. Special character set avoids quote, backslash,
        pipe, backtick, slash and whitespace.
        
        Guarantees:
        - At least $MinLower lowercase, $MinUpper uppercase, $MinDigit digit,
          and $MinSpecial special characters
        - Fisher-Yates shuffle so required-class positions aren't predictable
        
        Compatible with Windows PowerShell 5.1 and PowerShell 7+; does not
        rely on [RandomNumberGenerator]::GetInt32 (.NET Core / 5+ only).
    .EXAMPLE
        $pw = New-Password
    .EXAMPLE
        $pw = New-Password -Length 32 -MinSpecial 6
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [ValidateRange(12, 256)][int]$Length     = 20
        ,[ValidateRange(0, 64)][int]$MinUpper    = 1
        ,[ValidateRange(0, 64)][int]$MinLower    = 1
        ,[ValidateRange(0, 64)][int]$MinDigit    = 1
        ,[ValidateRange(0, 64)][int]$MinSpecial  = 4
    )
    $minTotal = $MinUpper + $MinLower + $MinDigit + $MinSpecial
    if ($minTotal -gt $Length) {
        throw "Minimum character requirements ($minTotal) exceed total password length ($Length)."
    }
    # Ambiguous chars removed: I O l 0 1
    $upper   = [char[]]'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lower   = [char[]]'abcdefghijkmnopqrstuvwxyz'
    $digit   = [char[]]'23456789'
    $special = [char[]]'!@#$%^&*()-_=+[]{};:,.<>?'
    $all     = $upper + $lower + $digit + $special
    $chars = [System.Collections.Generic.List[char]]::new()
    # Required minimums
    for ($i = 0; $i -lt $MinUpper;   $i++) { $chars.Add($upper[(Get-CryptoRandIndex $upper.Length)]) }
    for ($i = 0; $i -lt $MinLower;   $i++) { $chars.Add($lower[(Get-CryptoRandIndex $lower.Length)]) }
    for ($i = 0; $i -lt $MinDigit;   $i++) { $chars.Add($digit[(Get-CryptoRandIndex $digit.Length)]) }
    for ($i = 0; $i -lt $MinSpecial; $i++) { $chars.Add($special[(Get-CryptoRandIndex $special.Length)]) }
    # Fill the rest from the union
    while ($chars.Count -lt $Length) {
        $chars.Add($all[(Get-CryptoRandIndex $all.Length)])
    }
    # Fisher-Yates shuffle so the required-class chars aren't at fixed positions
    for ($i = $chars.Count - 1; $i -gt 0; $i--) {
        $j = Get-CryptoRandIndex ($i + 1)
        $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
    }
    -join $chars
}
#====================================================================

#====================================================================
# Test password against password policy
#====================================================================
function Test-Password {
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$Password
        ,[Parameter(Mandatory)][Int]$PasswordLength
    )
    #================================================================
    # Purpose:          Test password against password policy
    # Assumptions:      Password has been generated with enough characters for required groups
    # Effects:          Password should be valid
    # Inputs:           $LogFile - String of log location passed to Write-LogFile
    #                   $Password
    #                   $PasswordLength
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:            There are 5 requirements in the current policy, but this could change in future
    #================================================================
    $TestsPassed = 0
    if ($Password.length -ge ($PasswordLength)) {$TestsPassed ++} # Must be >= 20 characters in length
    if ($Password -cmatch "[a-z]") {$TestsPassed ++} # Must contain a lowercase letter
    if ($Password -cmatch "[A-Z]") {$TestsPassed ++} # Must contain an uppercase letter
    if ($Password -cmatch "[0-9]") {$TestsPassed ++} # Must contain a digit
    if ($Password -cmatch "[^a-zA-Z0-9]") {$TestsPassed ++} # Must contain a special character
    if ($TestsPassed -ge 5) {
        Write-LogFile -LogFile $LogFile -LogString "Password validated"
        Write-LogFile -LogFile $LogFile -LogString " "
    } else {
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString "ERROR: Password does not comply with the password policy, skipping user" -ForegroundColor Red
        Write-LogFile -LogFile $LogFile -LogString ("-" * 80) -ForegroundColor Red
        throw "Password does not comply with the password policy"
    }
}
#====================================================================

#====================================================================
# GPO link function
#====================================================================
function Add-GPOLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][String]$DCHostName
        ,[Parameter(Mandatory)][string]$GPOName
        ,[Parameter(Mandatory)][string]$GPOTarget
    )
    #================================================================
    # Purpose:          To link a GPO to an OU
    # Assumptions:      Parameters have been set correctly
    # Effects:          GPO will be linked to the OU
    # Inputs:           $GPOName - Name of GPO as set before calling the function
    #                   $GPOTarget - OU where GPO will be linked
    # Returns:
    # Notes:
    #================================================================
    Write-LogFile -LogFile $LogFile -LogString "Linking $GPOName to $GPOTarget"
    try {
        New-GPLink -name $GPOName -target $GPOTarget -LinkEnabled Yes -enforced yes -Order 1 -ErrorAction Stop -Server $DCHostName
        Write-LogFile -LogFile $LogFile -LogString "Linked $GPOName to $GPOTarget"
    } catch {
        $ex = $_.Exception
        if ($ex.Message -match "already linked") {
            Write-LogFile -LogFile $LogFile -LogString "'$GPOName' already linked to $GPOTarget" -ForegroundColor Green
        } else {
            throw
        }
    }
}
#====================================================================

#====================================================================
# AD Sync
#====================================================================
function Invoke-ADSync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][pscredential]$Cred
        ,[Parameter(Mandatory)][String]$AzureADConnect
        ,[Parameter(Mandatory)][String]$O365EmailSuffix
    )
    $ADConnectSession = $null
    try {
        $ADConnectSession = New-PSSession -Computername $AzureADConnect -Credential $Cred
        Invoke-Command -Session $ADConnectSession {Import-Module ADSync}
        Import-PSSession -Session $ADConnectSession -Module ADSync -AllowClobber
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector is currently busy, waiting 30 seconds before trying again"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if ($ADSyncLoop -ge 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector has returned a busy state for 5 minutes or more, if this continues, please contact the servicedesk to investigate further"
        } else {
            Write-LogFile -LogFile $LogFile -LogString "Attempting to run Azure AD Sync Cycle"
            Start-ADSyncSyncCycle -PolicyType Delta
        }
        $state = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
        $ADSyncLoop = 0
        while ($State -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync Connector is busy, waiting 30 seconds To allow sync to complete"
            Start-Sleep -Seconds 30
            $State = (Get-ADSyncConnectorRunStatus | Where-Object { $_.RunspaceId -eq (Get-ADSyncConnector -Name "$O365EmailSuffix - AAD").runspaceid })
            $ADSyncLoop++
        }
        if (!($state) -and $ADSyncLoop -le 10) {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync complete"
        } else {
            Write-LogFile -LogFile $LogFile -LogString "AD Sync has not completed within 5 minutes, please check log for issues relating to syncronization issues."
        }
    } catch {
        $e = $_.Exception
        Write-LogFile -LogFile $LogFile -LogString $e -ForegroundColor Red
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-LogFile -LogFile $LogFile -LogString $line -ForegroundColor Red
        $msg = $e.Message
        Write-LogFile -LogFile $LogFile -LogString $msg -ForegroundColor Red
        $Action = "Unable to Sync AD"
        Write-LogFile -LogFile $LogFile -LogString $Action -ForegroundColor Red
    } finally {
        if ($ADConnectSession) { Remove-PSSession $ADConnectSession }
    }
}
#====================================================================

#====================================================================
# Build the schema GUID map (lDAPDisplayName -> schemaIDGUID).
# Used by callers needing GUIDs for AccessRule construction.
#====================================================================
function Get-ADSchemaGuidMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )
    $rootdse = Get-ADRootDSE -Server $Server
    $map = @{}
    $params = @{
        SearchBase = $rootdse.SchemaNamingContext
        LDAPFilter = '(schemaidguid=*)'
        Properties = ('lDAPDisplayName', 'schemaIDGUID')
    }
    Get-ADObject @params -Server $Server | ForEach-Object {
        $map[$_.lDAPDisplayName] = [System.GUID]$_.schemaIDGUID
    }
    return $map
}
#====================================================================

#====================================================================
# Build the extended rights map (displayName -> rightsGuid).
#====================================================================
function Get-ADExtendedRightsMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server
    )
    $rootdse = Get-ADRootDSE -Server $Server
    $map = @{}
    $params = @{
        SearchBase = $rootdse.ConfigurationNamingContext
        LDAPFilter = '(&(objectclass=controlAccessRight)(rightsguid=*))'
        Properties = ('displayName', 'rightsGuid')
    }
    Get-ADObject @params -Server $Server | ForEach-Object {
        $map[$_.displayName] = [System.GUID]$_.rightsGuid
    }
    return $map
}
#====================================================================

#====================================================================
# Delegate permission on computer objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty, WriteDacl, plus
# validated writes (DNS host name, SPN) and Reset/Change Password
# extended rights, all scoped to computer objects under $TargetOU.
#====================================================================
function Grant-ComputerJoinDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['computer'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['computer'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteDacl',$AccessControlTypeAllow,'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'Self',$AccessControlTypeAllow,$ExtendedRightsMap['Validated write to DNS host name'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'Self',$AccessControlTypeAllow,$ExtendedRightsMap['Validated write to service principal name'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['computer']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ExtendedRightAccessRule $AdminGroupSID,$AccessControlTypeAllow,$ExtendedRightsMap['Change Password'],'Descendents',$GuidMap['computer']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on group objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty on group objects.
#====================================================================
function Grant-GroupDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['group'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['group'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['group']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate modify membership permission on group objects to a group.
# Grants: WriteProperty on the 'member' attribute of group objects.
#====================================================================
function Grant-GroupMembershipEditDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['member'],'Descendents',$GuidMap['group']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate password reset permission on user objects to a group.
# Grants: WriteProperty on pwdLastSet / lockoutTime, plus Reset Password
# extended right, on user objects under $TargetOU.
#====================================================================
function Grant-PasswordResetDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['pwdLastSet'],'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,$GuidMap['lockoutTime'],'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['user']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on user objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty, plus Reset Password
# extended right, on user objects under $TargetOU.
#====================================================================
function Grant-UserDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['user'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['user'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['user']))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Reset Password'],'Descendents',$GuidMap['user']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permission on OU objects to a group.
# Grants: CreateChild, DeleteChild, WriteProperty on OU objects.
#====================================================================
function Grant-OUDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetOU
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $Acl = Get-Acl "AD:\OU=$TargetOU,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'CreateChild',$AccessControlTypeAllow,$GuidMap['organizationalUnit'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'DeleteChild',$AccessControlTypeAllow,$GuidMap['organizationalUnit'],'All'))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'WriteProperty',$AccessControlTypeAllow,'Descendents',$GuidMap['organizationalUnit']))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS Operators.
# Grants: GenericRead, GenericExecute, GenericWrite, CreateChild,
# DeleteChild Allow on the DNS container.
#====================================================================
function Grant-DNSOperatorsPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] 'GenericRead'
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] 'GenericExecute'
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] 'GenericWrite'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGW,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to DNS ReadOnly group.
# Grants: GenericRead, GenericExecute Allow.
# Denies: GenericWrite, CreateChild, DeleteChild, WriteOwner, WriteDacl,
# DeleteTree, Delete.
# Then strips the implicit Deny on ReadControl (which GenericWrite bit
# would otherwise drag in) so the group can still inspect the ACL.
#====================================================================
function Grant-DNSReadOnlyPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGR = [System.DirectoryServices.ActiveDirectoryRights] 'GenericRead'
    $adRightsGE = [System.DirectoryServices.ActiveDirectoryRights] 'GenericExecute'
    $adRightsGW = [System.DirectoryServices.ActiveDirectoryRights] 'GenericWrite'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $adRightsRC = [System.DirectoryServices.ActiveDirectoryRights] 'ReadControl'
    $adRightsWO = [System.DirectoryServices.ActiveDirectoryRights] 'WriteOwner'
    $adRightsWD = [System.DirectoryServices.ActiveDirectoryRights] 'WriteDacl'
    $adRightsDT = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteTree'
    $adRightsDEL = [System.DirectoryServices.ActiveDirectoryRights] 'Delete'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $AccessControlTypeDeny = [System.Security.AccessControl.AccessControlType] 'Deny'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGR,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGE,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGW,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWO,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWD,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDT,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDEL,$AccessControlTypeDeny,$inheritanceTypeAll))
    # GenericWrite includes the ReadControl bit, so its Deny ACE implicitly denies ReadControl too; strip that so the group can still inspect the ACL.
    $Acl.RemoveAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRC,$AccessControlTypeDeny,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate AD Sites/Subnets/Transports admin permission to a group.
# Grants: GenericAll, CreateChild, DeleteChild Allow on $TargetDN.
#====================================================================
function Grant-ADObjectPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsGA = [System.DirectoryServices.ActiveDirectoryRights] 'GenericAll'
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $adRightsDC = [System.DirectoryServices.ActiveDirectoryRights] 'DeleteChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsGA,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsDC,$AccessControlTypeAllow,$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate GPO link/option/RSoP permissions to a group.
# Grants: ReadProperty + WriteProperty on gPLink and gPOptions schema
# attributes plus the two RSoP extended rights, on the domain root.
#====================================================================
function Grant-GPOPermissionDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsRP = [System.DirectoryServices.ActiveDirectoryRights] 'ReadProperty'
    $adRightsWP = [System.DirectoryServices.ActiveDirectoryRights] 'WriteProperty'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'All'
    $Acl = Get-Acl "AD:\$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap['gPLink'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap['gPLink'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsRP,$AccessControlTypeAllow,$GuidMap['gPOptions'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsWP,$AccessControlTypeAllow,$GuidMap['gPOptions'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Generate Resultant Set of Policy (Logging)'],$inheritanceTypeAll))
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,'ExtendedRight',$AccessControlTypeAllow,$ExtendedRightsMap['Generate Resultant Set of Policy (Planning)'],$inheritanceTypeAll))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
# Delegate permissions to create GPOs.
# Grants: CreateChild on $TargetDN with no inheritance.
#====================================================================
function Grant-GPOCreationDelegation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminGroupName
        ,[Parameter(Mandatory)][string]$TargetDN
        ,[Parameter(Mandatory)][string]$BaseDN
        ,[Parameter(Mandatory)][hashtable]$GuidMap
        ,[Parameter(Mandatory)][hashtable]$ExtendedRightsMap
    )
    $AdminGroupSID = New-Object System.Security.Principal.SecurityIdentifier (Get-ADGroup $AdminGroupName).SID
    $adRightsCC = [System.DirectoryServices.ActiveDirectoryRights] 'CreateChild'
    $AccessControlTypeAllow = [System.Security.AccessControl.AccessControlType] 'Allow'
    $inheritanceTypeNone = [System.DirectoryServices.ActiveDirectorySecurityInheritance] 'None'
    $Acl = Get-Acl "AD:\$TargetDN,$BaseDN"
    $Acl.AddAccessRule((New-Object System.DirectoryServices.ActiveDirectoryAccessRule $AdminGroupSID,$adRightsCC,$AccessControlTypeAllow,$inheritanceTypeNone))
    $Acl | Set-Acl
}
#====================================================================

#====================================================================
function ConvertTo-IntOrDefault {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Value
        ,[int]$Default = 0
    )
    $result = $Default
    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        if (-not [int]::TryParse($Value.Trim(), [ref]$result)) {
            $result = $Default
        }
    }
    $result
}
#====================================================================

#====================================================================
function ConvertTo-SafeName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)][AllowEmptyString()][string]$Value
    )
    process {
        # Strip characters illegal in Office 365 name fields: ? @ \ +
        # (\\ is an escaped backslash inside the regex character class).
        ("$Value").Trim() -replace '[?@\\+]', [String]::Empty
    }
}
#====================================================================

#====================================================================
function ConvertTo-SafeSamAccountName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)][AllowEmptyString()][string]$Value
        # Trusted prefix applied AFTER the value is sanitised but BEFORE
        # truncation (mirrors CreateITAdminUser's "admin." + $UserName then cap).
        ,[Parameter()][string]$Prefix = ''
        ,[Parameter()][ValidateRange(1, 256)][int]$MaxLength = 20
    )
    process {
        # Keep only SAM-safe characters: letters, digits, dot, hyphen.
        $clean = $Prefix + (("$Value").Trim() -replace '[^A-Za-z0-9.-]', [String]::Empty)
        if ($clean.Length -gt $MaxLength) {
            $clean = $clean.Substring(0, $MaxLength)
        }
        $clean
    }
}
#====================================================================

#====================================================================
function Send-NotificationEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$LogFile
        ,[Parameter(Mandatory)][string]$SMTPServer
        ,[Parameter(Mandatory)][string]$EmailTo
        ,[Parameter(Mandatory)][string]$EmailFrom
        ,[Parameter(Mandatory)][string]$EmailSubject
        ,[Parameter(Mandatory)][string]$EmailBody
    )
    #================================================================
    # Purpose:          To send an email
    # Assumptions:      Parameters have been set correctly
    # Effects:          Email will be sent
    # Inputs:           $LogFile
    #                   $EmailTo
    #                   $EmailFrom
    #                   $EmailBody
    #                   $EmailSubject
    #                   $SMTPServer
    # Calls:            Write-LogFile function
    # Returns:
    # Notes:
    #================================================================
    Import-Module Send-MailKitMessage
    $RecipientList = [MimeKit.InternetAddressList]::new();
    $RecipientList.Add([MimeKit.InternetAddress]$EmailTo);
    $Splat = @{
        RecipientList   = $RecipientList
        From            = $EmailFrom
        Body            = $EmailBody
        Subject         = $EmailSubject
        SmtpServer      = $SMTPServer
        UseSecureConnectionIfAvailable = $true
    }
    try {
       Send-MailKitMessage @Splat
       Write-LogFile -LogFile $LogFile -LogString "Notification email sent to $EmailTo"
    } catch {
        Write-LogFile -LogFile $LogFile -LogString "ERROR sending notification email to $EmailTo : $_" -ForegroundColor Red
    }
}
#====================================================================
