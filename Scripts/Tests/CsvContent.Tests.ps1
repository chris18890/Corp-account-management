#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.5.0' }

# Per-row content validity for CSV columns that are load-bearing in
# user-creation scripts but unenforced at the script layer. CSV header
# validity is covered separately by ScriptParameters.Tests.ps1; this
# file pins the *value shape* of load-bearing columns.
#
# - users.csv USERNAME: not strictly enforced (CreateUsers strips
#   non-SAM characters via -replace '[^A-Za-z0-9.-]' and truncates to 20
#   chars), but pinning the convention here catches roster-entry typos
#   at test time rather than at user-creation time, when the resulting
#   SAM may quietly differ from what the requester expected.

BeforeAll {
    $script:scriptsRoot = Join-Path $PSScriptRoot '..'
    $csvPathStaff = Join-Path $script:scriptsRoot 'Users\users.csv'
    $script:StaffRows = Import-Csv -Path $csvPathStaff
    $csvPathMovers = Join-Path $script:scriptsRoot 'Users\movers.csv'
    $script:MoverRows = Import-Csv -Path $csvPathMovers
}

Describe 'users.csv USERNAME column validity' {
    It 'every USERNAME matches ^[A-Za-z0-9.-]{1,20}$ (SAM-safe, max 20 chars)' {
        $bad = @($script:StaffRows |
            Where-Object { $_.USERNAME -notmatch '^[A-Za-z0-9.\-]{1,20}$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' LASTNAME=$($_.LASTNAME)" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'users.csv REALNAME column validity' {
    # REALNAME becomes the left-hand side of the primary SMTP address:
    # "$RealName@$EmailSuffix". A value with @ signs, spaces, or characters
    # illegal in an email local-part would silently produce a malformed address.
    It 'every non-blank REALNAME is a valid email local-part (no spaces, no @)' {
        $bad = @($script:StaffRows |
            Where-Object { $_.REALNAME -and $_.REALNAME -notmatch '^[A-Za-z0-9._%+\-]{1,64}$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' REALNAME='$($_.REALNAME)'" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'users.csv S/E/R column validity' {
    # CreateUsers.ps1 calls .ToUpper() and switches on S / E / R / (default).
    # Any other non-blank value would fall through to the default branch and be
    # treated as a standard user even if that was not intended.
    It 'every non-blank S/E/R value is S, E, or R' {
        $bad = @($script:StaffRows |
            Where-Object { $_.'S/E/R' -and $_.'S/E/R' -notmatch '^[SERser]$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' S/E/R='$($_.'S/E/R')'" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'users.csv HIPRIV column validity' {
    # CreateUsers.ps1 calls .ToUpper() and checks -eq "Y".
    # Only Y (yes) or blank (no) are valid. Any other value is silently treated as N.
    It 'every non-blank HIPRIV value is Y' {
        $bad = @($script:StaffRows |
            Where-Object { $_.HIPRIV -and $_.HIPRIV -notmatch '^[Yy]$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' HIPRIV='$($_.HIPRIV)'" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'users.csv PrivLevel column validity' {
    # CreateITAdminUser.ps1 accepts ValidateSet(1,2,3).
    # A blank PrivLevel on a HIPRIV=Y row would fail silently or prompt interactively.
    It 'every HIPRIV=Y row has PrivLevel of 1, 2, or 3' {
        $bad = @($script:StaffRows |
            Where-Object { $_.HIPRIV -match '^[Yy]$' -and $_.'PrivLevel' -notmatch '^[123]$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' PrivLevel='$($_.PrivLevel)'" })
        $bad | Should -BeNullOrEmpty
    }
    It 'every non-blank PrivLevel is 1, 2, or 3' {
        $bad = @($script:StaffRows |
            Where-Object { $_.PrivLevel -and $_.PrivLevel -notmatch '^[123]$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' PrivLevel='$($_.PrivLevel)'" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'users.csv USERNAME uniqueness' {
    It 'no USERNAME appears more than once' {
        $dupes = @($script:StaffRows |
            Group-Object USERNAME |
            Where-Object Count -gt 1 |
            ForEach-Object { "USERNAME='$($_.Name)' count=$($_.Count)" })
        $dupes | Should -BeNullOrEmpty
    }
}

Describe 'movers.csv USERNAME column validity' {
    It 'every USERNAME matches ^[A-Za-z0-9.-]{1,20}$ (SAM-safe, max 20 chars)' {
        $bad = @($script:MoverRows |
            Where-Object { $_.USERNAME -notmatch '^[A-Za-z0-9.\-]{1,20}$' } |
            ForEach-Object { "USERNAME='$($_.USERNAME)' LASTNAME=$($_.LASTNAME)" })
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'movers.csv USERNAME uniqueness' {
    It 'no USERNAME appears more than once' {
        $dupes = @($script:MoverRows |
            Group-Object USERNAME |
            Where-Object Count -gt 1 |
            ForEach-Object { "USERNAME='$($_.Name)' count=$($_.Count)" })
        $dupes | Should -BeNullOrEmpty
    }
}
