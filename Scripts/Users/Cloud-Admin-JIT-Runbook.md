# Cloud Admin Just-In-Time Elevation — Ops Runbook

## Why this exists

The tenant doesn't have Entra ID P2, so Privileged Identity Management (PIM) is unavailable. Cloud admin accounts (`ca.*` for Tier-0 cloud roles, `ga.*` for Global Administrator) are created by `CreateITCloudAdminUser.ps1` and `CreateITGlobalAdminUser.ps1` in a **disabled state by default**. Role assignments sit on a disabled account; the account is enabled only when the admin needs to do privileged work, then disabled again afterward.

This runbook formalises the enable/disable workflow and documents the wrapper scripts that automate it.

## Account naming

| Tier | Prefix | Purpose | Roles |
|---|---|---|---|
| Cloud | `ca.foo` | Tier-1 cloud administration | Whatever Entra roles `CreateITCloudAdminUser.ps1` assigned for the admin's privilege level |
| Global | `ga.foo` | Tier-0 Tenant-level emergencies only | Global Administrator |

The plain `admin.foo` (Hi-Priv on-prem) account is what an IT admin uses day-to-day. `ca.foo` and `ga.foo` are *separate accounts* held in reserve for cloud work — break-glass for `ga.*`, routine-but-elevated for `ca.*`.

## Workflow

### Enable an account before privileged work

```powershell
.\Enable-CloudAdmin.ps1 -UserName foo -EmailSuffix company.com -Tier Cloud -DurationMinutes 60 -Reason "Investigating tenant alert XYZ-123"
```

Parameters:
- **`-UserName`** — the bare username (no prefix). Script prefixes `ca.` or `ga.` based on `-Tier`.
- **`-EmailSuffix`** — same suffix used at account creation.
- **`-Tier`** — `Cloud` for `ca.*`, `Global` for `ga.*`.
- **`-DurationMinutes`** — how long to keep the account enabled. Defaults to `$Env.Security.MaxElevationMinutes` (currently 480 = 8 hours). Caps at the same value. Pass `30` or `60` for shorter, more PIM-like windows.
- **`-Reason`** — recorded in the audit trail. **Required** — if omitted on the command line, the script prompts.

What it does:
1. Verifies the caller is in `ADM_Task_HiPriv_Account_Admins` or `Domain Admins`.
2. Connects to Microsoft Graph with `User.ReadWrite.All` scope.
3. Sets `AccountEnabled = $true` on the target account in Entra ID.
4. Registers a one-time Windows Scheduled Task on the local workstation, set to fire at `Now + DurationMinutes`, which will run `Disable-CloudAdmin.ps1` for the same account.
5. Appends a row to `Scripts/Users/LogFiles/cloud-admin-elevations.csv`.

If anything fails after the Entra ID enable, the script rolls back (disables the account again) before throwing.

### Disable manually when work is done

```powershell
.\Disable-CloudAdmin.ps1 -UserName foo -EmailSuffix company.com -Tier Cloud -Reason "Investigation closed"
```

What it does:
1. Verifies the caller is in `ADM_Task_HiPriv_Account_Admins` or `Domain Admins`.
2. Sets `AccountEnabled = $false` on the target account in Entra ID.
3. Unregisters the pending auto-disable scheduled task if one exists.
4. Appends a row to the audit CSV.

Safe to run early. Safe to run if the account is already disabled (logs `AlreadyDisabled` to the audit trail). Safe to run when no auto-disable task is present (e.g., if the auto-disable already fired).

### Auto-disable

The scheduled task created by `Enable-CloudAdmin.ps1` calls `Disable-CloudAdmin.ps1` automatically at the deadline. The task is named `CloudAdminAutoDisable_<prefix><username>` and is visible in Windows Task Scheduler. Running `Enable-CloudAdmin.ps1` a second time for the same account replaces the existing task with one set to a new deadline.

## Audit trail

`Scripts/Users/LogFiles/cloud-admin-elevations.csv` is a single append-only CSV recording every enable and disable, with columns:

| Column | Notes |
|---|---|
| `Timestamp` | When the action ran. |
| `Action` | `Enable` or `Disable`. |
| `Operator` | `DOMAIN\username` of who ran the script. |
| `Account` | Full UPN of the target account. |
| `Tier` | `Cloud` or `Global`. |
| `DurationMinutes` | Enable only. The duration the operator requested (post-cap). |
| `DisableAt` | Enable only. The scheduled auto-disable time. |
| `Reason` | The justification supplied. |
| `Outcome` | `Success`, `AlreadyDisabled`, or `Failed`. |

Per-run logs also go to `Scripts/Users/LogFiles/<DOMAIN>_cloud_admin_enable_log-YYYYMMDD_N.log` and the equivalent `_disable_` file. The CSV is the durable trail; the per-run logs carry full detail including stack traces if anything failed.

## Authorisation model

Same as the user-creation scripts: a caller must be in `ADM_Task_HiPriv_Account_Admins` or `Domain Admins`. This means:
- An IT admin can elevate their own cloud accounts.
- An IT admin can also elevate another admin's account (e.g., for a help-desk-escalated emergency).
- The reason for the elevation is captured either way and lives in the audit CSV.

If you want stricter "you may only elevate your own account" enforcement, add a check that `$UserName -eq ($env:USERNAME -replace '^(admin\.|ca\.|ga\.)', '')` to `Enable-CloudAdmin.ps1`. The current design deliberately allows the more permissive flow because emergencies happen.

## Limitations vs. PIM

| Aspect | PIM | This workflow |
|---|---|---|
| Approval workflow | Optional approver before elevation | None — operator self-approves with `-Reason` |
| Auto-disable enforcement | Microsoft-side, can't be tampered with | Local Windows Scheduled Task; depends on the admin workstation being on at the deadline |
| Max duration | Configurable per role | Configurable in `environment.psd1` (`Security.MaxElevationMinutes`), uniformly applied |
| Audit trail | Entra audit log | Entra audit log (from the `Update-MgUser` calls) **plus** the local CSV |
| Hidden roles during off-time | Yes — JIT role activation | No — roles are permanently assigned; only the account is disabled |
| MFA on activation | Configurable | Falls back to whatever MFA policy applies at sign-in |
| Just-in-time access reviews | Built-in | Not provided |

If the maintainer ever licenses Entra ID P2 (it can be purchased per-user, not just tenant-wide), the equivalent PIM-eligible role assignments would replace this workflow for those licensed accounts. The scripts remain useful for any account that isn't covered.

## Auto-disable: what happens if the workstation is offline?

The scheduled task is registered with `-StartWhenAvailable`. If the workstation is off at the deadline:
- The task fires the next time the workstation is on and the task scheduler runs.
- The account stays enabled in the meantime.

If you need stronger guarantees, options are:
1. Run the scripts from a server that's always on (e.g., the DC). Tasks registered there will fire reliably. Requires the calling user to have Scheduled Task permission on that server.
2. Build a central sweeper: store the deadline in a custom Entra ID extension attribute on the account at enable time, and run a daily scheduled job somewhere that disables any account whose deadline has passed. More resilient but another piece of infrastructure to maintain.

For now the local-task approach is the documented default. Move to (1) or (2) if reliability becomes a concern.

## Day-to-day checklist

**Before privileged cloud work:**
1. Sign in to admin workstation as `admin.foo`.
2. `.\Enable-CloudAdmin.ps1 -UserName foo -EmailSuffix company.com -Tier Cloud -DurationMinutes 60 -Reason "..."`
3. Wait ~30 seconds for the Entra ID change to propagate.
4. Sign in to the target cloud admin portal as `ca.foo`.
5. Do the work.

**After (whenever the work finishes):**
6. `.\Disable-CloudAdmin.ps1 -UserName foo -EmailSuffix company.com -Tier Cloud -Reason "Work complete"`

If you forget step 6, the scheduled task disables the account at the deadline. The audit CSV will record the auto-disable.

## Verifying state

Quick check of all currently-enabled cloud admin accounts (run on any workstation with Microsoft.Graph installed):

```powershell
Connect-MgGraph -NoWelcome -Scopes "User.Read.All"
Get-MgUser -Filter "(startsWith(userPrincipalName,'ca.') or startsWith(userPrincipalName,'ga.')) and accountEnabled eq true" `
    -Property UserPrincipalName,DisplayName,AccountEnabled |
    Format-Table UserPrincipalName, DisplayName, AccountEnabled
```

Anything in that list outside the expected elevation windows is a leftover — investigate via the audit CSV.
