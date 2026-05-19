Corp-Account-Management
==============

Scripts for setting up an example network

These scripts make a couple of assumptions:

1 - you have (at least) 2 VMs running fresh installs of Windows Server, at least 2022, with C: and D: drives:

a - one to be a "router" that has 2 network interfaces, 1 public and 1 private. 
    Make sure that the interface names are set correctly in `\Scripts\machine1\setup.ps1` & `\Scripts\machine1\dhcp.ps1`

b - one to be a DC/file server

2 - you copy the `\Scripts` folder to `C:\Scripts` on each machine

3 - when running the setup scripts the `-Domain` mandatory param is the NetBIOS name of the domain
    only, and the `-DomainSuffix` param is the external domain. EG if you supply `-Domain Corp` and
    `-DomainSuffix company.com` your AD domain will be called "corp.company.com". Ideally make the
    `-EmailSuffix` the same as `-DomainSuffix` to give your users a sensible email address and UPN values,
    EG username@company.com for the UPN and FirstName.LastName@company.com for the email address.
    Users can be added by modifying the `\Scripts\Users\users.csv` file.

4 - The machine that `azure_buildout.ps1` runs on needs the Az PowerShell module installed

5 - The following PowerShell modules will need to be installed:
    For `-O365 E` Send-MailKitMessage
    For `-O365 H` ExchangeOnlineManagement, Microsoft.Graph, Send-MailKitMessage

6 - `Scripts\environment.psd1` defines deployment defaults (OU names, group names, network topology, security limits, Entra
    roles, etc.) and must be reviewed/edited before running anything

7 - The GPO backup directory is not included in this repository - either build your own under Scripts/Prelim/GPOs/ 
    or comment out the Import-GPO and Add-GPOLink calls in DomainSetup.ps1 and setup-wsus.ps1 for your environment. 

Demo data note
`Scripts/Users/users.csv` is shipped as a worked example with deliberately recognisable placeholder names. One row 
(Dobby.House-Elf) has HIPRIV=Y and PrivLevel=1 so a reader can see the HiPriv account-creation mechanism end-to-end 
without editing the file. Running CreateUsers.ps1 against this CSV unedited will create, on top of the bare user account:

admin.Dobby.House-Elf - Tier-1 administrative account (workstation/server admin scope; no domain or Tier-0 reach).
ca.Dobby.House-Elf (only if also running with -O365) - Entra ID cloud admin with the Level1 role set: Helpdesk Administrator, 
Service Support Administrator, Global Reader. Notably not Authentication Administrator, not Domain Admin, not Global Admin.

Before running against any real directory, edit the CSV to use your own user data and set HIPRIV=N on any row where you don't 
want the admin accounts created. The full privilege ladder is defined in EntraRoles inside Scripts/environment.psd1 (Level1/2/3).

To Run:

1 - on the DC, run `\Scripts\Machine1\setup.ps1` with the `-Domain` parameter set to your NetBIOS domain, `-Platform` set 
    to either "Azure" or "Local", and `-Role` set to "DC1" (eg `.\setup.ps1 -Domain Corp -Platform Local -Role DC1`); 
    after the reboot run `\Scripts\Machine1\dcpromo.ps1`, enter a NetBIOS domain name and Domain Suffix when prompted 

2 - Once AD has been set up, on the DC run `\Scripts\Prelim\DomainSetup.ps1` and enter an email suffix and a drive letter 
    for the data drive, then run `CreateUsers.ps1` with `-O365 N` (no Office 365) and the same email suffix as before 

3 - on the router VM, open `\Scripts\Machine1\setup.ps1` & `\Scripts\machine1\dhcp.ps1` in notepad to check that the 
    network interfaces are set correctly

4 - on the router VM, run `\Scripts\Machine1\setup.ps1` with the `-Domain` parameter set to your NetBIOS 
    domain, `-Platform` set to either "Azure" or "Local",  and `-Role` set to "RTR"

5 - on the router VM, run `\Scripts\Machine1\dhcp.ps1` with the `-Domain` parameter set to your NetBIOS domain, 
    and `-Platform` set to either "Azure" or "Local"; you'll be prompted for domain admin credentials in a pop-up 
    to join the machine to the domain 

6 - on the router VM, rerun `\Scripts\Machine1\dhcp.ps1` a second time with the `-Domain` parameter set to your 
    NetBIOS domain, and `-Platform` set to either "Azure" or "Local" 

7 - download the contents of 
    https://www.dropbox.com/sh/n76ntsil0zjaapn/AAAEp7KrtLvLZk9_EFaExJIza?dl=0 
    and put it in `\\Domain\Share\Software`

8 - Optional roles, additional servers (CA, Exchange, WSUS, DC2/DC3) - see `\Scripts\Additional` for additional servers; 
    run the `setup.ps1`in `\Scripts\Machine1` with the relevant value for the `-Role` paramater to bootstrap, then their 
    corresponding `setup-*.ps1` / `additional-*.ps1` to configure 

9 - Optional role, WDS - if running on a local platform - on the router VM run `\Scripts\WDS\wds.ps1`, then mount ISOs 
    for Server 2022 and run the corresponding .PS1 script to set up the install images. You can now create 
    additional VMs that will do network installs. Machines need to have a private network adaptor
