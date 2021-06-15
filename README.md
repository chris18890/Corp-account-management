Corp-Account-Management
==============

Scripts for setting up an example Network

These scripts make a couple of assumptions:

1 - you have (at least) 2 VMs running fresh installs of Windows Server, at least 2022, with C: and D: drives:

a - one to be a "router" that has 2 network interfaces, 1 public and 1 private. 
    Make sure that the interface names are set correctly in `\Scripts\machine1\setup-router.ps1`
    
b - one to be a DC/file server

2 - you copy the `\Scripts` folder to `C:\Scripts` on each machine

3 - when running the setup scripts the `-Domain` mandatory param is the NetBIOS name of the domain 
    only, and the `-DomainSuffix` param is the external domain. EG if you supply `-Domain Corp` and 
    `-DomainSuffix company.com` your AD domain will be called "corp.company.com". Ideally make the 
    `-EmailSuffix` the same as `-DomainSuffix` to give your users a sensible email address and UPN values, 
    EG username@company.com for the UPN and FirstName.LastName@company.com for the email address.
    Users can be added by modifying the `\Scripts\Users\users.csv` file.

To Run:

1 - on the DC, run `\Scripts\Machine1\setup.ps1` with the `-Domain` parameter set to your NetBIOS domain, and 
    `-Platform` set to either "Azure" or "Local" (eg `.\setup.ps1 -Domain Corp -Platform Local`); after the reboot 
    run `\Scripts\Machine1\dcpromo.ps1`, enter a NetBIOS domain name and Domain Suffix when prompted 

2 - Once AD has been set up, on the DC run `\Scripts\Prelim\DomainSetup.ps1` and enter 
    an email suffix and a drive letter for the data drive, `CD` to `\Scripts\Users` and run `CreateUsers.ps1`, 
    with `-O365 N` (no Office 365) and the same email suffix as before

3 - on the router VM, open `\Scripts\Machine1\setup-router.ps1` in notepad to check that the 
    network interfaces are set

4 - on the router VM, run `\Scripts\Machine1\setup-router.ps1` with the `-Domain` parameter set to your NetBIOS 
    domain, and `-Platform` set to either "Azure" or "Local" 

5 - on the router VM, run `\Scripts\Machine1\dhcp.ps1` with the `-Domain` parameter set to your NetBIOS domain, 
    and `-Platform` set to either "Azure" or "Local"; you'll be prompted for domain admin credentials in a pop-up 
    to join the machine to the domain 

6 - on the router VM, rerun `\Scripts\Machine1\dhcp.ps1` a second time with the `-Domain` parameter set to your 
    NetBIOS domain, and `-Platform` set to either "Azure" or "Local" 

7 - download the contents of 
    https://www.dropbox.com/sh/n76ntsil0zjaapn/AAAEp7KrtLvLZk9_EFaExJIza?dl=0 
    and put it in `\\Domain\Share\Software`
    
8 - Optional roles, additional servers (CA, Exchange, WSUS, DC2/DC3) - see `\Scripts\ADCS`, `\Scripts\Exchange`, `\Scripts\WSUS`, `\Scripts\Machine2`, `\Scripts\Machine3` for 
    additional servers; run their `setup-*.ps1` to bootstrap, then the corresponding `install-*.ps1` to configure 
    
9 - Optional role, WDS - if running on a local platform - on the router VM run `\Scripts\WDS\wds.ps1`, 
    then mount ISOs for Server 2022 and Win10, and run the corresponding .PS1 script to set 
    up the install images. You can now create additional VMs that will do network installs. 
    Machines need to have a private network adaptor
