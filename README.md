# Corp-account-management
==============

Scripts for setting up Network

These scripts make a couple of assumptions:

1 - you have a machine running fresh install of Windows Server, at least 2016 with a C: & D: drives
2 - you copy the \scripts folder to C:\Scripts on each machine
3 - when running the setup scripts the -Domain param is the NetBIOS name of the domain only, & the 
    -DomainSuffix param is the externaal domain. EG if you supply -Domain "Corp" & -DomainSuffix 
    "company.com" your AD domain will be called "corp.company.com". Ideally make the -EmailSuffix 
    the same as -DomainSuffix to give your users a sensible email addres & UPN values, EG 
    <username>@company.com for the UPN & FirstName.LastName@company.com for the email address.
    Users can be added by modifying the \Scripts\Users\Users.CSV file

To Run -
    1 - Copy the folder to C:\Scripts on the DC
    2 - run \Scripts\Prelim\DomainSetup.ps1 & enter an email suffix, CD to \Scripts\Users 
        & run CreateUsers.ps1, choose N for office365 & enter the same email suffix as before, 
        & a password when prompted
