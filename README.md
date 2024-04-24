# Get-TenantInfo.ps1
This PowerShell script assesses a Microsoft 365 tenant. Please read the synopsis in the script to get more information.

This script assesses a Microsoft 365 tenant, providing with information (written on the shell) about hosted mailboxes (number and size), Azure groups (number),
Distribution Lists (number), One Drive (number and size), SharePoint sites (number and size), Teams site (number and size) and users (number). You can choose 
what must be analyzed by specifying the correct switches.
You can specify also a domain registered on the tenant to get only the objects with that domain (i.e, mailboxes with that domain as primary SMTP and/or groups/
SharePoint/Teams sites having a user with that domain as owner).

Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/
