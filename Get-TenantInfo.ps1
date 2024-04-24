﻿<#
.SYNOPSIS
Get-TenantInfo.ps1 - This script assesses a Microsoft 365 tenant.

.DESCRIPTION 
This script assesses a Microsoft 365 tenant, providing with information (written on the shell) about hosted mailboxes (number and size), Azure groups (number),
Distribution Lists (number), One Drive (number and size), SharePoint sites (number and size), Teams site (number and size) and users (number). You can choose 
what must be analyzed by specifying the correct switches.
You can specify also a domain registered on the tenant to get only the objects with that domain (i.e, mailboxes with that domain as primary SMTP and/or groups/
SharePoint/Teams sites having a user with that domain as owner).

.INPUTS
You must have valid credentials for connecting to Exchange Online, AzureAD, SharePoint, and Teams.

.OUTPUTS
The result of the assessment will be displayed directly in the shell.

.PARAMETER Domain
Insert the domain (e.g., "@contoso.com") to filter the objects related to that domain.

.PARAMETER Mailbox
Insert this switch if you want mailboxes data.

.PARAMETER DL
Insert this switch if you want Distribution Lists data.

.PARAMETER Groups
Insert this switch if you want Groups (Microsoft 365 and Security) data.

.PARAMETER OneDrive
Insert this switch if you want OneDrive data.

.PARAMETER SharePoint
Insert this switch if you want SharePoint data.

.PARAMETER Teams
Insert this switch if you want Teams data.

.PARAMETER All
Insert this switch if you want all the information.

.EXAMPLE
.\Get-TenantInfo.ps1 -All --> Get all the data about all the objects in the tenant;
.\Get-TenantInfo.ps1 -All -Domain "@contoso.com" --> Get all the data about the objects with the @contoso.com domain;
.\Get-TenantInfo.ps1 -Mailbox -OneDrive -Domain "@contoso.com" --> Get Mailbox and OneDrive data about the objects with the @contoso.com domain;
.\Get-TenantInfo.ps1 -DL -Groups -Domain "@contoso.com" --> Get Distribution Lists and Groups data about the objects with the @contoso.com domain;
.\Get-TenantInfo.ps1 -Teams -SharePoint "@contoso.com" --> Get Teams and SharePoint data about the objects with the @contoso.com domain.

.NOTES
Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/

#>

param(
    [Parameter(ParameterSetName='Category')]
    [string]$Domain,

	[Parameter(ParameterSetName='Category')]
    [switch]$Mailbox,

    [Parameter(ParameterSetName='Category')]
    [switch]$DL,

    [Parameter(ParameterSetName='Category')]
    [switch]$Groups,

    [Parameter(ParameterSetName='Category')]
    [switch]$OneDrive,

    [Parameter(ParameterSetName='Category')]
    [switch]$SharePoint,

    [Parameter(ParameterSetName='Category')]
    [switch]$Teams,

	[Parameter(ParameterSetName='Category')]
    [switch]$All
)

Clear-Host

#Verify what you have selected

if($Mailbox -or $All){
    $MailboxChoice = "TRUE"
    }
else{
    $MailboxChoice = "FALSE"
}

if($DL -or $All){
    $DLChoice = "TRUE"
    }
else{
    $DLChoice = "FALSE"
}

if($Groups -or $All){
    $GroupsChoice = "TRUE"
    }
else{
    $GroupsChoice = "FALSE"
}

if($OneDrive -or $All){
    $ODChoice = "TRUE"
    }
else{
    $ODChoice = "FALSE"
}

if($SharePoint -or $All){
    $SPChoice = "TRUE"
    }
else{
    $SPChoice = "FALSE"
}

if($Teams -or $All){
    $TeamsChoice = "TRUE"
    }
else{
    $TeamsChoice = "FALSE"
}

Write-Host "You are going to proceed with the following exports:" -ForegroundColor Cyan
Write-Host "Mailboxes - $($MailboxChoice)" -ForegroundColor Cyan
Write-Host "DL - $($DLChoice)" -ForegroundColor Cyan
Write-Host "Groups - $($GroupsChoice)" -ForegroundColor Cyan
Write-Host "OneDrive - $($ODChoice)" -ForegroundColor Cyan
Write-Host "SharePoint - $($SPChoice)" -ForegroundColor Cyan
Write-Host "Teams - $($TeamsChoice)" -ForegroundColor Cyan
$Confirmation = Read-Host "Do you want to proceed? [Y/N] (default is Yes)"

if ($Confirmation -eq "N"){
    exit
}

$SPURL = Read-Host "Insert the sharepoint admin Url (e.g., https://contoso-admin.sharepoint.com) since it will be used later in the export"

# Connection to M365 services (Exchange, AzureAD, SharePoint, and Teams)

try{
    #Connect-ExchangeOnline -ErrorAction Stop
    Import-Module AzureAD -UseWindowsPowerShell
    Connect-AzureAD -ErrorAction Stop
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    Connect-SPOService -Url $SPURL
    Connect-MicrosoftTeams
    Write-Host "Connected to all M365 services!" -ForegroundColor Green
    Write-Host ""
}
catch{
    Write-Host "Connection to M365 services failed! Check if you have all the required installed modules" -ForegroundColor Red
    exit
}

Clear-Host

if($Mailbox -or $All){
    
    Write-Host "Start to fetch Mailbox Data" -ForegroundColor White
    Write-Host ""

    if($Domain){
        
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $PrimaryUserMailboxes = $AllMailboxes | Where-Object {$_.PrimarySMTPAddress -match $Domain -and $_.RecipientTypeDetails -eq "UserMailbox"}
        $ArchiveUserMailboxes = Get-Mailbox -ResultSize Unlimited -Archive | Where-Object {$_.PrimarySMTPAddress -match $Domain -and $_.RecipientTypeDetails -eq "UserMailbox"}
        $SharedMailboxes = $AllMailboxes | Where-Object {$_.PrimarySMTPAddress -match $Domain -and $_.RecipientTypeDetails -eq "SharedMailbox"}
        $TotPrimaryUserMailboxes = $PrimaryUserMailboxes.count #To Export
        $TotArchiveUserMailboxes = $ArchiveUserMailboxes.count #To Export
        $TotSharedMailboxes = $SharedMailboxes.count #To Export

        $PrimaryUserMailboxSizeTot = 0
        Foreach($PrimaryUserMailbox in $PrimaryUserMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $PrimaryUserMailboxData = Get-MailboxStatistics $PrimaryUserMailbox.PrimarySMTPAddress
            $PrimaryUserMailboxSize = [math]::Round(($PrimaryUserMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $PrimaryUserMailboxSizeTot =  $PrimaryUserMailboxSizeTot + $PrimaryUserMailboxSize
            $ErrorActionPreference = "Continue"
        }
        if($PrimaryUserMailboxes.count -ne 0){
            $PrimaryUserMailboxSizeMean = $PrimaryUserMailboxSizeTot/ $PrimaryUserMailboxes.count
            [string]$PrimaryUserMailboxSizeMean = $PrimaryUserMailboxSizeMean.ToString() + " GB" #ToExport
        }
        [string]$PrimaryUserMailboxSizeTot = $PrimaryUserMailboxSizeTot.ToString() + " GB" #ToExport

        $ArchiveUserMailboxSizeTot = 0
        Foreach($ArchiveUserMailbox in $ArchiveUserMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $ArchiveUserMailboxData = Get-MailboxStatistics $ArchiveUserMailbox.PrimarySMTPAddress
            $ArchiveUserMailboxSize = [math]::Round(($ArchiveUserMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $ArchiveUserMailboxSizeTot =  $ArchiveUserMailboxSizeTot + $ArchiveUserMailboxSize
            $ErrorActionPreference = "Continue"
        }
        $TotArchiveUserMailboxes
        if($ArchiveUserMailboxes.count -ne 0){
            $ArchiveUserMailboxSizeMean = $ArchiveUserMailboxSizeTot/ $ArchiveUserMailboxes.count
            [string]$ArchiveUserMailboxSizeMean = $ArchiveUserMailboxSizeMean.ToString() + " GB" #ToExport
        }
        $TotArchiveUserMailboxes
        [string]$ArchiveUserMailboxSizeTot = $ArchiveUserMailboxSizeTot.ToString() + " GB" #ToExport

        $SharedMailboxSizeTot = 0
        Foreach($SharedMailbox in $SharedMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $SharedMailboxData = Get-MailboxStatistics $SharedMailbox.PrimarySMTPAddress
            $SharedMailboxSize = [math]::Round(($SharedMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $SharedMailboxSizeTot =  $SharedMailboxSizeTot + $SharedMailboxSize
            $ErrorActionPreference = "Continue"
        }
        if($SharedMailboxes.count -ne 0){
            $SharedMailboxSizeMean = $SharedMailboxSizeTot/ $SharedMailboxes.count
            [string]$SharedMailboxSizeMean = $SharedMailboxSizeMean.ToString() + " GB" #ToExport
        }
        [string]$SharedMailboxSizeTot = $SharedMailboxSizeTot.ToString() + " GB" #ToExport

        Write-Host "The tenant hosts $($TotPrimaryUserMailboxes) Primary User Mailboxes of $($Domain) users" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($TotArchiveUserMailboxes) Archive User Mailboxes of $($Domain) users" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($TotSharedMailboxes) Shared Mailboxes of $($Domain) user" -ForegroundColor Yellow
        Write-Host "The total size of the $($Domain) Primary User Mailboxes is $($PrimaryUserMailboxSizeTot)" -ForegroundColor Yellow
        if($PrimaryUserMailboxSizeMean -and $TotPrimaryUserMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Primary User Mailboxes is $($PrimaryUserMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host "The total size of the $($Domain) Archive User Mailboxes is $($ArchiveUserMailboxSizeTot)" -ForegroundColor Yellow
        if($ArchiveUserMailboxSizeMean -and $TotArchiveUserMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Archive User Mailboxes is $($ArchiveUserMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host "The total size of the $($Domain) Shared Mailboxes is $($SharedMailboxSizeTot)" -ForegroundColor Yellow
        if($SharedMailboxSizeMean -and $TotSharedMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Shared Mailboxes is $($SharedMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host ""

    }
    else{
        
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $PrimaryUserMailboxes = $AllMailboxes | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
        $ArchiveUserMailboxes =  Get-Mailbox -ResultSize Unlimited -Archive | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
        $SharedMailboxes = $AllMailboxes | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
        $TotPrimaryUserMailboxes = $PrimaryUserMailboxes.count #To Export
        $TotArchiveUserMailboxes = $ArchiveUserMailboxes.count #To Export
        $TotSharedMailboxes = $SharedMailboxes.count #To Export

        $PrimaryUserMailboxSizeTot = 0
        Foreach($PrimaryUserMailbox in $PrimaryUserMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $PrimaryUserMailboxData = Get-MailboxStatistics $PrimaryUserMailbox.PrimarySMTPAddress
            $PrimaryUserMailboxSize = [math]::Round(($PrimaryUserMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $PrimaryUserMailboxSizeTot =  $PrimaryUserMailboxSizeTot + $PrimaryUserMailboxSize
            $ErrorActionPreference = "Continue"
        } 
        if($PrimaryUserMailboxes.count -ne 0){
            $PrimaryUserMailboxSizeMean = $PrimaryUserMailboxSizeTot/ $PrimaryUserMailboxes.count
            [string]$PrimaryUserMailboxSizeMean = $PrimaryUserMailboxSizeMean.ToString() + " GB" #ToExport
        }
        [string]$PrimaryUserMailboxSizeTot = $PrimaryUserMailboxSizeTot.ToString() + " GB" #ToExport

        $ArchiveUserMailboxSizeTot = 0
        Foreach($ArchiveUserMailbox in $ArchiveUserMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $ArchiveUserMailboxData = Get-MailboxStatistics $ArchiveUserMailbox.PrimarySMTPAddress
            $ArchiveUserMailboxSize = [math]::Round(($ArchiveUserMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $ArchiveUserMailboxSizeTot =  $ArchiveUserMailboxSizeTot + $ArchiveUserMailboxSize
            $ErrorActionPreference = "Continue"
        }
        if($ArchiveUserMailboxes.count -ne 0){
            $ArchiveUserMailboxSizeMean = $ArchiveUserMailboxSizeTot/ $ArchiveUserMailboxes.count
            [string]$ArchiveUserMailboxSizeMean = $ArchiveUserMailboxSizeMean.ToString() + " GB" #ToExport
        }
        [string]$ArchiveUserMailboxSizeTot = $ArchiveUserMailboxSizeTot.ToString() + " GB" #ToExport

        $SharedMailboxSizeTot = 0
        Foreach($SharedMailbox in $SharedMailboxes){
            $ErrorActionPreference = "SilentlyContinue"
            $SharedMailboxData = Get-MailboxStatistics $SharedMailbox.PrimarySMTPAddress
            $SharedMailboxSize = [math]::Round(($SharedMailboxData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            $SharedMailboxSizeTot =  $SharedMailboxSizeTot + $SharedMailboxSize
            $ErrorActionPreference = "Continue"
        }
        if($SharedMailboxes.count -ne 0){
            $SharedMailboxSizeMean = $SharedMailboxSizeTot/ $SharedMailboxes.count
            [string]$SharedMailboxSizeMean = $SharedMailboxSizeMean.ToString() + " GB" #ToExport
        }
        [string]$SharedMailboxSizeTot = $SharedMailboxSizeTot.ToString() + " GB" #ToExport

        Write-Host "The tenant hosts $($TotPrimaryUserMailboxes) Primary User Mailboxes" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($TotArchiveUserMailboxes) Archive User Mailboxes" -ForegroundColor Yellow
        Write-Host "The total size of the $($Domain) Primary User Mailboxes is $($PrimaryUserMailboxSizeTot)" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($TotSharedMailboxes) Shared Mailboxes" -ForegroundColor Yellow
        if($PrimaryUserMailboxSizeMean -and $TotPrimaryUserMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Primary User Mailboxes is $($PrimaryUserMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host "The total size of the $($Domain) Archive User Mailboxes is $($ArchiveUserMailboxSizeTot)" -ForegroundColor Yellow
        if($ArchiveUserMailboxSizeMean -and $TotArchiveUserMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Archive User Mailboxes is $($ArchiveUserMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host "The total size of the $($Domain) Shared Mailboxes is $($SharedMailboxSizeTot)" -ForegroundColor Yellow
        if($SharedMailboxSizeMean -and $TotSharedMailboxes -ne 0){
            Write-Host "The average size of the $($Domain) Shared Mailboxes is $($SharedMailboxSizeMean)" -ForegroundColor Yellow
        }
        Write-Host ""

    }

    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Write-Host "Mailbox Export Data Finished!" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    
}

if($DL -or $All){

    Write-Host "Start to fetch DLs Data" -ForegroundColor White
    Write-Host ""

    if($Domain){
        $DistributionLists = Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.PrimarySMTPAddress -match $Domain}
        $TotalNumberofDL = $DistributionLists.count #To Export

        Write-Host "The tenant hosts $($TotalNumberofDL) Distribution Lists having as owner a $($Domain) user" -ForegroundColor Yellow
        Write-Host ""

    }

    else{
        $DistributionLists = Get-DistributionGroups
        $TotalNumberofDL = $DistributionLists.count #To Export

        Write-Host "The tenant hosts $($TotalNumberofDL) Distribution Lists" -ForegroundColor Yellow
        Write-Host ""

    }

    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host "DLs Export Data Finished!" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan

}

if($Groups -or $All){

    Write-Host "Start to fetch Groups Data" -ForegroundColor White
    Write-Host ""

    $AllGroups = Get-AzureADGroup -All:$True | Where-Object {$_.DirSyncEnabled -ne $True}

    $SecurityGroups = $AllGroups | Where-Object {$_.SecurityEnabled -eq $True}
    $SecurityGroupsCount = $SecurityGroups.count
    $MailEnabledSecurityGroup = $AllGroups | Where-Object {$_.SecurityEnabled -eq $True -and $_.MailEnabled -eq $True}
    $MailEnabledSecurityGroupCount = $MailEnabledSecurityGroup.count
    $M365Groups = $AllGroups | Where-Object {$_.SecurityEnabled -eq $False}

    if ($Domain){
        $M365GroupsCount = 0
        Foreach($M365Group in $M365Groups){
            $Owner = Get-AzureADGroupOwner -ObjectId $M365Group.ObjectID
            if($Owner -match $Domain){
                $M365GroupsCount = $M365GroupsCount + 1
            }
        }

        Write-Host "The tenant hosts $($SecurityGroupsCount) Security Groups" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($MailEnabledSecurityGroupCount) Mail-Enabled Security Groups" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($M365GroupsCount) M365 Groups having a sowner a $($Domain) user" -ForegroundColor Yellow

    }
    else{

        $M365Groups = $AllGroups | Where-Object {$_.SecurityEnabled -eq $False}
        $M365GroupsCount = $M365Groups.count

        Write-Host "The tenant hosts $($SecurityGroupsCount) Security Groups" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($MailEnabledSecurityGroupCount) Mail-Enabled Security Groups" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($M365GroupsCount) M365 Groups" -ForegroundColor Yellow

    }

    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Groups Export Data Finished!" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan 

}

if($OneDrive -or $All){
    
    Write-Host "Start to fetch OneDrive Data" -ForegroundColor White
    Write-Host ""

    if($Domain){

    $ODURLs = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" | Where-Object {$_.Owner -match $Domain}
    $TotalNumberODSite = $ODURLs.count #To Export
    $ODTotalSpace = 0

        Foreach($ODURL in $ODURLs){
            $SingleODUsedSpace = [Math]::Round($ODURL.StorageUsageCurrent / 1024, 3)
            $ODTotalSpace = $ODTotalSpace + $SingleODUsedSpace
        }

    $ODSpaceMean = $ODTotalSpace / $TotalNumberODSite
    [string]$ODTotalSpace = $ODTotalSpace.ToString() + " GB" #ToExport 
    [string]$ODSpaceMean = $ODSpaceMean.ToString() + " GB" #ToExport

    Write-Host "The tenant hosts $($TotalNumberODSite) OneDriveSite owned by a $($Domain) user" -ForegroundColor Yellow
    Write-Host "The total size of the OneDrive Sites owned by a $($Domain) user is $($ODTotalSpace)" -ForegroundColor Yellow
    Write-Host "The average size of the OneDrive Sites owned by a $($Domain) user is $($ODSpaceMean)" -ForegroundColor Yellow
    Write-Host ""

    }

    else{

    $ODURLs = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
    $TotalNumberODSite = $ODURLs.count #To Export
    $ODTotalSpace = 0

        Foreach($ODURL in $ODURLs){
            $SingleODUsedSpace = [Math]::Round($ODURL.StorageUsageCurrent / 1024, 3)
            $ODTotalSpace = $ODTotalSpace + $SingleODUsedSpace
        }

    $ODSpaceMean = $ODTotalSpace / $TotalNumberODSite
    [string]$ODTotalSpace = $ODTotalSpace.ToString() + " GB" #ToExport 
    [string]$ODSpaceMean = $ODSpaceMean.ToString() + " GB" #ToExport

    Write-Host "The tenant hosts $($TotalNumberODSite) OneDriveSite" -ForegroundColor Yellow
    Write-Host "The total size of the OneDrive Sites is $($ODTotalSpace)" -ForegroundColor Yellow
    Write-Host "The average size of the OneDrive Sites is $($ODSpaceMean)" -ForegroundColor Yellow
    Write-Host ""

    }

    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host "OneDrive Export Data Finished!" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
}

if($SharePoint -or $All){
    
    Write-Host "Start to fetch SharePoint Site Data" -ForegroundColor White
    Write-Host ""

    if($Domain){

        $SPSites = Get-SPOSite -Limit ALL | Where-Object {$_.Url -notmatch "my.sharepoint.com/personal" -and $_.IsTeamsConnected -eq $False -and $_.owner -match $Domain}
        $TotalNumberSPSites = $SPSites.count #To Export

        $SPTotalSpace = 0
        Foreach($SPSite in $SPSites){
            $SingleSPUsedSpace = [Math]::Round($SPSite.StorageUsageCurrent / 1024, 3)
            $SPTotalSpace = $SPTotalSpace + $SingleSPUsedSpace
        }

        $SPSpaceMean = $SPTotalSpace / $TotalNumberSPSites
        [string]$SPTotalSpace = $SPTotalSpace.ToString() + " GB" #ToExport 
        [string]$SPSpaceMean = $SPSpaceMean.ToString() + " GB" #ToExport

        Write-Host "The tenant hosts $($TotalNumberSPSites) SharePoint Sites having as owner a $($Domain) user" -ForegroundColor Yellow
        Write-Host "The total size of the SharePoint Sites having as owner a $($Domain) user is $($SPTotalSpace)" -ForegroundColor Yellow
        Write-Host "The average size of the SharePoint Sites having as owner a $($Domain) user is $($SPSpaceMean)" -ForegroundColor Yellow
        Write-Host ""

    }

    else{

        $SPSites = Get-SPOSite -Limit ALL | Where-Object {$_.Url -notmatch "my.sharepoint.com/personal" -and $_.IsTeamsConnected -eq $False}
        $TotalNumberSPSites = $SPSites.count #To Export

        $SPTotalSpace = 0
        Foreach($SPSite in $SPSites){
            $SingleSPUsedSpace = [Math]::Round($SPSite.StorageUsageCurrent / 1024, 3)
            $SPTotalSpace = $SPTotalSpace + $SingleSPUsedSpace
        }

        $SPSpaceMean = $SPTotalSpace / $TotalNumberSPSites
        [string]$SPTotalSpace = $SPTotalSpace.ToString() + " GB" #ToExport 
        [string]$SPSpaceMean = $SPSpaceMean.ToString() + " GB" #ToExport

        Write-Host "The tenant hosts $($TotalNumberSPSites) SharePoint Sites" -ForegroundColor Yellow
        Write-Host "The total size of the SharePoint Sites is $($SPTotalSpace)" -ForegroundColor Yellow
        Write-Host "The average size of the SharePoint Sites is $($SPSpaceMean)" -ForegroundColor Yellow
        Write-Host ""

    }

    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host "SharePoint Site Export Data Finished!" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan

}

if($Teams -or $All){
    
    Write-Host "Start to fetch Teams Data" -ForegroundColor White
    Write-Host ""

    $AllTeamsGroups = Get-AzureADGroup -All:$True | Where-Object {$_.grouptypes -Contains "unified" -and $_.resourceProvisioningOptions -contains "Team"}

    if($Domain){

        $SPTeamsSites = Get-SPOSite -Limit ALL | Where-Object {$_.Url -notmatch "my.sharepoint.com/personal" -and $_.IsTeamsConnected -eq $True -and $_.owner -match $Domain}

        $SPTeamsTotalSpace = 0
        Foreach($SPTeamsSite in $SPTeamsSites){
            $SingleSPTeamsUsedSpace = [Math]::Round($SPTeamsSite.StorageUsageCurrent / 1024, 3)
            $SPTeamsTotalSpace = $SPTeamsTotalSpace + $SingleSPTeamsUsedSpace
        }

        $SPTeamsSpaceMean = $SPTeamsTotalSpace / $TotalNumberSPSites
        [string]$SPTeamsTotalSpace = $SPTeamsTotalSpace.ToString() + " GB" #ToExport 
        [string]$SPTeamsSpaceMean = $SPTeamsSpaceMean.ToString() + " GB" #ToExport

        $TeamsUser = Get-CSOnlineUser -AccountType User -ResultSize 2147483647 | Where-Object {$_.UserPrincipalName -match $Domain}
        $TotalTeamsUser = $TeamsUser.count #To Export

        $TeamsGroupCount = 0
        Foreach($TeamGroup in $AllTeamsGroups){
            $TeamOwners = Get-TeamUser -GroupId $TeamGroup.ID -Role Owner
            $TeamOwnersList = ($TeamOwners.User) -join "|"
            if($TeamOwnersList -match $Domain){
                $TeamsGroupCount = $TeamsGroupCount + 1
            }
        }

        Write-Host "There are $($TotalTeamsUser) Teams Users having a $($Domain) domain" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($TeamsGroupCount) Teams Sites having among the owners a $($Domain) user" -ForegroundColor Yellow
        Write-Host "The total size of the SharePoint Sites connected to Teams having as owner a $($Domain) user is $($SPTeamsTotalSpace)" -ForegroundColor Yellow
        Write-Host "The average size of the SharePoint Sites connected to Teams having as owner a $($Domain) user is $($SPTeamsSpaceMean)" -ForegroundColor Yellow
        Write-Host ""
    }
    else{

        $SPTeamsSites = Get-SPOSite -Limit ALL | Where-Object {$_.Url -notmatch "my.sharepoint.com/personal" -and $_.IsTeamsConnected -eq $True}

        $SPTeamsTotalSpace = 0
        Foreach($SPTeamsSite in $SPTeamsSites){
            $SingleSPTeamsUsedSpace = [Math]::Round($SPTeamsSite.StorageUsageCurrent / 1024, 3)
            $SPTeamsTotalSpace = $SPTeamsTotalSpace + $SingleSPTeamsUsedSpace
        }

        $SPTeamsSpaceMean = $SPTeamsTotalSpace / $TotalNumberSPSites
        [string]$SPTeamsTotalSpace = $SPTeamsTotalSpace.ToString() + " GB" #ToExport 
        [string]$SPTeamsSpaceMean = $SPTeamsSpaceMean.ToString() + " GB" #ToExport

        $TeamsUser = Get-CSOnlineUser -AccountType User -ResultSize 2147483647
        $TotalTeamsUser = $TeamsUser.count #To Export
        $SPTeamsSites = Get-SPOSite -Limit ALL | Where-Object {$_.Url -notmatch "my.sharepoint.com/personal" -and $_.IsTeamsConnected -eq $True}

        Write-Host "There are $($TotalTeamsUser) Teams Users in the tenant" -ForegroundColor Yellow
        Write-Host "The tenant hosts $($AllTeamsGroups.count) Teams Sites" -ForegroundColor Yellow
        Write-Host "The total size of the SharePoint Sites connected to Teams is $($SPTeamsTotalSpace)" -ForegroundColor Yellow
        Write-Host "The average size of the SharePoint Sites connected to Teams is $($SPTeamsSpaceMean)" -ForegroundColor Yellow
        Write-Host ""

    }

    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Teams Site Export Data Finished!" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
}