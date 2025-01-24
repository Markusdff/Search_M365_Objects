<#
.SYNOPSIS
This script searches Microsoft 365 and Azure AD for various object types using a keyword.

.DESCRIPTION
The script performs a comprehensive search across the following object types:
- Users
- Shared Mailboxes
- Distribution Lists
- Security Groups
- Microsoft Teams
- Dynamic Distribution Groups
- Azure AD Roles and Role Assignments
- Public Folders
- Rooms and Equipment
- Office 365 Groups (Unified Groups)

The script outputs matching results with the object type and additional information for easier identification.

.REQUIREMENTS
- ExchangeOnlineManagement PowerShell Module
- AzureAD or Microsoft Graph PowerShell Module
- MicrosoftTeams PowerShell Module
#>
CLS
# Prompt for the admin account
$adminAccount = Read-Host "Enter the admin account (e.g., admin@yourdomain.com)"

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -UserPrincipalName $adminAccount
} catch {
    Write-Host "Failed to connect to Exchange Online. Exiting..." -ForegroundColor Red
    exit
}

# Connect to Azure AD
try {
    Connect-AzureAD -AccountId $adminAccount
} catch {
    Write-Host "Failed to connect to Azure AD. Exiting..." -ForegroundColor Red
    exit
}

# Connect to Microsoft Teams interactively
try {
    Import-Module MicrosoftTeams
    Connect-MicrosoftTeams
} catch {
    Write-Host "Failed to connect to Microsoft Teams. Continuing with other services..." -ForegroundColor Yellow
}

# Prompt for the search keyword
$searchKeyword = Read-Host "Enter the keyword to search"

# Initialize a results array
$results = @()

# Search for users
try {
    $users = Get-User -Filter "DisplayName -like '*$searchKeyword*' -or UserPrincipalName -like '*$searchKeyword*'"
    foreach ($user in $users) {
        $results += [pscustomobject]@{
            Name = $user.DisplayName
            Type = "User"
            AdditionalInfo = $user.UserPrincipalName
        }
    }
} catch {
    Write-Host "Error retrieving users. Skipping..." -ForegroundColor Yellow
}

# Search for shared mailboxes
try {
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter "DisplayName -like '*$searchKeyword*'"
    foreach ($mailbox in $sharedMailboxes) {
        $results += [pscustomobject]@{
            Name = $mailbox.DisplayName
            Type = "Shared Mailbox"
            AdditionalInfo = $mailbox.PrimarySmtpAddress
        }
    }
} catch {
    Write-Host "Error retrieving shared mailboxes. Skipping..." -ForegroundColor Yellow
}

# Search for distribution lists
try {
    $distributionLists = Get-DistributionGroup -Filter "DisplayName -like '*$searchKeyword*'"
    foreach ($dl in $distributionLists) {
        $results += [pscustomobject]@{
            Name = $dl.DisplayName
            Type = "Distribution List"
            AdditionalInfo = $dl.PrimarySmtpAddress
        }
    }
} catch {
    Write-Host "Error retrieving distribution lists. Skipping..." -ForegroundColor Yellow
}

# Search for dynamic distribution groups
try {
    $dynamicGroups = Get-DynamicDistributionGroup -Filter "DisplayName -like '*$searchKeyword*'"
    foreach ($ddg in $dynamicGroups) {
        $results += [pscustomobject]@{
            Name = $ddg.DisplayName
            Type = "Dynamic Distribution Group"
            AdditionalInfo = $ddg.PrimarySmtpAddress
        }
    }
} catch {
    Write-Host "Error retrieving dynamic distribution groups. Skipping..." -ForegroundColor Yellow
}

# Search for security groups
try {
    $securityGroups = Get-AzureADGroup | Where-Object { $_.DisplayName -like "*$searchKeyword*" }
    foreach ($group in $securityGroups) {
        $results += [pscustomobject]@{
            Name = $group.DisplayName
            Type = "Security Group"
            AdditionalInfo = $group.ObjectId
        }
    }
} catch {
    Write-Host "Error retrieving security groups. Skipping..." -ForegroundColor Yellow
}

# Search for Teams
try {
    $teams = Get-Team | Where-Object { $_.DisplayName -like "*$searchKeyword*" }
    foreach ($team in $teams) {
        $results += [pscustomobject]@{
            Name = $team.DisplayName
            Type = "Microsoft Team"
            AdditionalInfo = $team.MailNickName
        }
    }
} catch {
    Write-Host "Error retrieving Microsoft Teams. Skipping..." -ForegroundColor Yellow
}

# Search for public folders
try {
    $publicFolders = Get-PublicFolder -Recurse | Where-Object { $_.Name -like "*$searchKeyword*" }
    foreach ($folder in $publicFolders) {
        $results += [pscustomobject]@{
            Name = $folder.Name
            Type = "Public Folder"
            AdditionalInfo = $folder.Path
        }
    }
} catch {
    Write-Host "Error retrieving public folders. Skipping..." -ForegroundColor Yellow
}

# Search for rooms and equipment
try {
    $roomsAndEquipment = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -Filter "DisplayName -like '*$searchKeyword*'"
    foreach ($room in $roomsAndEquipment) {
        $results += [pscustomobject]@{
            Name = $room.DisplayName
            Type = "Room/Equipment"
            AdditionalInfo = $room.PrimarySmtpAddress
        }
    }
} catch {
    Write-Host "Error retrieving rooms and equipment. Skipping..." -ForegroundColor Yellow
}

# Search for Office 365 Groups (Unified Groups)
try {
    $office365Groups = Get-UnifiedGroup -Filter "DisplayName -like '*$searchKeyword*'"
    foreach ($group in $office365Groups) {
        $results += [pscustomobject]@{
            Name = $group.DisplayName
            Type = "Office 365 Group"
            AdditionalInfo = $group.PrimarySmtpAddress
        }
    }
} catch {
    Write-Host "Error retrieving Office 365 Groups. Skipping..." -ForegroundColor Yellow
}

# Search for Azure AD roles and role assignments
try {
    $roles = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -like "*$searchKeyword*" }
    foreach ($role in $roles) {
        $results += [pscustomobject]@{
            Name = $role.DisplayName
            Type = "Azure AD Role"
            AdditionalInfo = $role.ObjectId
        }
    }
} catch {
    Write-Host "Error retrieving Azure AD roles. Skipping..." -ForegroundColor Yellow
}

# Display results
if ($results.Count -eq 0) {
    Write-Host "No matching results found for '$searchKeyword'."
} else {
    $results | Format-Table -AutoSize
}

# Disconnect from services
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD -Confirm:$false
Disconnect-MicrosoftTeams -Confirm:$false
