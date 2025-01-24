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
