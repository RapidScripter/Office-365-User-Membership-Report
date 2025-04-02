# Office 365 User Membership Report PowerShell Script

This PowerShell script generates a comprehensive report of Office 365 user memberships and exports the data to a CSV file.

## Features

- Retrieves group memberships for Office 365 users
- Filters for specific user types (guest users, disabled users, users not in any group)
- Reports on admin roles assigned to users
- Supports both interactive login and certificate-based app authentication
- Generates CSV reports with detailed user information

## Prerequisites

- PowerShell 5.1 or later
- Microsoft Graph PowerShell module
- Appropriate Azure AD permissions (Directory.Read.All at minimum)

## Installation

1. Clone this repository:
   ```powershell
   git clone https://github.com/RapidScripter/office365-user-membership-report.git
   cd office365-user-membership-report
   ```

2. Install the Microsoft Graph module (if not already installed):
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser -Force
   ```

## Usage

### Basic Usage (Interactive Login)
```powershell
.\UserMembershipReport.ps1
```

### Filter Options
```powershell
# Report only guest users
.\UserMembershipReport.ps1 -GuestUsersOnly

# Report only disabled users
.\UserMembershipReport.ps1 -DisabledUsersOnly

# Report users not in any groups
.\UserMembershipReport.ps1 -UsersNotinAnyGroup
```

### Process Specific Users
```powershell
# Process users from a CSV file (single column with header "UserIdentityValue")
.\UserMembershipReport.ps1 -UsersIdentityFile "path\to\users.csv"
```

### App Authentication
```powershell
.\UserMembershipReport.ps1 -TenantId "your-tenant-id" -ClientId "your-app-id" -CertificateThumbprint "cert-thumbprint"
```

## Output

The script generates a CSV file with the following columns:
- Display Name
- Email Address
- Group Name(s)
- License Status
- Account Status (Enabled/Disabled)
- Department
- Admin Roles

Output filename format: `UserMembershipReport_MMM-dd-hh-mm-ss-tt.csv`

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-UsersIdentityFile` | Path to CSV file containing user identities to process |
| `-GuestUsersOnly` | Process only guest users |
| `-DisabledUsersOnly` | Process only disabled users |
| `-UsersNotinAnyGroup` | Process only users not in any groups |
| `-TenantId` | Azure AD tenant ID for app authentication |
| `-ClientId` | Application ID for app authentication |
| `-CertificateThumbprint` | Certificate thumbprint for app authentication |

## Required Permissions

- Directory.Read.All
- User.Read.All
