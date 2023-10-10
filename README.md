# SpListModule

SpListModule reinvents the wheel. Partly. :-/ SpListModule simplifies managing on-prem SharePoint lists.

![Status](https://img.shields.io/badge/status-alpha-red.svg)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://mit-license.org/)

## Features

- Establish and terminate connections to a SharePoint site
- Retrieve and manipulate list data
- Get user information from the SharePoint site

## Getting Started

1. Clone or download this repository.
2. Import the module using `Import-Module ./path_to_module/SpListModule.psm1`
3. Use `Connect-SharePoint` cmdlet to establish a connection to your SharePoint site.

## Usage
### Connecting to SharePoint

```powershell
# Connect using default credentials
Connect-SharePoint -SiteUrl "https://your-sharepoint-site.com" -ListTitle "YourList" -UseDefaultCredentials

# Connect using a prompted credential
Connect-SharePoint -SiteUrl "https://your-sharepoint-site.com" -ListTitle "YourList"
```

### Getting User Information
```powershell
# Get a user's ID using their email address
$userId = Get-SharePointUserId -UserEmail "user@example.com"

# Get a user's display name using their ID
$userName = Get-SharePointUser -UserID $userId
```

### Working with List Items
```powershell
# Get all items from the connected list
$items = Get-SharePointListItems

# Create a new list item
$newItem = @{ Title = "New Item" }
New-SharePointListItem -RequestBody $newItem

# Update an existing list item
$updateItem = @{ Title = "Updated Title" }
Update-SharePointListItem -ItemId 1 -RequestBody $updateItem

# Delete a list item
Remove-SharePointListItem -ItemId 1
```

### Disconnecting from SharePoint
```powershell
# Disconnect from SharePoint and optionally display the functions being removed
Disconnect-SharePoint -ShowFunctions
```

## Thanks
Thank you for checking out SpListModule! :D

