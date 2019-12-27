# AddItemPermissions.ps1

This is a sample CSOM PowerShell Script to add an permission to the list item.

## Prerequitesite
You need to download SharePoint Online Client Components SDK to run this script.
https://www.microsoft.com/en-us/download/details.aspx?id=42038

You can also acquire the latest SharePoint Online Client SDK by Nuget as well.

1. You need to access the following site.
https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

2. Download the nupkg.
3. Change the file extension to *.zip.
4. Unzip and extract those file.
5. place them in the specified directory from the code. 

## How to Run - parameters
    $SiteUrl,
    $LibraryName,
    $ItemID,
    $UserName,
    $RoleName

-SiteUrl ... Target site collection (site) or site (web) URL.

-LibraryName ... List Name (Title)

-ItemID ... Item ID

-UserName ... Account name to add permissions.

-RoleName ... The role name of permissions.

### Example 
.\AddItemPermissions.ps1 -siteUrl 'https://tenant.sharepoint.com/sites/site' -LibraryName 'DocumentLibrary1' -ItemID '2' -UserName 'user1@tenant.onmicrosoft.com' -RoleName 'Edit'

## Remarks
If the item does not have unique role, assign only the user to roles. (Do not copy the role assignments from the parent obuject to this object.)
