# Report_M365_licenses_changes
## Description

This PowerShell script downloads a file from SharePoint Online to a local path, collects data about assigned licenses for users from Microsoft Graph, and exports the collected data to an Excel file. In the main sheet, it shows which licenses were added and which were deleted for specific users. Then the file is uploaded to SharePoint Online.

## Prerequisites

- PowerShell 5.1 or later installed on the system.
- Installed SharePointPnPPowerShellOnline module
- Installed Microsoft.Graph module
- Installed ImportExcel module

## Configuration
Set the following variables in the script before running:

   - `$filename`: Complete with the desired filename for the generated Excel file (e.g., Raport_M365_licenses_changes.xlsx).
   - `$localPath`: Complete with the local path where the file will be downloaded and saved (e.g., C:\Raporty\).
   - `$siteUrl`: Complete with the URL of the SharePoint site where the file is located (e.g., https://company.sharepoint.com/sites/it-dep).
   - `$onlinePath`: Complete with the path of the SharePoint location where the file will be added (e.g., Shared Documents/Global/).
   - `$tenant`: Complete with the name of your Microsoft 365 tenant (e.g., company.onmicrosoft.com).
   - `$appId`: Complete with the Client ID of the Azure AD application used for authentication.
   - `$thumbprint`: Complete with the Thumbprint of the certificate used for authentication.
