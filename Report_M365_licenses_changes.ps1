# Install required modules
#Install-Module SharePointPnPPowerShellOnline
#Install-Module Microsoft.Graph
#Install-Module ImportExcel

# Set variables
$filename 		= '' #Complete with filename (ex. Raport_M365_licenses_changes.xlsx)
$localPath 		= '' #Complete with local path (ex. C:\Raporty\)
$siteUrl		= '' #Complete with Url site (ex. https://company.sharepoint.com/sites/it-dep)
$onlinePath		= '' #Complete with path where file is on sharepoint (ex. Shared Documents/Global/)
$tenant			= '' #Complete with tenant name (ex. company.onmicrosoft.com)
$appId			= '' #Complete with ClientId (which is ID of application registered in Azure AD)
$thumbprint		= '' #Complete with Thumbprint (which is certificate thumbprint)

# Connect to SharePoint Online
$pnpConnectParams  = @{
    Url				=  $siteUrl
    Tenant			=  $tenant
    ClientId		=  $appId
    Thumbprint		=  $thumbprint
}
Connect-PnPOnline @pnpConnectParams

# Get the file from SharePoint Online
$getPnPFileParams = @{
    Url				= ($onlinePath + $filename)
    Path			= $localPath
    Filename		= $filename
    AsFile			= $true
    Force			= $true
}
Get-PnPFile @getPnPFileParams

Start-Sleep -s 3

# Connect to Microsoft Graph
$graphParams  = @{
    Tenant					= $tenant
    AppId					= $appId
    CertificateThumbprint	= $thumbprint
}
Connect-Graph @graphParams

# List all users and their services in assigned licenses
$users = Get-MgUser -All
$excel = Open-ExcelPackage -Path ($localPath + $filename)
$excel.old.Cells["A2:D500"].Clear()
$excel.old.Cells["A2:D500"].Value = $excel.new.Cells["A2:D500"].Value
$excel.new.Cells["A2:D500"].Clear()
$zmienna = 2

foreach ($usr in $users) {
    $licenses = Get-MgUserLicenseDetail -UserId $usr.id
    foreach ($lic in $licenses) {
		$excel.new.Cells["A$zmienna"].Value = $usr.Id + "_" + $lic.SkuPartNumber
        $excel.new.Cells["B$zmienna"].Value = $usr.DisplayName
        $excel.new.Cells["C$zmienna"].Value = $usr.Mail
        $excel.new.Cells["D$zmienna"].Value = $lic.SkuPartNumber
        $zmienna++
    }
}
Close-ExcelPackage -ExcelPackage $excel

# Add the file from SharePoint Online
$addPnPFileParams = @{
    Folder = $onlinePath	
    Path   = ($localPath + $filename)
}
Add-PnPFile @addPnPFileParams