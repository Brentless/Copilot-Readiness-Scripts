<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# This script will quickly extract Public M365 Groups and Teams for analysis

# Entra ID App Registration and Permissions required
# Graph > Groups.Read.All > Application Permissions
# Recommend a self signed certificate and thumbprint for authentication.

# ======================================================= #
# Configure variables
# ======================================================= #

$tenant = "YourTenantIdentifier" # from ___.onmicrosoft.com, recommend type in lower case

$appId = "YourEntraAppIDHere" # From Entra App Registration
$thumbprint = "YourCertThumbprintHere" # From self signed certificate

$timestamp = Get-Date -Format FileDateTime
$exportLocation = "C:\temp\groups_export_$($timestamp).csv"

# ======================================================= #
# Configure variables
# ======================================================= #

$AdminSiteURL = "https://$($tenant)-admin.sharepoint.com"
Connect-PnPOnline -Url $AdminSiteURL -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint

$m365Groups = Get-PnPMicrosoft365Group -IncludeSiteUrl | Where-Object { $_.Visibility -eq "Public"} | Select Id, DisplayName, SiteUrl, HasTeam, Visibility, CreatedDateTime

$m365Groups | Select-Object * | export-csv $exportLocation -Encoding UTF8 -NoTypeInformation 
Write-Host "Export Complete" -ForegroundColor Cyan