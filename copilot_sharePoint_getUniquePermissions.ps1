<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# This script will iterate through a particular document library on a particular site
# Usually the Shared Documents is the primary location, though sometimes a custom document library might also be created

# Entra ID App Registration and Permissions required
# SharePoint > Sites.FullControl.All > Application Permissions
# Graph > Sites.FullControl.All and Files.Read.All > Application Permissions
# Recommend a self signed certificate and thumbprint for authentication.

# Note this has not been tested with super large document libraries (greater than 5000 items may need to 
# rewrite the Get-PnPListItem with CAML query).  But this SHOULD work with the paging for large libraries.

# ======================================================= #
# Configure variables
# ======================================================= #

$tenant = "YourTenantIdentifier" # from ___.onmicrosoft.com, recommend type in lower case

$appId = "YourEntraAppIDHere" # From Entra App Registration
$thumbprint = "YourCertThumbprintHere" # From self signed certificate

$timestamp = Get-Date -Format FileDateTime
$exportLocation = "C:\temp\permission_export_$($timestamp).csv"

$Global:includeonlyBrokenPermissions = $false # set to true to only export items that have broken permissions, otherwise will get a full report of all content

# ======================================================= #
# Get files and check broken permissions
# ======================================================= #

function Iterate-FilePermissions {
    param($library)
    $batch = 0

    Get-PnPListItem -List $library -PageSize 1000 -ScriptBlock{
        param($ListItems)

        $batch += 1
        Write-Host "PROCESSING BATCH $($batch)" -ForegroundColor Cyan

        foreach($ListItem in $ListItems){
            Write-Host $ListItem.FieldValues.FileLeafRef # This is the file or folder name
            $hasUniquePermissions = Get-PnPProperty -ClientObject $ListItem HasUniqueRoleAssignments
            if($ListItem.FieldValues.ContentTypeId -like "0x012000*"){
                $category = "Folder"
            } else {
                $category = "File"
            }

            $sharingLinkCount = 0

            if($category -eq "File" -and $hasUniquePermissions){
                Write-Host "Checking File Sharing Links..." -f DarkGray
                $sharingLinks = Get-PnPFileSharingLink -Identity $ListItem.FieldValues.FileRef
                $sharingLinkCount = $sharingLinks.Count
            }
            
            if($category -eq "Folder" -and $hasUniquePermissions){
                Write-Host "Checking Folder Sharing Links..." -f DarkGray
                $sharingLinks = Get-PnPFolderSharingLink -Folder $ListItem.FieldValues.FileRef
                $sharingLinkCount = $sharingLinks.Count
            }         

            $siteCollection = $ListItem.FieldValues.FileDirRef -split "/"

            #Write-Host $ListItem.FieldValues.FileDirRef # This is the SharePoint site relative path
            #Write-Host $brokenPermissions # This is true or false based on if it has unique permissions
            #Write-Host $ListItem.FieldValues.Last_x0020_Modified # This is the last modified date

            if(!$Global:includeonlyBrokenPermissions -or ($Global:includeonlyBrokenPermissions -and $hasUniquePermissions)){
                $row = New-Object -TypeName PSObject
                Add-Member -InputObject $row -NotePropertyName 'Item' -NotePropertyValue $ListItem.FieldValues.FileLeafRef
                Add-Member -InputObject $row -NotePropertyName 'Path' -NotePropertyValue $ListItem.FieldValues.FileDirRef
                Add-Member -InputObject $row -NotePropertyName 'Site Collection Path' -NotePropertyValue "/$($siteCollection[1])/$($siteCollection[2])"
                Add-Member -InputObject $row -NotePropertyName 'Unique Permissions' -NotePropertyValue $hasUniquePermissions
                Add-Member -InputObject $row -NotePropertyName 'Custom Sharing Links' -NotePropertyValue $sharingLinkCount
                Add-Member -InputObject $row -NotePropertyName 'Category' -NotePropertyValue $category
                Add-Member -InputObject $row -NotePropertyName 'Last Modified' -NotePropertyValue $ListItem.FieldValues.Last_x0020_Modified               
                $Global:permissionsExport += $row
            }

        }
    } | Out-Null

}

# ======================================================= #
# OPTION 1: Iterate particular SharePoint Site and Shared Documents
# ======================================================= #

<#
# Select individual site and library
$targetURL =  "https://gov369830.sharepoint.com/sites/2018SafetyAudit/"
$libraryName = "Shared Documents"

# Connect to site
Connect-PnPOnline -Url $targetURL -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint
$context = Get-PnPContext

$permissionsExport = @()
Iterate-FilePermissions -library $libraryName
Write-Host "Exporting to CSV..." -ForegroundColor Cyan
$permissionsExport | Sort-Object Path | Select-Object * | export-csv $exportLocation -Encoding UTF8 -NoTypeInformation 
Write-Host "Export Complete" -ForegroundColor Cyan
#>

# ======================================================= #
# OPTION 2: Iterate All SharePoint Sites and Documents Libraries
# ======================================================= #

$AdminSiteURL = "https://$($tenant)-admin.sharepoint.com"
Connect-PnPOnline -Url $AdminSiteURL -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint

$sites = Get-PnPTenantSite

$Global:permissionsExport = @()

foreach($site in $sites){

    Write-Host "`n$($site.Url)" -f Green    
    Connect-PnPOnline -Url $site.Url -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint
    $context = Get-PnPContext

    #Get all document libraries - Exclude Hidden Libraries
    $excludedLibraries = @("Form Templates", "Site Assets", "Site Pages", "Style Library", "Theme Gallery", "Pages")
    $docLibraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -notin $excludedLibraries }

    foreach($library in $docLibraries){
        Write-Host "Processing $($library.Title)" -f Yellow
        Iterate-FilePermissions -library $library.Title
    }
}

$Global:permissionsExport | Sort-Object Path | Select-Object * | export-csv $exportLocation -Encoding UTF8 -NoTypeInformation 
Write-Host "Export Complete" -ForegroundColor Cyan




