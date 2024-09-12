<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# This script will generate a Site/Team inventory, with important information on each

# Entra ID App Registration and Permissions required
# Graph > Sites.FullControl.All and Groups.ReadWrite.All and Directory.Read.All > Application Permissions
# Recommend a self signed certificate and thumbprint for authentication.

# Note this may take a while to complete for large tenants

# ======================================================= #
# Configure variables
# ======================================================= #

$tenant = "YourTenantIdentifier" # from ___.onmicrosoft.com, recommend type in lower case

$appId = "YourEntraAppIDHere" # From Entra App Registration
$thumbprint = "YourCertThumbprintHere" # From self signed certificate

$timestamp = Get-Date -Format FileDateTime
$exportLocation = "C:\temp\site_export_$($timestamp).csv"

# ======================================================= #
# Get all sites and extract key details to a CSV
# ======================================================= #

$AdminSiteURL = "https://$($tenant)-admin.sharepoint.com"
Connect-PnPOnline -Url $AdminSiteURL -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint

$siteExport = @()

$excludedSiteTemplates = @("SPSMSITEHOST#0","SRCHCEN#0","APPCATALOG#0")
$sites = Get-PnPTenantSite | Where-Object { $_.Template -notin $excludedSiteTemplates }

foreach($site in $sites){

    #$site | select *

    Write-Host "`n$($site.Url)" -f Green    
    Connect-PnPOnline -Url $site.Url -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint
    $context = Get-PnPContext

    # Get site / group owners
    $siteOwners_email = @()
    $siteOwners_name = @()

    if ($site.RelatedGroupId.Guid -eq '00000000-0000-0000-0000-000000000000') {
        Write-Host "Getting SharePoint Owner Group"
        $ownerGroups = Get-PnPGroup -AssociatedOwnerGroup -Includes Users
        foreach($owner in $ownerGroups.Users){
            $siteOwners_email += $owner.Email
            $siteOwners_name += $owner.Title
        }
    } else {
        Write-Host "Getting M365 Group Owners"
        $groupOwners = Get-PnPMicrosoft365GroupOwner -Identity $site.RelatedGroupId.Guid
        foreach($owner in $groupOwners){
            $siteOwners_email += $owner.Email
            $siteOwners_name += $owner.DisplayName
        }    
    }
    
    # Write who the owners are
    #$siteOwners_name -join '; '
    #$siteOwners_email -join '; '

    # Get sensitivity label (if present)
    $siteLabel = Get-PnPSiteSensitivityLabel

    # Get some information from property bags
    $groupVisibility = Get-PnPPropertyBag -Key GroupType

    $row = New-Object -TypeName PSObject
    Add-Member -InputObject $row -NotePropertyName 'Site Title' -NotePropertyValue $site.Title
    Add-Member -InputObject $row -NotePropertyName 'Site Url' -NotePropertyValue $site.Url
    Add-Member -InputObject $row -NotePropertyName 'Site Collection Path' -NotePropertyValue ($site.Url).Replace("https://$($tenant.ToLower()).sharepoint.com","").Replace("https://$($tenant.ToLower())-my.sharepoint.com","")
    Add-Member -InputObject $row -NotePropertyName 'Sensitivity Label' -NotePropertyValue $siteLabel.DisplayName
    Add-Member -InputObject $row -NotePropertyName 'Site Template' -NotePropertyValue  $site.Template
    Add-Member -InputObject $row -NotePropertyName 'Team Connect' -NotePropertyValue  $site.IsTeamsConnected
    Add-Member -InputObject $row -NotePropertyName 'Team Channel Connected' -NotePropertyValue  $site.IsTeamsChannelConnected
    Add-Member -InputObject $row -NotePropertyName 'Related Group Id' -NotePropertyValue $site.RelatedGroupId
    Add-Member -InputObject $row -NotePropertyName 'Group Visibility' -NotePropertyValue $groupVisibility
    Add-Member -InputObject $row -NotePropertyName 'DefaultShareLinkScope' -NotePropertyValue $site.DefaultShareLinkScope
    Add-Member -InputObject $row -NotePropertyName 'DefaultSharingLinkType' -NotePropertyValue $site.DefaultSharingLinkType
    Add-Member -InputObject $row -NotePropertyName 'SharingCapability' -NotePropertyValue $site.SharingCapability
    Add-Member -InputObject $row -NotePropertyName 'Owner Name' -NotePropertyValue ($siteOwners_name -join '; ')
    Add-Member -InputObject $row -NotePropertyName 'Owner Email' -NotePropertyValue ($siteOwners_email -join '; ')       
    Add-Member -InputObject $row -NotePropertyName 'Last Modified' -NotePropertyValue $site.LastContentModifiedDate            
    $siteExport += $row

}

$siteExport | Sort-Object 'Site Collection Path' | Select-Object * | export-csv $exportLocation -Encoding UTF8 -NoTypeInformation 
Write-Host "Export Complete" -ForegroundColor Cyan