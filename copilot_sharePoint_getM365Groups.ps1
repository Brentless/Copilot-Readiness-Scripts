<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# This script will quickly extract Public M365 Groups and Teams for analysis, 
# along with Group/Team Owner Email Address (for any communications).

# Entra ID App Registration and Permissions required
# Graph > Application Permissions > Groups.Read.All and Directory.Read.All
# Recommend a self signed certificate and thumbprint for authentication.

# ======================================================= #
# Configure variables
# ======================================================= #

$tenant = "YourTenantIdentifier" # from ___.onmicrosoft.com, recommend type in lower case

$appId = "YourEntraAppIDHere" # From Entra App Registration
$thumbprint = "YourCertThumbprintHere" # From self signed certificate

$timestamp = Get-Date -Format FileDateTime
$exportLocation = "C:\temp\groups_export_$($timestamp).csv"
$exportLocation_owners = "C:\temp\groupOwners_export_$($timestamp).csv"

# ======================================================= #
# Configure variables
# ======================================================= #

$AdminSiteURL = "https://$($tenant)-admin.sharepoint.com"
Connect-PnPOnline -Url $AdminSiteURL -Tenant ("$($tenant).onmicrosoft.com") -ClientId $appId -Thumbprint $thumbprint

$m365Groups = Get-PnPMicrosoft365Group -IncludeSiteUrl -IncludeOwners  | Where-Object { $_.Visibility -eq "Public"} # GET ONLY PUBLIC GROUPS
#$m365Groups = Get-PnPMicrosoft365Group -IncludeSiteUrl -IncludeOwners  | Where-Object { $_.Visibility -eq "Private"} # GET ONLY PRIVATE GROUPS
#$m365Groups = Get-PnPMicrosoft365Group -IncludeSiteUrl -IncludeOwners # GET ALL GROUPS

$groupExport = @()
$groupOwners = @()

foreach($grp in $m365Groups){

    write-host "Processing $($grp.DisplayName)"

    # Get site / group owners
    $grpOwners_email = @()
    $grpOwners_name = @()

    foreach($owner in $grp.Owners){
        $grpOwners_email += $owner.UserPrincipalName
        $groupOwners += $owner.UserPrincipalName
        $grpOwners_name += $owner.DisplayName
    }
    
    $row = New-Object -TypeName PSObject
    Add-Member -InputObject $row -NotePropertyName 'DisplayName' -NotePropertyValue $grp.DisplayName
    Add-Member -InputObject $row -NotePropertyName 'SiteUrl' -NotePropertyValue $grp.SiteUrl
    Add-Member -InputObject $row -NotePropertyName 'Id' -NotePropertyValue $grp.Id
    Add-Member -InputObject $row -NotePropertyName 'HasTeam' -NotePropertyValue $grp.HasTeam
    Add-Member -InputObject $row -NotePropertyName 'Visibility' -NotePropertyValue $grp.Visibility   
    Add-Member -InputObject $row -NotePropertyName 'Owners (Name)' -NotePropertyValue ($grpOwners_name -join '; ')
    Add-Member -InputObject $row -NotePropertyName 'Owners (Email)' -NotePropertyValue ($grpOwners_email -join '; ')       
    Add-Member -InputObject $row -NotePropertyName 'CreatedDateTime' -NotePropertyValue $grp.CreatedDateTime     
    Add-Member -InputObject $row -NotePropertyName 'RenewedDateTime' -NotePropertyValue $grp.RenewedDateTime           
    $groupExport += $row    

}

# Remove Duplicates of Group Owners
$groupOwners = $groupOwners | select -uniq

# Export List of Owner Email Addresses
$groupOwners | Select-Object @{Name='Owner';Expression={$_}} | export-csv $exportLocation_owners -Encoding UTF8 -NoTypeInformation 
Write-Host "Export of Owners Complete" -ForegroundColor Cyan

# Export Groups Inventory
$groupExport | Select-Object * | export-csv $exportLocation -Encoding UTF8 -NoTypeInformation 
Write-Host "Export of Groups Complete" -ForegroundColor Cyan
