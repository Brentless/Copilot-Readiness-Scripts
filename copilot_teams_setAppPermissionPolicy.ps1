<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# Need either Global Admin Permission
# OR Need Teams Admin Permissions

# ---------------------------------------------

# Define the group IDs in an array (ensure these are valid GUIDs)

$groupIds = @(
    "3a4ef3e6-be5c-4644-b101-e2ac53b13e90", # Copilot - Licensed Users
    "3615ba62-aa86-4fa7-8f4e-a73a78bfb456"  # Copilot - Copilot Chat Pilot Users
)

# ---------------------------------------------

# Check for AzureAD and MicrosoftTeams Modules
# Install Required Modules if not found (run as admin may be required)

if ($null -eq (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Graph module" -ForegroundColor Cyan
    Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
}

if ($null -eq (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Installing MicrosoftTeams module" -ForegroundColor Cyan
    Install-Module MicrosoftTeams -Repository PSGallery -AllowClobber -Force
}

# ---------------------------------------------

# Connect to required services (login with MFA as needed)
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"
Connect-MicrosoftTeams

# ---------------------------------------------
# Create an array of all of the users who will have their Teams App Permission policy changed

$allUsers = @()

# Loop through groups and extract users
foreach ($groupId in $groupIds) {
    Write-Host "Getting membership of group $($groupId)" -ForegroundColor Cyan
    try {
        $members = Get-MgGroupMember -GroupId $groupId -All | Select -ExpandProperty AdditionalProperties
        foreach($member in $members){
            $allUsers += $member.userPrincipalName
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Failed to get members for group ID ${groupId}: ${errorMessage}"
    }
}

# Remove duplicates
Write-Host "`nRemoving duplicates from array" -ForegroundColor Cyan
$uniqueUsers = $allUsers | Get-Unique

# Display the users who will be reviewed
Write-Host "`nUser accounts ($($uniqueUsers.Count)):" -ForegroundColor Cyan
foreach ($user in $uniqueUsers){
    Write-Host "$($user.DisplayName) - $($user.UserPrincipalName)"
}
Write-Host "`nPreparing to process $($uniqueUsers.Count) user(s)`n" -ForegroundColor Cyan

# ---------------------------------------------

# Loop through each user and check their current Teams App Permission Policy
foreach ($user in $uniqueUsers) {
    try {
        $policy = Get-CsUserPolicyAssignment -Identity $user.UserPrincipalName | Where-Object { $_.PolicyType -eq "TeamsAppPermissionPolicy" }

        if (-not $policy) {
            # No explicit policy = Global
            Grant-CsTeamsAppPermissionPolicy -Identity $user.UserPrincipalName -PolicyName "Allow Copilot"
            Write-Host "Assigned 'Allow Copilot' policy to $($user.UserPrincipalName) (was using Global)" -ForegroundColor Green
        } else {
            Write-Host "$($user.UserPrincipalName) already has a custom policy: $($policy.PolicyName)" -ForegroundColor Yellow
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Error processing $($user.UserPrincipalName): ${errorMessage}"
    }
}

# ---------------------------------------------

<#

# Optional code block: Loop through each user and reset them to Global Teams App Permission Policy

foreach ($user in $uniqueUsers) {
    try {
        Grant-CsTeamsAppPermissionPolicy -Identity $user.UserPrincipalName -PolicyName $null
        Write-Host "Reset $($user.UserPrincipalName) to Global policy" -ForegroundColor Yellow
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Error processing $($user.UserPrincipalName): ${errorMessage}"
    }
}

#>


