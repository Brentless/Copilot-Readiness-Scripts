<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
#>

# This script is primarily built for GCC customers who need to manage Teams App Permissions in bulk for
# licensed M365 Copilot users as well as Copilot Chat Pilot Users.  Group-based license assignment is 
# not currently available in GCC.  Authorized user groups should also be enabled for Copilot in the 
# M365 Apps > Integrated Apps > Copilot app.

# Ideally this is temporary if you need to block access to the Copilot App in teams while piloting, with
# the goal of fully enabling the Copilot App in the Global policy so you do not have to manage this policy

# Recommend running this script in Windows PowerShell ISE, script may not run using PowerShell Core (e.g. in VS Code)
# This script requires both AzureAD and MicrosoftTeams modules
# Need either Global Admin Permission OR Need Teams Admin Permissions

# ---------------------------------------------

# Define the group IDs in an array (ensure these are valid GUIDs) - add as many as needed
# In this example I have a group for my Licensed users and another ones who are piloting Copilot Chat

$groupIds = @(
    "3a4ef3e6-be5c-4644-b101-e2ac53b13e90", # Copilot - Licensed Users
    "3615ba62-aa86-4fa7-8f4e-a73a78bfb456"  # Copilot - Copilot Chat Pilot Users
)

# ---------------------------------------------

# Check for AzureAD and MicrosoftTeams Modules
# Install Required Modules if not found (run as admin may be required)

if ($null -eq (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "Installing Azure AD module" -ForegroundColor Cyan
    Install-Module AzureAD -Repository PSGallery -AllowClobber -Force
}

if ($null -eq (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Installing MicrosoftTeams module" -ForegroundColor Cyan
    Install-Module MicrosoftTeams -Repository PSGallery -AllowClobber -Force
}

# ---------------------------------------------

# Connect to required services (login with MFA as needed)

Connect-AzureAD
Connect-MicrosoftTeams

# ---------------------------------------------

# Create an array of all of the users who will have their Teams App Permission policy changed

$allUsers = @()

# Loop through groups and extract users
foreach ($groupId in $groupIds) {
    Write-Host "Getting membership of group $($groupId)" -ForegroundColor Cyan
    try {
        $members = Get-AzureADGroupMember -ObjectId $groupId -All $true
        $allUsers += $members
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Failed to get members for group ID ${groupId}: ${errorMessage}"
    }
}

# Remove duplicates
Write-Host "`nRemoving duplicates from array" -ForegroundColor Cyan
$uniqueUsers = $allUsers | Sort-Object ObjectId -Unique

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
