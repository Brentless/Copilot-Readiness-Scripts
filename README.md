# Copilot-Readiness-Scripts
This is a collection of useful PowerShell scripts that can support data readiness exploration in Copilot for Microsoft 365.  These scripts will look across SharePoint, Teams, and M365 Groups.  There are certainly other products and solutions to gather some of this information, but these are quick and easy scripts to manually pull the data out so you can do some analysis yourself.  Sharing here so others might be able to take advantage of these.

**Pre-reqs**
Need an Entra ID App Registration with SharePoint-admin level rights

Graph (Application Permissions)
- Sites.FullControl.All
- Groups.ReadWrite.All
- Directory.Read.All
- Files.Read.All

SharePoint (Application Permission)
- Sites.FullControl.All

Recommend you set up a self-signed certificate to run this with a thumbprint, as there have been recent changes to PnP powershell's multi tenant application.  You should remove/revoke permissions on the app while it is not in use.

**Site and Team Inventory**

[copilot_sharePoint_getSitesAndTeamsInventory.ps1](copilot_sharePoint_getSitesAndTeamsInventory.ps1)

This script will iterate through all of your SharePoint sites, Teams, and M365 Group sites and provide a summary of useful information.
- Site name
- Site id
- Site urls
- Sensitivity labels (if applicable)
- Visibility - Public/Private (if applicable)
- Site level sharing controls
- Owners
- Last modified date

![image](https://github.com/user-attachments/assets/3f1ecaa3-53aa-4b95-899a-ed991d77bf92)

**Files with Unique Permissions (broken permission inheritance or sharing links)**

[copilot_sharePoint_getUniquePermissions.ps1](copilot_sharePoint_getUniquePermissions.ps1)

This script will iterate through all of the sites, and then all of the subsequent document libraries, and THEN all of the subsequent files and folders looking for the existance of broken permission inheritance or sharing links.  These represent places where users may be sharing content outside of the bounds of the team/site and may need investigation by content owners.

CUSTOM OPTIONS: There is an option to run this on just a particular SharePoint site and Document Library, to do targeted investigations.  Additionally, when running the script, you can choose to export just the files and folders that have unique permissions, or you can do a complete file and folder extract.

![image](https://github.com/user-attachments/assets/8a22684a-8067-4dae-aaac-d8dceab54ffb)


**Public Teams and M365 Groups**

[copilot_sharePoint_getPrivateM365Groups.ps1](copilot_sharePoint_getPrivateM365Groups.ps1)

This script will pull a list of Teams and M365 Groups that have been marked as "PUBLIC" for visibility.  Teams marked as Public will make content available for all users in the organization, and thus accessible for processing by Copilot.  Some Teams and Groups are ok to be marked as Public, but admins should review that those are approved for purpose.  This is a subset of data that can be pulled from the bigger Site and Team Inventory export.

This could be modified to also automatically set Public teams to Private when they are found.

![image](https://github.com/user-attachments/assets/6f937bd9-fd3f-47a2-90ff-d47986128fc7)


**Creating a Self Signed Certificate**

[copilot_entra_createSelfSignedCertificate.ps1](copilot_entra_createSelfSignedCertificate.ps1)

If you need a quick script to create a self signed certificate that can be used with some of the PowerShell scripts above, here is a quick one for reference.

[https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate](https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate)
