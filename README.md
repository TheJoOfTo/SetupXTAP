Microsoft 365 Cross-Tenant Access Policy Setup
`setupXTAP.ps1` is an interactive PowerShell script that configures Microsoft 365 cross-tenant access policy (XTAP) capabilities for a partner organization by using Microsoft Graph.
The script is Graph-only. It does not connect to Exchange Online or run Exchange Online cmdlets.
What's New
Partner tenant lookup by domain name
You no longer need to manually locate or enter the partner organization's Microsoft Entra tenant GUID.
The script now prompts for a verified domain belonging to the partner organization, such as:
```text
contoso.com
```
It then queries the Microsoft identity platform OpenID Connect discovery endpoint directly from PowerShell and extracts the partner organization's Microsoft Entra tenant ID/GUID.
The resolved domain and tenant ID are displayed before configuration begins so they can be reviewed before any changes are made.
If automatic lookup fails, the script displays the following fallback tenant lookup tool:
Microsoft Tenant Lookup Tool
What the Script Configures
The script can create or update the following Microsoft 365 cross-tenant capabilities:
Free/Busy access
Do not configure
Availability Only: time only
Limited Details: time, subject, and location
MailTips access
Do not configure
Limited
All
Calendar Sharing access
Do not configure
Free/Busy Simple: time only
Free/Busy Detail: time, subject, and location
Reviewer: full calendar details
Configuration Workflow
The script performs the following operations:
Prompts for a verified domain belonging to the partner organization.
Normalizes and validates the domain name.
Resolves the partner Microsoft Entra tenant ID from the domain.
Prompts for the desired Free/Busy, MailTips, and Calendar Sharing settings.
Displays the requested configuration and asks for confirmation.
Installs and imports the required Microsoft Graph PowerShell modules when necessary.
Connects to Microsoft Graph with the required delegated permissions.
Confirms that the connected tenant is different from the partner tenant.
Checks for the parent Entra cross-tenant access partner policy.
Creates the parent partner policy if it does not already exist.
Verifies that the parent partner policy can be retrieved before continuing.
Enables Microsoft 365 Collaboration inbound access for all users.
Discovers existing Microsoft 365 capabilities for the partner tenant.
Creates any selected capabilities that do not already exist.
Updates the selected capabilities to allow inbound access for all users.
Retrieves and displays the final configuration by using Microsoft Graph PowerShell Get cmdlets.
Displays a human-readable summary of the requested configuration.
Requirements
Windows PowerShell 5.1 or PowerShell 7.x
Internet access to Microsoft Graph, Microsoft identity platform endpoints, and PowerShell Gallery
Permission to install PowerShell modules for the current user, if the required modules are not already installed
An account authorized to manage cross-tenant access policies in the connected Microsoft Entra tenant
Administrative consent for the Microsoft Graph permissions requested by the script, where required by organizational policy
A verified domain belonging to the partner Microsoft Entra organization
Required PowerShell Modules
The script checks for, installs when missing, and imports these modules:
`Microsoft.Graph.Authentication`
`Microsoft.Graph.Identity.SignIns`
`Microsoft.Graph.Beta.Identity.SignIns`
The modules are installed from PowerShell Gallery with `CurrentUser` scope.
Microsoft Graph Permissions
The script requests these delegated Microsoft Graph scopes:
`Policy.Read.All`
`Policy.ReadWrite.CrossTenantAccess`
`Policy.ReadWrite.CrossTenantCapability`
Your account must be able to consent to these permissions, or the permissions must already be approved in the tenant.
How to Run
Download `setupXTAP.ps1`.
If Windows marked the downloaded file as blocked, unblock it:
```powershell
Unblock-File -Path .\setupXTAP.ps1
```
Open PowerShell.
Change to the folder containing the script.
Run the script:
```powershell
.\setupXTAP.ps1
```
Enter a verified partner domain when prompted.
Review the resolved partner tenant ID.
Select the desired Free/Busy, MailTips, and Calendar Sharing settings.
Review the configuration summary and confirm whether to continue.
Sign in to Microsoft Graph when prompted.
Review the final configuration shown by the script.
Example Partner Lookup
```text
Enter the partner organization domain: contoso.com

Resolving the Microsoft Entra tenant for contoso.com...

Partner Domain    : contoso.com
Partner Tenant ID : 00000000-0000-0000-0000-000000000000
```
The GUID above is only an example. The script displays the tenant ID returned for the domain entered at runtime.
Safety and Validation
The script includes several safeguards:
Rejects empty or invalid domain names.
Removes an entered `http://`, `https://`, `www.`, or trailing path before validating the domain.
Stops if the partner tenant ID cannot be resolved.
Displays the partner domain and resolved tenant ID before configuration.
Stops if the connected tenant and partner tenant are the same.
Makes no changes when no capabilities are selected.
Displays the complete requested configuration before asking for confirmation.
Makes no changes when the confirmation prompt is declined.
Creates the parent Entra partner policy only when it is missing.
Stops capability creation if the parent partner policy cannot be retrieved.
Detects existing capabilities before attempting to create missing capabilities.
Updates the selected capabilities after creation or discovery.
Final Verification Output
At the end of a successful run, the script retrieves and displays:
The Entra cross-tenant partner configuration
The Microsoft 365 Collaboration configuration
The configured Microsoft 365 capabilities
The connected tenant ID
The partner domain
The resolved partner tenant ID
The requested Free/Busy setting
The requested MailTips setting
The requested Calendar Sharing setting
The selected Microsoft Graph capability names
The final configuration is retrieved with Microsoft Graph PowerShell Get cmdlets rather than displayed only from the values submitted by the script.
Important Behavior
The script configures inbound Microsoft 365 Collaboration access for `AllUsers`.
Selected Microsoft 365 capabilities are configured with inbound access allowed for all users.
Selecting Do not configure skips that capability. It does not remove or disable an existing capability.
If all three capability categories are set to Do not configure, the script exits without making changes.
Existing Microsoft 365 capabilities that were not selected are not removed.
The Microsoft Graph connection uses process-scoped authentication context.
`Disconnect-MgGraph` is included as a commented line at the end of the script. The script does not automatically disconnect the Graph session.
Troubleshooting
The partner domain cannot be resolved
Confirm that:
The domain is a verified domain in the partner Microsoft Entra tenant.
The domain was entered without a path or other extra characters.
The computer can reach `login.microsoftonline.com`.
The Microsoft identity platform discovery endpoint is available for that domain.
The script also displays the fallback Microsoft Tenant Lookup Tool when automatic lookup fails.
A required command is unavailable
Close and reopen PowerShell, then confirm that the required Microsoft Graph modules are installed and import successfully.
```powershell
Get-Module Microsoft.Graph.Authentication -ListAvailable
Get-Module Microsoft.Graph.Identity.SignIns -ListAvailable
Get-Module Microsoft.Graph.Beta.Identity.SignIns -ListAvailable
```
Consent or authorization fails
Confirm that the signed-in account is permitted to manage cross-tenant access policies and that the requested Microsoft Graph scopes have been approved.
The script will not run after download
Unblock the file:
```powershell
Unblock-File -Path .\setupXTAP.ps1
```
Execution policy and script-signing requirements vary by organization. Follow the security requirements established for the environment where the script is being run.
Disclaimer
This script is provided as-is, without warranty of any kind. Review and test it in a non-production environment before using it in production. You are responsible for validating the resulting configuration and ensuring that it meets your organization's security, compliance, and operational requirements.
This is a community-provided script and is not an official Microsoft product or support offering.
License
See the repository's `LICENSE` file for licensing terms.
