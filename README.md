
Microsoft 365 Cross-Tenant Access Policy Setup
PowerShell utility for configuring Microsoft 365 Cross-Tenant Access Policy (XTAP) capabilities between two Microsoft Entra tenants.
The script provides an interactive setup experience for configuring:

- Free/Busy Access
- MailTips Access
- Calendar Sharing Access
Each capability can be configured independently or skipped entirely.
Disclaimer
This project is provided independently and is not an official Microsoft product, service, or supported offering.
The code is provided "as is" without warranty of any kind. Test and validate the script in your environment before using it in production. Use of this script is at your own risk.
Features
The script prompts for the partner Microsoft Entra Tenant ID and allows you to select the desired level of access for each Microsoft 365 cross-tenant capability.
Free/Busy
Available options:

- Do not configure
- Availability Only - Time only
- Limited Details - Time, subject, and location
MailTips
Available options:

- Do not configure
- Limited
- All
Calendar Sharing
Available options:

- Do not configure
- Free/Busy Simple - Time only
- Free/Busy Detail - Time, subject, and location
- Reviewer - Full calendar details
Choosing Do not configure for a capability causes the script to leave that capability unconfigured.
If all three capabilities are skipped, the script exits without making changes.
Prerequisites
Run the script from PowerShell using an account with the permissions necessary to configure Microsoft 365 Cross-Tenant Access Policy.
The script uses the Microsoft Graph PowerShell SDK and requires the following delegated Microsoft Graph permissions:

- Policy.Read.All
- Policy.ReadWrite.CrossTenantAccess
- Policy.ReadWrite.CrossTenantCapability
The script checks for the required Microsoft Graph Beta PowerShell module and installs it for the current user if it is not already installed.
Running the Script
Download setupxtap4.ps1 to your computer.
If the script was downloaded from GitHub or another Internet source, Windows may mark the file as originating from the Internet.
You can remove that mark with:
Unblock-File .\setupxtap4.ps1
Then run the script:
.\setupxtap4.ps1
The script will prompt you for the partner Microsoft Entra Tenant ID and the desired configuration for each capability.
Finding the Partner Tenant ID
You will need the Microsoft Entra Tenant ID (Tenant GUID) of the partner organization.
If you do not know it, the script provides the following tenant lookup utility:
Microsoft Tenant Partition Lookup
Example
A configuration could enable:

- Free/Busy: Limited Details
- MailTips: Do not configure
- Calendar Sharing: Reviewer
In this example, only the Free/Busy and Calendar Sharing capabilities are configured. MailTips is skipped.
What the Script Does
At a high level, the script:

1. Prompts for the partner Microsoft Entra Tenant ID.
2. Validates the Tenant ID format.
3. Prompts for the desired Free/Busy configuration.
4. Prompts for the desired MailTips configuration.
5. Prompts for the desired Calendar Sharing configuration.
6. Skips any capability where Do not configure was selected.
7. Connects to Microsoft Graph.
8. Creates the Microsoft 365 collaboration partner configuration.
9. Creates the selected Microsoft 365 cross-tenant capabilities.
10. Retrieves the resulting configuration for verification.
Existing Configuration
The script includes handling for scenarios where portions of the partner configuration already exist and attempts to continue processing the remaining selected capabilities.
As with any administrative script, review the selected configuration displayed by the script and validate the resulting configuration before production use.
Security
The script does not contain tenant credentials, application secrets, certificates, or other authentication material.
Authentication is performed interactively through Microsoft Graph PowerShell.
Do not modify the script to hard-code credentials or other sensitive information.
License
This project is licensed under the MIT License.
See the LICENSE file in this repository for the full license text.
Contributions and Issues
If you encounter an issue, open a GitHub issue with:

- The PowerShell version you are using
- The relevant error output
- The capability you were attempting to configure
- Steps necessary to reproduce the problem
Do not include tenant IDs, credentials, access tokens, customer names, or other confidential information in GitHub issues.
