$ErrorActionPreference = "Stop"

# ============================================================
# Microsoft 365 Cross-Tenant Access Policy Setup
#
# Graph only.
#
# Order of operations:
#
#   1. Resolve partner tenant ID
#   2. Connect to Microsoft Graph
#   3. Create/verify the parent Entra XTAP partner relationship
#   4. Configure M365 Collaboration inbound access
#   5. Discover existing M365 capabilities
#   6. Create missing capabilities
#   7. Update selected capabilities
#   8. Display final configuration using Graph Get- cmdlets
#
# No Exchange Online cmdlets are used.
# ============================================================

function Write-Section {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    do {
        $Answer = Read-Host "$Prompt [Y/N]"

        if ($null -ne $Answer) {
            $Answer = $Answer.Trim().ToUpperInvariant()
        }

        if ($Answer -notin @("Y", "N")) {
            Write-Host "Enter Y or N." -ForegroundColor Red
        }
    }
    until ($Answer -in @("Y", "N"))

    return ($Answer -eq "Y")
}

function Test-NotFoundError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $ErrorText = [string]$ErrorRecord.Exception.Message

    if ($ErrorRecord.ErrorDetails) {
        $ErrorText += " " + [string]$ErrorRecord.ErrorDetails.Message
    }

    return (
        $ErrorText -match "(?i)\b404\b" -or
        $ErrorText -match "(?i)notfound" -or
        $ErrorText -match "(?i)not found" -or
        $ErrorText -match "(?i)directory_objectnotfound" -or
        $ErrorText -match "(?i)unable to read the company information"
    )
}

function Resolve-PartnerTenantId {
    param(
        [Parameter(Mandatory)]
        [string]$Domain
    )

    $DiscoveryUri =
        "https://login.microsoftonline.com/$Domain/v2.0/.well-known/openid-configuration"

    $DiscoveryResponse = Invoke-RestMethod `
        -Uri $DiscoveryUri `
        -Method Get `
        -ErrorAction Stop

    if ($null -eq $DiscoveryResponse) {
        throw "The Microsoft identity platform returned an empty discovery response."
    }

    $DiscoveryValues = @(
        $DiscoveryResponse.issuer
        $DiscoveryResponse.authorization_endpoint
        $DiscoveryResponse.token_endpoint
        $DiscoveryResponse.jwks_uri
    ) | Where-Object {
        $null -ne $_ -and ([string]$_).Trim().Length -gt 0
    }

    $GuidPattern =
        "(?i)([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})"

    foreach ($DiscoveryValue in $DiscoveryValues) {
        if (([string]$DiscoveryValue) -match $GuidPattern) {
            return $Matches[1].ToLowerInvariant()
        }
    }

    throw "The discovery response did not contain a Microsoft Entra tenant GUID."
}

function Invoke-XtapGraphRequest {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("GET", "POST", "PATCH", "DELETE")]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [hashtable]$Body
    )

    $Parameters = @{
        Method      = $Method
        Uri         = $Uri
        ErrorAction = "Stop"
    }

    if ($null -ne $Body) {

        $Parameters["Body"] =
            $Body | ConvertTo-Json -Depth 20 -Compress

        $Parameters["ContentType"] =
            "application/json"
    }

    Invoke-MgGraphRequest @Parameters
}

function Get-CapabilityCollection {
    param(
        [Parameter(Mandatory)]
        [string]$CapabilitiesUri
    )

    try {

        $Response =
            Invoke-XtapGraphRequest `
                -Method GET `
                -Uri $CapabilitiesUri

        if ($null -eq $Response) {
            return @()
        }

        if (
            $Response -is [System.Collections.IDictionary] -and
            $Response.Contains("value")
        ) {
            return @($Response["value"])
        }

        if ($Response.PSObject.Properties.Name -contains "value") {
            return @($Response.value)
        }

        return @($Response)
    }
    catch {

        if (Test-NotFoundError -ErrorRecord $_) {
            return @()
        }

        throw
    }
}

function Get-CapabilityName {
    param(
        [Parameter(Mandatory)]
        [object]$CapabilityObject
    )

    $Name = $null

    if ($CapabilityObject -is [System.Collections.IDictionary]) {

        if ($CapabilityObject.Contains("name")) {
            $Name = [string]$CapabilityObject["name"]
        }

        if (-not $Name -and $CapabilityObject.Contains("id")) {
            $Name = [string]$CapabilityObject["id"]
        }

        if (
            -not $Name -and
            $CapabilityObject.Contains("@odata.type")
        ) {
            $Name =
                [string]$CapabilityObject["@odata.type"]
        }
    }
    else {

        if (
            $CapabilityObject.PSObject.Properties.Name `
                -contains "name"
        ) {
            $Name =
                [string]$CapabilityObject.name
        }

        if (
            -not $Name -and
            $CapabilityObject.PSObject.Properties.Name `
                -contains "id"
        ) {
            $Name =
                [string]$CapabilityObject.id
        }

        if (
            -not $Name -and
            $CapabilityObject.PSObject.Properties.Name `
                -contains "@odata.type"
        ) {
            $Name =
                [string]$CapabilityObject.'@odata.type'
        }
    }

    if ($Name) {

        $Name = $Name `
            -replace "^#microsoft\.graph\.", "" `
            -replace "^microsoft\.graph\.", ""
    }

    return $Name
}

# ============================================================
# Start
# ============================================================

Write-Section "MICROSOFT 365 CROSS-TENANT ACCESS POLICY SETUP"

# ============================================================
# Partner Domain
# ============================================================

Write-Host "Enter a verified domain belonging to the partner organization." -ForegroundColor Cyan
Write-Host "Example: contoso.com" -ForegroundColor DarkGray
Write-Host ""

do {

    $PartnerDomain =
        Read-Host "Enter the partner organization domain"

    if ($null -ne $PartnerDomain) {

        $PartnerDomain =
            $PartnerDomain.Trim().ToLowerInvariant()
    }

    if (-not $PartnerDomain) {

        Write-Host "Domain name cannot be empty." -ForegroundColor Red
        Write-Host ""

        continue
    }

    $PartnerDomain = $PartnerDomain `
        -replace "^https?://", "" `
        -replace "^www\.", "" `
        -replace "/.*$", ""

    if (
        -not $PartnerDomain -or
        $PartnerDomain.Length -gt 253 -or
        $PartnerDomain -notmatch "^(?=.{1,253}$)(?:(?!-)[a-z0-9-]{1,63}(?<!-)\.)+(?!-)[a-z0-9-]{2,63}(?<!-)$"
    ) {

        Write-Host "Enter a valid domain name, for example contoso.com." -ForegroundColor Red
        Write-Host ""

        $PartnerDomain = $null
    }
}
until ($PartnerDomain)

# ============================================================
# Resolve Partner Tenant ID
# ============================================================

Write-Host ""
Write-Host "Resolving the Microsoft Entra tenant for $PartnerDomain..." -ForegroundColor Yellow

try {

    $PartnerTenantId =
        Resolve-PartnerTenantId `
            -Domain $PartnerDomain
}
catch {

    Write-Host ""
    Write-Host "Unable to resolve the partner domain." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Fallback tenant lookup:" -ForegroundColor Yellow
    Write-Host "https://gettenantpartitionweb.azurewebsites.net/" -ForegroundColor Cyan

    throw
}

if (-not $PartnerTenantId) {
    throw "Partner Tenant ID is empty."
}

Write-Host ""
Write-Host "Partner Domain    : $PartnerDomain" -ForegroundColor Green
Write-Host "Partner Tenant ID : $PartnerTenantId" -ForegroundColor Green

# ============================================================
# Free/Busy Selection
# ============================================================

Write-Section "FREE/BUSY ACCESS"

Write-Host "  0. Do not configure Free/Busy"
Write-Host ""
Write-Host "  1. Availability Only"
Write-Host "     Time only"
Write-Host ""
Write-Host "  2. Limited Details"
Write-Host "     Time, subject, and location"
Write-Host ""

do {

    $FreeBusySelection =
        Read-Host "Select Free/Busy capability [0-2]"

    if ($FreeBusySelection -notin @("0", "1", "2")) {

        Write-Host "Invalid selection. Enter 0, 1, or 2." -ForegroundColor Red
    }
}
until ($FreeBusySelection -in @("0", "1", "2"))

switch ($FreeBusySelection) {

    "0" {

        $FreeBusyCapability = $null
        $FreeBusyDescription = "Not configured"
    }

    "1" {

        $FreeBusyCapability =
            "crossTenantCalendarAvailabilityBasic"

        $FreeBusyDescription =
            "Availability Only - Time"
    }

    "2" {

        $FreeBusyCapability =
            "crossTenantCalendarAvailabilityLimitedDetails"

        $FreeBusyDescription =
            "Limited Details - Time, Subject, Location"
    }
}

# ============================================================
# MailTips Selection
# ============================================================

Write-Section "MAILTIPS ACCESS"

Write-Host "  0. Do not configure MailTips"
Write-Host ""
Write-Host "  1. Limited"
Write-Host ""
Write-Host "  2. All"
Write-Host ""

do {

    $MailTipsSelection =
        Read-Host "Select MailTips capability [0-2]"

    if ($MailTipsSelection -notin @("0", "1", "2")) {

        Write-Host "Invalid selection. Enter 0, 1, or 2." -ForegroundColor Red
    }
}
until ($MailTipsSelection -in @("0", "1", "2"))

switch ($MailTipsSelection) {

    "0" {

        $MailTipsCapability = $null
        $MailTipsDescription = "Not configured"
    }

    "1" {

        $MailTipsCapability =
            "crossTenantMailTipsLimited"

        $MailTipsDescription =
            "Limited"
    }

    "2" {

        $MailTipsCapability =
            "crossTenantMailTipsAll"

        $MailTipsDescription =
            "All"
    }
}

# ============================================================
# Calendar Sharing Selection
# ============================================================

Write-Section "CALENDAR SHARING ACCESS"

Write-Host "  0. Do not configure Calendar Sharing"
Write-Host ""
Write-Host "  1. Free/Busy Simple"
Write-Host "     Time only"
Write-Host ""
Write-Host "  2. Free/Busy Detail"
Write-Host "     Time, subject, and location"
Write-Host ""
Write-Host "  3. Reviewer"
Write-Host "     Full calendar details"
Write-Host ""

do {

    $CalendarSelection =
        Read-Host "Select Calendar Sharing capability [0-3]"

    if ($CalendarSelection -notin @("0", "1", "2", "3")) {

        Write-Host "Invalid selection. Enter 0, 1, 2, or 3." -ForegroundColor Red
    }
}
until ($CalendarSelection -in @("0", "1", "2", "3"))

switch ($CalendarSelection) {

    "0" {

        $CalendarCapability = $null
        $CalendarDescription = "Not configured"
    }

    "1" {

        $CalendarCapability =
            "crossTenantCalendarSharingFreeBusySimple"

        $CalendarDescription =
            "Free/Busy Simple - Time"
    }

    "2" {

        $CalendarCapability =
            "crossTenantCalendarSharingFreeBusyDetail"

        $CalendarDescription =
            "Free/Busy Detail - Time, Subject, Location"
    }

    "3" {

        $CalendarCapability =
            "crossTenantCalendarSharingFreeBusyReviewer"

        $CalendarDescription =
            "Reviewer - Full Calendar Details"
    }
}

# ============================================================
# Build Capability List
# ============================================================

$Capabilities = @()

if ($FreeBusyCapability) {
    $Capabilities += $FreeBusyCapability
}

if ($CalendarCapability) {
    $Capabilities += $CalendarCapability
}

if ($MailTipsCapability) {
    $Capabilities += $MailTipsCapability
}

if ($Capabilities.Count -eq 0) {

    Write-Section "NO CAPABILITIES SELECTED"

    Write-Host "No changes have been made." -ForegroundColor Yellow

    return
}

# ============================================================
# Configuration Confirmation
# ============================================================

Write-Section "SELECTED CONFIGURATION"

Write-Host "Partner Domain    : $PartnerDomain"
Write-Host "Partner Tenant ID : $PartnerTenantId"
Write-Host ""

Write-Host "Free/Busy         : $FreeBusyDescription"
Write-Host "MailTips          : $MailTipsDescription"
Write-Host "Calendar Sharing  : $CalendarDescription"

Write-Host ""
Write-Host "Selected capabilities:" -ForegroundColor Cyan

foreach ($Capability in $Capabilities) {

    Write-Host "  $Capability"
}

Write-Host ""

if (-not (Read-YesNo "Continue with this configuration?")) {

    Write-Host ""
    Write-Host "Operation cancelled. No changes made." -ForegroundColor Yellow

    return
}

# ============================================================
# Microsoft Graph Modules
# ============================================================

Write-Section "MICROSOFT GRAPH MODULES"

$RequiredModules = @(
    "Microsoft.Graph.Authentication"
    "Microsoft.Graph.Identity.SignIns"
    "Microsoft.Graph.Beta.Identity.SignIns"
)

foreach ($RequiredModule in $RequiredModules) {

    if (-not (Get-Module -ListAvailable -Name $RequiredModule)) {

        Write-Host "Installing $RequiredModule..." -ForegroundColor Yellow

        Install-Module `
            -Name $RequiredModule `
            -Scope CurrentUser `
            -Repository PSGallery `
            -Force `
            -AllowClobber `
            -ErrorAction Stop
    }

    Write-Host "Importing $RequiredModule..." -ForegroundColor Yellow

    Import-Module `
        -Name $RequiredModule `
        -Force `
        -ErrorAction Stop
}

# ============================================================
# Confirm Required Commands
# ============================================================

$RequiredCommands = @(
    "Connect-MgGraph"
    "Get-MgContext"
    "Invoke-MgGraphRequest"
    "Get-MgPolicyCrossTenantAccessPolicyPartner"
    "New-MgPolicyCrossTenantAccessPolicyPartner"
    "Update-MgBetaPolicyCrossTenantAccessPolicyPartner"
    "Get-MgBetaPolicyCrossTenantAccessPolicyPartner"
    "Get-MgBetaPolicyCrossTenantAccessPolicyPartnerM365Capability"
)

foreach ($RequiredCommand in $RequiredCommands) {

    if (-not (Get-Command $RequiredCommand -ErrorAction SilentlyContinue)) {

        throw "Required command '$RequiredCommand' is unavailable."
    }
}

# ============================================================
# Microsoft Graph Connection
# ============================================================

Write-Section "MICROSOFT GRAPH CONNECTION"

$Scopes = @(
    "Policy.Read.All"
    "Policy.ReadWrite.CrossTenantAccess"
    "Policy.ReadWrite.CrossTenantCapability"
)

$ConnectParameters = @{

    Scopes       = $Scopes
    ContextScope = "Process"
    ErrorAction  = "Stop"
}

$ConnectCommand =
    Get-Command Connect-MgGraph -ErrorAction Stop

if ($ConnectCommand.Parameters.ContainsKey("NoWelcome")) {

    $ConnectParameters["NoWelcome"] = $true
}

Connect-MgGraph @ConnectParameters

$GraphContext =
    Get-MgContext -ErrorAction Stop

if ($null -eq $GraphContext) {

    throw "Microsoft Graph connection was not established."
}

$ConnectedTenantId =
    ([string]$GraphContext.TenantId).Trim()

if (-not $ConnectedTenantId) {

    throw "Microsoft Graph did not return the connected Tenant ID."
}

if ($ConnectedTenantId -eq $PartnerTenantId) {

    throw "The connected tenant and partner tenant cannot be the same."
}

Write-Host ""
Write-Host "Connected Account : $($GraphContext.Account)" -ForegroundColor Green
Write-Host "Connected Tenant  : $ConnectedTenantId" -ForegroundColor Green
Write-Host "Partner Tenant    : $PartnerTenantId" -ForegroundColor Green

# ============================================================
# Graph URLs
# ============================================================

$PartnerUri =
    "https://graph.microsoft.com/beta/policies/crossTenantAccessPolicy/partners/$PartnerTenantId"

$CapabilitiesUri =
    "$PartnerUri/m365Capabilities"

# ============================================================
# PARENT ENTRA XTAP POLICY
#
# This MUST exist before any M365 capability is created.
# ============================================================

Write-Section "ENTRA CROSS-TENANT PARTNER POLICY"

$ParentPartner = $null

Write-Host "Checking for parent Entra cross-tenant partner policy..." -ForegroundColor Yellow

try {

    $ParentPartner =
        Get-MgPolicyCrossTenantAccessPolicyPartner `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop
}
catch {

    if (Test-NotFoundError -ErrorRecord $_) {

        $ParentPartner = $null
    }
    else {

        throw
    }
}

# ============================================================
# Create parent when missing
# ============================================================

if ($null -eq $ParentPartner) {

    Write-Host "Parent Entra partner policy does not exist." -ForegroundColor Yellow
    Write-Host "Creating parent Entra partner policy..." -ForegroundColor Yellow

    try {

        New-MgPolicyCrossTenantAccessPolicyPartner `
            -TenantId $PartnerTenantId `
            -ErrorAction Stop |
            Out-Null
    }
    catch {

        Write-Host ""
        Write-Host "FAILED to create the parent Entra cross-tenant partner policy." -ForegroundColor Red
        Write-Host ""
        Write-Host $_.Exception.Message -ForegroundColor Red

        throw
    }

    Write-Host "Parent Entra partner policy creation command completed." -ForegroundColor Green
}
else {

    Write-Host "Parent Entra partner policy already exists." -ForegroundColor Green
}

# ============================================================
# CRITICAL CHECK
#
# Do NOT continue unless Graph can read the Entra parent.
# ============================================================

Write-Host ""
Write-Host "Verifying parent Entra partner policy..." -ForegroundColor Yellow

try {

    $ParentPartner =
        Get-MgPolicyCrossTenantAccessPolicyPartner `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop
}
catch {

    Write-Host ""
    Write-Host "STOPPING." -ForegroundColor Red
    Write-Host ""
    Write-Host "The parent Entra cross-tenant partner policy could not be retrieved." -ForegroundColor Red
    Write-Host "M365 capabilities will NOT be created." -ForegroundColor Red
    Write-Host ""
    Write-Host $_.Exception.Message -ForegroundColor Red

    throw
}

if ($null -eq $ParentPartner) {

    throw "The parent Entra cross-tenant partner policy was not returned by Microsoft Graph. M365 capability creation has been stopped."
}

Write-Host "Parent Entra cross-tenant partner policy verified." -ForegroundColor Green

Write-Host ""
Write-Host "Tenant ID         : $($ParentPartner.TenantId)" -ForegroundColor Cyan

# ============================================================
# Microsoft 365 Collaboration Inbound Trust
# ============================================================

Write-Section "MICROSOFT 365 COLLABORATION TRUST"

$M365CollaborationInbound = @{

    users = @{

        accessType = "allowed"

        targets = @(
            @{
                target     = "AllUsers"
                targetType = "user"
            }
        )
    }
}

Write-Host "Applying Microsoft 365 Collaboration inbound trust..." -ForegroundColor Yellow

Update-MgBetaPolicyCrossTenantAccessPolicyPartner `
    -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
    -M365CollaborationInbound $M365CollaborationInbound `
    -ErrorAction Stop

Write-Host "Microsoft 365 Collaboration inbound trust applied." -ForegroundColor Green

# ============================================================
# VERIFY PARENT AGAIN BEFORE CAPABILITY CREATION
# ============================================================

Write-Host ""
Write-Host "Rechecking parent Entra partner policy before capability creation..." -ForegroundColor Yellow

try {

    $ParentPartner =
        Get-MgPolicyCrossTenantAccessPolicyPartner `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop
}
catch {

    Write-Host ""
    Write-Host "STOPPING." -ForegroundColor Red
    Write-Host ""
    Write-Host "The required Entra parent policy is not readable." -ForegroundColor Red
    Write-Host "M365 capabilities will NOT be created." -ForegroundColor Red

    throw
}

if ($null -eq $ParentPartner) {

    throw "Required Entra parent policy was not returned. M365 capability creation stopped."
}

Write-Host "Parent Entra policy confirmed." -ForegroundColor Green

# ============================================================
# Existing M365 Capabilities
# ============================================================

Write-Section "EXISTING MICROSOFT 365 CAPABILITIES"

$ExistingCapabilities =
    Get-CapabilityCollection `
        -CapabilitiesUri $CapabilitiesUri

Write-Host "Existing capability count: $($ExistingCapabilities.Count)" -ForegroundColor Cyan

# ============================================================
# Create Missing Capabilities
# ============================================================

Write-Section "CREATE MISSING CAPABILITIES"

foreach ($Capability in $Capabilities) {

    $ExistingCapability =
        $ExistingCapabilities |
        Where-Object {

            (Get-CapabilityName -CapabilityObject $_) `
                -eq $Capability
        } |
        Select-Object -First 1

    if ($null -ne $ExistingCapability) {

        Write-Host "Already exists: $Capability" -ForegroundColor Green

        continue
    }

    $CreateBody = @{

        "@odata.type" =
            "#microsoft.graph.$Capability"

        inboundAccess = @{

            isAllowed = $true

            resourceScopes = @{

                included = @(
                    @{
                        resourceId   = "All"
                        resourceType = "user"
                    }
                )

                excluded = @()
            }
        }
    }

    Write-Host ""
    Write-Host "Creating: $Capability" -ForegroundColor Yellow

    Invoke-XtapGraphRequest `
        -Method POST `
        -Uri $CapabilitiesUri `
        -Body $CreateBody |
        Out-Null

    Write-Host "Created: $Capability" -ForegroundColor Green
}

# ============================================================
# Apply Selected Capability Settings
# ============================================================

Write-Section "APPLY SELECTED CAPABILITY SETTINGS"

foreach ($Capability in $Capabilities) {

    $CapabilityUri =
        "$CapabilitiesUri/$Capability"

    $UpdateBody = @{

        "@odata.type" =
            "#microsoft.graph.$Capability"

        inboundAccess = @{

            isAllowed = $true

            resourceScopes = @{

                included = @(
                    @{
                        resourceId   = "All"
                        resourceType = "user"
                    }
                )

                excluded = @()
            }
        }
    }

    Write-Host "Updating: $Capability" -ForegroundColor Yellow

    Invoke-XtapGraphRequest `
        -Method PATCH `
        -Uri $CapabilityUri `
        -Body $UpdateBody |
        Out-Null

    Write-Host "Updated: $Capability" -ForegroundColor Green
}

# ============================================================
# FINAL XTAP CONFIGURATION
#
# Use Graph PowerShell Get- cmdlets rather than REST JSON.
# ============================================================

Write-Section "FINAL XTAP CONFIGURATION"

Write-Host "Retrieving configuration using Microsoft Graph PowerShell..." -ForegroundColor Yellow
Write-Host ""

# ============================================================
# Entra Partner Configuration
# ============================================================

Write-Host "==================================================" -ForegroundColor Green
Write-Host " ENTRA PARTNER CONFIGURATION" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""

try {

    $FinalEntraPartner =
        Get-MgPolicyCrossTenantAccessPolicyPartner `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop

    $FinalEntraPartner |
        Format-List *
}
catch {

    Write-Host "Unable to retrieve the Entra partner configuration." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# ============================================================
# Microsoft 365 Collaboration Configuration
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host " MICROSOFT 365 COLLABORATION CONFIGURATION" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""

try {

    $FinalM365Partner =
        Get-MgBetaPolicyCrossTenantAccessPolicyPartner `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop

    $FinalM365Partner |
        Format-List *
}
catch {

    Write-Host "Unable to retrieve Microsoft 365 Collaboration configuration." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# ============================================================
# M365 Capabilities
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host " M365 CAPABILITIES" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""

try {

    $FinalCapabilities = @(
        Get-MgBetaPolicyCrossTenantAccessPolicyPartnerM365Capability `
            -CrossTenantAccessPolicyConfigurationPartnerTenantId $PartnerTenantId `
            -ErrorAction Stop
    )

    if ($FinalCapabilities.Count -eq 0) {

        Write-Host "No M365 capabilities returned." -ForegroundColor Yellow
    }
    else {

        $FinalCapabilities |
            Format-List *
    }
}
catch {

    Write-Host "Unable to retrieve M365 capabilities." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

# ============================================================
# Human-Readable Summary
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host " CONFIGURATION SUMMARY" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""

Write-Host "Connected Tenant  : $ConnectedTenantId"
Write-Host "Partner Domain    : $PartnerDomain"
Write-Host "Partner Tenant ID : $PartnerTenantId"

Write-Host ""
Write-Host "Requested Configuration" -ForegroundColor Cyan
Write-Host "-----------------------"

Write-Host "Free/Busy         : $FreeBusyDescription"
Write-Host "MailTips          : $MailTipsDescription"
Write-Host "Calendar Sharing  : $CalendarDescription"

Write-Host ""
Write-Host "Selected Graph Capabilities" -ForegroundColor Cyan
Write-Host "---------------------------"

foreach ($Capability in $Capabilities) {

    Write-Host "  $Capability"
}

Write-Host ""

# ============================================================
# Complete
# ============================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " XTAP CONFIGURATION COMPLETE" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "The parent Entra policy, Microsoft 365 Collaboration trust," -ForegroundColor Green
Write-Host "and selected M365 capabilities have been configured." -ForegroundColor Green
Write-Host ""

Write-Host "The configuration displayed above was retrieved using" -ForegroundColor Green
Write-Host "Microsoft Graph PowerShell Get cmdlets." -ForegroundColor Green
Write-Host ""

# Disconnect-MgGraph
