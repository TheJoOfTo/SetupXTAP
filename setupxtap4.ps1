$ErrorActionPreference = "Stop"

# ============================================================
# Microsoft 365 Cross-Tenant Access Policy Setup
# Free/Busy + Calendar Sharing + MailTips
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " MICROSOFT 365 CROSS-TENANT ACCESS POLICY SETUP" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""


# ============================================================
# Partner Tenant ID
# ============================================================

Write-Host "If you do not know the partner Tenant ID (Tenant GUID)," -ForegroundColor Cyan
Write-Host "use the Microsoft tenant lookup tool." -ForegroundColor Cyan
Write-Host ""

$PartnerTenantId = Read-Host "Enter the partner Microsoft Entra Tenant ID GUID"

if ($null -eq $PartnerTenantId) {
    throw "Partner Tenant ID cannot be empty."
}

$PartnerTenantId = $PartnerTenantId.Trim()

if ($PartnerTenantId -eq "") {
    throw "Partner Tenant ID cannot be empty."
}

$TenantIdPattern = "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"

if ($PartnerTenantId -notmatch $TenantIdPattern) {
    throw "Partner Tenant ID must be a valid GUID. Example: 2579bfc3-b142-4a7d-a6f4-175e4f8b581e"
}

Write-Host ""
Write-Host "Partner Tenant ID: $PartnerTenantId" -ForegroundColor Green
Write-Host ""


# ============================================================
# Free/Busy Selection
# ============================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " FREE/BUSY ACCESS" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "  0. Do not configure Free/Busy"
Write-Host ""
Write-Host "  1. Availability Only"
Write-Host "     Time only"
Write-Host ""
Write-Host "  2. Limited Details"
Write-Host "     Time, subject, and location"
Write-Host ""

do {
    $FreeBusySelection = Read-Host "Select Free/Busy capability [0-2]"

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
        $FreeBusyCapability = "crossTenantCalendarAvailabilityBasic"
        $FreeBusyDescription = "Availability Only - Time"
    }

    "2" {
        $FreeBusyCapability = "crossTenantCalendarAvailabilityLimitedDetails"
        $FreeBusyDescription = "Limited Details - Time, Subject, Location"
    }
}


# ============================================================
# MailTips Selection
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " MAILTIPS ACCESS" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "  0. Do not configure MailTips"
Write-Host ""
Write-Host "  1. Limited"
Write-Host "     Restricted subset of recipient information"
Write-Host ""
Write-Host "  2. All"
Write-Host "     Out-of-office status, automatic replies,"
Write-Host "     and all recipient-specific information"
Write-Host ""

do {
    $MailTipsSelection = Read-Host "Select MailTips capability [0-2]"

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
        $MailTipsCapability = "crossTenantMailTipsLimited"
        $MailTipsDescription = "Limited"
    }

    "2" {
        $MailTipsCapability = "crossTenantMailTipsAll"
        $MailTipsDescription = "All"
    }
}


# ============================================================
# Calendar Sharing Selection
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " CALENDAR SHARING ACCESS" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

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
    $CalendarSelection = Read-Host "Select Calendar Sharing capability [0-3]"

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
        $CalendarCapability = "crossTenantCalendarSharingFreeBusySimple"
        $CalendarDescription = "Free/Busy Simple - Time"
    }

    "2" {
        $CalendarCapability = "crossTenantCalendarSharingFreeBusyDetail"
        $CalendarDescription = "Free/Busy Detail - Time, Subject, Location"
    }

    "3" {
        $CalendarCapability = "crossTenantCalendarSharingFreeBusyReviewer"
        $CalendarDescription = "Reviewer - Full Calendar Details"
    }
}


# ============================================================
# Build Selected Capability List
# ============================================================

$Capabilities = @()

if ($null -ne $FreeBusyCapability) {
    $Capabilities += $FreeBusyCapability
}

if ($null -ne $CalendarCapability) {
    $Capabilities += $CalendarCapability
}

if ($null -ne $MailTipsCapability) {
    $Capabilities += $MailTipsCapability
}


# ============================================================
# Stop if Nothing Was Selected
# ============================================================

if ($Capabilities.Count -eq 0) {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host " NO CAPABILITIES SELECTED" -ForegroundColor Yellow
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Free/Busy        : Not configured"
    Write-Host "MailTips         : Not configured"
    Write-Host "Calendar Sharing : Not configured"
    Write-Host ""
    Write-Host "No changes have been made." -ForegroundColor Yellow
    Write-Host ""

    return
}


# ============================================================
# Display Selected Configuration
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " SELECTED CONFIGURATION" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Partner Tenant   : $PartnerTenantId"
Write-Host "Free/Busy        : $FreeBusyDescription"
Write-Host "MailTips         : $MailTipsDescription"
Write-Host "Calendar Sharing : $CalendarDescription"
Write-Host ""

Write-Host "Capabilities that will be created:" -ForegroundColor Cyan

foreach ($Capability in $Capabilities) {
    Write-Host "  $Capability"
}

Write-Host ""


# ============================================================
# Microsoft Graph Beta Module
# ============================================================

$ModuleName = "Microsoft.Graph.Beta.Identity.SignIns"

if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
    Write-Host "Microsoft Graph Beta module is not installed." -ForegroundColor Yellow
    Write-Host "Installing $ModuleName..." -ForegroundColor Yellow
    Write-Host ""

    Install-Module `
        -Name $ModuleName `
        -Scope CurrentUser `
        -Repository PSGallery `
        -Force `
        -AllowClobber
}

Write-Host "Importing $ModuleName..." -ForegroundColor Yellow

Import-Module -Name $ModuleName -Force


# ============================================================
# Connect to Microsoft Graph
# ============================================================

$Scopes = @(
    "Policy.Read.All"
    "Policy.ReadWrite.CrossTenantAccess"
    "Policy.ReadWrite.CrossTenantCapability"
)

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host ""

Connect-MgGraph -Scopes $Scopes -ContextScope Process

$Context = Get-MgContext

if ($null -eq $Context) {
    throw "Microsoft Graph connection was not established."
}

if ($null -eq $Context.TenantId) {
    throw "Microsoft Graph connected, but no Tenant ID was returned."
}

if ($Context.TenantId.ToString().Trim() -eq "") {
    throw "Microsoft Graph connected, but no Tenant ID was returned."
}

Write-Host ""
Write-Host "Microsoft Graph connection established." -ForegroundColor Green
Write-Host ""
Write-Host "Connected Tenant : $($Context.TenantId)" -ForegroundColor Cyan
Write-Host "Partner Tenant   : $PartnerTenantId" -ForegroundColor Cyan
Write-Host ""


# ============================================================
# Create Microsoft 365 Collaboration Partner Trust
# ============================================================

$PartnerBody = @{
    tenantId = $PartnerTenantId

    m365CollaborationInbound = @{
        users = @{
            accessType = "allowed"

            targets = @(
                @{
                    target = "AllUsers"
                    targetType = "user"
                }
            )
        }
    }
}

Write-Host "Creating Microsoft 365 Collaboration partner trust..." -ForegroundColor Yellow
Write-Host ""

try {
    New-MgBetaPolicyCrossTenantAccessPolicyPartner `
        -BodyParameter $PartnerBody `
        -ErrorAction Stop

    Write-Host "Partner trust created successfully." -ForegroundColor Green
}
catch {
    $PartnerError = $_.Exception.Message

    if (
        $PartnerError -match "already exists" -or
        $PartnerError -match "already exist" -or
        $PartnerError -match "conflict"
    ) {
        Write-Host "Partner trust already exists. Continuing." -ForegroundColor Yellow
    }
    else {
        Write-Host ""
        Write-Host "Failed to create the partner trust." -ForegroundColor Red
        Write-Host $PartnerError -ForegroundColor Red
        Write-Host ""

        throw
    }
}


# ============================================================
# Resource Scope
# ============================================================

$AllUsersScope = @{
    resourceId = "All"
    resourceType = "user"
}


# ============================================================
# Create Microsoft 365 Capabilities
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " CREATING MICROSOFT 365 CAPABILITIES" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

foreach ($Capability in $Capabilities) {
    Write-Host "Creating capability: $Capability" -ForegroundColor Yellow

    $CapabilityBody = @{
        "@odata.type" = "microsoft.graph.$Capability"

        inboundAccess = @{
            isAllowed = $true

            resourceScopes = @{
                included = @(
                    $AllUsersScope
                )

                excluded = @(
                    @{}
                )
            }
        }
    }

    $CapabilityParameters = @{
        CrossTenantAccessPolicyConfigurationPartnerTenantId = $PartnerTenantId
        BodyParameter = $CapabilityBody
        ErrorAction = "Stop"
    }

    try {
        New-MgBetaPolicyCrossTenantAccessPolicyPartnerM365Capability `
            @CapabilityParameters

        Write-Host "Created: $Capability" -ForegroundColor Green
    }
    catch {
        $CapabilityError = $_.Exception.Message

        if (
            $CapabilityError -match "already exists" -or
            $CapabilityError -match "already exist" -or
            $CapabilityError -match "conflict"
        ) {
            Write-Host "Capability already exists: $Capability" -ForegroundColor Yellow
            Write-Host "Continuing with remaining capabilities." -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            Write-Host "Failed to create capability: $Capability" -ForegroundColor Red
            Write-Host $CapabilityError -ForegroundColor Red
            Write-Host ""

            throw
        }
    }

    Write-Host ""
}


# ============================================================
# Verification
# ============================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " CONFIGURATION COMPLETE" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Partner Tenant   : $PartnerTenantId"
Write-Host ""
Write-Host "Selected Access:" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Free/Busy        : $FreeBusyDescription"
Write-Host "  MailTips         : $MailTipsDescription"
Write-Host "  Calendar Sharing : $CalendarDescription"
Write-Host ""

$GetParameters = @{
    CrossTenantAccessPolicyConfigurationPartnerTenantId = $PartnerTenantId
    ErrorAction = "Stop"
}

Write-Host "Retrieving configured Microsoft 365 capabilities..." -ForegroundColor Yellow
Write-Host ""

try {
    Get-MgBetaPolicyCrossTenantAccessPolicyPartnerM365Capability `
        @GetParameters |
        Format-List *
}
catch {
    Write-Host "Configuration completed, but verification failed." -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ""

# Disconnect-MgGraph