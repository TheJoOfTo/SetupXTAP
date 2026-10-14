$ErrorActionPreference = 'Stop'

# ============================================================
# Microsoft 365 Cross-Tenant Access Policy Management
# Microsoft Graph v1.0 only
# ============================================================

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host '==================================================' -ForegroundColor Cyan
    Write-Host ''
}

function Read-Number {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][int]$Minimum,
        [Parameter(Mandatory)][int]$Maximum
    )

    while ($true) {
        $Text = (Read-Host $Prompt).Trim()

        if ($Text -notmatch '^\d+$') {
            Write-Host "Enter a number from $Minimum through $Maximum." -ForegroundColor Red
            continue
        }

        try {
            $Number = [int]$Text
        }
        catch {
            Write-Host "Enter a number from $Minimum through $Maximum." -ForegroundColor Red
            continue
        }

        if ($Number -lt $Minimum -or $Number -gt $Maximum) {
            Write-Host "Enter a number from $Minimum through $Maximum." -ForegroundColor Red
            continue
        }

        return $Number
    }
}

function Read-YesNo {
    param([Parameter(Mandatory)][string]$Prompt)

    while ($true) {
        $Answer = (Read-Host "$Prompt [Y/N]").Trim().ToUpperInvariant()

        if ($Answer -eq 'Y') {
            return $true
        }

        if ($Answer -eq 'N') {
            return $false
        }

        Write-Host 'Enter Y or N.' -ForegroundColor Red
    }
}

function Write-GraphError {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $Exception = $ErrorRecord.Exception
    $Messages = @()

    while ($null -ne $Exception) {
        if ($Exception.Message) {
            if ($Exception.Message -notin $Messages) {
                $Messages += $Exception.Message
            }
        }

        $Exception = $Exception.InnerException
    }

    if ($ErrorRecord.ErrorDetails) {
        if ($ErrorRecord.ErrorDetails.Message) {
            if ($ErrorRecord.ErrorDetails.Message -notin $Messages) {
                $Messages += $ErrorRecord.ErrorDetails.Message
            }
        }
    }

    if ($Messages.Count -eq 0) {
        Write-Host 'Microsoft Graph returned an unspecified error.' -ForegroundColor Red
    }
    else {
        foreach ($Message in $Messages) {
            Write-Host $Message -ForegroundColor Red
        }
    }
}

function Test-NotFound {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $Text = [string]$ErrorRecord.Exception.Message

    if ($ErrorRecord.ErrorDetails) {
        $Text += ' ' + [string]$ErrorRecord.ErrorDetails.Message
    }

    if ($Text -match '(?i)404|notfound|not found') {
        return $true
    }

    return $false
}

function Get-Property {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }

        return $null
    }

    if ($Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }

    return $null
}

function Invoke-XtapGraph {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [hashtable]$Body
    )

    $Params = @{
        Method      = $Method
        Uri         = $Uri
        ErrorAction = 'Stop'
    }

    if ($null -ne $Body) {
        $Params['Body'] = $Body | ConvertTo-Json -Depth 20 -Compress
        $Params['ContentType'] = 'application/json'
    }

    Invoke-MgGraphRequest @Params
}

function Get-XtapCollection {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [switch]$EmptyOnNotFound
    )

    try {
        $Results = @()
        $Next = $Uri

        while ($Next) {
            $Response = Invoke-XtapGraph -Method GET -Uri $Next

            if ($null -eq $Response) {
                break
            }

            $Value = Get-Property -Object $Response -Name 'value'

            if ($null -ne $Value) {
                $Results += @($Value)
                $Next = Get-Property -Object $Response -Name '@odata.nextLink'
            }
            else {
                $Results += @($Response)
                $Next = $null
            }
        }

        return @($Results)
    }
    catch {
        if ($EmptyOnNotFound -and (Test-NotFound -ErrorRecord $_)) {
            return @()
        }

        throw
    }
}

function Initialize-Graph {
    Write-Section 'MICROSOFT GRAPH V1.0 MODULE'

    $Module = 'Microsoft.Graph.Authentication'

    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Host "Installing $Module..." -ForegroundColor Yellow

        $InstallParams = @{
            Name         = $Module
            Scope        = 'CurrentUser'
            Repository   = 'PSGallery'
            Force        = $true
            AllowClobber = $true
            ErrorAction  = 'Stop'
        }

        Install-Module @InstallParams
    }

    Import-Module $Module -Force -ErrorAction Stop

    foreach ($Command in @('Connect-MgGraph', 'Get-MgContext', 'Invoke-MgGraphRequest')) {
        if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
            throw "Required command '$Command' is unavailable."
        }
    }

    Write-Host 'Microsoft Graph module is available.' -ForegroundColor Green
}

function Connect-XtapGraph {
    Write-Section 'MICROSOFT GRAPH CONNECTION'

    $Params = @{
        Scopes = @(
            'Policy.Read.All'
            'Policy.ReadWrite.CrossTenantAccess'
            'Policy.ReadWrite.CrossTenantCapability'
        )
        ContextScope = 'Process'
        ErrorAction  = 'Stop'
    }

    $Command = Get-Command Connect-MgGraph -ErrorAction Stop

    if ($Command.Parameters.ContainsKey('NoWelcome')) {
        $Params['NoWelcome'] = $true
    }

    Connect-MgGraph @Params

    $Context = Get-MgContext -ErrorAction Stop

    if ($null -eq $Context) {
        throw 'Microsoft Graph connection was not established.'
    }

    if (-not $Context.TenantId) {
        throw 'Microsoft Graph did not return a tenant ID.'
    }

    $TenantId = ([string]$Context.TenantId).Trim().ToLowerInvariant()

    Write-Host "Connected Account : $($Context.Account)" -ForegroundColor Green
    Write-Host "Connected Tenant  : $TenantId" -ForegroundColor Green

    return $TenantId
}

function Resolve-PartnerTenantId {
    param([Parameter(Mandatory)][string]$Domain)

    $Uri = "https://login.microsoftonline.com/$Domain/v2.0/.well-known/openid-configuration"

    $Response = Invoke-RestMethod -Uri $Uri -Method Get -ErrorAction Stop

    $Values = @(
        $Response.issuer
        $Response.authorization_endpoint
        $Response.token_endpoint
        $Response.jwks_uri
    )

    $Pattern = '(?i)([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})'

    foreach ($Value in $Values) {
        if ($null -eq $Value) {
            continue
        }

        if (([string]$Value) -match $Pattern) {
            return $Matches[1].ToLowerInvariant()
        }
    }

    throw 'Unable to determine the Microsoft Entra tenant ID.'
}

function Read-PartnerDomain {
    Write-Section 'PARTNER DOMAIN'

    Write-Host 'Enter a verified domain belonging to the partner organization.'
    Write-Host 'Example: contoso.com' -ForegroundColor DarkGray
    Write-Host ''

    while ($true) {
        $Domain = (Read-Host 'Enter the partner organization domain').Trim().ToLowerInvariant()

        if ($Domain.StartsWith('https://')) {
            $Domain = $Domain.Substring(8)
        }

        if ($Domain.StartsWith('http://')) {
            $Domain = $Domain.Substring(7)
        }

        if ($Domain.StartsWith('www.')) {
            $Domain = $Domain.Substring(4)
        }

        $Slash = $Domain.IndexOf('/')

        if ($Slash -ge 0) {
            $Domain = $Domain.Substring(0, $Slash)
        }

        $Domain = $Domain.Trim().TrimEnd('.')

        if ($Domain -match '^[a-z0-9][a-z0-9.-]*\.[a-z]{2,63}$') {
            return $Domain
        }

        Write-Host 'Enter a valid domain, for example contoso.com.' -ForegroundColor Red
    }
}

function Get-PartnerTenantId {
    param([Parameter(Mandatory)][object]$Partner)

    $TenantId = Get-Property -Object $Partner -Name 'tenantId'

    if (-not $TenantId) {
        $TenantId = Get-Property -Object $Partner -Name 'id'
    }

    if ($TenantId) {
        return ([string]$TenantId).Trim().ToLowerInvariant()
    }

    return $null
}

function Get-PartnerDisplayName {
    param([Parameter(Mandatory)][object]$Partner)

    $Name = Get-Property -Object $Partner -Name 'displayName'

    if (-not $Name) {
        $Name = Get-Property -Object $Partner -Name 'tenantName'
    }

    if ($Name) {
        return ([string]$Name).Trim()
    }

    return $null
}

function Get-CapabilityName {
    param([Parameter(Mandatory)][object]$Capability)

    $Name = Get-Property -Object $Capability -Name 'name'

    if (-not $Name) {
        $Name = Get-Property -Object $Capability -Name 'id'
    }

    if (-not $Name) {
        $Name = Get-Property -Object $Capability -Name '@odata.type'
    }

    if ($Name) {
        return (([string]$Name) -replace '^#?microsoft\.graph\.', '')
    }

    return $null
}

function Test-CapabilityAllowed {
    param([Parameter(Mandatory)][object]$Capability)

    $Inbound = Get-Property -Object $Capability -Name 'inboundAccess'

    if ($null -eq $Inbound) {
        return $false
    }

    $Allowed = Get-Property -Object $Inbound -Name 'isAllowed'

    return ($Allowed -eq $true)
}

function Get-XtapPartners {
    $Uri = 'https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners'

    return @(Get-XtapCollection -Uri $Uri -EmptyOnNotFound)
}

function Get-XtapPartner {
    param([Parameter(Mandatory)][string]$PartnerTenantId)

    $Uri = "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/$PartnerTenantId"

    return Invoke-XtapGraph -Method GET -Uri $Uri
}

function Get-XtapCapabilities {
    param([Parameter(Mandatory)][string]$PartnerTenantId)

    $Uri = "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/$PartnerTenantId/m365Capabilities"

    return @(Get-XtapCollection -Uri $Uri -EmptyOnNotFound)
}

function Select-XtapPartner {
    param([Parameter(Mandatory)][array]$Partners)

    $List = @()

    foreach ($Partner in $Partners) {
        $TenantId = Get-PartnerTenantId -Partner $Partner

        if (-not $TenantId) {
            continue
        }

        $List += [PSCustomObject]@{
            TenantId    = $TenantId
            DisplayName = Get-PartnerDisplayName -Partner $Partner
        }
    }

    $List = @($List | Sort-Object DisplayName, TenantId)

    if ($List.Count -eq 0) {
        return $null
    }

    Write-Section 'EXISTING CROSS-TENANT PARTNERS'

    for ($Index = 0; $Index -lt $List.Count; $Index++) {
        $Number = $Index + 1
        $Partner = $List[$Index]

        if ($Partner.DisplayName) {
            Write-Host "  $Number. $($Partner.DisplayName)"
            Write-Host "     Tenant ID: $($Partner.TenantId)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "  $Number. $($Partner.TenantId)"
        }

        Write-Host ''
    }

    Write-Host '  0. Return to main menu'
    Write-Host ''

    $Choice = Read-Number -Prompt "Select a partner [0-$($List.Count)]" -Minimum 0 -Maximum $List.Count

    if ($Choice -eq 0) {
        return $null
    }

    return $List[$Choice - 1]
}

function Get-XtapSummary {
    param(
        [Parameter(Mandatory)][object]$Partner,
        [Parameter(Mandatory)][array]$Capabilities
    )

    $M365Inbound = Get-Property -Object $Partner -Name 'm365CollaborationInbound'
    $M365Outbound = Get-Property -Object $Partner -Name 'm365CollaborationOutbound'

    $InboundStatus = 'Not Configured'
    $OutboundStatus = 'Not Configured'
    $Users = 'Not Configured'

    if ($M365Inbound) {
        $InboundUsers = Get-Property -Object $M365Inbound -Name 'users'

        if ($InboundUsers) {
            $AccessType = Get-Property -Object $InboundUsers -Name 'accessType'

            if ($AccessType -eq 'allowed') {
                $InboundStatus = 'Allowed'
            }
            elseif ($AccessType) {
                $InboundStatus = [string]$AccessType
            }

            $Targets = Get-Property -Object $InboundUsers -Name 'targets'

            if ($Targets) {
                $TargetNames = @()

                foreach ($Target in @($Targets)) {
                    $TargetName = Get-Property -Object $Target -Name 'target'

                    if ($TargetName) {
                        $TargetNames += [string]$TargetName
                    }
                }

                if ($TargetNames.Count -gt 0) {
                    $Users = $TargetNames -join ', '
                }
            }
        }
    }

    if ($M365Outbound) {
        $OutboundUsers = Get-Property -Object $M365Outbound -Name 'users'

        if ($OutboundUsers) {
            $AccessType = Get-Property -Object $OutboundUsers -Name 'accessType'

            if ($AccessType -eq 'allowed') {
                $OutboundStatus = 'Allowed'
            }
            elseif ($AccessType) {
                $OutboundStatus = [string]$AccessType
            }
        }
    }

    $FreeBusy = 'Not Configured'
    $MailTips = 'Not Configured'
    $Calendar = 'Not Configured'

    foreach ($Capability in $Capabilities) {
        if (-not (Test-CapabilityAllowed -Capability $Capability)) {
            continue
        }

        $Name = Get-CapabilityName -Capability $Capability

        switch ($Name) {
            'crossTenantCalendarAvailabilityBasic' {
                $FreeBusy = 'Availability Only - Time'
            }

            'crossTenantCalendarAvailabilityLimitedDetails' {
                $FreeBusy = 'Limited Details - Time, Subject, Location'
            }

            'crossTenantMailTipsLimited' {
                $MailTips = 'Limited'
            }

            'crossTenantMailTipsAll' {
                $MailTips = 'All'
            }

            'crossTenantCalendarSharingFreeBusySimple' {
                $Calendar = 'Free/Busy Simple - Time'
            }

            'crossTenantCalendarSharingFreeBusyDetail' {
                $Calendar = 'Free/Busy Detail - Time, Subject, Location'
            }

            'crossTenantCalendarSharingFreeBusyReviewer' {
                $Calendar = 'Reviewer - Full Calendar Details'
            }
        }
    }

    return [PSCustomObject]@{
        InboundAccess    = $InboundStatus
        OutboundAccess   = $OutboundStatus
        Users            = $Users
        FreeBusy         = $FreeBusy
        MailTips         = $MailTips
        CalendarSharing  = $Calendar
    }
}

function Show-XtapConfiguration {
    param(
        [Parameter(Mandatory)][string]$PartnerTenantId,
        [string]$PartnerDisplayName,
        [string]$PartnerDomain,
        [string]$ConnectedTenantId
    )

    $Partner = Get-XtapPartner -PartnerTenantId $PartnerTenantId
    $Capabilities = @(Get-XtapCapabilities -PartnerTenantId $PartnerTenantId)

    $Summary = Get-XtapSummary -Partner $Partner -Capabilities $Capabilities

    Write-Section 'XTAP CONFIGURATION'

    if ($ConnectedTenantId) {
        Write-Host "Connected Tenant  : $ConnectedTenantId"
    }

    if ($PartnerDisplayName) {
        Write-Host "Partner Name      : $PartnerDisplayName"
    }

    if ($PartnerDomain) {
        Write-Host "Partner Domain    : $PartnerDomain"
    }

    Write-Host "Partner Tenant ID : $PartnerTenantId"

    Write-Section 'MICROSOFT 365 COLLABORATION'

    Write-Host "Inbound Access    : $($Summary.InboundAccess)"
    Write-Host "Users             : $($Summary.Users)"
    Write-Host "Outbound Access   : $($Summary.OutboundAccess)"

    Write-Section 'MICROSOFT 365 CAPABILITIES'

    Write-Host "Free/Busy         : $($Summary.FreeBusy)"
    Write-Host "MailTips          : $($Summary.MailTips)"
    Write-Host "Calendar Sharing  : $($Summary.CalendarSharing)"

    Write-Section 'XTAP CONFIGURATION SUMMARY'

    Write-Host "Partner Tenant ID : $PartnerTenantId"

    if ($Summary.InboundAccess -eq 'Allowed') {
        Write-Host 'M365 Collaboration: Enabled' -ForegroundColor Green
    }
    else {
        Write-Host 'M365 Collaboration: Not Configured' -ForegroundColor Yellow
    }

    Write-Host "Free/Busy         : $($Summary.FreeBusy)"
    Write-Host "MailTips          : $($Summary.MailTips)"
    Write-Host "Calendar Sharing  : $($Summary.CalendarSharing)"
    Write-Host ''
}

function Start-ViewWorkflow {
    param([Parameter(Mandatory)][string]$ConnectedTenantId)

    Write-Section 'VIEW EXISTING XTAP CONFIGURATION'

    Write-Host 'Retrieving existing cross-tenant partner relationships...' -ForegroundColor Yellow

    $Partners = @(Get-XtapPartners)

    if ($Partners.Count -eq 0) {
        Write-Host 'No partner relationships were returned.' -ForegroundColor Yellow
        return
    }

    $Selected = Select-XtapPartner -Partners $Partners

    if ($null -eq $Selected) {
        return
    }

    Show-XtapConfiguration `
        -PartnerTenantId $Selected.TenantId `
        -PartnerDisplayName $Selected.DisplayName `
        -ConnectedTenantId $ConnectedTenantId

    Write-Host 'View operation complete. No changes were made.' -ForegroundColor Green
}

function Read-CapabilitySelections {
    Write-Section 'FREE/BUSY ACCESS'

    Write-Host '  0. Do not configure Free/Busy'
    Write-Host '  1. Availability Only - Time'
    Write-Host '  2. Limited Details - Time, Subject, Location'
    Write-Host ''

    $FreeBusyChoice = Read-Number -Prompt 'Select Free/Busy capability [0-2]' -Minimum 0 -Maximum 2

    $FreeBusyCapability = $null
    $FreeBusyDescription = 'Not Configured'

    if ($FreeBusyChoice -eq 1) {
        $FreeBusyCapability = 'crossTenantCalendarAvailabilityBasic'
        $FreeBusyDescription = 'Availability Only - Time'
    }

    if ($FreeBusyChoice -eq 2) {
        $FreeBusyCapability = 'crossTenantCalendarAvailabilityLimitedDetails'
        $FreeBusyDescription = 'Limited Details - Time, Subject, Location'
    }

    Write-Section 'MAILTIPS ACCESS'

    Write-Host '  0. Do not configure MailTips'
    Write-Host '  1. Limited'
    Write-Host '  2. All'
    Write-Host ''

    $MailTipsChoice = Read-Number -Prompt 'Select MailTips capability [0-2]' -Minimum 0 -Maximum 2

    $MailTipsCapability = $null
    $MailTipsDescription = 'Not Configured'

    if ($MailTipsChoice -eq 1) {
        $MailTipsCapability = 'crossTenantMailTipsLimited'
        $MailTipsDescription = 'Limited'
    }

    if ($MailTipsChoice -eq 2) {
        $MailTipsCapability = 'crossTenantMailTipsAll'
        $MailTipsDescription = 'All'
    }

    Write-Section 'CALENDAR SHARING ACCESS'

    Write-Host '  0. Do not configure Calendar Sharing'
    Write-Host '  1. Free/Busy Simple - Time'
    Write-Host '  2. Free/Busy Detail - Time, Subject, Location'
    Write-Host '  3. Reviewer - Full Calendar Details'
    Write-Host ''

    $CalendarChoice = Read-Number -Prompt 'Select Calendar Sharing capability [0-3]' -Minimum 0 -Maximum 3

    $CalendarCapability = $null
    $CalendarDescription = 'Not Configured'

    if ($CalendarChoice -eq 1) {
        $CalendarCapability = 'crossTenantCalendarSharingFreeBusySimple'
        $CalendarDescription = 'Free/Busy Simple - Time'
    }

    if ($CalendarChoice -eq 2) {
        $CalendarCapability = 'crossTenantCalendarSharingFreeBusyDetail'
        $CalendarDescription = 'Free/Busy Detail - Time, Subject, Location'
    }

    if ($CalendarChoice -eq 3) {
        $CalendarCapability = 'crossTenantCalendarSharingFreeBusyReviewer'
        $CalendarDescription = 'Reviewer - Full Calendar Details'
    }

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

    return [PSCustomObject]@{
        Capabilities        = $Capabilities
        FreeBusyDescription = $FreeBusyDescription
        MailTipsDescription = $MailTipsDescription
        CalendarDescription = $CalendarDescription
    }
}

function Ensure-Partner {
    param([Parameter(Mandatory)][string]$PartnerTenantId)

    $PartnersUri = 'https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners'
    $PartnerUri = "$PartnersUri/$PartnerTenantId"

    Write-Section 'ENTRA CROSS-TENANT PARTNER POLICY'

    $Partner = $null

    try {
        $Partner = Invoke-XtapGraph -Method GET -Uri $PartnerUri
    }
    catch {
        if (-not (Test-NotFound -ErrorRecord $_)) {
            throw
        }
    }

    if ($null -eq $Partner) {
        Write-Host 'Creating parent partner policy...' -ForegroundColor Yellow

        $Body = @{
            tenantId = $PartnerTenantId
        }

        Invoke-XtapGraph -Method POST -Uri $PartnersUri -Body $Body | Out-Null

        Write-Host 'Parent partner policy created.' -ForegroundColor Green
    }
    else {
        Write-Host 'Parent partner policy already exists.' -ForegroundColor Green
    }

    $Partner = Invoke-XtapGraph -Method GET -Uri $PartnerUri

    if ($null -eq $Partner) {
        throw 'Parent partner policy could not be verified.'
    }

    Write-Host 'Parent partner policy verified.' -ForegroundColor Green
}

function Set-M365Trust {
    param([Parameter(Mandatory)][string]$PartnerTenantId)

    Write-Section 'MICROSOFT 365 COLLABORATION TRUST'

    $Uri = "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/$PartnerTenantId"

    $Body = @{
        m365CollaborationInbound = @{
            users = @{
                accessType = 'allowed'
                targets = @(
                    @{
                        target     = 'AllUsers'
                        targetType = 'user'
                    }
                )
            }
        }
    }

    Write-Host 'Applying Microsoft 365 Collaboration inbound trust...' -ForegroundColor Yellow

    Invoke-XtapGraph -Method PATCH -Uri $Uri -Body $Body | Out-Null

    Write-Host 'Microsoft 365 Collaboration inbound trust applied.' -ForegroundColor Green
}

function Set-M365Capabilities {
    param(
        [Parameter(Mandatory)][string]$PartnerTenantId,
        [Parameter(Mandatory)][array]$Capabilities
    )

    $Uri = "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/$PartnerTenantId/m365Capabilities"

    $Existing = @(Get-XtapCapabilities -PartnerTenantId $PartnerTenantId)

    Write-Section 'CONFIGURE M365 CAPABILITIES'

    foreach ($CapabilityName in $Capabilities) {
        $Found = $false

        foreach ($ExistingCapability in $Existing) {
            if ((Get-CapabilityName -Capability $ExistingCapability) -eq $CapabilityName) {
                $Found = $true
                break
            }
        }

        $Body = @{
            '@odata.type' = "#microsoft.graph.$CapabilityName"
            inboundAccess = @{
                isAllowed = $true
                resourceScopes = @{
                    included = @(
                        @{
                            resourceId   = 'All'
                            resourceType = 'user'
                        }
                    )
                    excluded = @()
                }
            }
        }

        if (-not $Found) {
            Write-Host "Creating: $CapabilityName" -ForegroundColor Yellow

            Invoke-XtapGraph -Method POST -Uri $Uri -Body $Body | Out-Null

            Write-Host "Created: $CapabilityName" -ForegroundColor Green
        }
        else {
            Write-Host "Already exists: $CapabilityName" -ForegroundColor Green
        }

        Write-Host "Updating: $CapabilityName" -ForegroundColor Yellow

        $CapabilityUri = "$Uri/$CapabilityName"

        Invoke-XtapGraph -Method PATCH -Uri $CapabilityUri -Body $Body | Out-Null

        Write-Host "Configured: $CapabilityName" -ForegroundColor Green
        Write-Host ''
    }
}

function Start-ConfigureWorkflow {
    param([Parameter(Mandatory)][string]$ConnectedTenantId)

    Write-Section 'CONFIGURE OR UPDATE PARTNER TENANT'

    $Domain = Read-PartnerDomain

    Write-Host ''
    Write-Host "Resolving the Microsoft Entra tenant for $Domain..." -ForegroundColor Yellow

    $PartnerTenantId = Resolve-PartnerTenantId -Domain $Domain

    if ($PartnerTenantId -eq $ConnectedTenantId) {
        throw 'The connected tenant and partner tenant cannot be the same.'
    }

    Write-Host ''
    Write-Host "Partner Domain    : $Domain" -ForegroundColor Green
    Write-Host "Partner Tenant ID : $PartnerTenantId" -ForegroundColor Green

    $Selection = Read-CapabilitySelections

    if ($Selection.Capabilities.Count -eq 0) {
        Write-Host 'No capabilities were selected. No changes made.' -ForegroundColor Yellow
        return
    }

    Write-Section 'SELECTED CONFIGURATION'

    Write-Host "Partner Domain    : $Domain"
    Write-Host "Partner Tenant ID : $PartnerTenantId"
    Write-Host ''
    Write-Host "Free/Busy         : $($Selection.FreeBusyDescription)"
    Write-Host "MailTips          : $($Selection.MailTipsDescription)"
    Write-Host "Calendar Sharing  : $($Selection.CalendarDescription)"
    Write-Host ''

    $Continue = Read-YesNo -Prompt 'Continue with this configuration?'

    if (-not $Continue) {
        Write-Host 'Operation cancelled. No changes made.' -ForegroundColor Yellow
        return
    }

    Ensure-Partner -PartnerTenantId $PartnerTenantId

    Set-M365Trust -PartnerTenantId $PartnerTenantId

    Set-M365Capabilities `
        -PartnerTenantId $PartnerTenantId `
        -Capabilities $Selection.Capabilities

    Write-Host ''
    Write-Host 'Retrieving final configuration...' -ForegroundColor Yellow

    Show-XtapConfiguration `
        -PartnerTenantId $PartnerTenantId `
        -PartnerDomain $Domain `
        -ConnectedTenantId $ConnectedTenantId

    Write-Host 'XTAP configuration complete.' -ForegroundColor Green
}

# ============================================================
# SCRIPT START
# ============================================================

Write-Section 'MICROSOFT 365 XTAP MANAGEMENT'

$InitialChoice = Read-Number `
    -Prompt '1 = View existing XTAP, 2 = Configure/update XTAP, 3 = Exit' `
    -Minimum 1 `
    -Maximum 3

if ($InitialChoice -eq 3) {
    Write-Host 'Exiting. No changes made.'
    return
}

Initialize-Graph

$ConnectedTenantId = Connect-XtapGraph

$CurrentChoice = $InitialChoice
$KeepRunning = $true

while ($KeepRunning) {
    try {
        if ($CurrentChoice -eq 1) {
            Start-ViewWorkflow -ConnectedTenantId $ConnectedTenantId
        }
        elseif ($CurrentChoice -eq 2) {
            Start-ConfigureWorkflow -ConnectedTenantId $ConnectedTenantId
        }
        elseif ($CurrentChoice -eq 3) {
            $KeepRunning = $false
        }
    }
    catch {
        Write-Host ''
        Write-Host 'Operation failed.' -ForegroundColor Red
        Write-GraphError -ErrorRecord $_
    }

    if (-not $KeepRunning) {
        break
    }

    Write-Host ''

    $Return = Read-YesNo -Prompt 'Return to the main menu?'

    if (-not $Return) {
        break
    }

    Write-Section 'MICROSOFT 365 CROSS-TENANT ACCESS POLICY'

    Write-Host '  1. View existing XTAP configuration'
    Write-Host '  2. Configure or update a partner tenant'
    Write-Host '  3. Exit'
    Write-Host ''

    $CurrentChoice = Read-Number `
        -Prompt 'Select an option [1-3]' `
        -Minimum 1 `
        -Maximum 3

    if ($CurrentChoice -eq 3) {
        $KeepRunning = $false
    }
}

Write-Host ''
Write-Host 'XTAP management session complete.' -ForegroundColor Cyan
Write-Host ''

# Disconnect-MgGraph
