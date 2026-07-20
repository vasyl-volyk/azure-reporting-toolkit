#Requires -Version 5.0
<#
.SYNOPSIS
    Collects detailed information about all devices connected to Cisco Meraki
    networks and exports the data to a CSV file.

.DESCRIPTION
    Uses the Meraki Dashboard REST API v1 to enumerate all organizations,
    networks, and devices accessible with the provided API key.
    Collected data includes device identity, network assignment, status,
    hardware details, addressing, and firmware information.

.PARAMETER ApiKey
    Your Meraki Dashboard API key (X-Cisco-Meraki-API-Key).
    Can also be set via the environment variable MERAKI_API_KEY.

.PARAMETER OutputPath
    Full path for the output CSV file.
    Defaults to "MerakiDevices_<timestamp>.csv" in the current directory.

.PARAMETER OrganizationId
    Optional. Limit collection to a single organization ID.
    If omitted, all organizations accessible to the API key are queried.

.PARAMETER IncludeClients
    Switch. When specified, also collects layer-2/3 client data
    (MAC, IP, VLAN, manufacturer, last-seen) per network and appends
    those rows to the CSV with DeviceType = "Client".

.EXAMPLE
    # Basic usage – all orgs, devices only
    .\Get-MerakiDevices.ps1 -ApiKey "YOUR_API_KEY"

.EXAMPLE
    # Include wireless/wired clients, single org, custom output path
    .\Get-MerakiDevices.ps1 -ApiKey "YOUR_API_KEY" `
        -OrganizationId "123456" `
        -IncludeClients `
        -OutputPath "C:\Reports\Meraki_$(Get-Date -f yyyyMMdd).csv"

.NOTES
    Requires PowerShell 5.0+, TLS 1.2 enforced automatically.
    Rate limit: Meraki API allows ~10 requests/second per org.
    This script inserts 120 ms delays between calls to stay well within limits.
#>

[CmdletBinding()]
param (
    [string]$ApiKey = $env:MerakiApiKey,
    [string]$OutputPath = "MerakiDevices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [string]$OrganizationId = "000000", # Your organization ID in Meraki
    [Parameter(Mandatory = $false)]
    [switch]$IncludeClients
)

# ─────────────────────────────────────────────
# 0. Pre-flight checks
# ─────────────────────────────────────────────
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Error "API key not supplied. Use -ApiKey or set the MERAKI_API_KEY environment variable."
    exit 1
}

$BaseUri   = "https://api.meraki.com/api/v1"
$Headers   = @{
    "X-Cisco-Meraki-API-Key" = $ApiKey
    "Content-Type"           = "application/json"
    "Accept"                 = "application/json"
}
$RateDelay = 120   # milliseconds between API calls

# ─────────────────────────────────────────────
# 1. Helper – Invoke Meraki API with retry
# ─────────────────────────────────────────────
function Invoke-MerakiApi {
    param (
        [string]$Endpoint,
        [hashtable]$Query = @{}
    )

    # Explicitly reference script-scope variables to avoid PS5 scoping issues
    $localBase    = $script:BaseUri
    $localHeaders = $script:Headers
    $localDelay   = $script:RateDelay

    # Build URI
    $uri = $localBase + $Endpoint

    # Build query string manually (PS5 has no built-in URI builder for this)
    if ($Query.Count -gt 0) {
        $parts = @()
        foreach ($kvp in $Query.GetEnumerator()) {
            $parts += "$($kvp.Key)=$([Uri]::EscapeDataString([string]$kvp.Value))"
        }
        $uri = $uri + "?" + ($parts -join "&")
    }

    $maxRetries = 3
    $attempt    = 0

    do {
        $attempt++
        try {
            Start-Sleep -Milliseconds $localDelay
            $response = Invoke-RestMethod -Uri $uri -Headers $localHeaders -Method Get -ErrorAction Stop
            return $response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 429 -and $attempt -lt $maxRetries) {
                Write-Warning "Rate limited on $Endpoint – waiting 5 s before retry $attempt/$maxRetries …"
                Start-Sleep -Seconds 5
            }
            elseif ($statusCode -eq 404) {
                Write-Verbose "404 Not Found: $uri – skipping."
                return $null
            }
            else {
                Write-Warning "API error on $Endpoint (HTTP $statusCode): $($_.Exception.Message)"
                return $null
            }
        }
    } while ($attempt -lt $maxRetries)

    return $null
}

# ─────────────────────────────────────────────
# 2. Resolve organizations
# ─────────────────────────────────────────────
Write-Host "`n[1/5] Fetching organizations …" -ForegroundColor Cyan

if ($OrganizationId -ne "") {
    $orgs = Invoke-MerakiApi -Endpoint "/organizations/$OrganizationId"
    if ($null -eq $orgs) {
        Write-Error "Could not retrieve organization '$OrganizationId'. Check the ID and API key permissions."
        exit 1
    }
    $orgs = @($orgs)   # wrap single object in array
}
else {
    $orgs = Invoke-MerakiApi -Endpoint "/organizations"
    if ($null -eq $orgs -or $orgs.Count -eq 0) {
        Write-Error "No organizations found for this API key."
        exit 1
    }
}

Write-Host "  Found $($orgs.Count) organization(s)." -ForegroundColor Green

# ─────────────────────────────────────────────
# 3. Fetch networks per org
# ─────────────────────────────────────────────
Write-Host "[2/5] Fetching networks …" -ForegroundColor Cyan

$allNetworks = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($org in $orgs) {
    $networks = Invoke-MerakiApi -Endpoint "/organizations/$($org.id)/networks"
    if ($null -ne $networks) {
        foreach ($net in $networks) {
            # Attach org name for convenience
            $net | Add-Member -NotePropertyName "organizationName" -NotePropertyValue $org.name -Force
            $allNetworks.Add($net)
        }
    }
}

Write-Host "  Found $($allNetworks.Count) network(s)." -ForegroundColor Green

# ─────────────────────────────────────────────
# 4. Fetch devices per org (inventory) + status
# ─────────────────────────────────────────────
Write-Host "[3/5] Fetching device inventory and statuses …" -ForegroundColor Cyan

# Build a lookup: networkId -> network object
$networkLookup = @{}
foreach ($net in $allNetworks) {
    $networkLookup[$net.id] = $net
}

$allRows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($org in $orgs) {
    Write-Host "  Org: $($org.name) ($($org.id))" -ForegroundColor Yellow

    # Inventory devices
    $devices = Invoke-MerakiApi -Endpoint "/organizations/$($org.id)/devices"
    if ($null -eq $devices) { continue }

    # Device statuses (bulk endpoint – much faster than per-device)
    $statuses = Invoke-MerakiApi -Endpoint "/organizations/$($org.id)/devices/statuses"

    # Build status lookup by serial
    $statusLookup = @{}
    if ($null -ne $statuses) {
        foreach ($s in $statuses) {
            $statusLookup[$s.serial] = $s
        }
    }

    foreach ($dev in $devices) {
        $stat = $statusLookup[$dev.serial]

        # Resolve network name
        $netName  = ""
        $netType  = ""
        $netTags  = ""
        $netTimeZ = ""
        if ($dev.networkId -and $networkLookup.ContainsKey($dev.networkId)) {
            $n        = $networkLookup[$dev.networkId]
            $netName  = $n.name
            $netType  = ($n.productTypes -join "; ")
            $netTags  = ($n.tags -join "; ")
            $netTimeZ = $n.timeZone
        }

        $row = [PSCustomObject]@{
            # ── Identity ──────────────────────────────────
            DeviceType          = "NetworkDevice"
            OrganizationId      = $org.id
            OrganizationName    = $org.name
            NetworkId           = $dev.networkId
            NetworkName         = $netName
            NetworkProductTypes = $netType
            NetworkTags         = $netTags
            NetworkTimeZone     = $netTimeZ
            # ── Hardware ──────────────────────────────────
            Serial              = $dev.serial
            Name                = $dev.name
            Model               = $dev.model
            ProductType         = $dev.productType
            # ── Addressing ────────────────────────────────
            MAC                 = $dev.mac
            LanIP               = $dev.lanIp
            WanIP1              = $dev.wan1Ip
            WanIP2              = $dev.wan2Ip
            PublicIP            = if ($stat) { $stat.publicIp }         else { "" }
            # ── Location ──────────────────────────────────
            Latitude            = $dev.lat
            Longitude           = $dev.lng
            Address             = $dev.address
            # ── Status ────────────────────────────────────
            Status              = if ($stat) { $stat.status }           else { "unknown" }
            LastReportedAt      = if ($stat) { $stat.lastReportedAt }   else { "" }
            UsingCellularFailover = if ($stat) { $stat.usingCellularFailover } else { "" }
            # ── Firmware ──────────────────────────────────
            Firmware            = $dev.firmware
            # ── Tags / Notes ──────────────────────────────
            Tags                = ($dev.tags -join "; ")
            Notes               = $dev.notes
            # ── Uplink (if present in status) ─────────────
            PrimaryUplink       = if ($stat -and $stat.primaryUplink) { $stat.primaryUplink } else { "" }
        }
        $allRows.Add($row)
    }
}

Write-Host "  Collected $($allRows.Count) network device(s)." -ForegroundColor Green

# ─────────────────────────────────────────────
# 5. (Optional) Fetch clients per network
# ─────────────────────────────────────────────
if ($IncludeClients) {
    Write-Host "[4/5] Fetching clients per network …" -ForegroundColor Cyan
    $clientCount = 0

    foreach ($net in $allNetworks) {
        # timespan = 86400 s = last 24 hours; increase as needed
        $clients = Invoke-MerakiApi -Endpoint "/networks/$($net.id)/clients" `
                                    -Query @{ timespan = "86400"; perPage = "1000" }

        if ($null -eq $clients) { continue }

        foreach ($c in $clients) {
            $row = [PSCustomObject]@{
                DeviceType          = "Client"
                OrganizationId      = ""
                OrganizationName    = $net.organizationName
                NetworkId           = $net.id
                NetworkName         = $net.name
                NetworkProductTypes = ($net.productTypes -join "; ")
                NetworkTags         = ($net.tags -join "; ")
                NetworkTimeZone     = $net.timeZone
                Serial              = ""
                Name                = $c.description
                Model               = ""
                ProductType         = $c.deviceTypePrediction
                MAC                 = $c.mac
                LanIP               = $c.ip
                WanIP1              = ""
                WanIP2              = ""
                PublicIP            = ""
                Latitude            = ""
                Longitude           = ""
                Address             = ""
                Status              = $c.status
                LastReportedAt      = $c.lastSeen
                UsingCellularFailover = ""
                Firmware            = ""
                Tags                = ""
                Notes               = ""
                PrimaryUplink       = ""
                # Client-specific extras (appended columns)
                ClientId            = $c.id
                VLAN                = $c.vlan
                Manufacturer        = $c.manufacturer
                OS                  = $c.os
                User                = $c.user
                SSID                = $c.ssid
                SwitchPort          = $c.switchport
                Usage_Sent_KB       = $c.usage.sent
                Usage_Recv_KB       = $c.usage.recv
            }
            $allRows.Add($row)
            $clientCount++
        }
    }
    Write-Host "  Collected $clientCount client(s)." -ForegroundColor Green
}
else {
    Write-Host "[4/5] Skipping client collection (use -IncludeClients to enable)." -ForegroundColor DarkGray
}

# ─────────────────────────────────────────────
# 6. Export to CSV
# ─────────────────────────────────────────────
Write-Host "[5/5] Exporting to CSV …" -ForegroundColor Cyan

if ($allRows.Count -eq 0) {
    Write-Warning "No data collected – CSV will not be written."
    exit 0
}

try {

    $sortedresults = $allRows | Select-Object NetworkName,Name,ProductType,Model,Serial,Firmware,MAC,LanIP,PublicIP,Status,Tags,DeviceType,OrganizationId,OrganizationName,NetworkId,NetworkProductTypes,NetworkTags,NetworkTimeZone,WanIP1,WanIP2,Address,LastReportedAt,UsingCellularFailover,Notes,PrimaryUplink | Sort-Object -Property NetworkName
    $sortedresults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "`n✔  Done! $($allRows.Count) rows written to:" -ForegroundColor Green
    Write-Host "   $((Resolve-Path $OutputPath).Path)" -ForegroundColor White
}
catch {
    Write-Error "Failed to write CSV: $($_.Exception.Message)"
    exit 1
}
