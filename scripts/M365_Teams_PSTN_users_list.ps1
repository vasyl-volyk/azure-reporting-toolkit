
param (
    [string]$ClientId     = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId     = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)

# ── Validate input ──────────────────────────────────────────────────────────
if ([string]::IsNullOrEmpty($ClientId) -or
    [string]::IsNullOrEmpty($ClientSecret) -or
    [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    exit 1
}

if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot "Teams_PSTN_Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# MicrosoftTeams 7.6+ fails to load Microsoft.IdentityModel.JsonWebTokens 8.3.0.
# Pin 7.5.0 from a script-local copy so Connect-MicrosoftTeams works reliably.
$teamsModuleRoot = Join-Path $PSScriptRoot '_modules\MicrosoftTeams'
$teamsModuleManifest = Join-Path $teamsModuleRoot '7.5.0\MicrosoftTeams.psd1'
if (-not (Test-Path $teamsModuleManifest)) {
    Write-Host "Installing MicrosoftTeams 7.5.0 to $teamsModuleRoot ..." -ForegroundColor Cyan
    if (-not (Test-Path $teamsModuleRoot)) {
        New-Item -ItemType Directory -Path $teamsModuleRoot -Force | Out-Null
    }
    Save-Module -Name MicrosoftTeams -RequiredVersion 7.5.0 -Path (Split-Path $teamsModuleRoot -Parent) -Force
}
Import-Module $teamsModuleManifest -Force

# ── Obtain access tokens via client credentials flow ────────────────────────
$tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$commonBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
}

# Token 1: MS Graph
$graphBody = $commonBody.Clone()
$graphBody["scope"] = "https://graph.microsoft.com/.default"
try {
    $graphToken = (Invoke-RestMethod -Uri $tokenUri -Method POST -Body $graphBody `
                   -ContentType "application/x-www-form-urlencoded").access_token
    Write-Host "[OK] Graph token acquired" -ForegroundColor Green
}
catch {
    Write-Error "Failed to obtain Graph token: $_"
    exit 1
}

# Token 2: Teams/Skype backend
$teamsBody = $commonBody.Clone()
$teamsBody["scope"] = "00000000-0000-0000-0000-000000000000/.default"
try {
    $teamsToken = (Invoke-RestMethod -Uri $tokenUri -Method POST -Body $teamsBody `
                   -ContentType "application/x-www-form-urlencoded").access_token
    Write-Host "[OK] Teams token acquired" -ForegroundColor Green
}
catch {
    Write-Error "Failed to obtain Teams token: $_"
    exit 1
}

# ── Connect to Microsoft Teams ──────────────────────────────────────────────
try {
    Connect-MicrosoftTeams -AccessTokens @($graphToken, $teamsToken) -ErrorAction Stop | Out-Null
    Write-Host "[OK] Connected to Microsoft Teams" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Teams: $_"
    Write-Host "TIP: Delete '$teamsModuleRoot' and re-run to reinstall MicrosoftTeams 7.5.0." -ForegroundColor Yellow
    exit 1
}

# ── Retrieve PSTN-capable users ─────────────────────────────────────────────
Write-Host "Retrieving Enterprise Voice-enabled users..." -ForegroundColor Cyan

$voiceUsers = @(Get-CsOnlineUser -Filter "EnterpriseVoiceEnabled -eq `$true" -ResultSize 99999)

# Further filter: only users with an assigned phone number (LineUri)
$pstnUsers = @($voiceUsers | Where-Object { -not [string]::IsNullOrEmpty($_.LineUri) })

# Total user count for reference
$totalCount = @(Get-CsOnlineUser -ResultSize 99999).Count

Write-Host "[OK] PSTN-capable users found: $($pstnUsers.Count) out of $totalCount total" -ForegroundColor Green

# ── Build results ───────────────────────────────────────────────────────────
if ($pstnUsers.Count -gt 0) {

    $results = foreach ($user in $pstnUsers) {

        # Clean phone number
        $phoneNumber = ($user.LineUri -replace '^tel:\+?', '+') -replace '^([^+])', '+$1'

        # Determine PSTN connectivity type
        $connectivity = if (-not [string]::IsNullOrEmpty($user.OnlineVoiceRoutingPolicy)) {
            "Direct Routing"
        }
        elseif ($user.LineUri -match "^tel:" -and
                [string]::IsNullOrEmpty($user.OnlineVoiceRoutingPolicy)) {
            "Calling Plan / Operator Connect"
        }
        else {
            "Unknown - verify in TAC"
        }

        [PSCustomObject]@{
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            AccountEnabled           = $user.AccountEnabled
            UsageLocation            = $user.UsageLocation
            PhoneNumber              = $phoneNumber
            EnterpriseVoiceEnabled   = $user.EnterpriseVoiceEnabled
            HostedVoiceMail          = $user.HostedVoiceMail
            OnlineVoiceRoutingPolicy = if ($user.OnlineVoiceRoutingPolicy) { $user.OnlineVoiceRoutingPolicy } else { "n/a" }
            TeamsCallingPolicy       = if ($user.TeamsCallingPolicy) { $user.TeamsCallingPolicy } else { "Global" }
            CallingLineIdentity      = if ($user.CallingLineIdentity) { $user.CallingLineIdentity } else { "n/a" }
            TenantDialPlan           = if ($user.TenantDialPlan) { $user.TenantDialPlan } else { "n/a" }
            Department               = if ($user.Department) { $user.Department } else { "n/a" }
            Title                    = if ($user.Title) { $user.Title } else { "n/a" }
            PSTNConnectivity         = $connectivity
        }
    }

    $results | Sort-Object DisplayName |
        Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "[OK] Report saved: $OutputPath" -ForegroundColor Green
    Write-Host "`nSample (first 5):" -ForegroundColor Cyan
    $results | Sort-Object DisplayName |
        Select-Object DisplayName, PhoneNumber, PSTNConnectivity, UsageLocation -First 5 |
        Format-Table -AutoSize
}
else {
    Write-Warning "No PSTN-capable users found. CSV was not created."
}

# ── Disconnect ──────────────────────────────────────────────────────────────
Disconnect-MicrosoftTeams | Out-Null
Write-Host "[OK] Disconnected from Microsoft Teams." -ForegroundColor Yellow
