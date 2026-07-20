param (
    [string]$ClientId     = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId     = $env:AZURE_TENANT_ID,
    [string]$OutputPath   = ".\O365_LicensedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


#region ── Validation ──────────────────────────────────────────────────────────
if ([string]::IsNullOrEmpty($ClientId) -or
    [string]::IsNullOrEmpty($ClientSecret) -or
    [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}
#endregion

#region ── Connect to Microsoft Graph ─────────────────────────────────────────
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

$SecuredPassword        = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object System.Management.Automation.PSCredential($ClientId, $SecuredPassword)

Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -ErrorAction Stop
Write-Host "Connected successfully." -ForegroundColor Green
#endregion

#region ── Fetch SKU friendly-name lookup table ────────────────────────────────
Write-Host "Fetching subscribed SKUs for friendly name lookup..." -ForegroundColor Cyan

$skuLookup = @{}
foreach ($sku in (Get-MgSubscribedSku -All)) {
    $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
}
#endregion

#region ── Fetch licensed users only ──────────────────────────────────────────
Write-Host "Fetching licensed users..." -ForegroundColor Cyan

$selectProps = "DisplayName,UserPrincipalName,JobTitle,AssignedLicenses"

$users = Get-MgUser -All `
    -Filter "assignedLicenses/`$count ne 0" `
    -ConsistencyLevel eventual `
    -CountVariable userCount `
    -Select $selectProps `
    -ErrorAction Stop

Write-Host "  Retrieved $($users.Count) licensed users." -ForegroundColor Green
#endregion

#region ── Resolve licenses per user ──────────────────────────────────────────
Write-Host "Resolving license names..." -ForegroundColor Cyan

# Intermediate list: one entry per user with resolved license name array
$userLicenseMap = foreach ($user in $users) {
    $names = foreach ($lic in $user.AssignedLicenses) {
        if ($skuLookup.ContainsKey($lic.SkuId)) { $skuLookup[$lic.SkuId] }
        else                                     { $lic.SkuId }
    }
    [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        JobTitle          = $user.JobTitle
        LicenseList       = @($names)   # keep as array for pivot logic
    }
}

# Collect all unique SKU names across all users — these become the pivot columns
$allSkus = $userLicenseMap |
    ForEach-Object { $_.LicenseList } |
    Sort-Object -Unique
#endregion

#region ── Build pivot matrix rows ────────────────────────────────────────────
Write-Host "Building pivot matrix..." -ForegroundColor Cyan

$results = foreach ($entry in $userLicenseMap) {

    # Start with the fixed identity columns
    $row = [ordered]@{
        DisplayName       = $entry.DisplayName
        UserPrincipalName = $entry.UserPrincipalName
        JobTitle          = $entry.JobTitle
    }

    # Add one YES/NO column per SKU
    foreach ($sku in $allSkus) {
        $row[$sku] = if ($entry.LicenseList -contains $sku) { "+" } else { "" }
    }

    [PSCustomObject]$row
}
#endregion

#region ── Export to CSV ───────────────────────────────────────────────────────
Write-Host "Exporting to: $OutputPath" -ForegroundColor Cyan

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Done. $($results.Count) users x $($allSkus.Count) license columns written to '$OutputPath'." -ForegroundColor Green
#endregion

#region ── Disconnect ─────────────────────────────────────────────────────────
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Cyan
#endregion
