param (
    [string]$APIKey = $env:Dcert_API_KEY,
    [string]$OutputPath
)

#region --- Validation ---
if (-not $APIKey) {
    Write-Error "API key is required. Provide it via -APIKey or set the PA_API_KEY environment variable."
    exit 1
}

if (-not $OutputPath) {
    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = "DigiCert_Domains_$timestamp.csv"
    Write-Warning "No -OutputPath specified. Defaulting to: $OutputPath"
}
#endregion

#region --- Configuration ---
$BaseUrl = "https://www.digicert.com/services/v2"
$Headers = @{
    "X-DC-DEVKEY" = $APIKey
    "Content-Type" = "application/json"
}
$PageLimit = 1000   # Max records per page (DigiCert maximum)
#endregion

#region --- Helper: Invoke DigiCert API ---
function Invoke-DigiCertAPI {
    param (
        [string]$Endpoint,
        [hashtable]$QueryParams = @{}
    )

    $uri = "$BaseUrl$Endpoint"

    if ($QueryParams.Count -gt 0) {
        $query = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        $uri = "${uri}?${query}"
    }

    try {
        $response = Invoke-RestMethod -Uri $uri -Method GET -Headers $Headers -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $message    = $_.ErrorDetails.Message

        Write-Error "API call failed [$statusCode] for '$uri': $message"
        return $null
    }
}
#endregion

#region --- Fetch All Domains (paginated) ---
Write-Host "Fetching domains from DigiCert..." -ForegroundColor Cyan

$allDomains = [System.Collections.Generic.List[object]]::new()
$offset     = 0
$page       = 1

do {
    Write-Host "  Fetching page $page (offset: $offset)..." -NoNewline

    $params = @{
        limit  = $PageLimit
        offset = $offset
    }

    $response = Invoke-DigiCertAPI -Endpoint "/domain" -QueryParams $params

    if (-not $response -or -not $response.domains) {
        Write-Host " No data returned. Stopping." -ForegroundColor Yellow
        break
    }

    $batch = $response.domains
    $allDomains.AddRange($batch)
    Write-Host " Retrieved $($batch.Count) domains." -ForegroundColor Green

    $total  = $response.page.total
    $offset += $PageLimit
    $page++

} while ($allDomains.Count -lt $total)

Write-Host "`nTotal domains fetched: $($allDomains.Count)" -ForegroundColor Cyan
#endregion

#region --- Flatten & Export to CSV ---
Write-Host "Processing and exporting to CSV..." -ForegroundColor Cyan

$csvData = foreach ($domain in $allDomains) {

    # Safely extract nested objects
    $org          = $domain.organization
    $container    = $domain.container
    $dcv          = $domain.dcv_method
    $validation   = $domain.validation

    # Flatten per-validation-type expiry dates (OV / EV)
    $ovExpiry = ($validation | Where-Object { $_.type -eq "ov" } | Select-Object -ExpandProperty verified_until -ErrorAction SilentlyContinue)
    $evExpiry = ($validation | Where-Object { $_.type -eq "ev" } | Select-Object -ExpandProperty verified_until -ErrorAction SilentlyContinue)

    [PSCustomObject]@{
        # Domain details
        DomainId             = $domain.id
        DomainName           = $domain.name
        Status               = $domain.status
        IsActive             = $domain.is_active
        DateCreated          = if ($domain.date_created) { [datetimeoffset]::Parse($domain.date_created).ToString("yyyy-MM-dd") } else { $null }

        # DCV (Domain Control Validation)
        DCVMethod            = $dcv
        DCVToken             = $domain.dcv_token.token
        DCVTokenExpiry       = $domain.dcv_token.expiration_date

        # Validation expiry per type
        OVVerifiedUntil      = $ovExpiry
        EVVerifiedUntil      = $evExpiryontinue

        # Organization
        OrgId                = $org.id
        OrgName              = $org.name

        # Container / Division
        ContainerId          = $container.id
        ContainerName        = $container.name
    }
}

try {
    $csvData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExport complete: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
    Write-Host "Records written: $($csvData.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to write CSV to '$OutputPath': $_"
    exit 1
}
#endregion