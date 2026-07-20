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
    $OutputPath = "DigiCert_Certificates_$timestamp.csv"
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

#region --- Helper: Parse date string to yyyy-MM-dd (PS5 safe) ---
function Format-Date {
    param ([string]$Value)
    if (-not $Value) { return $null }
    try { return [datetimeoffset]::Parse($Value).ToString("yyyy-MM-dd") }
    catch { return $Value }
}
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

#region --- Fetch All Issued Certificates (paginated) ---
Write-Host "Fetching issued certificates from DigiCert..." -ForegroundColor Cyan

$allCerts = [System.Collections.Generic.List[object]]::new()
$offset   = 0
$page     = 1

do {
    Write-Host "  Fetching page $page (offset: $offset)..." -NoNewline

    $params = @{
        limit         = $PageLimit
        offset        = $offset
        status        = "issued"
    }

    $response = Invoke-DigiCertAPI -Endpoint "/order/certificate" -QueryParams $params

    if (-not $response -or -not $response.orders) {
        Write-Host " No data returned. Stopping." -ForegroundColor Yellow
        break
    }

    $batch = $response.orders
    $allCerts.AddRange($batch)
    Write-Host " Retrieved $($batch.Count) certificates." -ForegroundColor Green

    $total  = $response.page.total
    $offset += $PageLimit
    $page++

} while ($allCerts.Count -lt $total)

Write-Host "`nTotal certificates fetched: $($allCerts.Count)" -ForegroundColor Cyan
#endregion

#region --- Flatten & Export to CSV ---
Write-Host "Processing and exporting to CSV..." -ForegroundColor Cyan

$csvData = foreach ($order in $allCerts) {

    $cert    = $order.certificate
    $org     = $cert.organization
    $product = $order.product

    [PSCustomObject]@{
        # Expiry countdown
        DaysLeft  = if ($cert.valid_till) { ([datetimeoffset]::Parse($cert.valid_till).Date - [datetime]::Today).Days.ToString("0000") } else { $null }
        ValidTill        = Format-Date $cert.valid_till

        # Order
        OrderId          = $order.id
        Status           = $order.status

        # Certificate
        CommonName       = $cert.common_name
        DNSNames         = ($cert.dns_names -join "; ")
        SerialNumber     = $cert.serial_number
        Thumbprint       = $cert.thumbprint
        SignatureHash    = $cert.signature_hash

        # Dates (sortable yyyy-MM-dd)
        ValidFrom        = Format-Date $cert.valid_from
#        ValidTill        = Format-Date $cert.valid_till
        DateCreated      = Format-Date $order.date_created

        # Product
        ProductName      = $product.name
        ProductType      = $product.type

        # Organization
        OrgId            = $org.id
        OrgName          = $org.name
    }
}

try {
    $csvData | Where-Object{$_.status -like "issued"} | Sort-object DaysLeft | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExport complete: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
    Write-Host "Records written: $($csvData.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to write CSV to '$OutputPath': $_"
    exit 1
}
#endregion