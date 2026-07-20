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
    $OutputPath = "DigiCert_Orders_$timestamp.csv"
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

#region --- Fetch All Orders (paginated) ---
Write-Host "Fetching orders from DigiCert..." -ForegroundColor Cyan

$allOrders = [System.Collections.Generic.List[object]]::new()
$offset    = 0
$page      = 1

do {
    Write-Host "  Fetching page $page (offset: $offset)..." -NoNewline

    $params = @{
        limit  = $PageLimit
        offset = $offset
    }

    $response = Invoke-DigiCertAPI -Endpoint "/order/certificate" -QueryParams $params

    if (-not $response -or -not $response.orders) {
        Write-Host " No data returned. Stopping." -ForegroundColor Yellow
        break
    }

    $batch = $response.orders
    $allOrders.AddRange($batch)
    Write-Host " Retrieved $($batch.Count) orders." -ForegroundColor Green

    $total  = $response.page.total
    $offset += $PageLimit
    $page++

} while ($allOrders.Count -lt $total)

Write-Host "`nTotal orders fetched: $($allOrders.Count)" -ForegroundColor Cyan
#endregion

#region --- Flatten & Export to CSV ---
Write-Host "Processing and exporting to CSV..." -ForegroundColor Cyan

$csvData = foreach ($order in $allOrders) {

    # Safely extract nested objects
    $cert         = $order.certificate
    $org          = $cert.organization
    $validity     = $cert.validity_years
    $product      = $order.product
    $container    = $order.container

    [PSCustomObject]@{
        # Order details
        OrderId              = $order.id
        Status               = $order.status
        OrderDate            = $order.date_created
        RenewalDate          = $order.date_renewed
        AutoRenew            = $order.auto_renew
        DisableRenewalNotif  = $order.disable_renewal_notifications

        # Product details
        ProductName          = $product.name
        ProductType          = $product.type
        ProductNameId        = $product.name_id

        # Certificate details
        CommonName           = $cert.common_name
        DNSNames             = ($cert.dns_names -join "; ")
        SerialNumber         = $cert.serial_number
        Thumbprint           = $cert.thumbprint
        SignatureHash        = $cert.signature_hash
        ValidFrom            = $cert.valid_from
        ValidTill            = $cert.valid_till
        ValidityYears        = $validity
        KeySize              = $cert.key_size
        CACertId             = $cert.ca_cert_id

        # Organization
        OrgId                = $org.id
        OrgName              = $org.name

        # Container / Division
        ContainerId          = $container.id
        ContainerName        = $container.name

        # Financial
        Amount               = $order.amount
        Currency             = "USD"
    }
}

$OrderedData = $csvData | select-object OrderId, Status, OrderDate,DisableRenewalNotif,ProductName,ProductType,ProductNameId,CommonName,DNSNames,SerialNumber,Thumbprint,SignatureHash,ValidFrom,ValidTill,ContainerId,ContainerName,ValidityYears,KeySize,CACertId,OrgId,Amount,OrgName,RenewalDate,AutoRenew 

try {
    $OrderedData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nExport complete: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
    Write-Host "Records written: $($csvData.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to write CSV to '$OutputPath': $_"
    exit 1
}
#endregion