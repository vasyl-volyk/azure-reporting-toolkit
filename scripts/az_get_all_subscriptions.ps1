param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)




# Validate input
if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    exit 1
}

# Connect
$securePassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ClientId, $securePassword)
Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $credential > $null

$results = @()

# Get billing accounts (MCA only)
$billingAccounts = Get-AzBillingAccount | Where-Object { $_.AgreementType -eq "MicrosoftCustomerAgreement" }

foreach ($billingAccount in $billingAccounts) {
    $billingAccountId = $billingAccount.Name

    $billingProfiles = Get-AzBillingProfile -BillingAccountName $billingAccountId

    foreach ($billingProfile in $billingProfiles) {
        $billingProfileId = $billingProfile.Name
        $invoiceSections = Get-AzInvoiceSection -BillingAccountName $billingAccountId -BillingProfileName $billingProfileId

        foreach ($invoiceSection in $invoiceSections) {
            $invoiceSectionId = $invoiceSection.Name

            $allSubs = @()

            # Initial URI
            $uri = "/providers/Microsoft.Billing/billingAccounts/$billingAccountId/billingProfiles/$billingProfileId/invoiceSections/$invoiceSectionId/billingSubscriptions?api-version=2024-04-01"

            do {
                # Make request
                $response = Invoke-AzRestMethod -Method GET -Path $uri
                $data = $response.Content | ConvertFrom-Json

                # Add current page results
                $allSubs += $data.value

                # Get nextLink, if present
                $uri = $data.nextLink

                # Remove the base URL if nextLink is full
                if ($uri -like "https://management.azure.com/*") {
                    $uri = $uri -replace "https://management.azure.com", ""
                }

            } while ($uri)


            if ($allSubs.properties) {
                foreach ($sub in $allSubs.properties) {

                    # Default value if lookup fails
                    $serviceTenantId = ""
    <#
                    try {
                        # Make ARM request to get tenant ID for subscription
                        $subDetails = Invoke-AzRestMethod -Method GET -Path "/subscriptions/$($sub.subscriptionId)?api-version=2020-01-01"
                        $subJson = $subDetails.Content | ConvertFrom-Json
                        $serviceTenantId = $subJson.tenantId
                    }
                    catch {
                        Write-Warning "Could not get tenantId for subscription $($sub.subscriptionId)"
                    }
                                        ServiceTenantId  = $serviceTenantId
    #>
                    $results += [PSCustomObject]@{
                        BillingScope     = $billingAccount.DisplayName
                        BillingProfile   = $sub.billingProfileDisplayName
                        InvoiceSection   = $sub.invoiceSectionDisplayName
                        SubscriptionName = $sub.displayName
                        SubscriptionId   = $sub.subscriptionId
                        skuDescription   = $sub.skuDescription
                        status           = $sub.status
                        billingFrequency = $sub.billingFrequency
                    }
                }
            }
            else {
                    $results += [PSCustomObject]@{
                        BillingScope     = $billingAccount.DisplayName
                        BillingProfile   = $billingProfile.DisplayName
                        InvoiceSection   = $invoiceSection.DisplayName
                        SubscriptionName = ""
                        SubscriptionId   = ""
                        skuDescription   = ""
                        status           = ""
                        billingFrequency = ""
                    }
            }
        }
    }
}


if ($OutputPath) {
    # Export
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to $OutputPath"
}
else
{ $results | Out-GridView }
