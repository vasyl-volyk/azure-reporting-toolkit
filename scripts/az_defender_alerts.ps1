param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)


# Validate input
If ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

# Connect to Graph
$SecuredPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPassword
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential

if (-not (Get-MgContext)) {
    Throw "Authentication failed. Make sure app registration has proper API permissions (SecurityEvents.Read.All, etc.)."
}

# Define date range (last 7 days)
$startDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Query Defender alerts
$alerts = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/alerts?$filter=createdDateTime ge $startDate&`$top=1000"

$alerts = $alerts.value

# Filter for "account disabled" type alerts
#$disabledAlerts = $alerts.value #| Where-Object {
#    ($_?.title -match 'account.*disabled') -or ($_?.description -match 'account.*disabled')
#}


# Extract basic account info
$results = $alerts | Select-Object `
    @{Name = 'Timestamp'; Expression = { $_.createdDateTime.ToString("yyyy-MM-ddTHH:mm:ss") } },
    @{Name = 'Account'; Expression = { (($_).userStates.userPrincipalName.split(",") -join "<br>") } },
    @{Name = 'AlertTitle'; Expression = { $_.title } },
    @{Name = 'Description'; Expression = { $_.description } },
    @{Name = 'Severity'; Expression = { $_.severity } },
    @{Name = 'AlertId'; Expression = { $_.id } }

# Show results in console
$results.count

# Optionally export to CSV
if ($OutputPath) {
    $results | Sort-Object Timestamp -Descending | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation
    Write-Host "Results exported to: $OutputPath"
}
