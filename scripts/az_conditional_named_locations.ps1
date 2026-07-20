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

# Import required module
try {
    Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
    Write-Host "Microsoft.Graph.Identity.SignIns module loaded successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import Microsoft.Graph.Identity.SignIns module. Please install it using: Install-Module Microsoft.Graph.Identity.SignIns"
    exit 1
}

# Connect to Graph
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
    
    if (-not $(Get-MgContext)) {
        throw "Authentication needed, call 'Connect-MgGraph -Scopes `"Policy.Read.All`"'"
    }
    
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    Write-Host "Tenant: $((Get-MgContext).TenantId)" -ForegroundColor Gray
    Write-Host "Account: $((Get-MgContext).Account)" -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# Retrieve Named Locations
try {
    Write-Host "Retrieving Named Locations..." -ForegroundColor Cyan
    $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All
    
    if ($null -eq $namedLocations -or $namedLocations.Count -eq 0) {
        Write-Warning "No Named Locations found in this tenant."
        Disconnect-MgGraph | Out-Null
        exit 0
    }
    
    Write-Host "Found $($namedLocations.Count) Named Location(s)." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Error "Failed to retrieve Named Locations: $_"
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Process and display Named Locations
$results = @()

foreach ($location in $namedLocations) {
    Write-Host "=" * 80 -ForegroundColor Yellow
    Write-Host "Named Location: $($location.DisplayName)" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Yellow
    
    # Determine location type
    $locationType = if ($location.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.ipNamedLocation') {
        "IP-based"
    }
    elseif ($location.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.countryNamedLocation') {
        "Country-based"
    }
    else {
        "Unknown"
    }
    
    Write-Host "ID                    : $($location.Id)" -ForegroundColor White
    Write-Host "Display Name          : $($location.DisplayName)" -ForegroundColor White
    Write-Host "Location Type         : $locationType" -ForegroundColor White
    Write-Host "Created Date/Time     : $($location.CreatedDateTime)" -ForegroundColor White
    Write-Host "Modified Date/Time    : $($location.ModifiedDateTime)" -ForegroundColor White
    
    # IP-based Named Location details
    if ($locationType -eq "IP-based") {
        $isTrusted = $location.AdditionalProperties['isTrusted']
        Write-Host "Is Trusted            : $isTrusted" -ForegroundColor White
        
        $ipRanges = $location.AdditionalProperties['ipRanges']
        if ($ipRanges -and $ipRanges.Count -gt 0) {
            Write-Host "IP Ranges             : $($ipRanges.Count) range(s)" -ForegroundColor White
            foreach ($ipRange in $ipRanges) {
                Write-Host "  - $($ipRange['cidrAddress'])" -ForegroundColor Gray
            }
            
            # Create result object for IP-based location
            $resultObj = [PSCustomObject]@{
                DisplayName      = $location.DisplayName
                Id               = $location.Id
                Type             = $locationType
                IsTrusted        = $isTrusted
                IPRanges         = ($ipRanges | ForEach-Object { $_['cidrAddress'] }) -join '<br>'
                Countries        = $null
                IncludeUnknown   = $null
                CreatedDateTime  = $location.CreatedDateTime
                ModifiedDateTime = $location.ModifiedDateTime
            }
        }
        else {
            Write-Host "IP Ranges             : None" -ForegroundColor Gray
            
            $resultObj = [PSCustomObject]@{
                DisplayName      = $location.DisplayName
                Id               = $location.Id
                Type             = $locationType
                IsTrusted        = $isTrusted
                IPRanges         = $null
                Countries        = $null
                IncludeUnknown   = $null
                CreatedDateTime  = $location.CreatedDateTime
                ModifiedDateTime = $location.ModifiedDateTime
            }
        }
    }
    # Country-based Named Location details
    elseif ($locationType -eq "Country-based") {
        $includeUnknown = $location.AdditionalProperties['includeUnknownCountriesAndRegions']
        Write-Host "Include Unknown       : $includeUnknown" -ForegroundColor White
        
        $countries = $location.AdditionalProperties['countriesAndRegions']
        if ($countries -and $countries.Count -gt 0) {
            Write-Host "Countries/Regions     : $($countries.Count) location(s)" -ForegroundColor White
            foreach ($country in $countries) {
                Write-Host "  - $country" -ForegroundColor Gray
            }
            
            # Create result object for country-based location
            $resultObj = [PSCustomObject]@{
                DisplayName      = $location.DisplayName
                Id               = $location.Id
                Type             = $locationType
                IsTrusted        = $null
                IPRanges         = $null
                Countries        = ($countries -join '; ')
                IncludeUnknown   = $includeUnknown
                CreatedDateTime  = $location.CreatedDateTime
                ModifiedDateTime = $location.ModifiedDateTime
            }
        }
        else {
            Write-Host "Countries/Regions     : None" -ForegroundColor Gray
            
            $resultObj = [PSCustomObject]@{
                DisplayName      = $location.DisplayName
                Id               = $location.Id
                Type             = $locationType
                IsTrusted        = $null
                IPRanges         = $null
                Countries        = $null
                IncludeUnknown   = $includeUnknown
                CreatedDateTime  = $location.CreatedDateTime
                ModifiedDateTime = $location.ModifiedDateTime
            }
        }
    }
    else {
        # Unknown type
        $resultObj = [PSCustomObject]@{
            DisplayName      = $location.DisplayName
            Id               = $location.Id
            Type             = $locationType
            IsTrusted        = $null
            IPRanges         = $null
            Countries        = $null
            IncludeUnknown   = $null
            CreatedDateTime  = $location.CreatedDateTime
            ModifiedDateTime = $location.ModifiedDateTime
        }
    }
    
    $results += $resultObj
    Write-Host ""
}


$results  = $results | Select-Object DisplayName, Type, IsTrusted, IPRanges, Countries, IncludeUnknown, CreatedDateTime, ModifiedDateTime, Id

# Export to CSV if OutputPath is specified
if (-not [string]::IsNullOrEmpty($OutputPath)) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export results to CSV: $_"
    }
}

# Display summary
Write-Host "=" * 80 -ForegroundColor Yellow
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Yellow
Write-Host "Total Named Locations : $($namedLocations.Count)" -ForegroundColor White

$ipLocations = $results | Where-Object { $_.Type -eq "IP-based" }
$countryLocations = $results | Where-Object { $_.Type -eq "Country-based" }
$trustedLocations = $ipLocations | Where-Object { $_.IsTrusted -eq $true }

Write-Host "IP-based Locations    : $($ipLocations.Count)" -ForegroundColor White
Write-Host "Country-based Locations: $($countryLocations.Count)" -ForegroundColor White
Write-Host "Trusted IP Locations  : $($trustedLocations.Count)" -ForegroundColor White
Write-Host ""

# Disconnect from Graph
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Gray