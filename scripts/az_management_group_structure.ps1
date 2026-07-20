param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$TargetMGroupId,
    [string]$OutputPath,
    [int]$MaxRetries = 5,
    [int]$BaseDelaySeconds = 2
)

write-host "TargetMGroupId -> ",$TargetMGroupId


    # Clean the ManagementGroupId
    if (-not [string]::IsNullOrEmpty($ManagementGroupId)) {
        $ManagementGroupId = $ManagementGroupId.Trim()
        # Remove any potential URI encoding or special characters
        $ManagementGroupId = [System.Uri]::UnescapeDataString($ManagementGroupId)
    }

$scope = "https://management.azure.com/.default"

# Function to invoke REST method with retry logic
function Invoke-RestMethodWithRetry {
    param (
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "Get",
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )
    
    $attempt = 0
    $success = $false
    $response = $null
    
    while (-not $success -and $attempt -lt $MaxRetries) {
        try {
            $attempt++
            $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
            $success = $true
            
            # Add a small delay between successful requests to avoid rate limiting
            Start-Sleep -Milliseconds 200
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            if ($statusCode -eq 429) {
                # Check for Retry-After header
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                
                if ($retryAfter) {
                    $waitTime = [int]$retryAfter
                    Write-Warning "Rate limit hit (429). Retry-After header suggests waiting $waitTime seconds. Attempt $attempt of $MaxRetries"
                } else {
                    # Exponential backoff: 2, 4, 8, 16, 32 seconds
                    $waitTime = [Math]::Pow($BaseDelay, $attempt)
                    Write-Warning "Rate limit hit (429). Using exponential backoff: $waitTime seconds. Attempt $attempt of $MaxRetries"
                }
                
                if ($attempt -lt $MaxRetries) {
                    Start-Sleep -Seconds $waitTime
                } else {
                    throw "Max retries ($MaxRetries) exceeded for URI: $Uri. Error: $($_.Exception.Message)"
                }
            }
            else {
                throw "API call failed with status code $statusCode. Error: $($_.Exception.Message)"
            }
        }
    }
    
    return $response
}

# Function to authenticate using client credentials grant flow
function Get-AzureToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Scope
    )

    $body = @{
        client_id     = $ClientId
        scope         = $Scope
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
    return $response.access_token
}

# Get management groups with expanded children information
function Get-ManagementGroups {
    param (
        [string]$AccessToken,
        [string]$ManagementGroupId = "",
        [switch]$IncludeNested = $true,
        [string]$RGPath = "./",
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    $uriBase = "https://management.azure.com/providers/Microsoft.Management/managementGroups"
    $apiVersion = "2021-04-01"

    if ([string]::IsNullOrEmpty($ManagementGroupId)) {
        $uri = "$uriBase?api-version=$apiVersion"
    } else {
        $uri = "$uriBase/$ManagementGroupId" + "?api-version=$apiVersion"
        if ($IncludeNested) {
            $uri += "&`$expand=children&`$recurse=true"
        }
    }

    try {
        $response = Invoke-RestMethodWithRetry -Uri $uri -Headers $headers -Method Get -MaxRetries $MaxRetries -BaseDelay $BaseDelay
    }
    catch {
        Write-Error "Failed to retrieve management groups. Error: $($_.Exception.Message)"
        return $null
    }

    Write-Host "Processing management group: $($response.name)"

    $managementGroups = @()

    if (-not $response.properties.children) {
        $managementGroups += [PSCustomObject]@{
            "Path"        = $RGPath
            "id"          = $response.id
            "type"        = $response.type
            "name"        = $response.name
            "displayName" = $response.properties.displayName
        }
    } 
    elseif ($IncludeNested -and $response.properties.children) {        
        $managementGroups += [PSCustomObject]@{
            "Path"        = $RGPath
            "id"          = $response.id
            "type"        = $response.type
            "name"        = $response.name
            "displayName" = $response.properties.displayName
        }
        
        foreach ($child in $response.properties.children) {
            if ($child.id -match "managementGroups") {
                $ChildID = $child.id -replace "/providers/Microsoft.Management/managementGroups/", ""
                $p = $RGPath + "/" + $response.name
                Write-Host "Processing nested management group: $ChildID at path: $p"
                $managementGroups += Get-ManagementGroups -AccessToken $AccessToken -ManagementGroupId $ChildID -IncludeNested $true -RGPath $p -MaxRetries $MaxRetries -BaseDelay $BaseDelay
            }
        }
    }
    return $managementGroups
}

# Get subscriptions under a management group with detailed information
function Get-Subscriptions {
    param (
        [string]$ManagementGroupId,
        [string]$AccessToken,
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    $ManagementGroupId = $ManagementGroupId -replace '^/providers/Microsoft\.Management/managementGroups/', ''
    $uri = "https://management.azure.com/providers/Microsoft.Management/managementGroups/$ManagementGroupId/subscriptions?api-version=2020-05-01"
    
    try {
        $response = Invoke-RestMethodWithRetry -Uri $uri -Headers $headers -Method Get -MaxRetries $MaxRetries -BaseDelay $BaseDelay
        return $response.value
    }
    catch {
        Write-Warning "Failed to retrieve subscriptions for management group $ManagementGroupId. Error: $($_.Exception.Message)"
        return @()
    }
}

# Get detailed subscription information
function Get-SubscriptionDetails {
    param (
        [string]$SubscriptionId,
        [string]$AccessToken,
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )
    
    $headers = @{ Authorization = "Bearer $AccessToken" }
    $uri = "https://management.azure.com/subscriptions/$SubscriptionId" + "?api-version=2020-01-01"
    
    try {
        $response = Invoke-RestMethodWithRetry -Uri $uri -Headers $headers -Method Get -MaxRetries $MaxRetries -BaseDelay $BaseDelay
        return $response
    }
    catch {
        Write-Warning "Failed to get details for subscription $SubscriptionId. Error: $($_.Exception.Message)"
        return $null
    }
}

# Get subscriptions from management groups recursively
function Get-SubscriptionsFromManagementGroups {
    param (
        [array]$ManagementGroups,
        [string]$AccessToken,
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )
    
    $allSubscriptions = @()
    
    foreach ($mg in $ManagementGroups) {
        Write-Host "Getting subscriptions from management group: $($mg.displayName) ($($mg.name))"
        
        # Get subscriptions directly under this management group
        $subscriptions = Get-Subscriptions -ManagementGroupId $mg.name -AccessToken $AccessToken -MaxRetries $MaxRetries -BaseDelay $BaseDelay
        
        foreach ($sub in $subscriptions) {
            Write-Host "Processing subscription: $($sub.properties.displayName) ($($sub.name))"
            
            # Get detailed subscription information
            $subDetails = Get-SubscriptionDetails -SubscriptionId $sub.subscriptionId -AccessToken $AccessToken -MaxRetries $MaxRetries -BaseDelay $BaseDelay
            
            if ( $mg.Path.contains("00000000-0000-0000-0000-000000000000") ){
                $subscriptionPath = $mg.Path.Replace("/00000000-0000-0000-0000-000000000000","") + "/" + $mg.displayName + "/" + $sub.displayName
            }
            else{
                $subscriptionPath = $mg.Path + "/" + $mg.displayName + "/" + $sub.displayName
            }
            
            $allSubscriptions += [PSCustomObject]@{
                "Path"             = $subscriptionPath
                "SubscriptionName" = $sub.properties.displayName
                "SubscriptionId"   = $sub.name
                "State"            = $sub.properties.state
            }
        }
    }
    
    return $allSubscriptions
}

# Get resource groups in a subscription
function Get-ResourceGroups {
    param (
        [string]$SubscriptionId,
        [string]$AccessToken,
        [int]$MaxRetries = 5,
        [int]$BaseDelay = 2
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    $uri = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups?api-version=2022-12-01"
    
    try {
        $response = Invoke-RestMethodWithRetry -Uri $uri -Headers $headers -Method Get -MaxRetries $MaxRetries -BaseDelay $BaseDelay
        return $response.value
    }
    catch {
        Write-Warning "Failed to retrieve resource groups for subscription $SubscriptionId. Error: $($_.Exception.Message)"
        return @()
    }
}

# Extract tag value
function Get-TagValue {
    param (
        [object]$Tags,
        [string]$TagName
    )

    if ($null -eq $Tags) {
        return $null
    }

    # If it's a hashtable
    if ($Tags -is [hashtable] -and $Tags.ContainsKey($TagName)) {
        return $Tags[$TagName]
    }

    # If it's an object (PSCustomObject), convert to hashtable
    if ($Tags -is [pscustomobject]) {
        $tagsHashtable = @{}
        foreach ($prop in $Tags.PSObject.Properties) {
            $tagsHashtable[$prop.Name] = $prop.Value
        }
        if ($tagsHashtable.ContainsKey($TagName)) {
            return $tagsHashtable[$TagName]
        }
    }

    return $null
}

# Main execution
try {
    Write-Host "Starting Azure Management Groups and Subscriptions inventory..."
    Write-Host "Using MaxRetries: $MaxRetries, BaseDelay: $BaseDelaySeconds seconds"
    
    # Get the access token
    Write-Host "Authenticating..."
    $accessToken = Get-AzureToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret -Scope $scope
    
    if ([string]::IsNullOrEmpty($accessToken)) {
        throw "Failed to obtain access token"
    }
    
    Write-Host "Getting management groups structure..."
    $managementGroups = Get-ManagementGroups -AccessToken $AccessToken -ManagementGroupId $TargetMGroupID -IncludeNested $true -RGpath "Root" -MaxRetries $MaxRetries -BaseDelay $BaseDelaySeconds
    
    if ($null -eq $managementGroups -or $managementGroups.Count -eq 0) {
        throw "No management groups found or failed to retrieve management groups"
    }
    
    Write-Host "Found $($managementGroups.Count) management groups"
    
    Write-Host "Collecting subscription information..."
    $results = Get-SubscriptionsFromManagementGroups -ManagementGroups $managementGroups -AccessToken $accessToken -MaxRetries $MaxRetries -BaseDelay $BaseDelaySeconds
    
    Write-Host "Found $($results.Count) subscriptions"
    
    # Display results
    Write-Host "`n=== SUBSCRIPTION INVENTORY RESULTS ===" -ForegroundColor Green
    $results | Format-Table -Property Path, SubscriptionName, SubscriptionId, State -AutoSize
    
    # Export results if OutputPath is specified
    if (-not [string]::IsNullOrEmpty($OutputPath)) {
        Write-Host "Exporting results to: $OutputPath"
        
        # Export to CSV
        $csvPath = $OutputPath -replace '\.json$', '.csv'
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV exported to: $csvPath"
    }
    
    # Return results
    return $results
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
    exit 1
}