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

# Helper function to resolve owner IDs to readable names
function Resolve-OwnerIds {
    param(
        [string[]]$Ids
    )

    $ownerTable = @{}
    if (-not $Ids -or $Ids.Count -eq 0) { 
        return $ownerTable 
    }

    Write-Host "Resolving $($Ids.Count) owner IDs..." -ForegroundColor Yellow

    # Process in chunks of 100 (API limit)
    for ($i = 0; $i -lt $Ids.Count; $i += 100) {
        $chunkEnd = [Math]::Min($i + 99, $Ids.Count - 1)
        $chunk = $Ids[$i..$chunkEnd]
        
        $requestBody = @{
            ids = $chunk
            types = @()  # Empty array means all types
        } | ConvertTo-Json -Depth 4

        try {
            Write-Host "  Processing chunk $([math]::Floor($i/100) + 1)..." -ForegroundColor Gray
            
            $response = Invoke-MgGraphRequest `
                -Method POST `
                -Uri 'https://graph.microsoft.com/v1.0/directoryObjects/getByIds' `
                -Body $requestBody `
                -ContentType 'application/json'

            foreach ($obj in $response.value) {
                $objectId = $obj.id
                $objectType = $obj.'@odata.type'

                switch ($objectType) {
                    '#microsoft.graph.user' {
                        $displayName = if ($obj.displayName) { $obj.displayName } else { "Unknown User" }
                        $upn = if ($obj.userPrincipalName) { $obj.userPrincipalName } else { "no-upn@unknown.com" }
                        $ownerTable[$objectId] = "$displayName ($upn)"
                    }
                    '#microsoft.graph.servicePrincipal' {
                        $displayName = if ($obj.displayName) { $obj.displayName } else { "Unknown Service Principal" }
                        $ownerTable[$objectId] = "$displayName [Service Principal]"
                    }
                    '#microsoft.graph.group' {
                        $displayName = if ($obj.displayName) { $obj.displayName } else { "Unknown Group" }
                        if ($obj.mail) {
                            $ownerTable[$objectId] = "$displayName [Group: $($obj.mail)]"
                        } else {
                            $ownerTable[$objectId] = "$displayName [Group]"
                        }
                    }
                    default {
                        $displayName = if ($obj.displayName) { $obj.displayName } else { "Unknown Object" }
                        $ownerTable[$objectId] = "$displayName [Other: $objectType]"
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to resolve owner chunk: $($_.Exception.Message)"
            # Add the failed IDs as unresolved to avoid nulls
            foreach ($failedId in $chunk) {
                if (-not $ownerTable.ContainsKey($failedId)) {
                    $ownerTable[$failedId] = "$failedId [Unresolved]"
                }
            }
        }
    }

    Write-Host "  Resolved $($ownerTable.Count) owners successfully" -ForegroundColor Green
    return $ownerTable
}

# Connect to Graph
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-MgGraph -Scopes ""Application.Read.All"", ""Group.Read.All"", ""Policy.Read.All"", ""RoleManagement.Read.Directory"", ""User.Read.All""'"
}

Write-Host "Connected to Microsoft Graph. Retrieving applications..." -ForegroundColor Green

# Initialize results ArrayList for better performance
$results = [System.Collections.ArrayList]::new()

try {
    # Get all applications with specific properties and expanded owners
    Write-Host "Fetching all applications with optimized properties..." -ForegroundColor Yellow
    $applications = Get-MgApplication -All -Property "Id,AppId,DisplayName,PublisherDomain,SignInAudience,CreatedDateTime,RequiredResourceAccess,AppRoles,Web,PublicClient,Spa,PasswordCredentials,KeyCredentials,Notes,Tags" -ExpandProperty "Owners"
    
    Write-Host "Found $($applications.Count) applications. Pre-loading reference data..." -ForegroundColor Yellow
    
    # Collect all unique ResourceAppIds for service principal loading
    $allResourceAppIds = @()
    foreach ($app in $applications) {
        if ($app.RequiredResourceAccess) {
            $allResourceAppIds += $app.RequiredResourceAccess.ResourceAppId
        }
    }
    $allResourceAppIds = $allResourceAppIds | Sort-Object -Unique | Where-Object { $_ }
    
    Write-Host "Loading $($allResourceAppIds.Count) unique service principals..." -ForegroundColor Cyan
    $servicePrincipals = @{}
    
    if ($allResourceAppIds.Count -gt 0) {
        # Special handling for Microsoft Graph (most common case)
        if ($allResourceAppIds -contains "00000003-0000-0000-c000-000000000000") {
            Write-Host "Loading Microsoft Graph service principal..." -ForegroundColor Yellow
            try {
                $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -Property "AppId,DisplayName,AppRoles,Oauth2PermissionScopes" -ErrorAction Stop
                if ($graphSp) {
                    # Handle case where Get-MgServicePrincipal returns an array
                    $graphSpObject = if ($graphSp -is [array]) { $graphSp[0] } else { $graphSp }
                    $servicePrincipals["00000003-0000-0000-c000-000000000000"] = $graphSpObject
                    Write-Host "  ✓ Microsoft Graph loaded: $($graphSpObject.DisplayName)" -ForegroundColor Green
                } else {
                    Write-Warning "Microsoft Graph service principal not found - attempting to create it..."
                    try {
                        $newGraphSp = New-MgServicePrincipal -AppId "00000003-0000-0000-c000-000000000000" -ErrorAction Stop
                        $servicePrincipals["00000003-0000-0000-c000-000000000000"] = $newGraphSp
                        Write-Host "  ✓ Microsoft Graph service principal created: $($newGraphSp.DisplayName)" -ForegroundColor Green
                        Start-Sleep -Seconds 3  # Allow time for propagation
                    } catch {
                        Write-Warning "Failed to create Microsoft Graph service principal: $($_.Exception.Message)"
                    }
                }
            } catch {
                Write-Warning "Failed to load Microsoft Graph SP: $($_.Exception.Message)"
                # Try alternative approach
                try {
                    Write-Host "Trying alternative query method..." -ForegroundColor Yellow
                    $graphSp = Get-MgServicePrincipal -All | Where-Object { $_.AppId -eq "00000003-0000-0000-c000-000000000000" } | Select-Object -First 1
                    if ($graphSp) {
                        $servicePrincipals["00000003-0000-0000-c000-000000000000"] = $graphSp
                        Write-Host "  ✓ Microsoft Graph found via alternative method: $($graphSp.DisplayName)" -ForegroundColor Green
                    }
                } catch {
                    Write-Warning "Alternative method also failed: $($_.Exception.Message)"
                }
            }
        }
        
        # Process other service principals in batches
        $otherAppIds = $allResourceAppIds | Where-Object { $_ -ne "00000003-0000-0000-c000-000000000000" }
        
        if ($otherAppIds.Count -gt 0) {
            Write-Host "Loading $($otherAppIds.Count) other service principals in batches..." -ForegroundColor Cyan
            for ($i = 0; $i -lt $otherAppIds.Count; $i += 15) {  # Reduced batch size for reliability
                $batchEnd = [Math]::Min($i + 14, $otherAppIds.Count - 1)
                $batch = $otherAppIds[$i..$batchEnd]
                $filterQuery = ($batch | ForEach-Object { "AppId eq '$_'" }) -join " or "
                
                try {
                    Write-Host "  Fetching batch $([math]::Floor($i/15) + 1)..." -ForegroundColor Gray
                    $spBatch = Get-MgServicePrincipal -Filter $filterQuery -Property "AppId,DisplayName,AppRoles,Oauth2PermissionScopes" -ErrorAction SilentlyContinue
                    if ($spBatch) {
                        # Handle both single results and arrays
                        $spArray = if ($spBatch -is [array]) { $spBatch } else { @($spBatch) }
                        $spArray | ForEach-Object { 
                            $servicePrincipals[$_.AppId] = $_
                            Write-Host "    Loaded: $($_.DisplayName) ($($_.AppId))" -ForegroundColor Gray
                        }
                    }
                } catch {
                    Write-Warning "Error fetching service principal batch: $($_.Exception.Message)"
                    # Try individual queries for this batch
                    foreach ($appId in $batch) {
                        try {
                            $sp = Get-MgServicePrincipal -Filter "AppId eq '$appId'" -Property "AppId,DisplayName,AppRoles,Oauth2PermissionScopes" -ErrorAction SilentlyContinue
                            if ($sp) {
                                $spObject = if ($sp -is [array]) { $sp[0] } else { $sp }
                                $servicePrincipals[$appId] = $spObject
                                Write-Host "    Individual query success: $($spObject.DisplayName)" -ForegroundColor Gray
                            }
                        } catch {
                            Write-Warning "Failed individual query for $appId"
                        }
                    }
                }
            }
        }
    }
    
    # Enhanced verification
    Write-Host "`nService principals cache verification:" -ForegroundColor Cyan
    Write-Host "  Total cached: $($servicePrincipals.Count)" -ForegroundColor White
    if ($servicePrincipals.ContainsKey("00000003-0000-0000-c000-000000000000")) {
        $graphSp = $servicePrincipals["00000003-0000-0000-c000-000000000000"]
        Write-Host "  ✓ Microsoft Graph: $($graphSp.DisplayName) (AppRoles: $($graphSp.AppRoles.Count), Scopes: $($graphSp.Oauth2PermissionScopes.Count))" -ForegroundColor Green
    } else {
        Write-Warning "  ✗ Microsoft Graph is NOT cached - permissions will show as GUIDs"
    }
    
    # IMPROVED OWNER RESOLUTION using directoryObjects/getByIds
    $allOwnerIds = @()
    foreach ($app in $applications) {
        if ($app.Owners) {
            $allOwnerIds += $app.Owners.Id
        }
    }
    $allOwnerIds = $allOwnerIds | Sort-Object -Unique | Where-Object { $_ }
    
    # Use the new function to resolve all owner IDs at once
    $ownerLookup = Resolve-OwnerIds -Ids $allOwnerIds
    
    Write-Host "`nProcessing applications with cached data..." -ForegroundColor Yellow
    
    # Process applications
    $processedCount = 0
    foreach ($app in $applications) {
        $processedCount++
        if ($processedCount % 25 -eq 0) {
            Write-Host "Processed $processedCount of $($applications.Count) applications..." -ForegroundColor Gray
        }
        
        # FIXED OWNER PROCESSING - Now uses the lookup table
        $ownerList = @()
        if ($app.Owners) {
            $ownerList = $app.Owners | ForEach-Object { 
                $resolvedOwner = $ownerLookup[$_.Id]
                if ($resolvedOwner) {
                    $resolvedOwner
                } else {
                    "$($_.Id) [Unresolved]"  # Fallback, should rarely happen
                }
            }
        }
        
        # Get API permissions with improved logic
        $apiPermissions = @()
        if ($app.RequiredResourceAccess) {
            foreach ($resource in $app.RequiredResourceAccess) {
                $servicePrincipal = $servicePrincipals[$resource.ResourceAppId]
                $resourceName = if ($servicePrincipal) { 
                    $servicePrincipal.DisplayName 
                } else { 
                    $resource.ResourceAppId 
                }
                
                foreach ($access in $resource.ResourceAccess) {
                    $permissionType = if ($access.Type -eq "Role") { "Application" } else { "Delegated" }
                    
                    # Default to GUID
                    $permissionName = $access.Id
                    $permissionDisplayName = $access.Id
                    
                    if ($servicePrincipal) {
                        if ($access.Type -eq "Role") {
                            # Application permission - look in AppRoles
                            $permission = $servicePrincipal.AppRoles | Where-Object { $_.Id -eq $access.Id }
                            if ($permission) { 
                                $permissionName = $permission.Value
                                $permissionDisplayName = if ($permission.DisplayName) { $permission.DisplayName } else { $permission.Value }
                            }
                        } else {
                            # Delegated permission - look in Oauth2PermissionScopes
                            $permission = $servicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Id -eq $access.Id }
                            if ($permission) { 
                                $permissionName = $permission.Value
                                $permissionDisplayName = if ($permission.AdminConsentDisplayName) { $permission.AdminConsentDisplayName } else { $permission.Value }
                            }
                        }
                    }
                    
                    $apiPermissions += "$resourceName : $permissionName ($permissionType)"
                }
            }
        }
        
        # Get app roles
        $appRolesList = @()
        if ($app.AppRoles) {
            $appRolesList = $app.AppRoles | ForEach-Object { 
                $roleInfo = "$($_.Value)"
                if ($_.DisplayName -and $_.DisplayName -ne $_.Value) {
                    $roleInfo += " - $($_.DisplayName)"
                }
                $roleInfo
            }
        }
        
        # Get redirect URIs
        $redirectUris = @()
        if ($app.Web -and $app.Web.RedirectUris) {
            $redirectUris += $app.Web.RedirectUris
        }
        if ($app.PublicClient -and $app.PublicClient.RedirectUris) {
            $redirectUris += $app.PublicClient.RedirectUris
        }
        if ($app.Spa -and $app.Spa.RedirectUris) {
            $redirectUris += $app.Spa.RedirectUris
        }
        
        # Check credentials
        $hasSecrets = ($app.PasswordCredentials -and $app.PasswordCredentials.Count -gt 0)
        $hasCertificates = ($app.KeyCredentials -and $app.KeyCredentials.Count -gt 0)
        
        # Get expiration dates
        $secretExpirations = @()
        $certExpirations = @()
        
        if ($app.PasswordCredentials) {
            $secretExpirations = $app.PasswordCredentials | ForEach-Object { 
                $expiryDate = $_.EndDateTime
                if ($expiryDate) {
                    $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
                    "$($expiryDate.ToString('yyyy-MM-dd')) ($daysUntilExpiry days)"
                } else {
                    "No expiration"
                }
            }
        }
        
        if ($app.KeyCredentials) {
            $certExpirations = $app.KeyCredentials | ForEach-Object { 
                $expiryDate = $_.EndDateTime
                if ($expiryDate) {
                    $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
                    "$($expiryDate.ToString('yyyy-MM-dd')) ($daysUntilExpiry days)"
                } else {
                    "No expiration"
                }
            }
        }
        
        # Create result object
        $result = [PSCustomObject]@{
            'Application Name' = $app.DisplayName
            'Application ID' = $app.AppId
            'Created Date' = $app.CreatedDateTime
            'Owners' = ($ownerList -join '<br>')
            'API Permissions' = ($apiPermissions -join '<br>')
            'App Roles' = ($appRolesList -join '<br>')
            'Redirect URIs' = ($redirectUris -join '<br>')
            'Has Client Secrets' = $hasSecrets
            'Has Certificates' = $hasCertificates
            'Secret Expiration Dates' = ($secretExpirations -join '<br>')
            'Certificate Expiration Dates' = ($certExpirations -join '<br>')
            'Homepage URL' = if ($app.Web) { $app.Web.HomePageUrl } else { "" }
            'Notes' = $app.Notes
            'Tags' = ($app.Tags -join '<br>')
            'Publisher Domain' = $app.PublisherDomain
            'Sign-in Audience' = $app.SignInAudience
            'Object ID' = $app.Id
        }
        
        [void]$results.Add($result)
    }
    
    # Export results
    Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export completed successfully!" -ForegroundColor Green
    Write-Host "Results saved to: $OutputPath" -ForegroundColor Green
    Write-Host "Total applications exported: $($results.Count)" -ForegroundColor Green
    
    # Enhanced summary
    Write-Host "`nSummary:" -ForegroundColor Yellow
    Write-Host "- Total Applications: $($results.Count)"
    Write-Host "- Applications with Secrets: $(($results | Where-Object { $_.'Has Client Secrets' -eq $true }).Count)"
    Write-Host "- Applications with Certificates: $(($results | Where-Object { $_.'Has Certificates' -eq $true }).Count)"
    Write-Host "- Applications with API Permissions: $(($results | Where-Object { $_.'API Permissions' -ne '' }).Count)"
    Write-Host "- Applications with Owners: $(($results | Where-Object { $_.'Owners' -ne '' }).Count)"
    Write-Host "- Owner objects resolved: $($ownerLookup.Count)"
    
} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Error $_.Exception.StackTrace
} finally {
    # Disconnect from Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "`nDisconnected from Microsoft Graph." -ForegroundColor Green
}
