param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)


if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

# Import required modules
try {
    Import-Module Az.Accounts -Force
    Import-Module Az.Compute -Force
    Import-Module Az.Resources -Force
    Write-Host "Azure PowerShell modules imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import Azure PowerShell modules. Please ensure they are installed: Install-Module Az"
    Exit 1
}

# Connect to Azure
try {
    $SecuredPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPassword
    
    Connect-AzAccount -ServicePrincipal -Credential $ClientSecretCredential -TenantId $TenantId
    Write-Host "Successfully connected to Azure." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Azure: $($_.Exception.Message)"
    Exit 1
}

$results = @()

try {
    # Get all Azure locations/regions
    Write-Host "Retrieving all Azure regions..." -ForegroundColor Yellow
    $locations = Get-AzLocation | Where-Object { $_.Providers -contains "Microsoft.Compute" }
    Write-Host "Found $($locations.Count) regions with compute capabilities." -ForegroundColor Green

    # Get all compute resource SKUs at once (more efficient than per-region calls)
    Write-Host "Retrieving all VM resource SKUs..." -ForegroundColor Yellow
    $allResourceSkus = Get-AzComputeResourceSku | Where-Object { $_.ResourceType -eq "virtualMachines" }
    Write-Host "Found $($allResourceSkus.Count) VM SKUs across all regions." -ForegroundColor Green

    $regionCount = 0
    foreach ($location in $locations) {
        $regionCount++
        $regionName = $location.Location
        $displayName = $location.DisplayName
        
        Write-Host "[$regionCount/$($locations.Count)] Processing region: $displayName ($regionName)" -ForegroundColor Cyan
        
        try {
            # Filter SKUs for this specific region
            $regionSkus = $allResourceSkus | Where-Object { $_.Locations -contains $regionName }
            
            if ($regionSkus.Count -eq 0) {
                Write-Warning "No VM SKUs found in region $regionName"
                continue
            }

            # Process each VM SKU in this region
            foreach ($sku in $regionSkus) {
                $vmSize = $sku.Name
                
                # Extract capabilities
                $capabilities = @{}
                if ($sku.Capabilities) {
                    foreach ($cap in $sku.Capabilities) {
                        $capabilities[$cap.Name] = $cap.Value
                    }
                }
                
                # Get VM specifications from capabilities
                $numberOfCores = if ($capabilities.ContainsKey("vCPUs")) { [int]$capabilities["vCPUs"] } else { 0 }
                $memoryInMB = if ($capabilities.ContainsKey("MemoryGB")) { [int]$capabilities["MemoryGB"] * 1024 } else { 0 }
                $maxDataDiskCount = if ($capabilities.ContainsKey("MaxDataDiskCount")) { [int]$capabilities["MaxDataDiskCount"] } else { 0 }
                $osDiskSizeInMB = if ($capabilities.ContainsKey("OSVhdSizeMB")) { [int]$capabilities["OSVhdSizeMB"] } else { 0 }
                $resourceDiskSizeInMB = if ($capabilities.ContainsKey("ResourceDiskSizeInMB")) { [int]$capabilities["ResourceDiskSizeInMB"] } else { 0 }
                
                # Check for availability zones
                $zones = @()
                if ($sku.LocationInfo) {
                    foreach ($locationInfo in $sku.LocationInfo) {
                        if ($locationInfo.Location -eq $regionName -and $locationInfo.Zones) {
                            $zones = $locationInfo.Zones
                            break
                        }
                    }
                }
                
                # Check for restrictions (important for zone availability)
                $hasRestrictions = $false
                if ($sku.Restrictions) {
                    $hasRestrictions = $true
                }
                
                if ($zones.Count -eq 0) {
                    # No zones available, add single entry
                    $results += [PSCustomObject]@{
                        Region = $regionName
                        RegionDisplayName = $displayName
                        AvailabilityZone = "N/A"
                        VMSize = $vmSize
                        NumberOfCores = $numberOfCores
                        MemoryInMB = $memoryInMB
                        MaxDataDiskCount = $maxDataDiskCount
                        OSDiskSizeInMB = $osDiskSizeInMB
                        ResourceDiskSizeInMB = $resourceDiskSizeInMB
                        HasRestrictions = $hasRestrictions
                        SKUTier = $sku.Tier
                        SKUFamily = $sku.Family
                    }
                }
                else {
                    # Add entry for each availability zone
                    $zone = $Zones -join ", "
                        $results += [PSCustomObject]@{
                            Region = $regionName
                            RegionDisplayName = $displayName
                            AvailabilityZone = $zone
                            VMSize = $vmSize
                            NumberOfCores = $numberOfCores
                            MemoryInMB = $memoryInMB
                            MaxDataDiskCount = $maxDataDiskCount
                            OSDiskSizeInMB = $osDiskSizeInMB
                            ResourceDiskSizeInMB = $resourceDiskSizeInMB
                            HasRestrictions = $hasRestrictions
                            SKUTier = $sku.Tier
                            SKUFamily = $sku.Family
                        }
                    
                }
            }
            
            Write-Host "  Found $($regionSkus.Count) VM SKUs" -ForegroundColor Gray
        }
        catch {
            Write-Warning "Error processing region $regionName : $($_.Exception.Message)"
        }
    }

    Write-Host "`nData collection completed. Total entries: $($results.Count)" -ForegroundColor Green

    # Display summary statistics
    $uniqueRegions = ($results | Select-Object -ExpandProperty Region -Unique).Count
    $uniqueSizes = ($results | Select-Object -ExpandProperty VMSize -Unique).Count
    $regionsWithZones = ($results | Where-Object { $_.AvailabilityZone -ne "N/A" } | Select-Object -ExpandProperty Region -Unique).Count
    $totalZones = ($results | Where-Object { $_.AvailabilityZone -ne "N/A" } | Select-Object -ExpandProperty AvailabilityZone -Unique).Count
    $restrictedSkus = ($results | Where-Object { $_.HasRestrictions -eq $true }).Count

    Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
    Write-Host "Unique Regions: $uniqueRegions" -ForegroundColor White
    Write-Host "Unique VM Sizes: $uniqueSizes" -ForegroundColor White
    Write-Host "Regions with Availability Zones: $regionsWithZones" -ForegroundColor White
    Write-Host "Total Availability Zones: $totalZones" -ForegroundColor White
    Write-Host "SKUs with Restrictions: $restrictedSkus" -ForegroundColor White
    Write-Host "Total Records: $($results.Count)" -ForegroundColor White

    # Output results
    if ($OutputPath) {
        try {
            $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to export to file: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "`nDisplaying first 10 results (use -OutputPath to save all results):" -ForegroundColor Yellow
        $results | Select-Object Region, AvailabilityZone, VMSize, NumberOfCores, MemoryInMB, SKUFamily -First 10 | Format-Table -AutoSize
        
        if ($results.Count -gt 10) {
            Write-Host "... and $($results.Count - 10) more entries" -ForegroundColor Gray
        }
    }

    # Show sample of regions with their zone availability
    Write-Host "`n=== SAMPLE ZONE AVAILABILITY ===" -ForegroundColor Magenta
    $zonesSample = $results | Where-Object { $_.AvailabilityZone -ne "N/A" } | 
                   Group-Object Region | Select-Object -First 5 | 
                   ForEach-Object { 
                       $zones = ($_.Group | Select-Object -ExpandProperty AvailabilityZone -Unique | Sort-Object) -join ", "
                       "$($_.Name): Zones $zones"
                   }
    
    if ($zonesSample.Count -gt 0) {
        $zonesSample | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
    }
    else {
        Write-Host "  No availability zones found in processed regions." -ForegroundColor Gray
    }

    # Show breakdown by VM size families
    Write-Host "`n=== VM SIZE FAMILIES (Top 10) ===" -ForegroundColor Magenta
    $familyBreakdown = $results | Where-Object { -not [string]::IsNullOrEmpty($_.SKUFamily) } |
                      Group-Object SKUFamily | Sort-Object Count -Descending | Select-Object -First 10

    foreach ($family in $familyBreakdown) {
        $uniqueSizesInFamily = ($family.Group | Select-Object -ExpandProperty VMSize -Unique).Count
        Write-Host "  $($family.Name): $uniqueSizesInFamily unique sizes ($($family.Count) total entries)" -ForegroundColor White
    }

    # Show regions with most VM size options
    Write-Host "`n=== TOP REGIONS BY VM SIZE VARIETY ===" -ForegroundColor Magenta
    $regionVariety = $results | Group-Object Region | 
                    ForEach-Object { 
                        [PSCustomObject]@{
                            Region = $_.Name
                            DisplayName = ($_.Group | Select-Object -First 1).RegionDisplayName
                            UniqueSizes = ($_.Group | Select-Object -ExpandProperty VMSize -Unique).Count
                            TotalEntries = $_.Count
                        }
                    } | Sort-Object UniqueSizes -Descending | Select-Object -First 5

    foreach ($region in $regionVariety) {
        Write-Host "  $($region.DisplayName): $($region.UniqueSizes) unique sizes" -ForegroundColor White
    }

}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Exit 1
}
finally {
    # Disconnect from Azure
    try {
        Disconnect-AzAccount | Out-Null
        Write-Host "`nDisconnected from Azure." -ForegroundColor Gray
    }
    catch {
        # Ignore disconnect errors
    }
}

Write-Host "`nScript completed successfully!" -ForegroundColor Green