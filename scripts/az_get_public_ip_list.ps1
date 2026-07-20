param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = ".\PublicIPReport.csv",
    [string]$ManagementGroupId = "Internal"  # Management Group ID to scan (includes all nested)
)


# Authenticate to Azure using Service Principal
try {
    $SecuredPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPassword
    
    Write-Host "Connecting to Azure..." -ForegroundColor Yellow
    Connect-AzAccount -ServicePrincipal -Credential $Credential -Tenant $TenantId -ErrorAction Stop | Out-Null
    Write-Host "Successfully authenticated to Azure" -ForegroundColor Green
}
catch {
    Write-Error "Failed to authenticate to Azure: $_"
    Exit 1
}

# Function to recursively get all subscriptions from a management group
function Get-SubscriptionsFromManagementGroup {
    param (
        [string]$ManagementGroupId
    )
    
    $allSubscriptions = New-Object System.Collections.ArrayList
    $processedGroups = @{}
    
    function Get-NestedSubscriptions {
        param ([string]$MgId)
        
        if ($processedGroups.ContainsKey($MgId)) {
            return
        }
        
        $processedGroups[$MgId] = $true
        Write-Host "  Scanning Management Group: $MgId" -ForegroundColor Gray
        
        try {
            $mg = Get-AzManagementGroup -GroupId $MgId -Expand -Recurse -ErrorAction Stop
            
            # Get direct subscriptions
            if ($mg.Children) {
                foreach ($child in $mg.Children) {
                    if ($child.Type -eq "/subscriptions") {
                        $sub = Get-AzSubscription -SubscriptionId $child.Name -ErrorAction SilentlyContinue
                        if ($sub -and $sub.State -eq "Enabled") {
                            [void]$allSubscriptions.Add($sub)
                            Write-Host "    Found subscription: $($sub.Name)" -ForegroundColor DarkGray
                        }
                    }
                    elseif ($child.Type -match "managementGroups") {
                        # Recursively process nested management groups
                        Get-NestedSubscriptions -MgId $child.Name
                    }
                }
            }
        }
        catch {
            Write-Warning "Error accessing management group $MgId : $_"
        }
    }
    
    Get-NestedSubscriptions -MgId $ManagementGroupId
    return $allSubscriptions.ToArray()
}

# Main script execution
Write-Host "`nRetrieving subscriptions from Management Group: $ManagementGroupId" -ForegroundColor Cyan
$subscriptions = Get-SubscriptionsFromManagementGroup -ManagementGroupId $ManagementGroupId

if ($subscriptions.Count -eq 0) {
    Write-Warning "No enabled subscriptions found in management group $ManagementGroupId"
    Exit 0
}

Write-Host "`nFound $($subscriptions.Count) enabled subscription(s)" -ForegroundColor Green
Write-Host "`nScanning for Public IP addresses..." -ForegroundColor Cyan

$publicIPReport = New-Object System.Collections.ArrayList

foreach ($subscription in $subscriptions) {
    Write-Host "`nProcessing subscription: $($subscription.Name) ($($subscription.Id))" -ForegroundColor Yellow
    
    try {
        # Set the context to the current subscription
        Set-AzContext -SubscriptionId $subscription.Id -ErrorAction Stop | Out-Null
        
        # Get all public IP addresses in the subscription
        $publicIPs = Get-AzPublicIpAddress -ErrorAction SilentlyContinue
        
        if ($publicIPs) {
            Write-Host "  Found $($publicIPs.Count) Public IP address(es)" -ForegroundColor Green
            
            foreach ($pip in $publicIPs) {
                # Get attached resource information
                $attachedTo = "Not Assigned"
                $attachedResourceName = ""
                $attachedResourceType = ""
                
                if ($pip.IpConfiguration) {
                    $attachedTo = "Assigned"
                    if ($pip.IpConfiguration.Id) {
                        $parts = $pip.IpConfiguration.Id -split '/'
                        $attachedResourceType = $parts[-3]
                        $attachedResourceName = $parts[-5]
                    }
                }
                
                $ipInfo = [PSCustomObject]@{
                    SubscriptionName = $subscription.Name
                    ResourceGroup = $pip.ResourceGroupName
                    PublicIPName = $pip.Name
                    IPAddress = $pip.IpAddress
                    AllocationMethod = $pip.PublicIpAllocationMethod
                    SKU = $pip.Sku.Name
                    Version = $pip.PublicIpAddressVersion
                    Status = $attachedTo
                    DnsSettings = if ($pip.DnsSettings.Fqdn) { $pip.DnsSettings.Fqdn } else { "" }
                    AttachedResourceName = $attachedResourceName
                    AttachedResourceType = $attachedResourceType
                    Location = $pip.Location
                    SubscriptionId = $subscription.Id
                    Tags = if ($pip.Tags) { ($pip.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ' } else { "" }
                }
                
                [void]$publicIPReport.Add($ipInfo)
                Write-Host "    - $($pip.Name): $($pip.IpAddress) [$attachedTo]" -ForegroundColor Gray
            }
        }
        else {
            Write-Host "  No Public IP addresses found" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Warning "Error processing subscription $($subscription.Name): $_"
    }
}

# Export results
if ($publicIPReport.Count -gt 0) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Total Public IP Addresses Found: $($publicIPReport.Count)" -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    try {
        $publicIPReport | Export-Csv -Path $OutputPath -NoTypeInformation -Force
        Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
        
        # Display summary statistics
        $assignedCount = ($publicIPReport | Where-Object { $_.Status -eq "Assigned" }).Count
        $unassignedCount = ($publicIPReport | Where-Object { $_.Status -eq "Not Assigned" }).Count
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "  Assigned IPs: $assignedCount" -ForegroundColor Green
        Write-Host "  Unassigned IPs: $unassignedCount" -ForegroundColor Yellow
        Write-Host "  Subscriptions scanned: $($subscriptions.Count)" -ForegroundColor Gray
    }
    catch {
        Write-Error "Failed to export report: $_"
    }
}
else {
    Write-Host "`nNo Public IP addresses found in any subscription" -ForegroundColor Yellow
}

Write-Host "`nScript completed successfully" -ForegroundColor Green