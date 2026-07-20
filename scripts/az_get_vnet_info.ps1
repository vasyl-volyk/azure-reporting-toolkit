param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = ".\VNetReport.csv",
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

# ─────────────────────────────────────────────────────────────────────────────
# Function: recursively collect all enabled subscriptions from a management group
# ─────────────────────────────────────────────────────────────────────────────
function Get-SubscriptionsFromManagementGroup {
    param (
        [string]$ManagementGroupId
    )

    $allSubscriptions = New-Object System.Collections.ArrayList
    $processedGroups  = @{}

    function Get-NestedSubscriptions {
        param ([string]$MgId)

        if ($processedGroups.ContainsKey($MgId)) { return }
        $processedGroups[$MgId] = $true
        Write-Host "  Scanning Management Group: $MgId" -ForegroundColor Gray

        try {
            $mg = Get-AzManagementGroup -GroupId $MgId -Expand -Recurse -ErrorAction Stop

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

# ─────────────────────────────────────────────────────────────────────────────
# Function: collect all VNet + subnet details for a single subscription
# ─────────────────────────────────────────────────────────────────────────────
function Get-VNetDetails {
    param (
        [string]$SubscriptionId,
        [string]$SubscriptionName
    )

    $results = New-Object System.Collections.ArrayList

    try {
        Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Could not switch context to subscription $SubscriptionName ($SubscriptionId): $_"
        return $results
    }

    try {
        $vnets = Get-AzVirtualNetwork -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve VNets in subscription $SubscriptionName : $_"
        return $results
    }

    # ── Cache resource group tags for this subscription (keyed by RG name) ──
    $rgTagCache = @{}
    try {
        Get-AzResourceGroup -ErrorAction Stop | ForEach-Object {
            $rgTagCache[$_.ResourceGroupName] = $_.Tags
        }
    }
    catch {
        Write-Warning "Could not retrieve resource groups in subscription $SubscriptionName : $_"
    }

    # Helper: case-insensitive tag lookup
    function Get-TagValue {
        param (
            [hashtable]$Tags,
            [string]$KeyName
        )
        if (-not $Tags) { return "" }
        $match = $Tags.Keys | Where-Object { $_ -ieq $KeyName } | Select-Object -First 1
        if ($match) { return $Tags[$match] }
        return ""
    }

    foreach ($vnet in $vnets) {

        # ── VNet-level properties ──────────────────────────────────────────────
        $resourceGroup   = $vnet.ResourceGroupName
        $rgCostCenter    = Get-TagValue -Tags $rgTagCache[$resourceGroup] -KeyName "Cost Center"
        $vnetName        = $vnet.Name
        $location        = $vnet.Location
        $vnetId          = $vnet.Id
        $provisionState  = $vnet.ProvisioningState
        $enableDdos      = $vnet.EnableDdosProtection
        $enableVmProt    = $vnet.EnableVmProtection
        $ddosPlanId      = if ($vnet.DdosProtectionPlan) { $vnet.DdosProtectionPlan.Id } else { "" }
        $dnsServers      = ($vnet.DhcpOptions.DnsServers -join ";")
        $addressSpaces   = ($vnet.AddressSpace.AddressPrefixes -join ";")
        $tags            = if ($vnet.Tag) {
                               ($vnet.Tag.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ";"
                           } else { "" }

        # Peerings
        $peeringNames    = ($vnet.VirtualNetworkPeerings | ForEach-Object { $_.Name }) -join ";"
        $peeringStates   = ($vnet.VirtualNetworkPeerings | ForEach-Object { $_.PeeringState }) -join ";"
        $peeringRemoteVnets = ($vnet.VirtualNetworkPeerings | ForEach-Object {
            ($_.RemoteVirtualNetwork.Id -split "/")[-1]
        }) -join ";"

        # ── Subnet-level properties ────────────────────────────────────────────
        if ($vnet.Subnets -and $vnet.Subnets.Count -gt 0) {
            foreach ($subnet in $vnet.Subnets) {

                $subnetName          = $subnet.Name
                $subnetPrefix        = ($subnet.AddressPrefix -join ";")   # supports multiple prefixes
                $subnetProvState     = $subnet.ProvisioningState
                $nsgName             = if ($subnet.NetworkSecurityGroup) {
                                           ($subnet.NetworkSecurityGroup.Id -split "/")[-1]
                                       } else { "" }
                $nsgId               = if ($subnet.NetworkSecurityGroup) { $subnet.NetworkSecurityGroup.Id } else { "" }
                $routeTableName      = if ($subnet.RouteTable) {
                                           ($subnet.RouteTable.Id -split "/")[-1]
                                       } else { "" }
                $routeTableId        = if ($subnet.RouteTable) { $subnet.RouteTable.Id } else { "" }
                $serviceEndpoints    = ($subnet.ServiceEndpoints | ForEach-Object { $_.Service }) -join ";"
                $delegations         = ($subnet.Delegations | ForEach-Object { $_.ServiceName }) -join ";"
                $privateEndptNetwork = $subnet.PrivateEndpointNetworkPolicies
                $privateLinkService  = $subnet.PrivateLinkServiceNetworkPolicies
                $natGateway          = if ($subnet.NatGateway) { ($subnet.NatGateway.Id -split "/")[-1] } else { "" }

                # Connected devices / IP configs attached to this subnet
                $connectedDeviceCount = if ($subnet.IpConfigurations) { $subnet.IpConfigurations.Count } else { 0 }

                [void]$results.Add([PSCustomObject]@{
                    # Subscription & Resource context
                    SubscriptionId               = $SubscriptionId
                    SubscriptionName             = $SubscriptionName
                    ResourceGroup                = $resourceGroup
                    RGCostCenter                 = $rgCostCenter

                    # VNet details
                    VNetName                     = $vnetName
                    VNetId                       = $vnetId
                    Location                     = $location
                    AddressSpaces                = $addressSpaces
                    ProvisioningState            = $provisionState
                    DnsServers                   = $dnsServers
                    EnableDdosProtection         = $enableDdos
                    DdosProtectionPlanId         = $ddosPlanId
                    EnableVmProtection           = $enableVmProt
                    VNetTags                     = $tags
                    PeeringNames                 = $peeringNames
                    PeeringStates                = $peeringStates
                    PeeringRemoteVNets           = $peeringRemoteVnets

                    # Subnet details
                    SubnetName                   = $subnetName
                    SubnetAddressPrefix          = $subnetPrefix
                    SubnetProvisioningState      = $subnetProvState
                    NetworkSecurityGroupName     = $nsgName
                    NetworkSecurityGroupId       = $nsgId
                    RouteTableName               = $routeTableName
                    RouteTableId                 = $routeTableId
                    ServiceEndpoints             = $serviceEndpoints
                    Delegations                  = $delegations
                    NatGateway                   = $natGateway
                    PrivateEndpointNetworkPolicy = $privateEndptNetwork
                    PrivateLinkServicePolicy     = $privateLinkService
                    ConnectedIPConfigCount       = $connectedDeviceCount
                })
            }
        }
        else {
            # VNet with no subnets — still record the VNet row
            [void]$results.Add([PSCustomObject]@{
                SubscriptionId               = $SubscriptionId
                SubscriptionName             = $SubscriptionName
                ResourceGroup                = $resourceGroup
                RGCostCenter                 = $rgCostCenter
                VNetName                     = $vnetName
                VNetId                       = $vnetId
                Location                     = $location
                AddressSpaces                = $addressSpaces
                ProvisioningState            = $provisionState
                DnsServers                   = $dnsServers
                EnableDdosProtection         = $enableDdos
                DdosProtectionPlanId         = $ddosPlanId
                EnableVmProtection           = $enableVmProt
                VNetTags                     = $tags
                PeeringNames                 = $peeringNames
                PeeringStates                = $peeringStates
                PeeringRemoteVNets           = $peeringRemoteVnets
                SubnetName                   = "(no subnets)"
                SubnetAddressPrefix          = ""
                SubnetProvisioningState      = ""
                NetworkSecurityGroupName     = ""
                NetworkSecurityGroupId       = ""
                RouteTableName               = ""
                RouteTableId                 = ""
                ServiceEndpoints             = ""
                Delegations                  = ""
                NatGateway                   = ""
                PrivateEndpointNetworkPolicy = ""
                PrivateLinkServicePolicy     = ""
                ConnectedIPConfigCount       = 0
            })
        }
    }

    return $results
}

# ─────────────────────────────────────────────────────────────────────────────
# Main execution
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`nRetrieving subscriptions from Management Group: $ManagementGroupId" -ForegroundColor Cyan
$subscriptions = Get-SubscriptionsFromManagementGroup -ManagementGroupId $ManagementGroupId

if ($subscriptions.Count -eq 0) {
    Write-Warning "No enabled subscriptions found in management group $ManagementGroupId"
    Exit 0
}

Write-Host "`nFound $($subscriptions.Count) enabled subscription(s). Starting VNet collection...`n" -ForegroundColor Cyan

$allVNetRecords = New-Object System.Collections.ArrayList
$subIndex = 0

foreach ($sub in $subscriptions) {
    $subIndex++
    Write-Host "[$subIndex/$($subscriptions.Count)] Processing subscription: $($sub.Name) ($($sub.Id))" -ForegroundColor Yellow

    $records = Get-VNetDetails -SubscriptionId $sub.Id -SubscriptionName $sub.Name

    if ($records.Count -gt 0) {
        foreach ($r in $records) { [void]$allVNetRecords.Add($r) }
        Write-Host "  -> Collected $($records.Count) subnet record(s)" -ForegroundColor Green
    }
    else {
        Write-Host "  -> No VNets found" -ForegroundColor DarkGray
    }
}

$allVNetRecords = $allVNetRecords | Select-Object -Property SubscriptionName, ResourceGroup, RGCostCenter, VNetName, Location, AddressSpaces, PeeringNames, PeeringStates, PeeringRemoteVNets, SubnetName, SubnetAddressPrefix, NetworkSecurityGroupName, RouteTableName, ServiceEndpoints, Delegations, ConnectedIPConfigCount, NatGateway

# ─────────────────────────────────────────────────────────────────────────────
# Export results
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`nTotal records collected: $($allVNetRecords.Count)" -ForegroundColor Cyan

if ($allVNetRecords.Count -gt 0) {
    try {
        $allVNetRecords | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export CSV: $_"
        Exit 1
    }
}
else {
    Write-Warning "No VNet data collected. CSV not created."
}

Write-Host "`nDone." -ForegroundColor Cyan