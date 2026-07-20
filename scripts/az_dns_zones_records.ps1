param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$SubscriptionId = $env:AZURE_SUBSCRIPTION_ID,
    [string]$OutputPath

)


# Login using service principal
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$azCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $secureSecret

Disconnect-AzAccount -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
Connect-AzAccount -ServicePrincipal -Credential $azCred -Tenant $TenantId

# Get all accessible subscriptions
$subscriptions = if ($SubscriptionId) { Get-AzSubscription -SubscriptionId $SubscriptionId } else { Get-AzSubscription }

$dnsZoneRecords = @()

foreach ($sub in $subscriptions) {
    Write-Host "Processing subscription: $($sub.Name)" -ForegroundColor Cyan
    Set-AzContext -SubscriptionId $sub.Id

    try {
        $zones = Get-AzDnsZone
    } catch {
        Write-Warning "Access denied or no DNS zones in $($sub.Name)"
        continue
    }

    foreach ($zone in $zones) {
        Write-Host "  Zone: $($zone.Name)" -ForegroundColor Green

        $recordSets = Get-AzDnsRecordSet -ZoneName $zone.Name -ResourceGroupName $zone.ResourceGroupName

        foreach ($record in $recordSets) {
            $recordType = $record.RecordType
            $recordSetName = $record.Name
            $columnName = "$($recordType)_$($recordSetName)" # Create a unique column name

            if($recordType -like "CAA")
            {
                Write-Host
            }
            
            $value =  switch ($recordType) {
                    "A"     { $record.Records.Ipv4Address }
                    "AAAA"  { $record.Records.Ipv6Address }
                    "CNAME" { $record.Records.Cname }
                    "MX"    { ($record.Records | ForEach-Object {
                                    $recordObj = $_
                                    $recordObj | Get-Member -MemberType Properties | ForEach-Object {
                                    $name = $_.Name
                                    $value = $recordObj.$name
                                    "$name=$value"
                                    }
                                }
                                ) -join ", "
                            }
                    "NS"    { $record.Records -join ", " }
                    "PTR"   {($record.Records | ForEach-Object {
                                    $recordObj = $_
                                    $recordObj | Get-Member -MemberType Properties | ForEach-Object {
                                    $name = $_.Name
                                    $value = $recordObj.$name
                                    "$value"
                                    }
                                }
                                ) -join ", "
                            }
                    "SRV"   { ($record.Records | ForEach-Object {
                                    $recordObj = $_
                                    $recordObj | Get-Member -MemberType Properties | ForEach-Object {
                                    $name = $_.Name
                                    $value = $recordObj.$name
                                    "$name=$value"
                                    }
                                }
                                ) -join ", " }
                    "TXT"   { ($record.Records.value -join " ") } # TXT records can have multiple strings
                    
                    "SOA"   {($record.Records | ForEach-Object {
                                    $recordObj = $_
                                    $recordObj | Get-Member -MemberType Properties | ForEach-Object {
                                    $name = $_.Name
                                    $value = $recordObj.$name
                                    "$name=$value"
                                    }
                                }
                                ) -join ", "
                            }
                    default { ($record.Records | ForEach-Object {
                                    $recordObj = $_
                                    $recordObj | Get-Member -MemberType Properties | ForEach-Object {
                                    $name = $_.Name
                                    $value = $recordObj.$name
                                    "$name=$value"
                                    }
                                }
                                ) -join ", "} # Catch-all for other types, take first line
                }
             
            $zoneRecordData = [ordered]@{
                "DNS Zone Name" = $zone.Name
                "Record Name"   = $record.Name
                "Type"          = $recordType
                "TTL"           = $record.Ttl
                "Value"         = $value
                "Etag"          = $zone.Etag
                "ResourceGroup" = $zone.ResourceGroupName
                "Subscription"  = $sub.Name
            }
            $dnsZoneRecords += [PSCustomObject]$zoneRecordData
            #$dnsZoneRecords
        }
    }
}

Disconnect-AzAccount

# Optional output
Write-Host $dnsZoneRecords.count, " records have been processed"
# $dnsZoneRecords | Export-Csv -Path "dns-zone-records.csv" -NoTypeInformation -Encoding UTF8

$dnsZoneRecords | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
