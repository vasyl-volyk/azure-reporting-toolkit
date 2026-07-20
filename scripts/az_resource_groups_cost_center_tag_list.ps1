param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$TargetMGroupId,
    [string]$OutputPath

)



$scope = "https://management.azure.com/.default"  

  
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

  
# Get management groups  
function Get-ManagementGroups {
    param (
        [string]$AccessToken,
        [string]$ManagementGroupId = "", # New parameter for specific MG
        [switch]$IncludeNested = $true, # New parameter to control nested retrieval
        [string]$RGPath = "./"
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    $uriBase = "https://management.azure.com/providers/Microsoft.Management/managementGroups"
    $apiVersion = "2021-04-01"

    if ([string]::IsNullOrEmpty($ManagementGroupId)) {
        # If no specific ManagementGroupId is provided, get all top-level MGs
        $uri = "$uriBase?api-version=$apiVersion"
    } else {
        # If a specific ManagementGroupId is provided
        $uri = "$uriBase/$ManagementGroupId"+"?api-version=$apiVersion"
        if ($IncludeNested) {
            $uri += "&`$expand=children"
        }
    }

    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    }
    catch {
        Write-Error "Failed to retrieve management groups. Error: $($_.Exception.Message)"
        return $null
    }

    Write-Host "response name ",$response.name

    $managementGroups = @()

    if ( -not $response.properties.children) {
        
                $managementGroups += [PSCustomObject]@{
                                "Path" = $RGPath
                                "id" =  $response.id
                                "type" = $response.type
                                "name" = $response.name
                                "displayName" = $response.properties.displayName
                            }
    } 
    elseif($IncludeNested -and $response.properties.children) {        
        $managementGroups += [PSCustomObject]@{
                                "Path" = $RGPath
                                "id" =  $response.id
                                "type" = $response.type
                                "name" = $response.name
                                "displayName" = $response.properties.displayName
                            }
        foreach ($child in $response.properties.children) {
            if ($child.id -match "managementGroups"){
                $ChildID = $child.id -replace "/providers/Microsoft.Management/managementGroups/", ""
                $p = $RGPath+"/"+$response.name
                Write-Host $p 
                $managementGroups += Get-ManagementGroups -AccessToken $AccessToken -ManagementGroupId $ChildID -IncludeNested $true -RGPath $p
            }
        
        }
    }
    return $managementGroups
}

  
# Get subscriptions under a management group  
function Get-Subscriptions {  
    param (  
        [string]$ManagementGroupId,  
        [string]$AccessToken  
    )  
  
    $headers = @{ Authorization = "Bearer $AccessToken" }  
    $ManagementGroupId = $ManagementGroupId -replace '^/providers/Microsoft\.Management/managementGroups/', ''
    $uri = "https://management.azure.com/providers/Microsoft.Management/managementGroups/$ManagementGroupId/subscriptions?api-version=2020-05-01"  
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get  
    return $response.value  
}  
  
# Get resource groups in a subscription  
function Get-ResourceGroups {  
    param (  
        [string]$SubscriptionId,  
        [string]$AccessToken  
    )  
  
    $headers = @{ Authorization = "Bearer $AccessToken" }  
    $uri = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups?api-version=2022-12-01"  
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get  
    return $response.value  
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

    # Якщо це hashtable
    if ($Tags -is [hashtable] -and $Tags.ContainsKey($TagName)) {
        return $Tags[$TagName]
    }

    # Якщо це об'єкт (PSCustomObject), перетворимо в хеш
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
 
  
# Get the access token  
$accessToken = Get-AzureToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret -Scope $scope  

$managementGroups = Get-ManagementGroups -AccessToken $AccessToken -ManagementGroupId $TargetMGroupID -IncludeNested $true -RGpath "Root"
  
$ListRGandTags = @()
foreach ($mg in $managementGroups) {  
        $mgId = $mg.id  
        $mgName = $mg.name
        $RGpath = $mg.Path
  
        Write-Host "Processing Management Group: $mgName"
  
        $subscriptions = Get-Subscriptions -ManagementGroupId $mgId -AccessToken $AccessToken  
  
        foreach ($sub in $subscriptions) {  
            $subId = $sub.name  
            $subDisplayName = $sub.properties.displayName
            Write-Host "`tSubscription: $subDisplayName"

            $resourceGroups = Get-ResourceGroups -SubscriptionId $subId -AccessToken $AccessToken  
  
            foreach ($rg in $resourceGroups) {  
                $costCenter = Get-TagValue -Tags $rg.tags -TagName "Cost center"  
  
                $ListRGandTags += [PSCustomObject]@{  
                    Path               = $RGpath
                    ManagementGroup    = $mgName  
                    SubscriptionName   = $subDisplayName  
                    SubscriptionId     = $subId  
                    ResourceGroupName  = $rg.name  
                    CostCenter         = $costCenter  
                }  
            }  
        }  
}

$ListRGandTags | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
