param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)


# Get token
$Body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $ClientId
    client_secret = $ClientSecret
}
$TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $Body
$AccessToken = $TokenResponse.access_token
$Headers = @{ Authorization = "Bearer $AccessToken" }

# Retrieve all groups (no $select filter to ensure all properties come through)
$AllGroups = @()
$Uri = "https://graph.microsoft.com/v1.0/groups"

do {
    $Response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
    $AllGroups += $Response.value
    $Uri = $Response.'@odata.nextLink'
} while ($Uri)

# Filter dynamic groups and map to PSObject
$DynamicGroups = $AllGroups | Where-Object { $_.membershipRule -ne $null } | ForEach-Object {
    [PSCustomObject]@{
        DisplayName                    = $_.displayName
        MembershipRule                 = $_.membershipRule
        MembershipRuleProcessingState = $_.membershipRuleProcessingState
        SecurityEnabled                = $_.securityEnabled
        MailEnabled                    = $_.mailEnabled
        GroupTypes                     = ($_.groupTypes -join ", ")
        Description                    = $_.description
        Id                             = $_.id
    }
}

# Export to CSV
$DynamicGroups | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "✔ Exported $($DynamicGroups.Count) dynamic groups to $OutputPath"