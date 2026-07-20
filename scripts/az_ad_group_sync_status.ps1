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

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..."
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential

if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-Graph -Scopes `"Application.Read.All`", `"Group.Read.All`"" 
}

# Fetch Azure AD groups
Write-Host "Fetching Azure Entra ID groups..."
$azureGroups = Get-MgGroup -All -Property "DisplayName,Id"

# Fetch on-prem AD groups with extensionAttribute15
Write-Host "Fetching Active Directory groups..."
Import-Module ActiveDirectory
$adGroups = Get-ADGroup -Filter * -Properties extensionAttribute15 | Where-Object { $_.extensionAttribute15 -and $_.extensionAttribute15 -ne '' }

# Build lookup hashtables
$adGroupsByExt15 = @{}
foreach ($group in $adGroups) {
    if (-not $adGroupsByExt15.ContainsKey($group.extensionAttribute15)) {
        $adGroupsByExt15[$group.extensionAttribute15] = $group
    }
}

$azureGroupsById = @{}
foreach ($group in $azureGroups) {
    if ($group.Id -and $group.Id -ne '') {
        if (-not $azureGroupsById.ContainsKey($group.Id)) {
            $azureGroupsById[$group.Id] = $group
        }
    }
}

# Create normal result array
$syncResults = @()

# Compare Azure -> AD
foreach ($group in $azureGroups) {
    if ($null -ne $group.Id -and $group.Id -ne '' -and $adGroupsByExt15.ContainsKey($group.Id)) {
        $adGroup = $adGroupsByExt15[$group.Id]
        $syncResults += [PSCustomObject]@{
            SyncStatus        = "SyncReady"
            AzureGroupName    = $group.DisplayName
            AzureGroupId      = $group.Id
            ADGroupName       = $adGroup.Name
            ADGroupExtAttr15  = $adGroup.extensionAttribute15
        }
    }
    else {
        $syncResults += [PSCustomObject]@{
            SyncStatus        = "Azure only"
            AzureGroupName    = $group.DisplayName
            AzureGroupId      = $group.Id
            ADGroupName       = ""
            ADGroupExtAttr15  = ""
        }
    }
}

# Compare AD -> Azure
foreach ($group in $adGroups) {
    $extAttr = $group.extensionAttribute15
    if ($extAttr -and $extAttr -ne '' -and -not $azureGroupsById.ContainsKey($extAttr)) {
        $syncResults += [PSCustomObject]@{
            SyncStatus        = "AD only"
            AzureGroupName    = ""
            AzureGroupId      = ""
            ADGroupName       = $group.Name
            ADGroupExtAttr15  = $extAttr
        }
    }
}

# Export results to CSV if OutputPath specified
if ($OutputPath) {
    Write-Host "Exporting results to $OutputPath..."
    $syncResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
}
Disconnect-MgGraph
Write-Host "`nComparison complete. Total records: $($syncResults.Count)"

