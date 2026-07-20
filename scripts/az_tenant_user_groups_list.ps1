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

# Connect to Graph
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential


if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-Graph -Scopes `"Application.Read.All`", `"Group.Read.All`", `"Policy.Read.All`", `"RoleManagement.Read.Directory`", `"User.Read.All`""
}


# Get all users
$allUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Id,UserType,AccountEnabled,CreatedDateTime"

# Filter out guests
$usersToProcess = $allUsers | Where-Object { $_.UserType -ne "Guest" }

$results = @()

foreach ($user in $usersToProcess) {
    Write-Host "Processing $($user.DisplayName) <$($user.UserPrincipalName)>..."

    try {
        $userGroups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
        $groupList = ($userGroups.AdditionalProperties.displayName | Sort-Object) -join '<br>'

        $results += [PSCustomObject]@{
            DisplayName = $user.DisplayName
            Email       = $user.UserPrincipalName
            Groups      = $groupList
        }
    }
    catch {
        Write-Warning "Failed to get groups for user $($user.UserPrincipalName): $_"
    }
}

if ($OutputPath) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to $OutputPath"
}
else {
    $results | Format-Table -AutoSize
}

Disconnect-MgGraph
