param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath,
    [int]$BatchSize = 50  # Process users in batches for progress updates
)

# Validate input
If ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "ClientId, ClientSecret, and TenantId must be set."
    Exit 1
}

# Connect to Graph using Service Principal
Write-Host "Connecting to Microsoft Graph using Service Principal..." -ForegroundColor Cyan
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential
Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green

# Get all users from the "Employees - All" mail-enabled security group
Write-Host "`nRetrieving users from 'GRP - All Active Company Users' group..." -ForegroundColor Cyan

try {
    # Find the group by display name
    $groupName = "GRP - All Active Company Users"
    $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
    
    if (-not $group) {
        Write-Error "Group '$groupName' not found in Entra ID"
        Disconnect-MgGraph
        Exit 1
    }
    
    Write-Host "Found group: $($group.DisplayName) (ID: $($group.Id))" -ForegroundColor Green
    
    # Get all members of the group in one call with expanded properties
    Write-Host "Retrieving all group members..." -ForegroundColor Yellow
    $groupMembers = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
    
    Write-Host "Found $($groupMembers.Count) total members in the group" -ForegroundColor Yellow
    
    # Extract user IDs for batch retrieval
    $userIds = @()
    foreach ($member in $groupMembers) {
        if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
            $userIds += $member.Id
        }
    }
    
    Write-Host "Identified $($userIds.Count) user members" -ForegroundColor Green
    
    # Batch retrieve user details
    Write-Host "Retrieving user details in batches..." -ForegroundColor Yellow
    $allUsers = @()
    for ($i = 0; $i -lt $userIds.Count; $i += $BatchSize) {
        $batch = $userIds[$i..[Math]::Min($i + $BatchSize - 1, $userIds.Count - 1)]
        Write-Host "Fetching users $($i + 1) to $($i + $batch.Count)..." -ForegroundColor Yellow
        
        foreach ($userId in $batch) {
            $userDetails = Get-MgUser -UserId $userId -Property UserPrincipalName, DisplayName, JobTitle, UserType, Id -ErrorAction SilentlyContinue
            if ($userDetails) {
                $allUsers += $userDetails
            }
        }
    }
    
    Write-Host "Retrieved details for $($allUsers.Count) users" -ForegroundColor Green
    
} catch {
    Write-Error "Failed to retrieve group members: $_"
    Disconnect-MgGraph
    Exit 1
}

# Pre-cache all manager relationships and user details to minimize API calls
Write-Host "`nPre-caching all manager relationships..." -ForegroundColor Cyan
$managerCache = @{}
$userDetailsCache = @{}

$cacheCounter = 0
$totalUsers = $allUsers.Count

foreach ($user in $allUsers) {
    $cacheCounter++
    
    if ($cacheCounter % 50 -eq 0 -or $cacheCounter -eq $totalUsers) {
        $percentComplete = [Math]::Round(($cacheCounter / $totalUsers) * 100, 1)
        Write-Host "Cached $cacheCounter of $totalUsers users ($percentComplete%)..." -ForegroundColor Yellow
    }
    
    # Cache user details by both UPN and ID for fast lookup
    $userDetailsCache[$user.UserPrincipalName] = @{
        DisplayName = $user.DisplayName
        JobTitle = $user.JobTitle
        UPN = $user.UserPrincipalName
        Id = $user.Id
    }
    
    $userDetailsCache[$user.Id] = @{
        DisplayName = $user.DisplayName
        JobTitle = $user.JobTitle
        UPN = $user.UserPrincipalName
        Id = $user.Id
    }
    
    # Cache manager relationship
    try {
        $manager = Get-MgUserManager -UserId $user.UserPrincipalName -ErrorAction SilentlyContinue
        if ($manager) {
            $managerCache[$user.UserPrincipalName] = $manager.Id
            $managerCache[$user.Id] = $manager.Id
        }
    } catch {
        # No manager or error - skip
    }
}

Write-Host "Caching complete! Cached $($managerCache.Count / 2) manager relationships" -ForegroundColor Green

# Function to get the full manager chain using cached data
function Get-ManagerChainFromCache {
    param (
        [string]$UserIdentifier,
        [string]$OriginalUserName,
        [hashtable]$ManagerCache,
        [hashtable]$UserCache
    )
    
    $managerChain = @()
    $currentUser = $UserIdentifier
    $maxIterations = 20
    $iteration = 0
    $seenIdentifiers = @{}
    
    while ($currentUser -and $iteration -lt $maxIterations) {
        $iteration++
        
        # Check if we've already seen this user (loop detection)
        if ($seenIdentifiers.ContainsKey($currentUser)) {
            break
        }
        
        $seenIdentifiers[$currentUser] = $true
        
        # Check if manager exists in cache
        if ($ManagerCache.ContainsKey($currentUser)) {
            $managerId = $ManagerCache[$currentUser]
            
            # Get manager details from cache
            if ($UserCache.ContainsKey($managerId)) {
                $managerDetails = $UserCache[$managerId]
                
                # Check if manager has the same display name as the original user
                if ($managerDetails.DisplayName -eq $OriginalUserName) {
                    break
                }
                
                # Check if this manager's display name already exists in the chain
                if ($managerChain -contains $managerDetails.DisplayName) {
                    break
                }
                
                $managerChain += $managerDetails.DisplayName
                $currentUser = $managerId
            } else {
                break
            }
        } else {
            # No manager found
            break
        }
    }
    
    return ($managerChain -join " -> ")
}

# Process all users and build manager chains
Write-Host "`nBuilding manager chains from cached data..." -ForegroundColor Cyan

$results = @()
$processCounter = 0

foreach ($user in $allUsers) {
    $processCounter++
    
    if ($processCounter % 100 -eq 0 -or $processCounter -eq $totalUsers) {
        $percentComplete = [Math]::Round(($processCounter / $totalUsers) * 100, 1)
        Write-Host "Processing: $processCounter of $totalUsers users ($percentComplete%)..." -ForegroundColor Yellow
    }
    
    $upn = $user.UserPrincipalName
    
    try {
        # Get manager chain from cache
        $managerChain = Get-ManagerChainFromCache -UserIdentifier $user.Id -OriginalUserName $user.DisplayName -ManagerCache $managerCache -UserCache $userDetailsCache
        
        # Create result object
        $results += [PSCustomObject]@{
            'Primary UPN' = $user.UserPrincipalName
            'Full Name' = $user.DisplayName
            'Title' = $user.JobTitle
            'Manager Chain' = if ($managerChain) { $managerChain } else { "No manager assigned" }
        }
    } catch {
        Write-Warning "Failed to process $upn : $_"
        $results += [PSCustomObject]@{
            'Primary UPN' = $upn
            'Full Name' = $user.DisplayName
            'Title' = $user.JobTitle
            'Manager Chain' = "Failed to retrieve manager chain"
        }
    }
}

# Display summary
Write-Host "`n=== PROCESSING COMPLETE ===" -ForegroundColor Green
Write-Host "Total users processed: $($results.Count)" -ForegroundColor Cyan
Write-Host "Users with managers: $(($results | Where-Object { $_.'Manager Chain' -ne 'No manager assigned' -and $_.'Manager Chain' -ne 'Failed to retrieve manager chain' }).Count)" -ForegroundColor Cyan
Write-Host "Users without managers: $(($results | Where-Object { $_.'Manager Chain' -eq 'No manager assigned' }).Count)" -ForegroundColor Cyan

# Display first 10 results as preview
Write-Host "`n=== PREVIEW (First 10 Results) ===" -ForegroundColor Green
$results | Select-Object -First 10 | Format-Table -AutoSize -Wrap

# Export results to CSV
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = "UserManagerChain_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Green

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Cyan