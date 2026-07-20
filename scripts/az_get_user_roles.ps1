param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = "EntraID_RoleAssignments_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)


# Validate input
If ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

# Connect to Graph
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword

Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-MgGraph' with appropriate scopes: RoleManagement.Read.Directory, User.Read.All, Directory.Read.All"
}

Write-Host "Successfully connected to tenant: $TenantId" -ForegroundColor Green

# Initialize results array
$results = @()

# Create a hashtable to cache role definitions
$roleDefinitions = @{}

# Function to get principal details
function Get-PrincipalDetails {
    param (
        [string]$PrincipalId
    )
    
    $principalInfo = @{
        DisplayName = "N/A"
        PrincipalName = "N/A"
        Email = "N/A"
        PrincipalId = $PrincipalId
        AccountEnabled = "Unknown"
        ObjectType = "Unknown"
    }
    
    try {
        # Try to get as user first
        try {
            $user = Get-MgUser -UserId $PrincipalId -Property "Id,DisplayName,UserPrincipalName,Mail,AccountEnabled" -ErrorAction Stop
            if ($user) {
                Write-Host "  User found: $($user.DisplayName) / $($user.UserPrincipalName)" -ForegroundColor DarkGray
                $principalInfo.DisplayName = if ($user.DisplayName) { $user.DisplayName } else { "N/A" }
                $principalInfo.PrincipalName = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "N/A" }
                $principalInfo.Email = if ($user.Mail) { $user.Mail } else { "N/A" }
                $principalInfo.PrincipalId = if ($user.Id) { $user.Id } else { $PrincipalId }
                $principalInfo.AccountEnabled = if ($null -ne $user.AccountEnabled) { $user.AccountEnabled.ToString() } else { "Unknown" }
                $principalInfo.ObjectType = "User"
                return $principalInfo
            }
        } catch {}
        
        # Try service principal
        try {
            $sp = Get-MgServicePrincipal -ServicePrincipalId $PrincipalId -ErrorAction Stop
            if ($sp) {
                $principalInfo.DisplayName = if ($sp.DisplayName) { $sp.DisplayName } else { "N/A" }
                $principalInfo.PrincipalName = if ($sp.AppId) { $sp.AppId } else { "N/A" }
                $principalInfo.Email = "N/A"
                $principalInfo.PrincipalId = if ($sp.Id) { $sp.Id } else { $PrincipalId }
                $principalInfo.AccountEnabled = if ($null -ne $sp.AccountEnabled) { $sp.AccountEnabled.ToString() } else { "Unknown" }
                $principalInfo.ObjectType = "Service Principal"
                return $principalInfo
            }
        } catch {}
        
        # Try group
        try {
            $grp = Get-MgGroup -GroupId $PrincipalId -ErrorAction Stop
            if ($grp) {
                $principalInfo.DisplayName = if ($grp.DisplayName) { $grp.DisplayName } else { "N/A" }
                $principalInfo.PrincipalName = "N/A"
                $principalInfo.Email = if ($grp.Mail) { $grp.Mail } else { "N/A" }
                $principalInfo.PrincipalId = if ($grp.Id) { $grp.Id } else { $PrincipalId }
                $principalInfo.AccountEnabled = "N/A"
                $principalInfo.ObjectType = "Group"
                return $principalInfo
            }
        } catch {}
        
        # Unknown type
        $principalInfo.DisplayName = "Unknown"
        $principalInfo.PrincipalName = $PrincipalId
        return $principalInfo
        
    } catch {
        Write-Warning "Error getting principal details for ${PrincipalId}: $($_.Exception.Message)"
        return $principalInfo
    }
}

# Function to get or cache role definition
function Get-RoleDefinitionCached {
    param (
        [string]$RoleDefinitionId
    )
    
    if (-not $roleDefinitions.ContainsKey($RoleDefinitionId)) {
        try {
            $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $RoleDefinitionId
            $roleDefinitions[$RoleDefinitionId] = $roleDefinition
        } catch {
            Write-Warning "Error getting role definition for ${RoleDefinitionId}: $($_.Exception.Message)"
            return $null
        }
    }
    
    return $roleDefinitions[$RoleDefinitionId]
}

Write-Host "`nRetrieving ACTIVE role assignments..." -ForegroundColor Cyan

# Get all active directory role assignments
$activeAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All

Write-Host "Found $($activeAssignments.Count) active role assignments" -ForegroundColor Yellow

Write-Host "`nRetrieving ELIGIBLE role assignments..." -ForegroundColor Cyan

# Get all eligible role assignments (PIM)
try {
    $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ErrorAction Stop
    Write-Host "Found $($eligibleAssignments.Count) eligible role assignments" -ForegroundColor Yellow
} catch {
    if ($_.Exception.Message -like "*PermissionScopeNotGranted*" -or $_.Exception.Message -like "*Authorization failed*") {
        Write-Host "⚠️  Unable to retrieve eligible assignments - Missing permissions" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "To include ELIGIBLE role assignments, grant one of these permissions to your service principal:" -ForegroundColor Cyan
        Write-Host "  • RoleEligibilitySchedule.Read.Directory (Recommended - Read Only)" -ForegroundColor White
        Write-Host "  • RoleManagement.Read.Directory" -ForegroundColor White
        Write-Host "  • RoleManagement.Read.All" -ForegroundColor White
        Write-Host ""
        Write-Host "How to add the permission:" -ForegroundColor Cyan
        Write-Host "  1. Go to Azure Portal → Microsoft Entra ID → App registrations" -ForegroundColor White
        Write-Host "  2. Find your app registration (Client ID: $ClientId)" -ForegroundColor White
        Write-Host "  3. Click 'API permissions' → 'Add a permission'" -ForegroundColor White
        Write-Host "  4. Select 'Microsoft Graph' → 'Application permissions'" -ForegroundColor White
        Write-Host "  5. Search for and add 'RoleEligibilitySchedule.Read.Directory'" -ForegroundColor White
        Write-Host "  6. Click 'Grant admin consent' at the top" -ForegroundColor White
        Write-Host ""
        Write-Host "Continuing with ACTIVE assignments only..." -ForegroundColor Yellow
        Write-Host ""
    } else {
        Write-Warning "Unexpected error retrieving eligible assignments: $($_.Exception.Message)"
    }
    $eligibleAssignments = @()
}

# Process Active Assignments
Write-Host "`nProcessing ACTIVE role assignments..." -ForegroundColor Cyan
$counter = 0

foreach ($assignment in $activeAssignments) {
    $counter++
    if ($counter % 50 -eq 0) {
        Write-Host "Processed $counter of $($activeAssignments.Count) active assignments..." -ForegroundColor Gray
    }
    
    try {
        # Get principal details
        $principalInfo = Get-PrincipalDetails -PrincipalId $assignment.PrincipalId
        
        # Get role definition
        $role = Get-RoleDefinitionCached -RoleDefinitionId $assignment.RoleDefinitionId
        
        # Add to results
        $results += [PSCustomObject]@{
            DisplayName         = $principalInfo.DisplayName
            PrincipalName       = $principalInfo.PrincipalName
            AccountEnabled      = $principalInfo.AccountEnabled
            ObjectType          = $principalInfo.ObjectType
            RoleName            = if ($role.DisplayName) { $role.DisplayName } else { "N/A" }
            RoleDescription     = if ($role.Description) { $role.Description } else { "N/A" }
            AssignmentType      = "Active"
            StartDateTime       = "N/A"
            EndDateTime         = "N/A"
        }
    }
    catch {
        Write-Warning "Error processing active assignment $($assignment.Id) for principal $($assignment.PrincipalId): $_"
        continue
    }
}

# Process Eligible Assignments
Write-Host "`nProcessing ELIGIBLE role assignments..." -ForegroundColor Cyan
$counter = 0

foreach ($assignment in $eligibleAssignments) {
    $counter++
    if ($counter % 50 -eq 0) {
        Write-Host "Processed $counter of $($eligibleAssignments.Count) eligible assignments..." -ForegroundColor Gray
    }
    
    try {
        # Get principal details
        $principalInfo = Get-PrincipalDetails -PrincipalId $assignment.PrincipalId
        
        # Get role definition
        $role = Get-RoleDefinitionCached -RoleDefinitionId $assignment.RoleDefinitionId
        
        # Format dates
        $startDate = if ($assignment.StartDateTime) { $assignment.StartDateTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
        $endDate = if ($assignment.EndDateTime) { $assignment.EndDateTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
        
        # Add to results
        $results += [PSCustomObject]@{
            DisplayName         = $principalInfo.DisplayName
            PrincipalName       = $principalInfo.PrincipalName
            AccountEnabled      = $principalInfo.AccountEnabled
            ObjectType          = $principalInfo.ObjectType
            RoleName            = if ($role.DisplayName) { $role.DisplayName } else { "N/A" }
            AssignmentType      = "Eligible"
            StartDateTime       = $startDate
            EndDateTime         = $endDate
            RoleDescription     = if ($role.Description) { $role.Description } else { "N/A" }
        }
    }
    catch {
        Write-Warning "Error processing eligible assignment $($assignment.Id) for principal $($assignment.PrincipalId): $_"
        continue
    }
}

Write-Host "`nProcessing complete!" -ForegroundColor Green
Write-Host "Found $($results.Count) total role assignments (active + eligible)" -ForegroundColor Yellow

# Export to CSV
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Green
    
    # Display summary
    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host "Total role assignments: $($results.Count)" -ForegroundColor White
    
    # Assignment type summary
    $assignmentTypeSummary = $results | Group-Object -Property AssignmentType | Sort-Object Count -Descending
    Write-Host "`nAssignment Type Distribution:" -ForegroundColor Cyan
    foreach ($type in $assignmentTypeSummary) {
        Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
    }
    
    # Object type summary
    $typeSummary = $results | Group-Object -Property ObjectType | Sort-Object Count -Descending
    Write-Host "`nObject Type Distribution:" -ForegroundColor Cyan
    foreach ($type in $typeSummary) {
        Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
    }
    
    # Role summary
    $roleSummary = $results | Group-Object -Property RoleName | Sort-Object Count -Descending | Select-Object -First 10
    Write-Host "`nTop 10 Roles:" -ForegroundColor Cyan
    foreach ($role in $roleSummary) {
        Write-Host "  $($role.Name): $($role.Count)" -ForegroundColor White
    }
}
else {
    Write-Host "`nNo role assignments found." -ForegroundColor Yellow
}

# Disconnect
Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Cyan