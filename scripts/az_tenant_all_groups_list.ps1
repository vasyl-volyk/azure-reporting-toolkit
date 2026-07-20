
param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath,
    [int]$BatchSize = 20,
    [int]$MaxJobs = 10,
    [switch]$SkipMembers,
    [switch]$SkipOwners,
    [switch]$DebugMode,
    # Graph does not return owners for Exchange-managed DL / mail-enabled security groups; EXO ManagedBy fills the gap. Default: on. Use -EnrichOwnersFromExchangeOnline:$false to skip (automation with no EXO session).
    [bool]$EnrichOwnersFromExchangeOnline = $true,
    # Optional: same app registration as Graph, with Exchange app-only auth (certificate on the host).
    [string]$ExchangeCertificateThumbprint,
    # Optional: e.g. contoso.onmicrosoft.com (defaults to tenant initial verified domain from Graph).
    [string]$ExchangeOrganization,
    # When no certificate is configured, use browser/device-code sign-in for EXO (off by default so automation does not hang).
    [switch]$ExchangeOnlineAllowInteractiveLogin
)

# Microsoft Graph does not return /groups/{id}/owners for distribution lists created in Exchange,
# many mail-enabled groups managed in EAC, and some on-premises–synced groups (see List group owners docs).
# Exchange enrichment uses the same app as Graph: (1) -ExchangeCertificateThumbprint, or (2) client credentials token
# (AZURE_CLIENT_SECRET + app permission Exchange.ManageAsApp on Office 365 Exchange Online + Exchange RBAC for the app).


function Normalize-CsvField {
    param([string]$Value)

    if ($null -eq $Value) { return "" }

    # Replace CR/LF and normalize commas
    $Value -replace "`r|`n", " " -replace ",", ";"
}

function Resolve-ExchangeManagedByToOwnerObjects {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $ManagedByProperty,
        [Parameter(Mandatory)]
        [hashtable]$RecipientCache
    )

    $dns = @()
    if ($null -eq $ManagedByProperty) { return @() }
    if ($ManagedByProperty -is [string]) {
        if (-not [string]::IsNullOrWhiteSpace($ManagedByProperty)) { $dns += $ManagedByProperty.Trim() }
    } elseif ($ManagedByProperty -is [System.Collections.IEnumerable] -and -not ($ManagedByProperty -is [string])) {
        foreach ($item in $ManagedByProperty) {
            if ($null -eq $item) { continue }
            $dn = if ($item -is [string]) { $item.Trim() }
                  elseif ($null -ne $item.DistinguishedName) { [string]$item.DistinguishedName }
                  else { $item.ToString().Trim() }
            if ($dn.Length -gt 0) { $dns += $dn }
        }
    } else {
        $single = if ($null -ne $ManagedByProperty.DistinguishedName) { [string]$ManagedByProperty.DistinguishedName } else { $ManagedByProperty.ToString().Trim() }
        if ($single.Length -gt 0) { $dns += $single }
    }

    $owners = [System.Collections.Generic.List[object]]::new()
    foreach ($dn in ($dns | Select-Object -Unique)) {
        if ($RecipientCache.ContainsKey($dn)) {
            $cached = $RecipientCache[$dn]
            if ($cached) { $owners.Add($cached) }
            continue
        }
        try {
            $rec = Get-Recipient -Identity $dn -ErrorAction Stop
            $obj = [PSCustomObject]@{
                Id                 = $rec.ExternalDirectoryObjectId
                DisplayName        = $rec.DisplayName
                UserPrincipalName  = if ($rec.UserPrincipalName) { $rec.UserPrincipalName } else { $rec.PrimarySmtpAddress }
                ObjectType         = if ($null -ne $rec.RecipientTypeDetails) { $rec.RecipientTypeDetails.ToString() } else { "Recipient" }
            }
            $RecipientCache[$dn] = $obj
            $owners.Add($obj)
        }
        catch {
            $fallback = [PSCustomObject]@{
                Id                 = ""
                DisplayName        = $dn
                UserPrincipalName  = ""
                ObjectType         = "UnresolvedRecipient"
            }
            $RecipientCache[$dn] = $fallback
            $owners.Add($fallback)
        }
    }
    return ,$owners.ToArray()
}

function Get-ExchangeDistributionGroupManagedByProperty {
    <#
    Graph often returns no owners for DL / mail-enabled security groups; EXO stores managers on ManagedBy.
    Try several identities because bulk enumeration can miss objects or Id vs SMTP differs.
    #>
    param(
        [string]$GraphGroupId,
        [string]$PrimarySmtpAddress
    )
    foreach ($identity in @($GraphGroupId, $PrimarySmtpAddress)) {
        if ([string]::IsNullOrWhiteSpace($identity)) { continue }
        try {
            $dg = Get-DistributionGroup -Identity $identity -ErrorAction Stop
            if ($null -ne $dg.ManagedBy) { return $dg.ManagedBy }
        }
        catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($GraphGroupId)) {
        try {
            $dg = @(Get-DistributionGroup -Filter "ExternalDirectoryObjectId -eq '$GraphGroupId'" -ErrorAction Stop)
            if ($dg.Count -gt 0 -and $null -ne $dg[0].ManagedBy) { return $dg[0].ManagedBy }
        }
        catch { }
    }
    return $null
}

function Get-ExchangeOnlineAccessTokenClientCredentials {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )
    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $scope = [uri]::EscapeDataString("https://outlook.office365.com/.default")
    $body = "client_id=$([uri]::EscapeDataString($ClientId))&client_secret=$([uri]::EscapeDataString($ClientSecret))&scope=$scope&grant_type=client_credentials"
    $response = Invoke-RestMethod -Uri $tokenUri -Method Post -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($response.access_token)) {
        throw "Token response did not include access_token."
    }
    return [string]$response.access_token
}

# Validate input
if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion.Major
Write-Host "Running on PowerShell $psVersion" -ForegroundColor Yellow

# Connect to Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword

try {
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Exit 1
}

if (-not $(Get-MgContext)) {
    Write-Error "Authentication failed. Please verify credentials and permissions."
    Exit 1
}

# Pre-load all service principal app assignments once (AppRoleAssignedTo on each SP is expensive per-group;
# instead we query all app role assignments across all SPs and build a lookup keyed by group ID)
Write-Host "`nPre-loading application assignments for all groups..." -ForegroundColor Green
$groupAppAssignmentMap = @{}
try {
    $allServicePrincipals = Get-MgServicePrincipal -All -Property "Id,DisplayName,AppId" -ErrorAction Stop
    Write-Host "Found $($allServicePrincipals.Count) service principals. Collecting group assignments..." -ForegroundColor Yellow

    foreach ($sp in $allServicePrincipals) {
        try {
            $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue
            foreach ($assignment in $assignments) {
                # PrincipalType == "Group" means the assignment is to a group
                if ($assignment.PrincipalType -eq "Group") {
                    $gid = $assignment.PrincipalId
                    if (-not $groupAppAssignmentMap.ContainsKey($gid)) {
                        $groupAppAssignmentMap[$gid] = [System.Collections.Generic.List[string]]::new()
                    }
                    if (-not $groupAppAssignmentMap[$gid].Contains($sp.DisplayName)) {
                        $groupAppAssignmentMap[$gid].Add($sp.DisplayName)
                    }
                }
            }
        }
        catch {
            if ($DebugMode) {
                Write-Warning "Could not retrieve assignments for SP '$($sp.DisplayName)': $($_.Exception.Message)"
            }
        }
    }
    Write-Host "Application assignment map built. Groups with assignments: $($groupAppAssignmentMap.Count)" -ForegroundColor Green
}
catch {
    Write-Warning "Failed to pre-load application assignments: $($_.Exception.Message)"
}

# Get all groups with detailed properties
Write-Host "`nCollecting all groups from Entra ID..." -ForegroundColor Green
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $allGroups = Get-MgGroup -All -Property "Id,DisplayName,Description,GroupTypes,SecurityEnabled,MailEnabled,Mail,MailNickname,CreatedDateTime,Visibility,MembershipRule,MembershipRuleProcessingState,ProxyAddresses,OnPremisesSyncEnabled,OnPremisesSecurityIdentifier" -ErrorAction Stop
    Write-Host "Found $($allGroups.Count) groups." -ForegroundColor Yellow
}
catch {
    Write-Error "Failed to retrieve groups: $($_.Exception.Message)"
    Exit 1
}

if ($allGroups.Count -eq 0) {
    Write-Warning "No groups found. Exiting."
    Exit 0
}

# Initialize tracking
$script:processedGroups = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
$script:failedGroups = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
$script:jobErrors = [System.Collections.Concurrent.ConcurrentBag[string]]::new()

# For PowerShell 5, use background jobs for parallel processing
Write-Host "Using PowerShell 5 optimized batch processing with background jobs..." -ForegroundColor Green

# Split groups into batches
$groupBatches = @()
for ($i = 0; $i -lt $allGroups.Count; $i += $BatchSize) {
    $batch = $allGroups[$i..([Math]::Min($i + $BatchSize - 1, $allGroups.Count - 1))]
    $groupBatches += ,@{
        Batch = $batch
        BatchNumber = [Math]::Floor($i / $BatchSize) + 1
    }
}

Write-Host "Processing $($groupBatches.Count) batches with batch size $BatchSize (Max concurrent jobs: $MaxJobs)..." -ForegroundColor Yellow

$allResults = @()
$jobCounter = 0
$jobList = @()

foreach ($batchInfo in $groupBatches) {
    # Wait if we have too many jobs running
    while ((Get-Job -State Running).Count -ge $MaxJobs) {
        Start-Sleep -Milliseconds 500
        
        # Collect completed jobs
        $completedJobs = Get-Job -State Completed
        if ($completedJobs) {
            foreach ($job in $completedJobs) {
                try {
                    $jobResult = Receive-Job -Job $job -ErrorAction Stop
                    if ($jobResult) {
                        $allResults += $jobResult
                    }
                }
                catch {
                    Write-Warning "Error receiving job $($job.Id) results: $($_.Exception.Message)"
                }
                Remove-Job -Job $job -Force
            }
        }
        
        # Check for failed jobs
        $failedJobs = Get-Job -State Failed
        if ($failedJobs) {
            foreach ($job in $failedJobs) {
                Write-Warning "Job $($job.Id) failed"
                try {
                    Receive-Job -Job $job -ErrorAction Continue
                }
                catch {
                    Write-Warning "Job error: $($_.Exception.Message)"
                }
                Remove-Job -Job $job -Force
            }
        }
    }
    
    # Start new job for this batch
    $jobCounter++
    $batchNum = $batchInfo.BatchNumber
    Write-Host "Starting job $jobCounter/$($groupBatches.Count) (Batch #$batchNum with $($batchInfo.Batch.Count) groups)..." -ForegroundColor Cyan
    
    $job = Start-Job -ScriptBlock {
        param($GroupBatch, $SkipMembers, $SkipOwners, $ClientId, $ClientSecret, $TenantId, $BatchNumber, $DebugMode)
        
        # Re-establish Graph connection in job
        try {
            $SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome -ErrorAction Stop
        }
        catch {
            Write-Error "Job Batch $BatchNumber - Failed to connect to Graph: $($_.Exception.Message)"
            return @()
        }
        
        $batchResults = @()
        $groupIndex = 0
        
        foreach ($group in $GroupBatch) {
            $groupIndex++
            try {
                if ($DebugMode) {
                    Write-Verbose "Job Batch $BatchNumber - Processing group $groupIndex/$($GroupBatch.Count): $($group.DisplayName)" -Verbose
                }
                
                $memberDetails = @()
                $ownerDetails = @()
                
                # Get members if not skipped
                if (-not $SkipMembers) {
                    try {
                        $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
                        if ($members) {
                            foreach ($member in $members) {
                                try {
                                    $memberDetails += [PSCustomObject]@{
                                        Id = $member.Id
                                        DisplayName = $member.AdditionalProperties.displayName
                                        UserPrincipalName = $member.AdditionalProperties.userPrincipalName
                                        ObjectType = ($member.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                                    }
                                }
                                catch {
                                    if ($DebugMode) {
                                        Write-Warning "Job Batch $BatchNumber - Failed to process member in group $($group.DisplayName): $($_.Exception.Message)"
                                    }
                                    continue
                                }
                            }
                        }
                    }
                    catch {
                        if ($DebugMode) {
                            Write-Warning "Job Batch $BatchNumber - Failed to get members for group $($group.DisplayName): $($_.Exception.Message)"
                        }
                    }
                }
                
                # Get owners if not skipped
                if (-not $SkipOwners) {
                    try {
                        $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop
                        if ($owners) {
                            foreach ($owner in $owners) {
                                try {
                                    $ownerDetails += [PSCustomObject]@{
                                        Id = $owner.Id
                                        DisplayName = $owner.AdditionalProperties.displayName
                                        UserPrincipalName = $owner.AdditionalProperties.userPrincipalName
                                        ObjectType = ($owner.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                                    }
                                }
                                catch {
                                    if ($DebugMode) {
                                        Write-Warning "Job Batch $BatchNumber - Failed to process owner in group $($group.DisplayName): $($_.Exception.Message)"
                                    }
                                    continue
                                }
                            }
                        }
                    }
                    catch {
                        if ($DebugMode) {
                            Write-Warning "Job Batch $BatchNumber - Failed to get owners for group $($group.DisplayName): $($_.Exception.Message)"
                        }
                    }
                }

                # FIX: Determine group type once here, used by both the mail-owner block below
                # and the stored GroupType value. The original code had two separate type-detection
                # blocks ($groupTypeCheck and $groupType) which could theoretically diverge.
                $groupType = "Security Group"
                if ($group.GroupTypes -contains "Unified") {
                    $groupType = "Microsoft 365 Group"
                } elseif ($group.MailEnabled -and -not $group.SecurityEnabled) {
                    $groupType = "Distribution Group"
                } elseif ($group.MailEnabled -and $group.SecurityEnabled) {
                    $groupType = "Mail-enabled Security Group"
                } elseif ($group.GroupTypes -contains "DynamicMembership") {
                    $groupType = "Dynamic Security Group"
                }

                # For Distribution Groups and Mail-enabled Security Groups, ensure owners are always
                # collected regardless of -SkipOwners, because these types require owner visibility.
                # Results go into $mailGroupOwnerDetails to avoid duplicating $ownerDetails.
                $mailGroupOwnerDetails = @()
                if ($groupType -in @("Distribution Group", "Mail-enabled Security Group")) {
                    if ($SkipOwners) {
                        # Owners were not fetched above — fetch now specifically for these group types.
                        try {
                            $mailOwners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop
                            if ($mailOwners) {
                                foreach ($owner in $mailOwners) {
                                    try {
                                        $mailGroupOwnerDetails += [PSCustomObject]@{
                                            Id = $owner.Id
                                            DisplayName = $owner.AdditionalProperties.displayName
                                            UserPrincipalName = $owner.AdditionalProperties.userPrincipalName
                                            ObjectType = ($owner.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                                        }
                                    }
                                    catch {
                                        if ($DebugMode) {
                                            Write-Warning "Job Batch $BatchNumber - Failed to process mail-group owner in group $($group.DisplayName): $($_.Exception.Message)"
                                        }
                                        continue
                                    }
                                }
                            }
                        }
                        catch {
                            if ($DebugMode) {
                                Write-Warning "Job Batch $BatchNumber - Failed to get mail-group owners for $($group.DisplayName): $($_.Exception.Message)"
                            }
                        }
                    } else {
                        # Owners already collected in $ownerDetails — just reference them.
                        $mailGroupOwnerDetails = $ownerDetails
                    }
                }

                # FIX: OwnerCount must reflect the owners that were actually collected.
                # When -SkipOwners is set for DG / Mail-enabled SG, $ownerDetails is empty but
                # $mailGroupOwnerDetails has the data — use that count instead.
                $effectiveOwnerCount = $ownerDetails.Count
                if ($groupType -in @("Distribution Group", "Mail-enabled Security Group") -and $SkipOwners) {
                    $effectiveOwnerCount = $mailGroupOwnerDetails.Count
                }
                
                # Create detailed group object
                $groupDetails = [PSCustomObject]@{
                    Id = $group.Id
                    DisplayName = $group.DisplayName
                    Description = $group.Description
                    GroupType = $groupType
                    SecurityEnabled = $group.SecurityEnabled
                    MailEnabled = $group.MailEnabled
                    Mail = $group.Mail
                    MailNickname = $group.MailNickname
                    CreatedDateTime = $group.CreatedDateTime
                    Visibility = $group.Visibility
                    MembershipRule = $group.MembershipRule
                    MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                    ProxyAddresses = $group.ProxyAddresses -join "; "
                    OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                    OnPremisesSecurityIdentifier = $group.OnPremisesSecurityIdentifier
                    IsDynamic = $group.GroupTypes -contains "DynamicMembership"
                    IsUnified = $group.GroupTypes -contains "Unified"
                    MemberCount = $memberDetails.Count
                    OwnerCount = $effectiveOwnerCount
                    Members = $memberDetails
                    Owners = $ownerDetails
                    MailGroupOwners = $mailGroupOwnerDetails
                    GroupTypes = $group.GroupTypes -join ", "
                    ProcessedInBatch = $BatchNumber
                }
                
                $batchResults += $groupDetails
            }
            catch {
                Write-Error "Job Batch $BatchNumber - Error processing group $($group.DisplayName) (ID: $($group.Id)): $($_.Exception.Message)"
                # Still try to add basic info even if processing failed
                $batchResults += [PSCustomObject]@{
                    Id = $group.Id
                    DisplayName = $group.DisplayName
                    Description = $group.Description
                    GroupType = "Unknown (Processing Failed)"
                    SecurityEnabled = $group.SecurityEnabled
                    MailEnabled = $group.MailEnabled
                    Mail = $group.Mail
                    MailNickname = $group.MailNickname
                    CreatedDateTime = $group.CreatedDateTime
                    Visibility = $group.Visibility
                    MembershipRule = $group.MembershipRule
                    MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                    ProxyAddresses = $group.ProxyAddresses -join "; "
                    OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                    OnPremisesSecurityIdentifier = $group.OnPremisesSecurityIdentifier
                    IsDynamic = $group.GroupTypes -contains "DynamicMembership"
                    IsUnified = $group.GroupTypes -contains "Unified"
                    MemberCount = 0
                    OwnerCount = 0
                    Members = @()
                    Owners = @()
                    MailGroupOwners = @()
                    GroupTypes = $group.GroupTypes -join ", "
                    ProcessedInBatch = $BatchNumber
                    ProcessingError = $_.Exception.Message
                }
            }
        }
        
        if ($DebugMode) {
            Write-Verbose "Job Batch $BatchNumber - Completed with $($batchResults.Count) results" -Verbose
        }
        
        return $batchResults
    } -ArgumentList $batchInfo.Batch, $SkipMembers.IsPresent, $SkipOwners.IsPresent, $ClientId, $ClientSecret, $TenantId, $batchNum, $DebugMode.IsPresent
    
    $jobList += $job
}

# Wait for all jobs to complete and collect results
Write-Host "`nWaiting for all jobs to complete..." -ForegroundColor Yellow
$lastRunningCount = -1

while ((Get-Job -State Running).Count -gt 0) {
    $runningCount = (Get-Job -State Running).Count
    
    if ($runningCount -ne $lastRunningCount) {
        Write-Host "Jobs still running: $runningCount" -ForegroundColor Cyan
        $lastRunningCount = $runningCount
    }
    
    Start-Sleep -Seconds 2
    
    # Collect completed jobs
    $completedJobs = Get-Job -State Completed
    if ($completedJobs) {
        foreach ($job in $completedJobs) {
            try {
                $jobResult = Receive-Job -Job $job -ErrorAction Stop
                if ($jobResult) {
                    $allResults += $jobResult
                    Write-Host "Collected $($jobResult.Count) results from completed job" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "Error receiving job results: $($_.Exception.Message)"
            }
            Remove-Job -Job $job -Force
        }
    }
    
    # Check for failed jobs
    $failedJobs = Get-Job -State Failed
    if ($failedJobs) {
        foreach ($job in $failedJobs) {
            Write-Warning "Job $($job.Id) failed!"
            try {
                Receive-Job -Job $job -ErrorAction Continue
            }
            catch {
                Write-Warning "Job error details: $($_.Exception.Message)"
            }
            Remove-Job -Job $job -Force
        }
    }
}

# Final collection of any remaining results
Write-Host "`nCollecting final results..." -ForegroundColor Yellow
$remainingJobs = Get-Job
foreach ($job in $remainingJobs) {
    try {
        $jobResult = Receive-Job -Job $job -ErrorAction Stop
        if ($jobResult) {
            $allResults += $jobResult
        }
    }
    catch {
        Write-Warning "Error in final collection: $($_.Exception.Message)"
    }
}

# Clean up all jobs
Get-Job | Remove-Job -Force

$results = $allResults

$stopwatch.Stop()

# Verify all groups were processed
Write-Host "`n=== PROCESSING VERIFICATION ===" -ForegroundColor Magenta
$processedGroupIds = $results | Select-Object -ExpandProperty Id -Unique
$missingGroups = $allGroups | Where-Object { $processedGroupIds -notcontains $_.Id }

if ($missingGroups) {
    Write-Host "`nWARNING: $($missingGroups.Count) groups were not processed!" -ForegroundColor Red
    Write-Host "Missing groups:" -ForegroundColor Red
    $missingGroups | Select-Object DisplayName, Id | Format-Table -AutoSize
    
    # Attempt to reprocess missing groups
    Write-Host "`nAttempting to reprocess missing groups individually..." -ForegroundColor Yellow
    
    foreach ($group in $missingGroups) {
        try {
            Write-Host "Reprocessing: $($group.DisplayName)" -ForegroundColor Cyan
            
            $memberDetails = @()
            $ownerDetails = @()
            
            if (-not $SkipMembers) {
                try {
                    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
                    foreach ($member in $members) {
                        $memberDetails += [PSCustomObject]@{
                            Id = $member.Id
                            DisplayName = $member.AdditionalProperties.displayName
                            UserPrincipalName = $member.AdditionalProperties.userPrincipalName
                            ObjectType = ($member.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                        }
                    }
                }
                catch {
                    Write-Warning "Could not retrieve members: $($_.Exception.Message)"
                }
            }
            
            if (-not $SkipOwners) {
                try {
                    $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue
                    foreach ($owner in $owners) {
                        $ownerDetails += [PSCustomObject]@{
                            Id = $owner.Id
                            DisplayName = $owner.AdditionalProperties.displayName
                            UserPrincipalName = $owner.AdditionalProperties.userPrincipalName
                            ObjectType = ($owner.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                        }
                    }
                }
                catch {
                    Write-Warning "Could not retrieve owners: $($_.Exception.Message)"
                }
            }

            # FIX: Retry block now uses the same full group type detection as the job block,
            # including the DynamicMembership branch that was missing in the original —
            # without it, Dynamic Security Groups fell through to "Security Group", and more
            # importantly Mail-enabled Security Groups could be misclassified, preventing
            # their mail-group owner fetch from ever running.
            $retryGroupType = "Security Group"
            if ($group.GroupTypes -contains "Unified") {
                $retryGroupType = "Microsoft 365 Group"
            } elseif ($group.MailEnabled -and -not $group.SecurityEnabled) {
                $retryGroupType = "Distribution Group"
            } elseif ($group.MailEnabled -and $group.SecurityEnabled) {
                $retryGroupType = "Mail-enabled Security Group"
            } elseif ($group.GroupTypes -contains "DynamicMembership") {
                $retryGroupType = "Dynamic Security Group"
            }

            # Also fetch mail-group owners during retry for DG / Mail-enabled SG
            $mailGroupOwnerDetails = @()
            if ($retryGroupType -in @("Distribution Group", "Mail-enabled Security Group")) {
                if ($SkipOwners) {
                    try {
                        $mailOwners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction SilentlyContinue
                        foreach ($owner in $mailOwners) {
                            $mailGroupOwnerDetails += [PSCustomObject]@{
                                Id = $owner.Id
                                DisplayName = $owner.AdditionalProperties.displayName
                                UserPrincipalName = $owner.AdditionalProperties.userPrincipalName
                                ObjectType = ($owner.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.','') -replace '^$','Unknown'
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not retrieve mail-group owners: $($_.Exception.Message)"
                    }
                } else {
                    $mailGroupOwnerDetails = $ownerDetails
                }
            }

            # FIX: same effective owner count logic as in the job block
            $effectiveOwnerCount = $ownerDetails.Count
            if ($retryGroupType -in @("Distribution Group", "Mail-enabled Security Group") -and $SkipOwners) {
                $effectiveOwnerCount = $mailGroupOwnerDetails.Count
            }
            
            $recoveredGroup = [PSCustomObject]@{
                Id = $group.Id
                DisplayName = $group.DisplayName
                Description = $group.Description
                GroupType = $retryGroupType
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled = $group.MailEnabled
                Mail = $group.Mail
                MailNickname = $group.MailNickname
                CreatedDateTime = $group.CreatedDateTime
                Visibility = $group.Visibility
                MembershipRule = $group.MembershipRule
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                ProxyAddresses = $group.ProxyAddresses -join "; "
                OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                OnPremisesSecurityIdentifier = $group.OnPremisesSecurityIdentifier
                IsDynamic = $group.GroupTypes -contains "DynamicMembership"
                IsUnified = $group.GroupTypes -contains "Unified"
                MemberCount = $memberDetails.Count
                OwnerCount = $effectiveOwnerCount
                Members = $memberDetails
                Owners = $ownerDetails
                MailGroupOwners = $mailGroupOwnerDetails
                GroupTypes = $group.GroupTypes -join ", "
                ProcessedInBatch = "Retry"
            }
            
            $results += $recoveredGroup
            Write-Host "Successfully recovered: $($group.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to reprocess $($group.DisplayName): $($_.Exception.Message)"
        }
    }
} else {
    Write-Host "All groups successfully processed!" -ForegroundColor Green
}

if ($EnrichOwnersFromExchangeOnline) {
    Write-Host "`nEnriching owners from Exchange Online (ManagedBy) where Graph returned no owners (normal for Exchange-managed DL / mail-enabled security groups)..." -ForegroundColor Green
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Warning "ExchangeOnlineManagement module is not installed. Install with: Install-Module ExchangeOnlineManagement. Skipping Exchange owner enrichment."
    } else {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $exoOrg = $ExchangeOrganization
        if ([string]::IsNullOrWhiteSpace($exoOrg)) {
            try {
                $orgInfo = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
                $exoOrg = ($orgInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1 -ExpandProperty Name)
            }
            catch {
                Write-Warning "Could not read organization from Graph: $($_.Exception.Message)"
            }
        }
        if ([string]::IsNullOrWhiteSpace($exoOrg)) {
            Write-Warning "Exchange organization not determined. Set -ExchangeOrganization (e.g. contoso.onmicrosoft.com). Skipping Exchange owner enrichment."
        } else {
            $exoConnected = $false
            try {
                if (-not [string]::IsNullOrWhiteSpace($ExchangeCertificateThumbprint)) {
                    Write-Host "Connecting to Exchange Online (app-only certificate)..." -ForegroundColor Yellow
                    Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $ExchangeCertificateThumbprint `
                        -Organization $exoOrg -ShowBanner:$false -ErrorAction Stop
                }
                elseif (-not [string]::IsNullOrWhiteSpace($ClientSecret)) {
                    Write-Host "Connecting to Exchange Online (app-only client credentials / same secret as Graph)..." -ForegroundColor Yellow
                    $exoToken = Get-ExchangeOnlineAccessTokenClientCredentials -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
                    Connect-ExchangeOnline -AccessToken $exoToken -Organization $exoOrg -ShowBanner:$false -ErrorAction Stop
                }
                elseif ($ExchangeOnlineAllowInteractiveLogin) {
                    Write-Host "Connecting to Exchange Online interactively (sign in when prompted)..." -ForegroundColor Yellow
                    Connect-ExchangeOnline -Organization $exoOrg -ShowBanner:$false -ErrorAction Stop
                }
                else {
                    throw "No Exchange auth method: set AZURE_CLIENT_SECRET for unattended EXO (app needs Exchange.ManageAsApp + Exchange RBAC), or -ExchangeCertificateThumbprint, or -ExchangeOnlineAllowInteractiveLogin."
                }
                $exoConnected = $true

                Write-Host "Loading distribution groups and mail-enabled security groups from Exchange..." -ForegroundColor Yellow
                try {
                    $exoRecipients = @(Get-DistributionGroup -ResultSize Unlimited -Property ExternalDirectoryObjectId, ManagedBy -ErrorAction Stop)
                }
                catch {
                    $exoRecipients = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
                }
                $managedByByExtId = @{}
                foreach ($exo in $exoRecipients) {
                    $extId = if ($null -ne $exo.ExternalDirectoryObjectId) { $exo.ExternalDirectoryObjectId.ToString() } else { $null }
                    if ([string]::IsNullOrWhiteSpace($extId)) { continue }
                    $managedByByExtId[$extId] = $exo.ManagedBy
                }

                $recipientCache = @{}
                $enrichedCount = 0

                function Apply-ExchangeOwnersToGroup {
                    param($GroupRow, $ManagedByRaw)
                    $exoOwners = Resolve-ExchangeManagedByToOwnerObjects -ManagedByProperty $ManagedByRaw -RecipientCache $recipientCache
                    if (-not $exoOwners -or $exoOwners.Count -eq 0) { return $false }
                    $GroupRow.Owners = @($exoOwners)
                    $GroupRow.OwnerCount = $exoOwners.Count
                    if ($GroupRow.GroupType -in @("Distribution Group", "Mail-enabled Security Group")) {
                        $GroupRow.MailGroupOwners = @($exoOwners)
                    }
                    return $true
                }

                foreach ($g in $results) {
                    $hasGraphOwners = ($g.Owners -and $g.Owners.Count -gt 0)
                    if ($hasGraphOwners) { continue }

                    $key = [string]$g.Id
                    $applied = $false

                    if ($managedByByExtId.ContainsKey($key)) {
                        if (Apply-ExchangeOwnersToGroup -GroupRow $g -ManagedByRaw $managedByByExtId[$key]) {
                            $enrichedCount++
                            $applied = $true
                        }
                    }

                    # Bulk map missed or ManagedBy was empty — resolve the DL / mail-enabled SG directly in EXO (GUID or primary SMTP).
                    if (-not $applied -and $g.GroupType -in @("Distribution Group", "Mail-enabled Security Group")) {
                        $directMb = Get-ExchangeDistributionGroupManagedByProperty -GraphGroupId $key -PrimarySmtpAddress ([string]$g.Mail)
                        if (Apply-ExchangeOwnersToGroup -GroupRow $g -ManagedByRaw $directMb) {
                            $enrichedCount++
                        }
                    }
                }
                Write-Host "Exchange ManagedBy applied to $enrichedCount group(s) with no Graph owners." -ForegroundColor Green
            }
            catch {
                Write-Warning "Exchange Online owner enrichment failed: $($_.Exception.Message)"
                Write-Warning "For client-secret auth, the app registration needs application permission Exchange.ManageAsApp (Office 365 Exchange Online) with admin consent, and an Exchange Online admin role assigned to the enterprise app (Exchange RBAC for applications)."
            }
            finally {
                if ($exoConnected) {
                    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
                }
            }
        }
    }
}
elseif (-not $EnrichOwnersFromExchangeOnline) {
    Write-Host "`nExchange Online owner enrichment skipped (-EnrichOwnersFromExchangeOnline:`$false). Distribution / mail-enabled security manager rows may stay empty (Graph does not expose them)." -ForegroundColor Yellow
}

Write-Host "Processing completed in $($stopwatch.Elapsed.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Green
Write-Host "Successfully processed $($results.Count) groups." -ForegroundColor Green

# Export results if OutputPath is specified
if (-not [string]::IsNullOrEmpty($OutputPath)) {
    Write-Host "`nExporting results to: $OutputPath" -ForegroundColor Yellow
    
    # Ensure directory exists
    $directory = Split-Path $OutputPath -Parent
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    # Prepare export data
    $exportResults = @()
    foreach ($group in $results) {
        $ownerNames = ""
        if ($group.Owners -and $group.Owners.Count -gt 0) {
            $ownerNames = ($group.Owners | ForEach-Object {
                $n = $_.DisplayName
                if ([string]::IsNullOrWhiteSpace($n)) { $n = $_.UserPrincipalName }
                if ([string]::IsNullOrWhiteSpace($n)) { $n = [string]$_.Id }
                $n
            }) -join ", "
        }
        
        $memberNames = ""
        if ($group.Members -and $group.Members.Count -gt 0) {
            $memberNames = ($group.Members | ForEach-Object { $_.DisplayName }) -join ", "
        }

        # Resolve assigned applications from the pre-built lookup map
        $assignedApps = ""
        if ($groupAppAssignmentMap.ContainsKey($group.Id)) {
            $assignedApps = $groupAppAssignmentMap[$group.Id] -join "; "
        }

        # Resolve mail-group owner names (Distribution Group / Mail-enabled Security Group)
        $mailGroupOwnerNames = ""
        if ($group.MailGroupOwners -and $group.MailGroupOwners.Count -gt 0) {
            $mailGroupOwnerNames = ($group.MailGroupOwners | ForEach-Object {
                $n = $_.DisplayName
                if ([string]::IsNullOrWhiteSpace($n)) { $n = $_.UserPrincipalName }
                if ([string]::IsNullOrWhiteSpace($n)) { $n = [string]$_.Id }
                $n
            }) -join ", "
        }

        # FIX: Build OwnerNames without duplication or a leading comma.
        # When -SkipOwners is NOT set: $ownerNames and $mailGroupOwnerNames hold the same data
        #   ($mailGroupOwnerDetails = $ownerDetails in the job), so use $ownerNames only.
        # When -SkipOwners IS set for DG / Mail-enabled SG: $ownerNames is empty and
        #   $mailGroupOwnerNames has the data, so fall back to $mailGroupOwnerNames.
        # For all other group types $mailGroupOwnerNames is always empty, so this is safe universally.
        $resolvedOwnerNames = if ($ownerNames) { $ownerNames } else { $mailGroupOwnerNames }

        $exportResults += [PSCustomObject]@{
            DisplayName = Normalize-CsvField $group.DisplayName
            GroupType = $group.GroupType
            SecurityEnabled = $group.SecurityEnabled
            MailEnabled = $group.MailEnabled
            Mail = $group.Mail
            CreatedDateTime = $group.CreatedDateTime
            Visibility = $group.Visibility
            MemberCount = $group.MemberCount
            OwnerCount = $group.OwnerCount
            OwnerNames = (Normalize-CsvField $resolvedOwnerNames)
#            MemberNames = Normalize-CsvField $memberNames
#            MailGroupOwnerNames = Normalize-CsvField $mailGroupOwnerNames
            AssignedApplications = Normalize-CsvField $assignedApps
            GroupTypes = ($group.GroupTypes -replace ",", ";")
            IsDynamic = $group.IsDynamic
            IsUnified = $group.IsUnified
            Id = $group.Id
            Description = Normalize-CsvField $group.Description
            MailNickname = $group.MailNickname
            OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
            ProcessedInBatch = $group.ProcessedInBatch
        }
    }
    
    # Export to CSV
    try {
        $exportResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Successfully exported to: $OutputPath" -ForegroundColor Green
        
        # Also export detailed JSON
        $jsonPath = $OutputPath -replace '\.csv$', '_Detailed.json'
        $results | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Host "Detailed data exported to: $jsonPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# Display summary
Write-Host "`n=== GROUP SUMMARY ===" -ForegroundColor Green
Write-Host "PowerShell Version: $psVersion" -ForegroundColor Yellow
Write-Host "Total Groups Found: $($allGroups.Count)" -ForegroundColor Yellow
Write-Host "Total Groups Processed: $($results.Count)" -ForegroundColor Yellow
Write-Host "Processing Time: $($stopwatch.Elapsed.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Yellow

if ($results.Count -gt 0) {
    Write-Host "Average Time per Group: $((($stopwatch.Elapsed.TotalSeconds) / $results.Count).ToString('F2')) seconds" -ForegroundColor Yellow
}

Write-Host "`nGroup Type Breakdown:" -ForegroundColor Cyan
$groupTypeCounts = $results | Group-Object GroupType | Sort-Object Count -Descending
foreach ($typeGroup in $groupTypeCounts) {
    Write-Host "  $($typeGroup.Name): $($typeGroup.Count)" -ForegroundColor White
}

# Show groups with errors if any
$groupsWithErrors = $results | Where-Object { $_.PSObject.Properties.Name -contains 'ProcessingError' }
if ($groupsWithErrors) {
    Write-Host "`nGroups with Processing Errors: $($groupsWithErrors.Count)" -ForegroundColor Red
    $groupsWithErrors | Select-Object DisplayName, ProcessingError | Format-Table -AutoSize
}

Write-Host "`n=== PERFORMANCE TIPS ===" -ForegroundColor Magenta
Write-Host "* Use -BatchSize parameter to adjust batch processing (current: $BatchSize)"
Write-Host "* Use -MaxJobs parameter to control concurrent jobs (current: $MaxJobs)"
Write-Host "* Use -SkipMembers and -SkipOwners switches for faster processing"
Write-Host "* Use -DebugMode for detailed logging during processing"
Write-Host "* Exchange ManagedBy enrichment: unattended uses the same AZURE_CLIENT_* client secret as Graph (app needs Exchange.ManageAsApp + Exchange RBAC), or -ExchangeCertificateThumbprint; use -ExchangeOnlineAllowInteractiveLogin for browser sign-in"
Write-Host "* Consider upgrading to PowerShell 7 for better parallel processing"

Write-Host "`nScript completed successfully!" -ForegroundColor Green
