param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = "group_review_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [int]$BatchSize = 50,  # Process reviews in batches for progress updates
    [switch]$IncludeHtml
)


If ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "ClientId, ClientSecret, and TenantId must be set."
    Exit 1
}

# ---------------------------------------------------------------------------
# Ensure required modules are available
# ---------------------------------------------------------------------------
$RequiredModules = @('Microsoft.Graph.Authentication')
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Host "Installing required module: $Module" -ForegroundColor Yellow
        Install-Module -Name $Module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Module -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# Connect to Graph using Service Principal
# ---------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph using Service Principal..." -ForegroundColor Cyan
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Helper: resolve a group's display name (cached to avoid repeat lookups)
# ---------------------------------------------------------------------------
$GroupNameCache = @{}
function Get-GroupDisplayName {
    param([string]$GroupId)

    if ([string]::IsNullOrEmpty($GroupId)) { return $null }

    if ($GroupNameCache.ContainsKey($GroupId)) {
        return $GroupNameCache[$GroupId]
    }

    try {
        $Uri = "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=id,displayName,mailNickname"
        $Group = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        $GroupNameCache[$GroupId] = $Group.displayName
        return $Group.displayName
    }
    catch {
        $GroupNameCache[$GroupId] = "<Group not found / deleted: $GroupId>"
        return $GroupNameCache[$GroupId]
    }
}

# ---------------------------------------------------------------------------
# Helper: extract Group ID(s) from an Access Review scope object
#   Scope shapes seen in the Graph API:
#     - Single group review:
#         query = "/groups/{id}/transitiveMembers/microsoft.graph.user"
#         or query = "/groups/{id}/owners"
#     - Multi-group ("all groups" / dynamic scope) reviews expose
#         resourceScopes[] with the same query pattern per group,
#         or use queryType "MicrosoftGraph" with query "/groups"
#         plus a filter (e.g. group-membership-based dynamic reviews).
# ---------------------------------------------------------------------------
function Get-ScopeGroupIds {
    param($Scope)

    $ids = New-Object System.Collections.Generic.List[string]

    if ($null -eq $Scope) { return $ids }

    $queries = @()
    if ($Scope.query) { $queries += $Scope.query }
    if ($Scope.resourceScopes) {
        foreach ($rs in $Scope.resourceScopes) {
            if ($rs.query) { $queries += $rs.query }
        }
    }

    foreach ($q in $queries) {
        if ($q -match '/groups/([0-9a-fA-F-]{36})') {
            $ids.Add($Matches[1]) | Out-Null
        }
    }

    return $ids | Select-Object -Unique
}

# ---------------------------------------------------------------------------
# Helper: determine if a review targets groups at all (vs. Azure resources,
# access packages, etc.) based on scope query / resourceScopes content
# ---------------------------------------------------------------------------
function Test-IsGroupReview {
    param($Definition)

    $scope = $Definition.scope
    if ($null -eq $scope) { return $false }

    if ($scope.'@odata.type' -match 'accessReviewQueryScope' -or $scope.query) {
        if ($scope.query -match '/groups/') { return $true }
    }
    if ($scope.resourceScopes) {
        foreach ($rs in $scope.resourceScopes) {
            if ($rs.query -match '/groups/') { return $true }
        }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Helper: friendly frequency string from settings.recurrence
# ---------------------------------------------------------------------------
function Get-FrequencyDescription {
    param($Settings)

    if ($null -eq $Settings -or $null -eq $Settings.recurrence) { return "One-time / Unknown" }

    $pattern = $Settings.recurrence.pattern
    if ($null -eq $pattern) { return "One-time" }

    $type = $pattern.type
    $interval = $pattern.interval

    switch ($type) {
        'absoluteMonthly' {
            if ($interval -eq 1) { return "Monthly" }
            elseif ($interval -eq 3) { return "Quarterly" }
            elseif ($interval -eq 6) { return "Semi-Annually" }
            elseif ($interval -eq 12) { return "Annually" }
            else { return "Every $interval months" }
        }
        'absoluteYearly' { return "Annually (every $interval year(s))" }
        'weekly' { return "Every $interval week(s)" }
        'daily' { return "Every $interval day(s)" }
        default {
            if ($Settings.recurrence.range -and $Settings.recurrence.range.type -eq 'noEnd' -and -not $type) {
                return "One-time"
            }
            return "$type (interval: $interval)"
        }
    }
}

# ---------------------------------------------------------------------------
# Helper: friendly reviewer configuration string
# ---------------------------------------------------------------------------
function Get-ReviewerDescription {
    param($Definition)

    $reviewers = $Definition.reviewers
    $fallback = $Definition.fallbackReviewers
    $settings = $Definition.settings

    $descriptions = New-Object System.Collections.Generic.List[string]

    # Self-review flag (settings.reviewerType or presence of "reviewers" being empty w/ self-review)
    if ($settings -and $settings.reviewerType -eq 'Self') {
        $descriptions.Add("Self-Review (members review their own access)") | Out-Null
    }

    if ($reviewers -and $reviewers.Count -gt 0) {
        foreach ($r in $reviewers) {
            $type = $r.'@odata.type' -replace '#microsoft.graph.', ''
            switch ($type) {
                'groupMembersReviewers' { $descriptions.Add("Group Owners") | Out-Null }
                'userIdentity' {
                    $name = if ($r.displayName) { $r.displayName } else { $r.id }
                    $descriptions.Add("Specific User: $name") | Out-Null
                }
                default {
                    if ($r.displayName) { $descriptions.Add("$type : $($r.displayName)") | Out-Null }
                    else { $descriptions.Add("$type") | Out-Null }
                }
            }
        }
    }

    if ($settings -and $settings.PSObject.Properties.Name -contains 'reviewerType' -and $settings.reviewerType -and $settings.reviewerType -ne 'Self') {
        # e.g. "GroupOwners", "Manager", "Reviewers", "Delegated"
        if ($descriptions -notcontains $settings.reviewerType) {
            $descriptions.Add($settings.reviewerType) | Out-Null
        }
    }

    if ($fallback -and $fallback.Count -gt 0) {
        foreach ($fb in $fallback) {
            $type = $fb.'@odata.type' -replace '#microsoft.graph.', ''
            $name = if ($fb.displayName) { $fb.displayName } else { $fb.id }
            $descriptions.Add("Fallback: $type $name") | Out-Null
        }
    }

    if ($descriptions.Count -eq 0) { return "Not specified / No reviewers configured" }

    return ($descriptions -join "; ")
}

# ---------------------------------------------------------------------------
# Retrieve all Access Review definitions (paged), expanding useful properties
# ---------------------------------------------------------------------------
Write-Host "Retrieving Access Review definitions from Microsoft Graph..." -ForegroundColor Cyan

$AllDefinitions = New-Object System.Collections.Generic.List[object]
$Uri = "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?`$expand=instances"

do {
    try {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to retrieve access review definitions: $_"
        Disconnect-MgGraph | Out-Null
        Exit 1
    }

    if ($Response.value) {
        $AllDefinitions.AddRange($Response.value)
    }

    $Uri = $Response.'@odata.nextLink'
} while ($Uri)

Write-Host "Retrieved $($AllDefinitions.Count) total access review definitions." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Filter to Group-scoped reviews and build report rows
# ---------------------------------------------------------------------------
$ReportRows = New-Object System.Collections.Generic.List[object]
$Counter = 0
$GroupReviewDefinitions = $AllDefinitions | Where-Object { Test-IsGroupReview -Definition $_ }

Write-Host "Found $($GroupReviewDefinitions.Count) access review definitions scoped to Groups." -ForegroundColor Cyan

foreach ($Definition in $GroupReviewDefinitions) {

    $Counter++
    if ($Counter % $BatchSize -eq 0) {
        Write-Host "  Processed $Counter / $($GroupReviewDefinitions.Count) group reviews..." -ForegroundColor DarkCyan
    }

    # Need full detail (settings/reviewers aren't always fully populated on
    # the list response) - fetch the single definition for completeness
    try {
        $DetailUri = "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions/$($Definition.id)?`$expand=instances"
        $Detail = Invoke-MgGraphRequest -Method GET -Uri $DetailUri -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve full detail for review '$($Definition.displayName)' ($($Definition.id)): $_"
        $Detail = $Definition
    }

    $GroupIds = Get-ScopeGroupIds -Scope $Detail.scope
    if ($GroupIds.Count -eq 0) { $GroupIds = @('Unknown') }

    $GroupNames = $GroupIds | ForEach-Object {
        if ($_ -eq 'Unknown') { 'Unknown / Dynamic scope' } else { Get-GroupDisplayName -GroupId $_ }
    }

    $Settings = $Detail.settings

    $DurationDays = if ($Settings -and $Settings.instanceDurationInDays) { "$($Settings.instanceDurationInDays) day(s)" } else { "Not specified" }
    $Frequency = Get-FrequencyDescription -Settings $Settings
    $AutoApply = if ($Settings) { [bool]$Settings.autoApplyDecisionsEnabled } else { $false }
    $ReviewerConfig = Get-ReviewerDescription -Definition $Detail

    $StartDate = $null
    $EndDate = $null
    if ($Detail.instances -and $Detail.instances.Count -gt 0) {
        $StartDate = $Detail.instances[0].startDateTime
        $EndDate = $Detail.instances[0].endDateTime
    }
    $Status = $Detail.status

    $RowsPerGroup = $GroupNames

    foreach ($GName in $RowsPerGroup) {
        $ReportRows.Add([PSCustomObject]@{
            'Review Name'             = $Detail.displayName
            'Group Name'              = $GName
            'Status'                  = $Status
            'Duration'                = $DurationDays
            'Frequency'               = $Frequency
            'Auto-Apply Enabled'      = $AutoApply
            'Reviewer Configuration'  = $ReviewerConfig
            'Decisions Requiring Justification' = if ($Settings) { [bool]$Settings.justificationRequiredOnApproval } else { $null }
            'Recommendations Enabled' = if ($Settings) { [bool]$Settings.recommendationsEnabled } else { $null }
            'Mail Notifications Enabled' = if ($Settings) { [bool]$Settings.mailNotificationsEnabled } else { $null }
            'Reminders Enabled'       = if ($Settings) { [bool]$Settings.remindersEnabled } else { $null }
            'Current Instance Start'  = $StartDate
            'Current Instance End'    = $EndDate
            'Review Definition Id'    = $Detail.id
        }) | Out-Null
    }
}

Write-Host "Report generation complete. $($ReportRows.Count) row(s) produced." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Export report
# ---------------------------------------------------------------------------
i
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = "review_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

$ReportRows | Sort-Object 'Review Name', 'Group Name' | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV report written to: $OutputPath" -ForegroundColor Green


# ---------------------------------------------------------------------------
# Console summary
# ---------------------------------------------------------------------------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Total access review definitions in tenant : $($AllDefinitions.Count)"
Write-Host "Group-scoped access review definitions    : $($GroupReviewDefinitions.Count)"
Write-Host "Report rows (1 per review/group pair)     : $($ReportRows.Count)"

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph." -ForegroundColor Cyan