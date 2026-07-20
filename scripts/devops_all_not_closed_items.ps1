param(
    [string]$OutputPath,
    [string]$Organization,
    [string]$PersonalAccessToken = $env:PAT,
    [string]$UserEmails
)


# Function to create authorization header
function Get-AuthHeader {
    param([string]$token)
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$token"))
    return @{Authorization = "Basic $base64AuthInfo"}
}

# Function to get work item details
function Get-WorkItemDetails {
    param(
        [int]$workItemId,
        [hashtable]$headers,
        [string]$organization
    )
    
    $uri = "https://dev.azure.com/$organization/_apis/wit/workitems/$workItemId" +
           "?`$expand=relations&api-version=7.0"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        return $response
    }
    catch {
        Write-Warning "Failed to get work item $workItemId : $($_.Exception.Message)"
        return $null
    }
}

# Function to build hierarchy path
function Get-WorkItemHierarchyPath {
    param(
        [object]$workItem,
        [hashtable]$headers,
        [string]$organization,
        [hashtable]$workItemCache = @{}
    )
    
    $path = @()
    $currentItem = $workItem
    
    while ($currentItem) {
        $workItemType = $currentItem.fields.'System.WorkItemType'
        $id = $currentItem.id
        $title = $currentItem.fields.'System.Title'
        
        # Add current item to path (prepend to build hierarchy from root)
        $itemInfo = "$workItemType $id"
        $path = @($itemInfo) + $path
        
        # Find parent relationship
        $parentId = $null
        if ($currentItem.relations) {
            foreach ($relation in $currentItem.relations) {
                if ($relation.rel -eq "System.LinkTypes.Hierarchy-Reverse") {
                    # Extract work item ID from URL
                    if ($relation.url -match '/(\d+)$') {
                        $parentId = [int]$matches[1]
                        break
                    }
                }
            }
        }
        
        if ($parentId) {
            # Check cache first
            if ($workItemCache.ContainsKey($parentId)) {
                $currentItem = $workItemCache[$parentId]
            }
            else {
                $currentItem = Get-WorkItemDetails -workItemId $parentId -headers $headers -organization $organization
                if ($currentItem) {
                    $workItemCache[$parentId] = $currentItem
                }
            }
        }
        else {
            $currentItem = $null
        }
    }
    
    return ($path -join "/")
}

$UserArray = $UserEmails -split "," -replace " ",""
"xxxxxxxxxxxxxxxxxxx"
$UserArray
Write-Host "PAT length: $($PersonalAccessToken.Length)"
Write-Host "PAT length: $(($env:PAT).Length)"

# Main execution
try {
    Write-Host "Starting Azure DevOps work items collection..." -ForegroundColor Green
    
    $headers = Get-AuthHeader -token $PersonalAccessToken
    $results = @()
    $workItemCache = @{}
    
    # Get all projects in the organization
    Write-Host "Getting projects..." -ForegroundColor Yellow
    $projectsUri = "https://dev.azure.com/$Organization/_apis/projects?api-version=7.0"
    $projects = Invoke-RestMethod -Uri $projectsUri -Headers $headers -Method Get
    
    # Process each user individually
    foreach ($userEmail in $UserArray) {
        Write-Host "`nProcessing user: $userEmail" -ForegroundColor Magenta
        
        foreach ($project in $projects.value) {
            $projectName = $project.name
            Write-Host "  Processing project: $projectName for user: $userEmail" -ForegroundColor Cyan
            
            # Build WIQL query for open work items assigned to current user
            $wiqlQuery = @"
SELECT [System.Id], [System.Title], [System.WorkItemType], [System.State], 
       [System.AssignedTo], [System.CreatedDate], [System.ChangedDate],
       [System.AreaPath], [System.IterationPath]
FROM workitems 
WHERE [System.TeamProject] = '$projectName' 
  AND [System.State] <> 'Closed' 
  AND [System.State] <> 'Done' 
  AND [System.State] <> 'Removed'
  AND [System.AssignedTo] = '$userEmail'
ORDER BY [System.Id]
"@
            
            # Execute WIQL query
            $wiqlUri = "https://dev.azure.com/$Organization/$projectName/_apis/wit/wiql?api-version=7.0"
            $wiqlBody = @{ query = $wiqlQuery } | ConvertTo-Json
            
            try {
                $queryResult = Invoke-RestMethod -Uri $wiqlUri -Headers $headers -Method Post -Body $wiqlBody -ContentType "application/json"
                
                if ($queryResult.workItems.Count -eq 0) {
                    Write-Host "    No work items found for user $userEmail in project $projectName" -ForegroundColor Gray
                    continue
                }
                
                Write-Host "    Found $($queryResult.workItems.Count) work items for user $userEmail" -ForegroundColor Green
                
                # Get detailed work item information
                $workItemIds = $queryResult.workItems.id -join ","
                $detailsUri = "https://dev.azure.com/$Organization/_apis/wit/workitems?ids=$workItemIds&`$expand=relations&api-version=7.0"
                $workItemDetails = Invoke-RestMethod -Uri $detailsUri -Headers $headers -Method Get
                
                foreach ($workItem in $workItemDetails.value) {
                    Write-Host "      Processing work item $($workItem.id)..." -ForegroundColor Gray
                    
                    # Get hierarchy path
                    $hierarchyPath = Get-WorkItemHierarchyPath -workItem $workItem -headers $headers -organization $Organization -workItemCache $workItemCache
                    
                    # Create result object
                    $link = "https://dev.azure.com/$Organization/$projectName/_workitems/edit/$($workItem.id)"

                    $resultItem = [PSCustomObject]@{
                        Id = "<a href='$link' target='_blank' rel='noopener' onclick='window.open(this.href, `"_blank`"); return false;'>$($workItem.id)</a>"
                        HierarchyPath = $hierarchyPath
                        Title = $workItem.fields.'System.Title'
                        WorkItemType = $workItem.fields.'System.WorkItemType'
                        State = $workItem.fields.'System.State'
                        AssignedTo = if ($workItem.fields.'System.AssignedTo') { $workItem.fields.'System.AssignedTo'.displayName } else { "Unassigned" }
                        IterationPath = $workItem.fields.'System.IterationPath'
                        AssignedToEmail = if ($workItem.fields.'System.AssignedTo') { $workItem.fields.'System.AssignedTo'.uniqueName } else { "" }
                        Project = $projectName
                        AreaPath = $workItem.fields.'System.AreaPath'                        
                        CreatedDate = $workItem.fields.'System.CreatedDate'
                        ChangedDate = $workItem.fields.'System.ChangedDate'
#                        Link = $link
                    }

                    $results += $resultItem
                }
            }
            catch {
                Write-Warning "##[error]Failed to query work items in project $projectName for user $userEmail : $($_.Exception.Message)"
            }
        }
        
        Write-Host "  Completed processing for user: $userEmail. Running total: $($results.Count) work items" -ForegroundColor Green
    }
    
    Write-Host "`nCollection completed. Found $($results.Count) total work items across all users." -ForegroundColor Green
    
    # Export results
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Display summary
        Write-Host "`nSummary by Work Item Type:" -ForegroundColor Yellow
        $results | Group-Object WorkItemType | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor White
        }
        
        Write-Host "`nSummary by Assigned User:" -ForegroundColor Yellow
        $results | Group-Object AssignedTo | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor White
        }
        
        Write-Host "`nSummary by User Email:" -ForegroundColor Yellow
        $results | Group-Object AssignedToEmail | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor White
        }
    }
    else {
        Write-Host "No work items found matching the criteria." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
}

Write-Host "`nScript completed. Results are available in `$result variable and exported to $OutputPath" -ForegroundColor Green