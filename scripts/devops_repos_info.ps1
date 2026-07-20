# Azure DevOps Repository Documentation Generator
# This script collects information about repositories across Azure DevOps projects
# and generates AI-powered summaries using Azure OpenAI


param(
    #[Parameter(Mandatory = $true)]
    [string]$Organization,
    
    #[Parameter(Mandatory = $true)]
    [string]$PersonalAccessToken = $env:PAT,
    
    #[Parameter(Mandatory = $true)]
    [string]$AzureOpenAIEndpoint,
    
    #[Parameter(Mandatory = $true)]
    [string]$AzureOpenAIKey = $env:AzureOpenAIKey,
    
    #[Parameter(Mandatory = $true)]
    [string]$DeploymentName,
    
    #[Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    #[Parameter(Mandatory = $false)]
    [array]$ProjectFilter = @(),
    
    #[Parameter(Mandatory = $false)]
    [int]$MaxFileSizeKB = 100
)


    $Organization= "YourOrg"
    
    
    $AzureOpenAIEndpoint= "https://your-openai-resource.openai.azure.com/"
    
    
    $DeploymentName= "gpt-4.1-nano"
    
    $OutputFile = "Repository-Documentation.md"
    
    $ProjectFilter = @()



# Azure DevOps PowerShell Files Collector - Version 2
# This script collects information about PowerShell files across all repositories in an organization


# Function to create authentication header
function Get-AuthHeader {
    param([string]$token)
    
    $encodedPAT = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$token"))
    return @{
        Authorization = "Basic $encodedPAT"
        'Content-Type' = 'application/json'
    }
}

# Function to insert <br> every 18 words using regex
function Insert-LineBreaks {
    param([string]$summary)
    
    # Check if less than 18 words, return unchanged
    if (($summary -split '\s+').Count -lt 18) {
        return $summary
    }
    
    # Regex pattern to match 18 words followed by whitespace
    # (\S+\s+) matches a word followed by whitespace
    # {17} matches exactly 17 occurrences of the above
    # (\S+) matches the 18th word without trailing space
    $pattern = '((?:\S+\s+){17}\S+)'
    
    # Replace with captured group + <br>
    $result = $summary -replace $pattern, '$1<br>'
    
    # Clean up any trailing <br> and extra spaces
    #$result = $result -replace '<br>\s*, '' -replace '\s+<br>', '<br>'
    
    return $result
}

# Initialize results array
$results = @()

# Create authentication header
$headers = Get-AuthHeader -token $PersonalAccessToken

Write-Host "Starting collection for organization: $Organization" -ForegroundColor Green

try {
    # Get all projects in the organization
    Write-Host "Fetching projects..." -ForegroundColor Yellow
    $projectsUri = "https://dev.azure.com/$Organization/_apis/projects?api-version=7.0"
    $projects = Invoke-RestMethod -Uri $projectsUri -Headers $headers -Method Get
    
    Write-Host "Found $($projects.count) projects" -ForegroundColor Cyan
    
    foreach ($project in $projects.value) {
        Write-Host "Processing project: $($project.name)" -ForegroundColor Yellow
        
        # Get all repositories in the project
        $reposUri = "https://dev.azure.com/$Organization/$($project.id)/_apis/git/repositories?api-version=7.0"
        $repos = Invoke-RestMethod -Uri $reposUri -Headers $headers -Method Get
        
        Write-Host "  Found $($repos.count) repositories in project $($project.name)" -ForegroundColor Cyan
        
        foreach ($repo in $repos.value) {
            Write-Host "    Processing repository: $($repo.name)" -ForegroundColor Gray
            
            try {
                # Get repository items (files and folders) recursively
                $itemsUri = "https://dev.azure.com/$Organization/$($project.id)/_apis/git/repositories/$($repo.id)/items?recursionLevel=full&api-version=7.0"
                $items = Invoke-RestMethod -Uri $itemsUri -Headers $headers -Method Get
                
                # Filter for PowerShell files (.ps1, .psm1, .psd1)
                $psFiles = $items.value | Where-Object { 
                    $_.gitObjectType -eq "blob" -and 
                    ($_.path -match "\.ps1$|\.psm1$|\.psd1$")
                }
                
                Write-Host "      Found $($psFiles.Count) PowerShell files" -ForegroundColor DarkCyan
                
                foreach ($file in $psFiles) {
                    Write-Host "        Processing file: $($file.path)" -ForegroundColor DarkGray
                    
                    try {
                        # Improve path encoding
                        $encodedPath = $file.path.TrimStart('/')  # Remove leading slash
                        $encodedPath = [System.Web.HttpUtility]::UrlPathEncode($encodedPath)  # Better URL encoding
                        
                        Write-Verbose "Encoded path: $encodedPath"  # Add this for debugging
                        
                        $fileContentUri = "https://dev.azure.com/$Organization/$($project.id)/_apis/git/repositories/$($repo.id)/items?path=$encodedPath&api-version=7.0"
                        
                        # Add error details for debugging
                        try {
                            $fileContent = Invoke-RestMethod -Uri $fileContentUri -Headers $headers -Method Get -ErrorVariable responseErr
                        }
                        catch {
                            Write-Warning "API Error Details: $($responseErr | ConvertTo-Json)"
                            Write-Warning "Request URL: $fileContentUri"
                            throw
                        }
                        
                        # Get summary from Azure OpenAI ----------------------------------------------
                        $body = @{
                            messages = @(
                                @{
                                    role = "system"
                                    content = "You are a PowerShell expert. Provide a brief, one-line summary of what this PowerShell script does. Do not start the summary with 'This PowerShell script' or similar phrases."
                                },
                                @{
                                    role = "user"
                                    content = $fileContent
                                }
                            )
                            max_tokens = 100
                            temperature = 0.3
                        } | ConvertTo-Json

                        $openAiHeaders = @{
                            'api-key' = $AzureOpenAIKey
                            'Content-Type' = 'application/json'
                        }

                        $openAiUri = "$AzureOpenAIEndpoint/openai/deployments/$DeploymentName/chat/completions?api-version=2023-05-15"
                        
                        $aiResponse = Invoke-RestMethod -Uri $openAiUri -Headers $openAiHeaders -Method Post -Body $body
                        $summary = $aiResponse.choices[0].message.content
                        
                        # Extract just the filename from the full path
                        $fileName = Split-Path $file.path -Leaf
                        
                        $summary = Insert-LineBreaks -summary $summary
                        # ---------------------------------------------------------------------------------

                        # Get Powershel 7 compatibility from Azure OpenAI ----------------------------------------------
                        $body = @{
                            messages = @(
                                @{
                                    role = "system"
                                    content = "You are a PowerShell expert. ckeck this script for compatibility with Powershell 7. Will it work in Powershell 7 or not. Give me one word answer YES or NO"
                                },
                                @{
                                    role = "user"
                                    content = $fileContent
                                }
                            )
                            max_tokens = 100
                            temperature = 0.1
                        } | ConvertTo-Json

                        $openAiHeaders = @{
                            'api-key' = $AzureOpenAIKey
                            'Content-Type' = 'application/json'
                        }

                        $openAiUri = "$AzureOpenAIEndpoint/openai/deployments/$DeploymentName/chat/completions?api-version=2023-05-15"
                        
                        $aiResponse = Invoke-RestMethod -Uri $openAiUri -Headers $openAiHeaders -Method Post -Body $body
                        $summarycompatibility = $aiResponse.choices[0].message.content
                        
                        # ---------------------------------------------------------------------------------

                        # Create custom object with required properties
                        $fileInfo = [PSCustomObject]@{
                            ProjectName = $project.name
                            RepositoryName = $repo.name
                            FullPath = $file.path
#                            FileName = $fileName
                            Summary = $summary
                            PS7_readiness = $summarycompatibility
                        }
                        
                        $results += $fileInfo
                        
                    } catch {
                        Write-Warning "Failed to process file $($file.path): $($_.Exception.Message)"
                        
                        # Add entry with error message if processing fails
                        $fileName = Split-Path $file.path -Leaf
                        $fileInfo = [PSCustomObject]@{
                            ProjectName = $project.name
                            RepositoryName = $repo.name
                            FullPath = $file.path
 #                           FileName = $fileName
                            Summary = "ERROR: Could not process file content"
                            PS7_readiness = "n/a"
                        }
                        $results += $fileInfo
                    }
                }
            } catch {
                Write-Warning "Failed to get items for repository $($repo.name): $($_.Exception.Message)"
            }
        }
    }
    
    Write-Host "`nCollection completed successfully!" -ForegroundColor Green
    Write-Host "Total PowerShell files found: $($results.Count)" -ForegroundColor Green
    
    # Display results summary
    Write-Host "`nResults Summary:" -ForegroundColor Yellow
    $results | Group-Object ProjectName | ForEach-Object {
        Write-Host "  Project '$($_.Name)': $($_.Count) files" -ForegroundColor Cyan
    }

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
    }
    else {
        Write-Host "No work items found matching the criteria." -ForegroundColor Yellow
    }
    
} catch {
    Write-Error "Failed to execute script: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
}

