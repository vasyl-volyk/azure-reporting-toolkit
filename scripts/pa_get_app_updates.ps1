# Palo Alto Panorama - Get Dynamic Updates Applications (Panorama Only)
# API Key and URL configuration

param (
    [string]$APIKey = $env:PA_API_KEY,    
    [string]$OutputPath
)

$baseUrl = "https://panorama.yourcompany.com/api/"

# The specific command URL-encoded
$contentInfoCmd = "%3Crequest%3E%3Ccontent%3E%3Cupgrade%3E%3Cinfo%3E%3C/info%3E%3C/upgrade%3E%3C/content%3E%3C/request%3E"
$apiUrl = "$baseUrl" + "?type=op&cmd=$contentInfoCmd&key=$APIkey"

# Headers for API requests
$headers = @{
    'Content-Type' = 'application/x-www-form-urlencoded'
}

# Initialize results array
$results = @()

# Function to parse content update entries
function Parse-ContentUpdateEntry {
    param([System.Xml.XmlElement]$Entry, [string]$LastUpdatedAt, [string]$DeviceName = "Unknown")
    
    $entryResult = [PSCustomObject]@{
        # Device Information
        DeviceName = $DeviceName
        LastUpdatedAt = $LastUpdatedAt
        CollectionTime = (Get-Date).ToString()
        
        # Version Information
        Version = $Entry.version
        AppVersion = $Entry.'app-version'
        Filename = $Entry.filename
        
        # Size Information
        SizeMB = $Entry.size
        SizeKB = $Entry.'size-kb'
        
        # Release Information
        ReleasedOn = $Entry.'released-on'
        ReleaseNotes = "<a href=`"$($Entry.'release-notes'.'#cdata-section')`">Applications Content Release Notes</a>"
        
        # Status Information
        Downloaded = ($Entry.downloaded -eq "yes")
        Current = ($Entry.current -eq "yes")
        Previous = ($Entry.previous -eq "yes")
        Installing = ($Entry.installing -eq "yes")
        
        # Content Information
        Features = $Entry.features
        UpdateType = $Entry.'update-type'
        FeatureDesc = $Entry.'feature-desc'
        
        # Security
        SHA256 = $Entry.sha256
        
        # Derived Information
        IsLatest = ($Entry.current -eq "yes")
        IsAvailable = ($Entry.downloaded -eq "yes")
        Status = if ($Entry.current -eq "yes") { "Current" } 
                elseif ($Entry.previous -eq "yes") { "Previous" }
                elseif ($Entry.installing -eq "yes") { "Installing" }
                elseif ($Entry.downloaded -eq "yes") { "Downloaded" }
                else { "Available" }
    }
    
    return $entryResult
}

# Function to get device name from system info
function Get-DeviceName {
    try {
        $systemInfoUrl = "$baseUrl" + "?type=op&cmd=<show><s><info></info></s></show>&key=$APIkey"
        $systemResponse = Invoke-RestMethod -Method Get -Uri $systemInfoUrl -Headers $headers
        
        if ($systemResponse.response.result.system.hostname) {
            return $systemResponse.response.result.system.hostname
        }
        return "Unknown-Device"
    }
    catch {
        Write-Host "Warning: Could not retrieve device name: $($_.Exception.Message)" -ForegroundColor Yellow
        return "Unknown-Device"
    }
}

try {
    Write-Host "=== Palo Alto Firewall Content Updates Reader ===" -ForegroundColor Magenta
    Write-Host "Connecting to firewall API..." -ForegroundColor Green
    Write-Host "API URL: $apiUrl" -ForegroundColor Gray
    
    # Get device name for context
    Write-Host "Retrieving device information..." -ForegroundColor Yellow
    $deviceName = Get-DeviceName
    Write-Host "Device Name: $deviceName" -ForegroundColor White
    
    # Execute the content info command
    Write-Host "Fetching content update information..." -ForegroundColor Yellow
    $response = Invoke-RestMethod -Method Get -Uri $apiUrl -Headers $headers
    
    # Check if the response is successful
    if ($response.response.status -ne "success") {
        Write-Host "API call failed with status: $($response.response.status)" -ForegroundColor Red
        if ($response.response.msg) {
            Write-Host "Error message: $($response.response.msg)" -ForegroundColor Red
        }
        exit 1
    }
    
    Write-Host "Successfully retrieved content update information" -ForegroundColor Green
    
    # Extract the content-updates information
    $contentUpdates = $response.response.result.'content-updates'
    $lastUpdatedAt = $contentUpdates.'last-updated-at'
    
    Write-Host "Content Updates Last Updated: $lastUpdatedAt" -ForegroundColor Green
    
    # Process each entry
    $entries = $contentUpdates.entry
    if ($entries -is [System.Xml.XmlElement]) {
        $entries = @($entries)  # Convert single entry to array
    }
    
    if ($entries -and $entries.Count -gt 0) {
        Write-Host "Processing $($entries.Count) content update entries..." -ForegroundColor Yellow
        
        foreach ($entry in $entries) {
            $parsedEntry = Parse-ContentUpdateEntry -Entry $entry -LastUpdatedAt $lastUpdatedAt -DeviceName $deviceName
            $results += $parsedEntry
            
            # Display entry summary
            $statusColor = switch ($parsedEntry.Status) {
                "Current" { "Green" }
                "Previous" { "Yellow" }
                "Installing" { "Cyan" }
                "Downloaded" { "Blue" }
                default { "Gray" }
            }
            
            Write-Host "  Entry: $($parsedEntry.Version) - Status: $($parsedEntry.Status) - Released: $($parsedEntry.ReleasedOn)" -ForegroundColor $statusColor
        }
    }
    else {
        Write-Host "No content update entries found in the response" -ForegroundColor Yellow
        
        # Create a summary entry even if no specific entries found
        $summaryEntry = [PSCustomObject]@{
            DeviceName = $deviceName
            LastUpdatedAt = $lastUpdatedAt
            CollectionTime = (Get-Date).ToString()
            Version = "NO_ENTRIES_FOUND"
            Status = "No Updates Available"
            Message = "No content update entries returned from API"
        }
        $results += $summaryEntry
    }
    
    # Display summary
    Write-Host "`n=== COLLECTION SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Device: $deviceName" -ForegroundColor White
    Write-Host "Total Entries Collected: $($results.Count)" -ForegroundColor Green
    Write-Host "Last Updated: $lastUpdatedAt" -ForegroundColor Green
    
    # Show current and available versions
    $currentVersion = $results | Where-Object { $_.Current -eq $true } | Select-Object -First 1
    $latestVersion = $results | Sort-Object ReleasedOn -Descending | Select-Object -First 1
    $downloadedVersions = $results | Where-Object { $_.Downloaded -eq $true }
    
    if ($currentVersion) {
        Write-Host "Current Version: $($currentVersion.Version) (Released: $($currentVersion.ReleasedOn))" -ForegroundColor Green
    }
    
    if ($latestVersion) {
        Write-Host "Latest Available Version: $($latestVersion.Version) (Released: $($latestVersion.ReleasedOn))" -ForegroundColor Yellow
    }
    
    Write-Host "Downloaded Versions: $($downloadedVersions.Count)" -ForegroundColor Blue
    
    # Show status breakdown
    Write-Host "`nStatus Breakdown:" -ForegroundColor White
    $statusGroups = $results | Group-Object Status
    foreach ($group in $statusGroups) {
        Write-Host "  $($group.Name): $($group.Count)" -ForegroundColor Gray
    }
    
    # Display sample entries
    if ($results.Count -gt 0 -and $results[0].Version -ne "NO_ENTRIES_FOUND") {
        Write-Host "`n=== SAMPLE ENTRIES ===" -ForegroundColor Cyan
        $results | Sort-Object ReleasedOn -Descending | Select-Object -First 3 | ForEach-Object {
            Write-Host "Version: $($_.Version)" -ForegroundColor Yellow
            Write-Host "  Status: $($_.Status) | Size: $($_.SizeMB) MB | Released: $($_.ReleasedOn)" -ForegroundColor Gray
            Write-Host "  Features: $($_.Features) | Type: $($_.UpdateType)" -ForegroundColor Gray
            if ($_.ReleaseNotes -and $_.ReleaseNotes.Trim() -ne "") {
                $notesUrl = $_.ReleaseNotes.Trim()
                if ($notesUrl.StartsWith("https://")) {
                    Write-Host "  Release Notes: Available at URL" -ForegroundColor Gray
                }
            }
            Write-Host ""
        }
    }
    
    # Export results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    $results = $results | Select-Object Filename,ReleasedOn,CollectionTime,LastUpdatedAt,Status,ReleaseNotes

    # Export to CSV
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to CSV: FirewallContentUpdates_$timestamp.csv" -ForegroundColor Green
    }
    
    Write-Host "`n=== RESULTS VARIABLE STATUS ===" -ForegroundColor Magenta
    Write-Host "`$results array populated with $($results.Count) entries" -ForegroundColor Green
    Write-Host "Each entry contains complete content update information from: $deviceName" -ForegroundColor Green
    
    # Display results structure
    if ($results.Count -gt 0) {
        Write-Host "`nSample `$results[0] properties:" -ForegroundColor White
        $results[0].PSObject.Properties.Name | Sort-Object | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Gray
        }
    }

} catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details:" -ForegroundColor Red
    Write-Host $_.Exception -ForegroundColor Red
    
    # Create error entry in results
    $errorEntry = [PSCustomObject]@{
        DeviceName = "Unknown"
        CollectionTime = (Get-Date).ToString()
        Version = "ERROR"
        Status = "Collection Failed"
        ErrorMessage = $_.Exception.Message
    }
    $results += $errorEntry
}

Write-Host "`nScript completed. Check `$results variable for collected data." -ForegroundColor Green

# Display final results count
Write-Host "Final `$results count: $($results.Count)" -ForegroundColor Magenta