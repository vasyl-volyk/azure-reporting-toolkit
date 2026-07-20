param (
    [string]$ConfigPath = "scripts/config.json",
    [string]$ContainerName = $env:AZURE_STORAGE_CONTAINER,
    [string]$StorageAccount = $env:AZURE_STORAGE_ACCOUNT,
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID
)

# Log in to Azure using the Service Principal
Write-Host "Logging in to Azure using Service Principal..."

az login --service-principal --username $ClientId --password $ClientSecret --tenant $TenantId | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Error "##error Azure login failed."
    exit 1
}

# Initialize timings collection
$timings = [System.Collections.Generic.List[PSCustomObject]]::new()

# Read the job configuration from JSON
$json = Get-Content $ConfigPath | ConvertFrom-Json

foreach ($job in $json) {
    Write-Host "##[group] Running $($job.name)..."
    Write-Host "========================================================"
    Write-Host "Running $($job.name)..."
    $scriptPath = $job.script
    $params = $job.parameters
    $dateString = (Get-Date -Format 'yyyy-MM-dd')
    $outputFileName = $job.outputName -replace '\$\(\w+\)', $dateString
    $outputFile = Join-Path -Path $env:TEMP -ChildPath $outputFileName

    if (-not $outputFile) {
        Write-Error "Output path is empty for $($job.name)"
        continue
    }

    # Build argument list for the data-generating script
    $arguments = @{}

    foreach ($prop in $params.PSObject.Properties) {
        $arguments[$prop.Name] = $prop.Value
    }

    $arguments["OutputPath"] = $outputFile

    # Execute the data-generating script and measure time
    $scriptStart = Get-Date
    & $scriptPath @arguments
    $scriptEnd = Get-Date
    $scriptStatus = if ($LASTEXITCODE -ne 0) { "Failed" } else { "Success" }
    $scriptDuration = ($scriptEnd - $scriptStart).TotalSeconds

    # Record timing entry
    $timings.Add([PSCustomObject]@{
        ScriptName = $job.script
#        StartTime  = $scriptStart.ToString("yyyy-MM-dd HH:mm:ss")
#        EndTime    = $scriptEnd.ToString("yyyy-MM-dd HH:mm:ss")
        Duration   = [TimeSpan]::FromSeconds($scriptDuration).ToString("hh\:mm\:ss")
        Status     = $scriptStatus
    })

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "##error Script $($job.name) failed."
        continue
    }

    # Define blob path with nested folders and date substitution
    $blobPath = "$($job.targetFolder)/$outputFileName"

    # Upload the CSV file to Azure Blob Storage
    Write-Host "Uploading $outputFileName to blob storage at $blobPath ..."

    az storage blob upload `
        --account-name $StorageAccount `
        --container-name $ContainerName `
        --name $blobPath `
        --file $outputFile `
        --auth-mode login `
        --overwrite true | Out-Null

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "##error Upload of $outputFileName failed."
        continue
    }

    Write-Host "Upload completed."
    Write-Host "##[endgroup]"
}

# Export timings to CSV and upload to blob storage root
Write-Host "Saving and uploading script timings..."
$timingsCsvPath = Join-Path -Path $env:TEMP -ChildPath "scriptstimings.csv"
$timings | Export-Csv -Path $timingsCsvPath -NoTypeInformation -Encoding UTF8

az storage blob upload `
    --account-name $StorageAccount `
    --container-name $ContainerName `
    --name "scriptstimings.csv" `
    --file $timingsCsvPath `
    --auth-mode login `
    --overwrite true | Out-Null

if ($LASTEXITCODE -ne 0) {
    Write-Warning "##error Upload of scriptstimings.csv failed."
} else {
    Write-Host "Script timings uploaded to scriptstimings.csv"
}

# (Optional) Log out from Azure
az logout | Out-Null