param (
    [string]$APIKey = $env:PA_API_KEY,    
    [string]$OutputPath,
    [int]$MaxLogs = 500
)



$PanoramaIP = "panorama.yourcompany.com" # e.g., "192.168.1.100" or "panorama.example.com"


# Step 1: Submit the log query (returns job ID)
$LogQueryUrl = "https://$PanoramaIP/api/?type=log&log-type=config&nlogs=$MaxLogs&key=$APIKey"

try {
    Write-Host "Submitting config log query to Panorama..."
    $InitialResponse = Invoke-RestMethod -Uri $LogQueryUrl -Method Get -ContentType "application/xml"

    $JobID = $InitialResponse.response.result.job
    if (-not $JobID) {
        throw "No job ID returned. Response: $($InitialResponse | Out-String)"
    }

    Write-Host "Log query job enqueued. Job ID: $JobID"

    # Step 2: Poll for job completion
    $JobStatusUrl = "https://$PanoramaIP/api/?type=log&action=get&job-id=$JobID&key=$APIKey"
    $MaxTries = 20
    $DelaySec = 3
    for ($i = 0; $i -lt $MaxTries; $i++) {
        Start-Sleep -Seconds $DelaySec
        $StatusResponse = Invoke-RestMethod -Uri $JobStatusUrl -Method Get -ContentType "application/xml"
        $Status = $StatusResponse.response.result.job.status

        Write-Host "Job status: $Status"
        if ($Status -eq "FIN") {
            # Step 3: Process final result
            $Logs = $StatusResponse.response.result.log.logs.entry
            Write-Host "Log job finished. Parsing $($Logs.Count) entries..."
            $Logs | ForEach-Object {
                [PSCustomObject]@{
                    TimeGenerated = $_.time_generated
                    Admin         = $_.admin
                    Command       = $_.cmd
                    ConfigPath    = $_.path
#                    Description   = $_.description
                    IPAddress     = $_.client
                    Result        = $_.result
                }
            } | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            return
        }
    }

    throw "Job did not finish in expected time."

} catch {
    Write-Error "Error during log query: $_"
}
