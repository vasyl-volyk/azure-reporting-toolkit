[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ApiKey = $env:PA_API_KEY,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$PanoramaHost = "panorama.yourcompany.com",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Query = "(eventid eq 'auth-fail')",

    [Parameter()]
    [datetime]$Since = (Get-Date).AddDays(-7),

    [Parameter()]
    [datetime]$Until,

    [Parameter()]
    [switch]$NoTimeFilter,

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$BatchSize = 5000,

    [Parameter()]
    [ValidateRange(1, 1000000)]
    [int]$MaxLogs = 50000,

    [Parameter()]
    [ValidateSet("backward", "forward")]
    [string]$Direction = "backward",

    [Parameter()]
    [switch]$SkipCertificateCheck,

    [Parameter()]
    [ValidateRange(1, 60)]
    [int]$PollIntervalSec = 2,

    [Parameter()]
    [ValidateRange(10, 3600)]
    [int]$JobTimeoutSec = 180
)



Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Format-PanReceiveTime {
    param([Parameter(Mandatory)][datetime]$Date)
    # PAN-OS examples use: yyyy/MM/dd HH:mm:ss
    return $Date.ToString("yyyy/MM/dd HH:mm:ss")
}

function New-PanoramaApiUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][hashtable]$Parameters
    )

    $pairs = foreach ($k in $Parameters.Keys) {
        $v = [string]$Parameters[$k]
        "{0}={1}" -f $k, ([System.Uri]::EscapeDataString($v))
    }

    "https://{0}/api/?{1}" -f $HostName, ($pairs -join "&")
}

function Invoke-PanoramaApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter()][ValidateSet("GET","POST")][string]$Method = "GET",
        [Parameter()][switch]$SkipCertCheck
    )

    # TLS 1.2 for older Windows PowerShell
    try {
        if ($PSVersionTable.PSVersion.Major -lt 6) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
    } catch { }

    $oldCallback = $null
    $usedLegacyBypass = $false

    try {
        $iwrSplat = @{
            Uri    = $Uri
            Method = $Method
        }

        # PowerShell 5 uses -UseBasicParsing; PowerShell 6+ removed it.
        if ($PSVersionTable.PSVersion.Major -lt 6) {
            $iwrSplat["UseBasicParsing"] = $true
        }

        if ($SkipCertCheck) {
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $iwrSplat["SkipCertificateCheck"] = $true
            }
            elseif ($PSVersionTable.PSVersion.Major -lt 6) {
                # Legacy bypass for Windows PowerShell (use with caution)
                $oldCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
                [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                $usedLegacyBypass = $true
            }
        }

        $resp = Invoke-WebRequest @iwrSplat
        return [xml]$resp.Content
    }
    finally {
        if ($usedLegacyBypass -and $PSVersionTable.PSVersion.Major -lt 6) {
            [Net.ServicePointManager]::ServerCertificateValidationCallback = $oldCallback
        }
    }
}

function Assert-PanSuccess {
    param([Parameter(Mandatory)][xml]$Xml, [string]$Context = "API call")

    $status = $Xml.response.status
    if ($status -ne "success") {
        $msg = $Xml.response.msg.line
        if (-not $msg) { $msg = $Xml.response.msg }
        if (-not $msg) { $msg = $Xml.OuterXml }
        throw "$Context failed. Panorama returned status '$status'. Message: $msg"
    }
}

function Start-PanoramaSystemLogJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$LogQuery,
        [Parameter(Mandatory)][int]$NLogs,
        [Parameter(Mandatory)][int]$Skip,
        [Parameter(Mandatory)][string]$Dir,
        [Parameter()][switch]$SkipCertCheck
    )

    $uri = New-PanoramaApiUri -HostName $HostName -Parameters @{
        "type"     = "log"
        "log-type" = "system"
        "query"    = $LogQuery
        "nlogs"    = "$NLogs"
        "skip"     = "$Skip"
        "dir"      = $Dir
        "key"      = $Key
    }

    $xml = Invoke-PanoramaApiRequest -Uri $uri -Method "GET" -SkipCertCheck:$SkipCertCheck
    Assert-PanSuccess -Xml $xml -Context "Start log job"

    $jobId = $xml.response.result.job
    if (-not $jobId) {
        throw "Start log job succeeded but no job-id was returned. Raw response: $($xml.OuterXml)"
    }

    [PSCustomObject]@{
        JobId = [int]$jobId
        Raw   = $xml
    }
}

function Get-PanoramaLogJobResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][int]$JobId,
        [Parameter(Mandatory)][int]$TimeoutSec,
        [Parameter(Mandatory)][int]$PollSec,
        [Parameter()][switch]$SkipCertCheck
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSec)

    while ($true) {
        $uri = New-PanoramaApiUri -HostName $HostName -Parameters @{
            "type"   = "log"
            "action" = "get"
            "job-id" = "$JobId"
            "key"    = $Key
        }

        $xml = Invoke-PanoramaApiRequest -Uri $uri -Method "GET" -SkipCertCheck:$SkipCertCheck
        Assert-PanSuccess -Xml $xml -Context "Get log job result (job-id=$JobId)"

        $status = $xml.response.result.job.status
        if ($status -eq "FIN") {
            return $xml
        }

        if ((Get-Date) -ge $deadline) {
            throw "Timed out waiting for log job $JobId to finish. Last known status: '$status'"
        }

        Start-Sleep -Seconds $PollSec
    }
}

function Convert-PanLogEntriesToObjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Entries
    )

    foreach ($entry in $Entries) {
        # Flatten all child elements into a dictionary
        $ht = [ordered]@{}
        foreach ($node in $entry.ChildNodes) {
            if ($node.NodeType -ne "Element") { continue }
            $name = $node.Name
            $value = $node.InnerText
            if (-not $ht.Contains($name)) { $ht[$name] = $value }
        }

        # Standardize a few common fields when present
        $message = $null
        if ($ht.Contains("opaque") -and $ht["opaque"]) { $message = $ht["opaque"] }
        elseif ($ht.Contains("description") -and $ht["description"]) { $message = $ht["description"] }

        # Parse common patterns from auth-fail messages (best-effort)
        $parsedUser  = $null
        $parsedIp    = $null
        $parsedReason = $null

        if ($message) {
            if ($message -match "user\s+'([^']+)'" ) { $parsedUser = $Matches[1] }
            elseif ($message -match 'username\s+"([^"]+)"') { $parsedUser = $Matches[1] }

            if ($message -match "From:\s*([0-9]{1,3}(\.[0-9]{1,3}){3})") { $parsedIp = $Matches[1] }

            if ($message -match "Reason:\s*([^\.]+)\.") { $parsedReason = $Matches[1].Trim() }
            elseif ($message -match "reply message\s+'([^']+)'" ) { $parsedReason = $Matches[1].Trim() }
        }

        # Emit a stable CSV schema (include flattened + parsed + message)
        [PSCustomObject]([ordered]@{
            receive_time     = $ht["receive_time"]
#            time_generated   = $ht["time_generated"]
            device_name      = $ht["device_name"]
#            serial           = $ht["serial"]
            type             = $ht["type"]
#            subtype          = $ht["subtype"]
            eventid          = $ht["eventid"]
#            severity         = $ht["severity"]
#            module           = $ht["module"]
#            vsys_id          = $ht["vsys_id"]
#            admin            = $ht["admin"]
            parsed_user      = $parsedUser
            parsed_source_ip = $parsedIp
            object           = $ht["object"]            
#            parsed_reason    = $parsedReason
            message          = $message

        })
    }
}

try {
    if (-not $ApiKey) { throw "ApiKey is empty. Set -ApiKey or env:PA_API_KEY." }

    if (-not $OutputPath) {
        $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $OutputPath = Join-Path -Path (Get-Location) -ChildPath "panorama_failed_auth_${stamp}.csv"
    }

    # Build final log query
    $finalQuery = $Query

    if (-not $NoTimeFilter) {
        $clauses = @()

        if ($Since) {
            $clauses += "(receive_time geq '$(Format-PanReceiveTime $Since)')"
        }
        if ($Until) {
            $clauses += "(receive_time leq '$(Format-PanReceiveTime $Until)')"
        }

        if ($clauses.Count -gt 0) {
            # Combine: (baseQuery) and (timeClause1) and (timeClause2)
            $timeExpr = ($clauses -join " and ")
            $finalQuery = "($finalQuery) and $timeExpr"
        }
    }

    Write-Host "Panorama host : $PanoramaHost"
    Write-Host "Output CSV    : $OutputPath"
    Write-Host "Query         : $finalQuery"
    Write-Host "BatchSize     : $BatchSize, MaxLogs: $MaxLogs, Direction: $Direction"
    Write-Host ""

    $all = New-Object System.Collections.Generic.List[object]
    $skip = 0

    while ($all.Count -lt $MaxLogs) {
        $remaining = $MaxLogs - $all.Count
        $thisBatch = [Math]::Min($BatchSize, $remaining)

        Write-Host ("Starting log job (skip={0}, nlogs={1})..." -f $skip, $thisBatch)

        $job = Start-PanoramaSystemLogJob -HostName $PanoramaHost -Key $ApiKey -LogQuery $finalQuery `
                                         -NLogs $thisBatch -Skip $skip -Dir $Direction -SkipCertCheck:$SkipCertificateCheck

        $xml = Get-PanoramaLogJobResult -HostName $PanoramaHost -Key $ApiKey -JobId $job.JobId `
                                       -TimeoutSec $JobTimeoutSec -PollSec $PollIntervalSec -SkipCertCheck:$SkipCertificateCheck

        # Extract entries (handle variation in response shape)
        $entries = @()
        if ($xml.response.result.log.logs.entry) { $entries = @($xml.response.result.log.logs.entry) }
        elseif ($xml.response.result.log.entry)  { $entries = @($xml.response.result.log.entry) }

        $countAttr = $xml.response.result.log.logs.count
        if (-not $countAttr) { $countAttr = $entries.Count }

        Write-Host ("Job {0} finished. Received {1} entries." -f $job.JobId, $entries.Count)

        if (-not $entries -or $entries.Count -eq 0) { break }

        $objs = Convert-PanLogEntriesToObjects -Entries $entries
        foreach ($o in $objs) { $null = $all.Add($o) }

        # If fewer than requested returned, we reached the end
        if ($entries.Count -lt $thisBatch) { break }

        $skip += $thisBatch
    }

    Write-Host ""
    Write-Host ("Total collected entries: {0}" -f $all.Count)

    # Export
    $all | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host ("Export completed: {0}" -f $OutputPath)

    # Output objects for pipeline use
    $all
}
catch {
    Write-Error ("Failed to export Panorama failed authentication logs. {0}" -f $_.Exception.Message)
    throw
}