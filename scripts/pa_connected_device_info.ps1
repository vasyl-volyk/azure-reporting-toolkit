param (
    [string]$ApiKey = $env:PA_API_KEY,    
    [string]$OutputPath
)

# Panorama API endpoint and API key
$PanoramaHost = "panorama.yourcompany.com"

function Invoke-PanoramaOpCommand {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HostName,
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [Parameter(Mandatory = $true)]
        [string]$XmlCommand
    )

    $encodedCmd = [System.Uri]::EscapeDataString($XmlCommand)
    $uri = "https://$HostName/api/?type=op&cmd=$encodedCmd&key=$Key"

    # -SkipCertificateCheck is available in newer PowerShell versions.
    # Keep this explicit switch to support Panorama environments with self-signed certs.
    $raw = Invoke-WebRequest -Uri $uri -Method Get -UseBasicParsing #-SkipCertificateCheck -ErrorAction Stop
    return [xml]$raw.Content
}

try {
    if ($ApiKey -eq "REPLACE_WITH_YOUR_PANORAMA_API_KEY") {
        throw "Set your Panorama API key in the `$ApiKey variable before running this script."
    }

    $cmd = "<show><devices><connected></connected></devices></show>"
    $xml = Invoke-PanoramaOpCommand -HostName $PanoramaHost -Key $ApiKey -XmlCommand $cmd

    if ($xml.response.status -ne "success") {
        $errorText = $xml.response.msg.line
        if (-not $errorText) { $errorText = $xml.response.msg.error }
        throw "Panorama API returned status '$($xml.response.status)'. Error: $errorText"
    }

    # Panorama responses can vary by version; collect entries from known paths.
    $entries = @()
    if ($xml.response.result.devices.entry) { $entries += $xml.response.result.devices.entry }
    if ($xml.response.result.devices.connected.entry) { $entries += $xml.response.result.devices.connected.entry }

    if (-not $entries -or $entries.Count -eq 0) {
        Write-Host "No connected firewalls found in Panorama."
        $deviceInfo = @()
    }
    else {
        $deviceInfo = foreach ($d in $entries) {
            [PSCustomObject]@{
                Hostname          = $d.hostname
                SerialNumber      = $d.serial
                ManagementIP      = $d.'ip-address'
#                IPv6Address       = $d.'ipv6-address'
                Model             = $d.model
                SoftwareVersion   = $d.'sw-version'
                Connected         = $d.connected
                HAState           = $d.ha.state
#                SharedPolicySynced = $d.'shared-policy-status'
#                LastHeartbeat     = $d.'last-masterkey-push-status'
                Uptime            = $d.uptime
                DeviceCertificate = $d.'device-cert-present'
                CertificatEexpiry = $d.'certificate-expiry'
                AppVersion        = $d.'app-version'
                AVVersion         = $d.'av-version'
                ThreatVersion     = $d.'threat-version'
                WildFireVersion   = $d.'wildfire-version'
            }
        }

        # Remove duplicate devices if Panorama returns from multiple nodes/paths.
        $deviceInfo = $deviceInfo | Sort-Object SerialNumber -Unique

        Write-Host ("Found {0} connected firewall(s)." -f $deviceInfo.Count)
        $deviceInfo | Sort-Object -Property { $_.Hostname.Substring($_.Hostname.Length - 4)} | Format-Table -AutoSize
    }

    # Keep objects available for downstream scripting/pipeline use.
    $deviceInfo
    $deviceInfo | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
}
catch {
    Write-Error "Failed to read connected firewalls from Panorama. $($_.Exception.Message)"
}
