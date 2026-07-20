param (
    [string]$APIKey = $env:PA_API_KEY,    
    [string]$OutputPath
)


$PanoramaIP = "panorama.yourcompany.com" # e.g., "192.168.1.100" or "panorama.example.com"

$DeviceGroupNames = "Mobile_User_Device_Group","Developer Sites" # The name of the device group you want to query

$AllSecurityRules = @()

# Ensure TLS 1.2 is used for Invoke-RestMethod for secure connections
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

foreach($DeviceGroupName in $DeviceGroupNames){ # This is the correct loop variable
    Write-Host "Fetching security rules for Device Group: $($DeviceGroupName)..." # Use $DeviceGroupName here

    # CONSTRUCT XPATH AND URL INSIDE THE LOOP for each device group
    # The device group name inside the XPath needs to be URL-encoded.
    $EncodedDeviceGroupName = [uri]::EscapeDataString($DeviceGroupName)
    $XPath = "/config/devices/entry/device-group/entry[@name='$EncodedDeviceGroupName']/pre-rulebase/security/rules"

    $URL = "https://$PanoramaIP/api/?type=config&action=get&xpath=$XPath&key=$APIKey"

    try {
        # Make the API call
        write-host "Calling URL: $URL" # Added more descriptive logging
        $Response = Invoke-RestMethod -Uri $URL -Method Get -ContentType "application/xml"

        if ($Response.response.status -eq "success") {
            Write-Host "Successfully retrieved security rules for $($DeviceGroupName)."

            # Parse the XML response
            $SecurityRules = $Response.response.result.rules.entry
        
            $ruleNo = 1
            if ($SecurityRules) {
                Write-Host "`nSecurity Rules in Device Group '$($DeviceGroupName)':" # Use $DeviceGroupName here
                foreach ($rule in $SecurityRules) {
                    Write-Host "$ruleNo, " -NoNewline
                    $AllSecurityRules +=
                    [PSCustomObject]@{
                        Location           = $DeviceGroupName # Use $DeviceGroupName here
                        No                 = $ruleNo
                        Name               = $($rule.name)
                        SourceZone         = $rule.'from'.member
                        DestinationZone    = $rule.'to'.member
                        SourceAddress      = $($rule.source.member -join '<br>')
                        DestinationAddress = $($rule.destination.member -join '<br>')
                        Application        = $($rule.application.member -join '<br>')
                        Service            = $($rule.service.member -join '<br>')
                        Action             = $($rule.action)
                        Description        = $($rule.description)
                        LogSetting         = $($rule.'log-setting')
                    }
                    $ruleNo++
                }
                Write-Host "`n--------------------------------------------------"
            } else {
                Write-Warning "No security rules found for Device Group '$($DeviceGroupName)' or XPath was incorrect." # Use $DeviceGroupName here
            }
        } else {
            Write-Error "API call failed for device group '$($DeviceGroupName)'. Status: $($Response.response.status), Code: $($Response.response.code), Message: $($Response.response.msg)" # Use $DeviceGroupName here
            if ($Response.response.msg) {
                $ErrorDetail = $Response.response.msg.error
                if ($ErrorDetail) {
                    Write-Error "Detailed Error: $($ErrorDetail)"
                }
            }
        }
    }
    catch {
        Write-Error "An error occurred during the API call for device group '$($DeviceGroupName)': $($_.Exception.Message)" # Use $DeviceGroupName here
        Write-Error "Check Panorama IP, API Key, and network connectivity."
    }
}

$AllSecurityRules | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8