param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath

)


# Validate input
If ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret) -or [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

# Connect to Graph
$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential


if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-Graph -Scopes `"Application.Read.All`", `"Group.Read.All`", `"Policy.Read.All`", `"RoleManagement.Read.Directory`", `"User.Read.All`", `"AuditLog.Read.All`""
}


# Get all users with sign-in activity
$allUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Id,UserType,AccountEnabled,CreatedDateTime,JobTitle,OfficeLocation,MobilePhone,EmployeeType,signInActivity,Mail,EmployeeId"

# Filter out guests
$usersToProcess = $allUsers #| Where-Object { $_.UserType -ne "Guest" }

$results = @()

foreach ($user in $usersToProcess) {
    Write-Host "Processing $($user.DisplayName) <$($user.UserPrincipalName)>"
    $mfaStatus = "Not Registered"
    $mfaDetails = @()

    # Get sign-in activity
    $lastInteractiveSignIn = $null
    $lastNonInteractiveSignIn = $null
    
    if ($user.SignInActivity) {
        $lastInteractiveSignIn = $user.SignInActivity.LastSignInDateTime
        $lastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
    }

    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
        $methods = $methods | Where-Object { $_.AdditionalProperties.'@odata.type' -ne "#microsoft.graph.passwordAuthenticationMethod" }

        $userMFAMethods=""
        if ($methods.Count -gt 0) {
            $mfaStatus = "Registered"
                        $userMFAMethods =  $methods.AdditionalProperties | ForEach-Object {
                            $method = $_
                            switch ($_.'@odata.type') {
                                "#microsoft.graph.emailAuthenticationMethod" {
                                    ("Method: Email, address: ",$method.emailAddress) -join ""
                                }
                                "#microsoft.graph.phoneAuthenticationMethod" {
                                    ("Method: ",$method.phoneType,",  Number: ",$method.phoneNumber) -join ""
                                }
                                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                                    ("Method: Windows Hello for Business, device: ",$method.displayName) -join ""
                                }
                                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                                    ("Method: MS Authenticator, App Version: ",$method.phoneAppVersion,", Phone: ",$method.displayName) -join ""
                                }
                                "#microsoft.graph.fido2AuthenticationMethod" {
                                    ("Method: Hardware key FIDO2 ",$method.model,", Device: ",$method.displayName) -join ""
                                }
                                "#microsoft.graph.passwordAuthenticationMethod" {}
                                "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                                    ("Method: Temp Access Pass, Is isUsableOnce ",$method.isUsableOnce,", LifeTimeInMinutes: ",$method.lifetimeInMinutes   ) -join ""
                                }
                                "#microsoft.graph.hardwareOathAuthenticationMethod" {
                                    ("Method: Hardware key ",$method.model,", Device: ",$method.manufacturer." serial num:",$method.serialNumber) -join ""
                                }
                                "#microsoft.graph.softwareOathAuthenticationMethod" {
                                    ("Method:  software OATH token" ) -join ""
                                }

                                default {
                                    $_
                                }
                            }
                        }
            
        }
    }
    catch {
        Write-Warning "Error retrieving methods for $($user.UserPrincipalName): $($_.Exception.Message)"
    }

    $results += [PSCustomObject]@{
        DisplayName              = $user.DisplayName
        UserPrincipalName        = $user.UserPrincipalName
        UserType                 = $user.UserType
        EmployeeType             = $user.EmployeeType
        EmployeeId               = $user.EmployeeId
        PrimaryEmail             = $user.Mail
        AccountEnabled           = $user.AccountEnabled
        CreatedDateTime          = $user.CreatedDateTime
        LastInteractiveSignIn    = if ($lastInteractiveSignIn) { $lastInteractiveSignIn.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
        LastNonInteractiveSignIn = if ($lastNonInteractiveSignIn) { $lastNonInteractiveSignIn.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
        MFAStatus                = $mfaStatus
        MFAMethods               = ($userMFAMethods -join "<br>")
        JobTitle                 = $user.JobTitle
        OfficeLocation           = $user.OfficeLocation
        MobilePhone              = $user.MobilePhone
        ObjectId                 = $user.Id
    }
}

# Export to CSV
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to $OutputPath"
} else {
    Write-Warning "No results to export."
}

Disconnect-MgGraph
Write-Host "Script completed."
