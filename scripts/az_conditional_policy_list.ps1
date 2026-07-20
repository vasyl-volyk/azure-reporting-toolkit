param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath
)



function Test-Guid {
    [Cmdletbinding()]
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$InputObject
    )
    process {
        return [guid]::TryParse($InputObject, $([ref][guid]::Empty))
    }
}



$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential


if (-not $(Get-MgContext)) {
    Throw "Authentication needed, call 'Connect-Graph -Scopes `"Application.Read.All`", `"Group.Read.All`", `"Policy.Read.All`", `"RoleManagement.Read.Directory`", `"User.Read.All`""
}

# Get Conditional Access Policies
$conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
#Get Conditional Access Named / Trusted Locations
$namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
# Get Azure AD Directory Role Templates
$directoryRoleTemplates = Get-MgDirectoryRoleTemplate -ErrorAction Stop
# Get Azure AD Service Principals
$servicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop
# Init report
$conditionalAccessDocumentation = [System.Collections.Generic.List[Object]]::new()

# Process all Conditional Access Policies
foreach ($conditionalAccessPolicy in $conditionalAccessPolicies) {

    # Display some progress (based on policy count)
    $currentIndex = $conditionalAccessPolicies.indexOf($conditionalAccessPolicy)
    Write-Progress -Activity "Generating Conditional Access Documentation..." -PercentComplete (($currentIndex + 1) / $conditionalAccessPolicies.Count * 100) `
        -CurrentOperation "Processing Policy '$($conditionalAccessPolicy.DisplayName)' ($currentIndex/$($conditionalAccessPolicies.Count))"

    try {
        # Resolve object IDs of included users
        $includeUsers = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.IncludeUsers | ForEach-Object {
            if (Test-Guid $PSItem) {
                    #$includeUsers.Add( $(Get-MgUser -userId $PSItem | Select-Object -ExpandProperty DisplayName -ErrorAction Stop))
                    try {
                            $user = Get-MgUser -UserId $PSItem -ErrorAction Stop
                        } catch {
                            # Помилка подавлена
                            $user = $null
                        }
                    
                    if ($user) { $includeUsers.Add( $($user | Select-Object -ExpandProperty DisplayName -ErrorAction Stop)) }
                    else { $includeUsers.Add( $PSItem) }
            }
        }
        # Resolve object IDs of excluded users
        $excludeUsers = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.ExcludeUsers | ForEach-Object {
            if (Test-Guid $PSItem) {
                #$excludeUsers.Add($(Get-MgUser -userId $PSItem | Select-Object -ExpandProperty DisplayName -ErrorAction Stop))
                    try {
                                $user = Get-MgUser -UserId $PSItem -ErrorAction Stop
                            } catch {
                                # Помилка подавлена
                                $user = $null
                            }
                    
                        if ($user) { $excludeUsers.Add( $($user | Select-Object -ExpandProperty DisplayName -ErrorAction Stop)) }
                        else { $excludeUsers.Add( $PSItem) }
                }
        }
        # Resolve object IDs of included groups
        $includeGroups = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.IncludeGroups | ForEach-Object {
            #$includeGroups.Add($(Get-MgGroup -GroupId $PSItem | Select-Object -ExpandProperty DisplayName))
                    try {
                            $group = Get-MgGroup -GroupId $PSItem -ErrorAction Stop
                        } catch {
                            # Помилка подавлена
                            $group = $null
                        }
                    if ($group) { $includeGroups.Add($(Get-MgGroup -GroupId $PSItem | Select-Object -ExpandProperty DisplayName)) }
                    else { $includeGroups.Add( $PSItem) }
                }
        
        # Resolve object IDs of excluded groups
        $excludeGroups = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.ExcludeGroups | ForEach-Object {
            #$excludeGroups.Add( $(Get-MgGroup -GroupId $PSItem | Select-Object -ExpandProperty DisplayName))
                    try {
                            $group = Get-MgGroup -GroupId $PSItem -ErrorAction Stop
                        } catch {
                            # Помилка подавлена
                            $group = $null
                        }
                    if ($group) { $includeGroups.Add($(Get-MgGroup -GroupId $PSItem | Select-Object -ExpandProperty DisplayName)) }
                        else { $includeGroups.Add( $PSItem) }
                    
                }
        
        # Resolve object IDs of included roles
        $includeRoles = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.IncludeRoles | ForEach-Object {
            $roleId = $PSItem
            $includeRoles.Add( $($directoryRoleTemplates | Where-Object { $PSItem.Id -eq $roleId } | Select-Object -ExpandProperty DisplayName))
        }

        # Resolve object IDs of excluded roles
        $excludeRoles = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Users.ExcludeRoles | ForEach-Object {
            $roleId = $PSItem
            $excludeRoles.Add( $($directoryRoleTemplates | Where-Object { $PSItem.Id -eq $roleId } | Select-Object -ExpandProperty DisplayName ))
        }
        # Resolve object IDs of included apps
        $includeApps = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Applications.IncludeApplications | ForEach-Object {
            $servicePrincipalId = $PSItem
            if (Test-Guid $PSItem) {
                $res = $servicePrincipals | Where-Object { $PSItem.AppId -eq $servicePrincipalId } | Select-Object -ExpandProperty DisplayName
                if ($null -ne $res) {
                    $includeApps.Add($res)
                }
                else {
                    $includeApps.Add($servicePrincipalId)
                }
            }
            else {
                $includeApps.Add($servicePrincipalId)
            }
        }
        # Resolve object IDs of excluded apps
        $excludeApps = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Applications.ExcludeApplications | ForEach-Object {
            $servicePrincipalId = $PSItem
            if (Test-Guid $PSItem) {
                $res = $servicePrincipals | Where-Object { $PSItem.AppId -eq $servicePrincipalId } | Select-Object -ExpandProperty DisplayName
                if ($null -ne $res) {
                    $excludeApps.Add($res)
                }
                else {
                    $excludeApps.Add($servicePrincipalId)
                }
            }
            else {
                $excludeApps.Add($servicePrincipalId)
            }
        }
        # Resolve object IDs of included locations
        $includeLocations = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Locations.IncludeLocations | ForEach-Object {
            $locationId = $PSItem
            if (Test-Guid $PSItem) {
                $includeLocations.Add( $($namedLocations | Where-Object { $PSItem.Id -eq $locationId } | Select-Object -ExpandProperty DisplayName))
            }
            else {
                $includeLocations.Add($locationId)
            }
        }
        # Resolve object IDs of excluded locations
        $excludeLocations = [System.Collections.Generic.List[Object]]::new()
        $conditionalAccessPolicy.Conditions.Locations.ExcludeLocations | ForEach-Object {
            $locationId = $PSItem
            if (Test-Guid $PSItem) {
                $excludeLocations.Add( $($namedLocations | Where-Object { $PSItem.Id -eq $locationId } | Select-Object -ExpandProperty DisplayName))
            }
            else {
                $excludeLocations.Add($locationId)
            }
        }

        # delimiter for arrays in csv report
        $separator = "<br>"
        if ($conditionalAccessPolicy.GrantControls.TermsOfUse) { $conditionalAccessPolicy.GrantControls.BuiltInControls += "TermsOfUse" }
        $conditionalAccessDocumentation.Add(
            [PSCustomObject]@{
                Name                            = $conditionalAccessPolicy.DisplayName
                State                           = $conditionalAccessPolicy.State

                IncludeUsers                    = $includeUsers -join $separator
                IncludeGroups                   = $includeGroups -join $separator
                IncludeRoles                    = $includeRoles -join $separator

                ExcludeUsers                    = $excludeUsers -join $separator
                ExcludeGroups                   = $excludeGroups -join $separator
                ExcludeRoles                    = $excludeRoles -join $separator

                IncludeApps                     = $includeApps -join $separator
                ExcludeApps                     = $excludeApps -join $separator

                IncludeUserActions              = $conditionalAccessPolicy.Conditions.Applications.IncludeUserActions -join $separator
                ClientAppTypes                  = $conditionalAccessPolicy.Conditions.ClientAppTypes -join $separator

                IncludePlatforms                = $conditionalAccessPolicy.Conditions.Platforms.IncludePlatforms -join $separator
                ExcludePlatforms                = $conditionalAccessPolicy.Conditions.Platforms.ExcludePlatforms -join $separator

                IncludeLocations                = $includeLocations -join $separator
                ExcludeLocations                = $excludeLocations -join $separator

                DeviceFilterMode                = $conditionalAccessPolicy.Conditions.Devices.DeviceFilter.Mode
                DeviceFilterRule                = $conditionalAccessPolicy.Conditions.Devices.DeviceFilter.Rule

                GrantControls                   = $conditionalAccessPolicy.GrantControls.BuiltInControls -join $separator
                GrantControlsOperator           = $conditionalAccessPolicy.GrantControls.Operator

                SignInRiskLevels                = $conditionalAccessPolicy.Conditions.SignInRiskLevels -join $separator
                UserRiskLevels                  = $conditionalAccessPolicy.Conditions.UserRiskLevels -join $separator

                ApplicationEnforcedRestrictions = $conditionalAccessPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
                CloudAppSecurity                = $conditionalAccessPolicy.SessionControls.CloudAppSecurity.IsEnabled
                PersistentBrowser               = $conditionalAccessPolicy.SessionControls.PersistentBrowser.Mode
                SignInFrequency                 = "$($conditionalAccessPolicy.SessionControls.SignInFrequency.Value) $($conditionalAccessPolicy.SessionControls.SignInFrequency.Type)"
            }
        )
    }
    catch {
        Write-Error $PSItem
    }
}

Disconnect-MgGraph

# Build export path (script directory)
$exportPath = Join-Path $PSScriptRoot "ConditionalAccessDocumentation.csv"
# Export report as csv
$conditionalAccessDocumentation | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Output "Exported Documentation to '$($exportPath)'"