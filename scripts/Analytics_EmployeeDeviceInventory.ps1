param (
    [string]$ClientId     = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId     = $env:AZURE_TENANT_ID,
    [string]$OutputPath   = ".\EntraUserDeviceReport.csv"
)


# ======================================================================
# Helper: unique List[string] (case-insensitive)
# ======================================================================
function Get-UniqueDeviceNames {
    param([System.Collections.Generic.List[string]]$DeviceList)
    $seen   = @{}
    $unique = New-Object 'System.Collections.Generic.List[string]'
    foreach ($name in $DeviceList) {
        if ([string]::IsNullOrEmpty($name)) { continue }
        $key = $name.ToLower()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            [void]$unique.Add($key)
        }
    }
    return ,$unique
}

# ======================================================================
# Helper: short hostname from FQDN
# ======================================================================
function Get-ShortName {
    param([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return $Name }
    return (($Name -split '\.')[0]).ToLower()
}

# ----------------------------------------------------------------------
# Validate input
# ----------------------------------------------------------------------
if ([string]::IsNullOrEmpty($ClientId) -or
    [string]::IsNullOrEmpty($ClientSecret) -or
    [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    Exit 1
}

# ----------------------------------------------------------------------
# Connect to Microsoft Graph
# ----------------------------------------------------------------------
Write-Host "[1/7] Connecting to Microsoft Graph..." -ForegroundColor Cyan

$SecuredPassword        = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential `
                            -ArgumentList $ClientId, $SecuredPassword

Connect-MgGraph -TenantId $TenantId `
                -ClientSecretCredential $ClientSecretCredential `
                -NoWelcome

# ----------------------------------------------------------------------
# STEP 1 - Users with EmployeeId
# ----------------------------------------------------------------------
Write-Host "[2/7] Retrieving Entra ID users (EmployeeId is not empty)..." -ForegroundColor Cyan

$allUsers = Get-MgUser -All `
    -Filter "employeeId ne null" `
    -CountVariable userCount `
    -ConsistencyLevel eventual `
    -Property Id, DisplayName, UserPrincipalName, AccountEnabled, EmployeeId `
    -PageSize 999

Write-Host ("       Found {0} users." -f $allUsers.Count) -ForegroundColor Green

# ----------------------------------------------------------------------
# STEP 2 - Entra devices (Windows + macOS) with owners
# ----------------------------------------------------------------------
Write-Host "[3/7] Retrieving Entra ID registered devices..." -ForegroundColor Cyan

$entraDevices = Get-MgDevice -All `
    -Filter "operatingSystem eq 'Windows' or operatingSystem eq 'MacMDM' or operatingSystem eq 'macOS'" `
    -ExpandProperty registeredOwners `
    -Property Id, DeviceId, DisplayName, OperatingSystem, RegisteredOwners `
    -PageSize 999

Write-Host ("       Found {0} Entra devices." -f $entraDevices.Count) -ForegroundColor Green

# UPN -> List[deviceName]  (primary: UPN lookup)
# All device names flat     (always searched by name pattern too)
$entraDevicesByUpn   = @{}
$entraAllDeviceNames = New-Object 'System.Collections.Generic.List[string]'

foreach ($dev in $entraDevices) {
    $ownerUpn = $null
    if ($dev.RegisteredOwners -and $dev.RegisteredOwners.Count -gt 0) {
        foreach ($owner in $dev.RegisteredOwners) {
            if ($owner.AdditionalProperties -and $owner.AdditionalProperties.ContainsKey('userPrincipalName')) {
                $ownerUpn = $owner.AdditionalProperties['userPrincipalName'].ToString().ToLower()
                break
            }
        }
    }
    $devName = ($dev.DisplayName.ToString()).ToLower()
    [void]$entraAllDeviceNames.Add($devName)
    if ($ownerUpn) {
        if (-not $entraDevicesByUpn.ContainsKey($ownerUpn)) {
            $entraDevicesByUpn[$ownerUpn] = New-Object 'System.Collections.Generic.List[string]'
        }
        $entraDevicesByUpn[$ownerUpn].Add($devName)
    }
}
$entraAllDeviceNames = Get-UniqueDeviceNames -DeviceList $entraAllDeviceNames

Write-Host ("       Entra devices with UPN: {0} users mapped." -f $entraDevicesByUpn.Count) -ForegroundColor Gray
Write-Host ("       Entra devices total   : {0} devices." -f $entraAllDeviceNames.Count) -ForegroundColor Gray

# ----------------------------------------------------------------------
# STEP 3 - Intune managed devices (Windows + macOS)
# ----------------------------------------------------------------------
Write-Host "[4/7] Retrieving Intune managed devices..." -ForegroundColor Cyan

$intuneDevices = Get-MgDeviceManagementManagedDevice -All `
    -Filter "operatingSystem eq 'Windows' or operatingSystem eq 'macOS'" `
    -Property Id, DeviceName, OperatingSystem, UserPrincipalName, AzureADDeviceId `
    -PageSize 999

Write-Host ("       Found {0} Intune devices." -f $intuneDevices.Count) -ForegroundColor Green

# UPN -> List[deviceName]  (primary: UPN lookup)
# All device names flat     (always searched by name pattern too)
$intuneDevicesByUpn   = @{}
$intuneAllDeviceNames = New-Object 'System.Collections.Generic.List[string]'

foreach ($dev in $intuneDevices) {
    $devName = ($dev.DeviceName.ToString()).ToLower()
    [void]$intuneAllDeviceNames.Add($devName)
    if (-not [string]::IsNullOrEmpty($dev.UserPrincipalName)) {
        $key = $dev.UserPrincipalName.ToLower()
        if (-not $intuneDevicesByUpn.ContainsKey($key)) {
            $intuneDevicesByUpn[$key] = New-Object 'System.Collections.Generic.List[string]'
        }
        $intuneDevicesByUpn[$key].Add($devName)
    }
}
$intuneAllDeviceNames = Get-UniqueDeviceNames -DeviceList $intuneAllDeviceNames

Write-Host ("       Intune devices with UPN: {0} users mapped." -f $intuneDevicesByUpn.Count) -ForegroundColor Gray
Write-Host ("       Intune devices total   : {0} devices." -f $intuneAllDeviceNames.Count) -ForegroundColor Gray

# ----------------------------------------------------------------------
# STEP 4 - Defender for Endpoint machines (get ALL)
# ----------------------------------------------------------------------
Write-Host "[5/7] Retrieving Defender for Endpoint machines..." -ForegroundColor Cyan

$mdeTokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://api.securitycenter.microsoft.com/.default"
}

$mdeDevices = @()

try {
    $tokenResponse = Invoke-RestMethod `
        -Uri ("https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $TenantId) `
        -Method POST `
        -Body $mdeTokenBody

    $mdeHeaders = @{
        Authorization  = "Bearer $($tokenResponse.access_token)"
        'Content-Type' = 'application/json'
    }

    $mdeUri = "https://api.securitycenter.microsoft.com/api/machines"
    do {
        $response    = Invoke-RestMethod -Uri $mdeUri -Headers $mdeHeaders -Method GET
        $mdeDevices += $response.value
        $mdeUri      = $response.'@odata.nextLink'
    } while ($mdeUri)

    Write-Host ("       Found {0} Defender machines." -f $mdeDevices.Count) -ForegroundColor Green
}
catch {
    Write-Warning "Could not retrieve Defender for Endpoint machines. Ensure the app has Machine.Read.All permission on https://api.securitycenter.microsoft.com."
    Write-Warning ("Error: {0}" -f $_)
}

# Build Entra deviceId -> ownerUPN map for MDE cross-reference
$entraDeviceIdToUpn = @{}
foreach ($dev in $entraDevices) {
    $ownerUpn = $null
    if ($dev.RegisteredOwners -and $dev.RegisteredOwners.Count -gt 0) {
        foreach ($owner in $dev.RegisteredOwners) {
            if ($owner.AdditionalProperties -and $owner.AdditionalProperties.ContainsKey('userPrincipalName')) {
                $ownerUpn = $owner.AdditionalProperties['userPrincipalName'].ToString().ToLower()
                break
            }
        }
    }
    if ($ownerUpn -and $dev.DeviceId) {
        $entraDeviceIdToUpn[$dev.DeviceId.ToString().ToLower()] = $ownerUpn
    }
}

# UPN -> List[shortName]  (primary: UPN lookup)
# All short names flat     (always searched by name pattern too)
$mdeDevicesByUpn   = @{}
$mdeAllDeviceNames = New-Object 'System.Collections.Generic.List[string]'

foreach ($m in $mdeDevices) {
    $fullName = $null
    if ($m.computerDnsName) { $fullName = $m.computerDnsName.ToString().ToLower() }
    elseif ($m.machineName) { $fullName = $m.machineName.ToString().ToLower() }
    if ([string]::IsNullOrEmpty($fullName)) { continue }

    $shortName = Get-ShortName -Name $fullName
    [void]$mdeAllDeviceNames.Add($shortName)

    # Resolve UPN
    $upn = $null
    if ($m.PSObject.Properties['userPrincipalName'] -and -not [string]::IsNullOrEmpty($m.userPrincipalName)) {
        $upn = $m.userPrincipalName.ToLower()
    }
    if (-not $upn -and $m.aadDeviceId) {
        $aadId = $m.aadDeviceId.ToString().ToLower()
        if ($entraDeviceIdToUpn.ContainsKey($aadId)) {
            $upn = $entraDeviceIdToUpn[$aadId]
        }
    }

    if ($upn) {
        if (-not $mdeDevicesByUpn.ContainsKey($upn)) {
            $mdeDevicesByUpn[$upn] = New-Object 'System.Collections.Generic.List[string]'
        }
        $mdeDevicesByUpn[$upn].Add($shortName)
    }
}
$mdeAllDeviceNames = Get-UniqueDeviceNames -DeviceList $mdeAllDeviceNames

Write-Host ("       Defender devices with UPN: {0} users mapped." -f $mdeDevicesByUpn.Count) -ForegroundColor Gray
Write-Host ("       Defender devices total   : {0} devices." -f $mdeAllDeviceNames.Count) -ForegroundColor Gray

# ----------------------------------------------------------------------
# STEP 5 - Active Directory computers
# ----------------------------------------------------------------------
Write-Host "[6/7] Retrieving Active Directory computers..." -ForegroundColor Cyan

$adComputerNames = New-Object 'System.Collections.Generic.List[string]'
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $adComputers = Get-ADComputer -Filter * -Property Name
    foreach ($pc in $adComputers) {
        if ($pc.Name) { [void]$adComputerNames.Add(($pc.Name.ToString()).ToLower()) }
    }
    Write-Host ("       Found {0} AD computers." -f $adComputerNames.Count) -ForegroundColor Green
}
catch {
    Write-Warning "Could not retrieve Active Directory computers. Ensure the ActiveDirectory module is installed and the machine has domain connectivity."
    Write-Warning ("Error: {0}" -f $_)
}

# ----------------------------------------------------------------------
# STEP 6 - Build report
# ----------------------------------------------------------------------
Write-Host "[7/7] Building report..." -ForegroundColor Cyan

$report = New-Object 'System.Collections.Generic.List[psobject]'

foreach ($user in $allUsers) {
    $upn      = $user.UserPrincipalName.ToLower()
    $userName = ($upn -split '@')[0]
    $status   = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }

    $namePattern = "^(lt|osx)-$([regex]::Escape($userName))([^a-zA-Z0-9]|$)"

    $entraMatched  = New-Object 'System.Collections.Generic.List[string]'
    $intuneMatched = New-Object 'System.Collections.Generic.List[string]'
    $mdeMatched    = New-Object 'System.Collections.Generic.List[string]'
    $adMatched     = New-Object 'System.Collections.Generic.List[string]'

    # ------------------------------------------------------------------
    # Entra: UPN lookup + name search always run; results merged
    # ------------------------------------------------------------------
    if ($entraDevicesByUpn.ContainsKey($upn)) {
        foreach ($d in $entraDevicesByUpn[$upn]) { [void]$entraMatched.Add($d) }
    }
    foreach ($devName in $entraAllDeviceNames) {
        if ($devName -imatch $namePattern) { [void]$entraMatched.Add($devName) }
    }
    $entraMatched = Get-UniqueDeviceNames -DeviceList $entraMatched

    # ------------------------------------------------------------------
    # Intune: UPN lookup + name search always run; results merged
    # ------------------------------------------------------------------
    if ($intuneDevicesByUpn.ContainsKey($upn)) {
        foreach ($d in $intuneDevicesByUpn[$upn]) { [void]$intuneMatched.Add($d) }
    }
    foreach ($devName in $intuneAllDeviceNames) {
        if ($devName -imatch $namePattern) { [void]$intuneMatched.Add($devName) }
    }
    $intuneMatched = Get-UniqueDeviceNames -DeviceList $intuneMatched

    # ------------------------------------------------------------------
    # Defender: UPN lookup + name search always run; results merged
    # ------------------------------------------------------------------
    if ($mdeDevicesByUpn.ContainsKey($upn)) {
        foreach ($d in $mdeDevicesByUpn[$upn]) { [void]$mdeMatched.Add($d) }
    }
    foreach ($mdeName in $mdeAllDeviceNames) {
        if ($mdeName -imatch $namePattern) { [void]$mdeMatched.Add($mdeName) }
    }
    $mdeMatched = Get-UniqueDeviceNames -DeviceList $mdeMatched

    # ------------------------------------------------------------------
    # AD: name-based only (AD computer objects have no user UPN field)
    # ------------------------------------------------------------------
    foreach ($adName in $adComputerNames) {
        if ($adName -imatch $namePattern) { [void]$adMatched.Add($adName) }
    }
    $adMatched = Get-UniqueDeviceNames -DeviceList $adMatched

    # Union device names across all sources
    $allDevicesSeen = @{}
    $allDeviceNames = New-Object 'System.Collections.Generic.List[string]'

    foreach ($list in @($entraMatched, $intuneMatched, $mdeMatched, $adMatched)) {
        foreach ($name in $list) {
            $key = $name.ToLower()
            if (-not $allDevicesSeen.ContainsKey($key)) {
                $allDevicesSeen[$key] = $true
                [void]$allDeviceNames.Add($key)
            }
        }
    }

    # Sets for membership check
    $entraSet  = @{}; foreach ($n in $entraMatched)  { $entraSet[$n]  = $true }
    $intuneSet = @{}; foreach ($n in $intuneMatched) { $intuneSet[$n] = $true }
    $mdeSet    = @{}; foreach ($n in $mdeMatched)    { $mdeSet[$n]    = $true }
    $adSet     = @{}; foreach ($n in $adMatched)     { $adSet[$n]     = $true }

    if ($allDeviceNames.Count -eq 0) {
        $report.Add([PSCustomObject]@{
            UserFullName          = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            UserStatus     = $status
            DeviceName        = ""
            EntraDevice       = "-"
            IntuneDevice      = "-"
            DefenderDevice    = "-"
            ADdevice          = "-"
        })
    }
    else {
        foreach ($devName in $allDeviceNames) {
            $report.Add([PSCustomObject]@{
                UserFullName          = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                UserStatus     = $status
                DeviceName        = $devName
                EntraDevice       = if ($entraSet.ContainsKey($devName))  { "+" } else { "-" }
                IntuneDevice      = if ($intuneSet.ContainsKey($devName)) { "+" } else { "-" }
                DefenderDevice    = if ($mdeSet.ContainsKey($devName))    { "+" } else { "-" }
                ADdevice          = if ($adSet.ContainsKey($devName))     { "+" } else { "-" }
            })
        }
    }
}

# ----------------------------------------------------------------------
# EXPORT CSV
# ----------------------------------------------------------------------
$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "`nReport saved to: $OutputPath" -ForegroundColor Green
Write-Host ("Total rows: {0}" -f $report.Count) -ForegroundColor Green

Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Cyan
