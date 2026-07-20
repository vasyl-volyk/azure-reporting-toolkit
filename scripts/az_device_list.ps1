param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = "$PSScriptRoot\az_device_list.csv",
    [bool]$IncludeKeyValue = $true,
    [switch]$Quiet
)


function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')]
        [string]$Level = 'Info'
    )
    if ($Quiet -and $Level -eq 'Debug') { return }
    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        'Debug'   { 'Gray' }
        default   { 'White' }
    }
    Write-Host $Message -ForegroundColor $color
}

if ([string]::IsNullOrEmpty($ClientId) -or
    [string]::IsNullOrEmpty($ClientSecret) -or
    [string]::IsNullOrEmpty($TenantId)) {
    Write-Error "AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    exit 1
}

$SecuredPasswordPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPasswordPassword

Disconnect-MgGraph -ErrorAction SilentlyContinue

Write-Output "Connecting to Microsoft Graph..."
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

if (-not (Get-MgContext)) {
    throw "Authentication failed. Ensure Connect-MgGraph succeeded."
}

Write-Output "Checking Graph permissions..."
$context = Get-MgContext
Write-Output "  Connected as: $($context.Account)"
Write-Output "  App ID: $($context.ClientId)"
Write-Output ""

Write-Output "Note: To retrieve recovery key values, the app registration needs:"
Write-Output "  Windows BitLocker keys:"
Write-Output "    - BitLockerKey.Read.All (Application permission)"
Write-Output "  macOS FileVault keys:"
Write-Output "    - DeviceManagementManagedDevices.PrivilegedOperations.All (Application permission)"
Write-Output "    - FileVault personal recovery key escrow enabled in Intune macOS disk encryption profile"
Write-Output ""

function Invoke-GraphPaged {
    param([string]$Url)
    $items = @()
    $next = $Url
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($null -eq $resp) { break }
        if ($resp.value) { $items += $resp.value }
        $next = $resp.'@odata.nextLink'
    }
    return $items
}

function Get-RecoveryKeyValue {
    param([string]$RecoveryKeyId)
    if (-not $RecoveryKeyId) { return $null }

    $uri = "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/${RecoveryKeyId}?`$select=key"

    try {
        Write-Log "  [DEBUG] Attempting to retrieve BitLocker key for: $RecoveryKeyId" -Level Debug
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        if ($resp.key) {
            Write-Log "  [SUCCESS] BitLocker key retrieved!" -Level Success
            return $resp.key
        }

        Write-Log "  [WARNING] Response received but 'key' property is missing or empty" -Level Warning
        return $null
    } catch {
        Write-Log "  [ERROR] Failed to retrieve BitLocker key: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Get-FileVaultKeyValue {
    param([string]$ManagedDeviceId)
    if (-not $ManagedDeviceId) { return $null }

    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$ManagedDeviceId/getFileVaultKey"

    try {
        Write-Log "  [DEBUG] Attempting FileVault key for Intune device: $ManagedDeviceId" -Level Debug
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

        $keyValue = if ($resp.value) { $resp.value } elseif ($resp -is [string]) { $resp } else { $null }
        if ($keyValue -and $keyValue.ToString().Trim() -ne "") {
            Write-Log "  [SUCCESS] FileVault key retrieved!" -Level Success
            return $keyValue.ToString().Trim()
        }

        Write-Log "  [WARNING] FileVault response was empty (key may not be escrowed yet)" -Level Warning
        return $null
    } catch {
        $errorDetails = $_.Exception.Message
        $errorBody    = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { "" }

        if ($errorDetails -like "*Forbidden*" -or $errorDetails -like "*403*" -or
            $errorBody -like "*PrivilegedOperations.All*") {
            Write-Log "  [ERROR] Grant DeviceManagementManagedDevices.PrivilegedOperations.All + admin consent" -Level Error
        } elseif ($errorDetails -like "*NotFound*" -or $errorDetails -like "*404*") {
            Write-Log "  [WARNING] No FileVault key escrowed for this device (404)" -Level Warning
        } else {
            Write-Log "  [ERROR] Failed to retrieve FileVault key: $errorDetails" -Level Error
        }
        return $null
    }
}

function Get-NormalizedDeviceName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    return $Name.Trim().ToLowerInvariant()
}

function Resolve-IntuneManagedDevice {
    param(
        $AadDevice,
        [hashtable]$ManagedByAadDeviceId,
        [hashtable]$ManagedByAadObjectId,
        [hashtable]$ManagedByDeviceName
    )

    if ($AadDevice.DeviceId -and $ManagedByAadDeviceId.ContainsKey($AadDevice.DeviceId)) {
        return $ManagedByAadDeviceId[$AadDevice.DeviceId]
    }
    if ($AadDevice.Id -and $ManagedByAadObjectId.ContainsKey($AadDevice.Id)) {
        return $ManagedByAadObjectId[$AadDevice.Id]
    }

    $normalizedName = Get-NormalizedDeviceName -Name $AadDevice.DisplayName
    if ($normalizedName -and $ManagedByDeviceName.ContainsKey($normalizedName)) {
        return $ManagedByDeviceName[$normalizedName]
    }

    return $null
}

function Test-IsMacOsDevice {
    param([string]$OperatingSystem)
    return ($OperatingSystem -in @('macOS', 'MacMDM')) -or ($OperatingSystem -like '*Mac*')
}

function Format-MacRecoveryKeyColumn {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return '' }
    return $Key.Trim()
}

$results = @()

Write-Output "1) Retrieving Azure AD devices..."
$aadDevices = Get-MgDevice -All
Write-Output "   Found $($aadDevices.Count) total devices"

$aadDevicesById = @{}
foreach ($d in $aadDevices) {
    $aadDevicesById[$d.Id] = $d
    if ($d.DeviceId) { $aadDevicesById[$d.DeviceId] = $d }
}

Write-Output "2) Retrieving Intune managed devices..."
try {
    $managedDevices = Get-MgDeviceManagementManagedDevice -All
} catch {
    $managedDevices = Invoke-GraphPaged -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
}

$managedByAadDeviceId = @{}
$managedByAadObjectId   = @{}
$managedByDeviceName    = @{}
$intuneMacDevices       = @()

foreach ($m in $managedDevices) {
    if ($m.azureADDeviceId) {
        $managedByAadDeviceId[$m.azureADDeviceId] = $m
    }
    if ($m.azureActiveDirectoryDeviceId) {
        $managedByAadObjectId[$m.azureActiveDirectoryDeviceId] = $m
    }

    $intuneName = if ($m.deviceName) { $m.deviceName } else { $m.managedDeviceName }
    $normalized = Get-NormalizedDeviceName -Name $intuneName
    if ($normalized -and -not $managedByDeviceName.ContainsKey($normalized)) {
        $managedByDeviceName[$normalized] = $m
    }

    if (Test-IsMacOsDevice -OperatingSystem $m.operatingSystem) {
        $intuneMacDevices += $m
    }
}

Write-Output "   Intune macOS devices: $($intuneMacDevices.Count)"

Write-Output "3) Retrieving BitLocker recovery key metadata (Windows)..."
$recoveryKeys = Invoke-GraphPaged -Url "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys"
Write-Output "   Found $($recoveryKeys.Count) BitLocker recovery keys"

if ($recoveryKeys.Count -gt 0 -and $IncludeKeyValue) {
    Write-Host ""
    Write-Host "TESTING BITLOCKER KEY RETRIEVAL" -ForegroundColor Cyan
    $testKey = Get-RecoveryKeyValue -RecoveryKeyId $recoveryKeys[0].id
    if ($testKey) {
        Write-Host "  BitLocker test: SUCCESS" -ForegroundColor Green
    } else {
        Write-Host "  BitLocker test: FAILED" -ForegroundColor Red
    }
    Write-Host ""
}

$firstMac = $intuneMacDevices | Select-Object -First 1
if ($firstMac -and $IncludeKeyValue) {
    Write-Host "TESTING FILEVAULT KEY RETRIEVAL" -ForegroundColor Cyan
    $testFvKey = Get-FileVaultKeyValue -ManagedDeviceId $firstMac.id
    if ($testFvKey) {
        Write-Host "  FileVault test: SUCCESS" -ForegroundColor Green
    } else {
        Write-Host "  FileVault test: FAILED or no key escrowed" -ForegroundColor Yellow
    }
    Write-Host ""
}

$keysByDeviceId = @{}
foreach ($rk in $recoveryKeys) {
    $did = if ($rk.deviceId) { $rk.deviceId } else { $rk.backedUpToDeviceId }
    if ($did) {
        if (-not $keysByDeviceId.ContainsKey($did)) { $keysByDeviceId[$did] = @() }
        $keysByDeviceId[$did] += $rk
    }
}

Write-Output "4) Correlating devices and keys..."
$keyRetrievalSuccessCount = 0
$devicesWithKeys          = 0
$devicesWithoutKeys       = 0
$processedIntuneMacIds    = @{}
$fileVaultCache           = @{}

function Get-CachedFileVaultKey {
    param([string]$ManagedDeviceId)
    if (-not $ManagedDeviceId) { return $null }
    if ($fileVaultCache.ContainsKey($ManagedDeviceId)) {
        return $fileVaultCache[$ManagedDeviceId]
    }
    $key = Get-FileVaultKeyValue -ManagedDeviceId $ManagedDeviceId
    $fileVaultCache[$ManagedDeviceId] = $key
    return $key
}

foreach ($device in $aadDevices) {
    $deviceObjectId = $device.Id
    $deviceId       = $device.DeviceId
    $deviceName     = $device.DisplayName
    $deviceOSType   = $device.OperatingSystem
    $managed        = Resolve-IntuneManagedDevice -AadDevice $device `
        -ManagedByAadDeviceId $managedByAadDeviceId `
        -ManagedByAadObjectId $managedByAadObjectId `
        -ManagedByDeviceName $managedByDeviceName
    $isManaged      = $null -ne $managed
    $isMac          = Test-IsMacOsDevice -OperatingSystem $deviceOSType
    $consolidatedKeys = ''
    $hasKey         = $false
    $keyCount       = 0

    if ($isMac) {
        if ($isManaged -and $IncludeKeyValue) {
            $intuneDeviceId = $managed.id
            $processedIntuneMacIds[$intuneDeviceId] = $true
            Write-Log "  [macOS] Retrieving FileVault key for: $deviceName (Intune ID: $intuneDeviceId)"

            $consolidatedKeys = Format-MacRecoveryKeyColumn -Key (Get-CachedFileVaultKey -ManagedDeviceId $intuneDeviceId)

            if ($consolidatedKeys) {
                $keyRetrievalSuccessCount++
                $hasKey   = $true
                $keyCount = 1
                Write-Log "    FileVault key retrieved for $deviceName" -Level Success
            } else {
                Write-Log "    No FileVault key escrowed for $deviceName — RecoveryKeys left blank" -Level Debug
            }
        } elseif (-not $isManaged) {
            Write-Log "  [macOS] $deviceName is not Intune-managed" -Level Debug
        }
    } else {
        $deviceKeys = @()
        if ($deviceId -and $keysByDeviceId.ContainsKey($deviceId)) {
            $deviceKeys = $keysByDeviceId[$deviceId]
            Write-Log "Found $($deviceKeys.Count) BitLocker key(s) for: $deviceName"
        }

        if ($deviceKeys.Count -gt 0) {
            $keyStrings = @()
            foreach ($rk in $deviceKeys) {
                $keyValue = $null
                if ($IncludeKeyValue) {
                    $keyValue = Get-RecoveryKeyValue -RecoveryKeyId $rk.id
                    if ($keyValue) { $keyRetrievalSuccessCount++ }
                }

                $createdDate = ""
                if ($rk.createdDateTime) {
                    try {
                        $createdDate = ([DateTime]::Parse($rk.createdDateTime)).ToString("MM/dd/yyyy HH:mm:ss")
                    } catch {
                        $createdDate = $rk.createdDateTime
                    }
                }
                $keyStrings += "[Created: $createdDate] Key: $keyValue"
            }

            $consolidatedKeys = $keyStrings -join "<br>"
            $hasKey   = $true
            $keyCount = $deviceKeys.Count
        }
    }

    if ($hasKey) { $devicesWithKeys++ } else { $devicesWithoutKeys++ }

    $results += [PSCustomObject]@{
        DeviceName             = $deviceName
        AccountEnabled         = $device.AccountEnabled
        DeviceType             = $device.DeviceType
        OperatingSystem        = $deviceOSType
        OperatingSystemVersion = $device.OperatingSystemVersion
        IsIntuneManaged        = $isManaged
        HasBitlockerKey        = $hasKey
        RecoveryKeys           = $consolidatedKeys
        KeyCount               = $keyCount
        DeviceId               = $deviceObjectId
        DeviceIdGuid           = $deviceId
        IntuneDeviceId         = $(if ($managed) { $managed.id } else { $null })
        Source                 = 'EntraDevice'
    }
}

Write-Output "5) Processing Intune macOS devices not matched to Entra..."
$orphanMacCount = 0
foreach ($m in $intuneMacDevices) {
    if ($processedIntuneMacIds.ContainsKey($m.id)) { continue }
    $orphanMacCount++

    $deviceName = if ($m.deviceName) { $m.deviceName } else { $m.managedDeviceName }
    $keys       = ''
    $hasKey     = $false
    $keyCount   = 0

    if ($IncludeKeyValue) {
        Write-Log "  [macOS/Intune-only] Retrieving FileVault key for: $deviceName"
        $keys = Format-MacRecoveryKeyColumn -Key (Get-CachedFileVaultKey -ManagedDeviceId $m.id)
        if ($keys) {
            $keyRetrievalSuccessCount++
            $hasKey   = $true
            $keyCount = 1
            Write-Log "    FileVault key retrieved for $deviceName" -Level Success
        } else {
            Write-Log "    No FileVault key escrowed for $deviceName — RecoveryKeys left blank" -Level Debug
        }
    }

    if ($hasKey) { $devicesWithKeys++ } else { $devicesWithoutKeys++ }

    $results += [PSCustomObject]@{
        DeviceName             = $deviceName
        AccountEnabled         = $null
        DeviceType             = 'IntuneManagedDevice'
        OperatingSystem        = $m.operatingSystem
        OperatingSystemVersion = $m.osVersion
        IsIntuneManaged        = $true
        HasBitlockerKey        = $hasKey
        RecoveryKeys           = $keys
        KeyCount               = $keyCount
        DeviceId               = $m.azureADDeviceId
        DeviceIdGuid           = $m.azureADDeviceId
        IntuneDeviceId         = $m.id
        Source                 = 'IntuneOnly'
    }
}
Write-Output "   Added $orphanMacCount Intune-only macOS device row(s)"

Write-Output "Writing results to $OutputPath"

$csvRows = foreach ($row in $results) {
    $recoveryKeys = $row.RecoveryKeys
    if (-not $row.HasBitlockerKey) {
        $recoveryKeys = ''
    } elseif (Test-IsMacOsDevice -OperatingSystem $row.OperatingSystem) {
        $recoveryKeys = Format-MacRecoveryKeyColumn -Key $recoveryKeys
    }

    [PSCustomObject]@{
        DeviceName             = $row.DeviceName
        AccountEnabled         = $row.AccountEnabled
        OperatingSystem        = $row.OperatingSystem
        OperatingSystemVersion = $row.OperatingSystemVersion
        IsIntuneManaged        = $row.IsIntuneManaged
        HasBitlockerKey        = $row.HasBitlockerKey
        RecoveryKeys           = $recoveryKeys
        KeyCount               = $row.KeyCount
        DeviceId               = $row.DeviceId
        DeviceIdGuid           = $row.DeviceIdGuid
        IntuneDeviceId         = $row.IntuneDeviceId
        Source                 = $row.Source
    }
}

$csvRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$macWithKey = @($results | Where-Object { (Test-IsMacOsDevice $_.OperatingSystem) -and $_.HasBitlockerKey }).Count
$macIntune  = @($results | Where-Object { (Test-IsMacOsDevice $_.OperatingSystem) -and $_.IsIntuneManaged }).Count
$winWithKey = @($results | Where-Object { -not (Test-IsMacOsDevice $_.OperatingSystem) -and $_.HasBitlockerKey }).Count

Write-Output ""
Write-Output "========== SUMMARY =========="
Write-Output "Done. CSV: $OutputPath"
Write-Output "Total rows: $($results.Count)"
Write-Output "  With recovery keys: $devicesWithKeys"
Write-Output "  Without recovery keys: $devicesWithoutKeys"
Write-Output "  Windows/other with keys: $winWithKey"
Write-Output "  macOS Intune-managed: $macIntune"
Write-Output "  macOS with FileVault key: $macWithKey"
Write-Output "Key API calls succeeded: $keyRetrievalSuccessCount"
Write-Output "============================="
