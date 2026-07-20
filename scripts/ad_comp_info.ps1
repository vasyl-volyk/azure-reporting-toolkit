param (
    [string]$TargetOU,
    [string]$OutputPath
)
# Requires the Active Directory module: Install-WindowsFeature -Name RSAT-AD-PowerShell
# Define the OU and Domain
#$TargetOU = "DC=yourcompany,DC=com"
$ouPath = $TargetOU
$domain = "dk.yourcompany.com"
# Define the properties to retrieve
$properties = @(
    "Name",
    "OperatingSystem",
    "OperatingSystemVersion",
    "distinguishedName", # For the full path
    "ms-Mcs-AdmPwd",     # For LAPS password
    "CanonicalName",      # To get the canonical name as well
    "WhenCreated"       # creation date
)
Write-Host "Connecting to Active Directory domain: $domain"
Write-Host "Searching for computer objects in OU: $ouPath`n"
# Initialize an empty array to store computer information (as hashtables)
$allComputerData = @()
try {
    # Get all computer objects from the specified OU
    $computers = Get-ADComputer -Filter * -SearchBase $ouPath -Properties $properties -ErrorAction Stop
    if ($computers) {
        Write-Host "Found $($computers.Count) computer objects."
        Write-Host "Processing data and storing in hashtables..."
        # Iterate through each computer object and create a hashtable for it
        foreach ($computer in $computers) {
            # Retrieve BitLocker recovery keys for the computer
            $bitlockerKeys = @()
            try {
                if ($computer.DistinguishedName) {
                    $recoveryInfo = Get-ADObject -Filter {objectClass -eq 'msFVE-RecoveryInformation'} `
                        -SearchBase $computer.DistinguishedName `
                        -SearchScope OneLevel `
                        -Properties msFVE-RecoveryPassword, msFVE-RecoveryGuid, WhenCreated `
                        -ErrorAction SilentlyContinue
                    
                    if ($recoveryInfo) {
                        foreach ($key in $recoveryInfo) {
                            $bitlockerKeys += "[Created: $($key.WhenCreated)] Key: $($key.'msFVE-RecoveryPassword')"
                        }
                    }
                }
            }
            catch {
                # Silently continue if BitLocker info cannot be retrieved
            }
            
            # Join all BitLocker keys with separator if multiple exist
            $bitlockerKeysString = if ($bitlockerKeys.Count -gt 0) { 
                $bitlockerKeys -join "<br>" 
            } else { 
                "" 
            }
            
            $computerInfo = [PSCustomObject]@{
                "ComputerName"          = $computer.Name
                "Enabled"               = $computer.Enabled
                "When created"          = $computer.WhenCreated
                "CanonicalName"         = $computer.CanonicalName
                "OperatingSystem"       = $computer.OperatingSystem
                "OperatingSystemVersion"= $computer.OperatingSystemVersion
                "LAPS_Password"         = $computer."ms-Mcs-AdmPwd"
                "BitLocker_Keys"        = $bitlockerKeysString
            }
            # Add the hashtable for the current computer to the array
            $allComputerData += $computerInfo
        }
        Write-Host "`nData collection complete. Here's a summary:`n"
        # You can now work with $allComputerData.
        # For example, to display it:
        #$allComputerData | Format-Table -AutoSize
        Write-Host "Records found", $allComputerData.Count
        # Or to export it to a CSV file:
        $allComputerData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nData exported to: $OutputPath"
    } else {
        Write-Host "No computer objects found in the '$ouPath' OU."
    }
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Write-Error "The specified OU '$ouPath' was not found in the domain '$domain'. Please check the OU path."
}
catch [System.Management.Automation.RuntimeException] {
    Write-Error "An error occurred while connecting to Active Directory or retrieving data. Ensure the Active Directory module is installed and you have sufficient permissions. Error: $($_.Exception.Message)"
}
catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
}
# The $allComputerData variable now holds an array of hashtables,
# each representing a computer object with its collected information.
Out-File -FilePath c -Encoding utf8