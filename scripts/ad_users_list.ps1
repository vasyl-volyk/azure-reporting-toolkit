param (
    [string[]]$TargetOU,
    [string]$OutputPath
)

# Requires the Active Directory module to be installed on the machine
# Install with: Install-Module -Name ActiveDirectory

# --- Configuration ---
# IMPORTANT: Replace "OU=YourUsers,DC=YourDomain,DC=com" with the actual distinguished name(s) of your OU(s).
# If you have multiple OUs to process, list them in this array.

$targetOUs = $TargetOU

#$targetOUs = "DC=yourcompany,DC=com"

$allCollectedData = @() # Initialize an empty array to store ALL user hash tables from all OUs

foreach ($ouPath in $targetOUs) {
    Write-Host "Collecting user information from: $ouPath"
    
    $usersInfoForCurrentOU = @() # Initialize an empty array for users in the current OU

    try {
        # Get all users from the specified OU
        $adUsers = Get-ADUser -Filter * -SearchBase $ouPath -Properties `
            DisplayName, GivenName, Surname, SamAccountName, AccountExpirationDate, `
            Office, Country, Title, Department, Company, Manager, Mail, ProxyAddresses, LastLogonDate, workerType, WhenCreated, Enabled, PasswordLastSet,employeeType

            if ($adUsers) {
                foreach ($user in $adUsers) {
                        Write-Host $user.DisplayName

                        # Format AccountExpirationDate
                        if ($user.AccountExpirationDate) {
                            $UserExp = $user.AccountExpirationDate.ToShortDateString()
                        } else {
                            $UserExp = "Never"
                        }

                        # Get Manager's Display Name
                        if ($user.Manager) {
                            try {
                                $manager = Get-ADUser -Identity $user.Manager -Properties DisplayName
                                $UserManager = $manager.DisplayName
                            }
                            catch {
                                $UserManager = "Could not resolve manager: $($user.Manager)"
                            }
                        } else {
                            $UserManager = "N/A"
                        }

                        # Extract email aliases from ProxyAddresses
                        $emailAliases = ""
                        if ($user.ProxyAddresses) {
                            foreach ($proxyAddress in $user.ProxyAddresses) {
                                # Filter for SIP and SMTP addresses that are not the primary SMTP address
                                if (($proxyAddress -cnotlike "SMTP:*") -and ($proxyAddress -notlike "x500:*") -and ($proxyAddress -notlike "x400:*") -and ($proxyAddress -notmatch "\byourcompany\b")) {
                                        $emailAliases += $proxyAddress.Replace("smtp:", "") + ", "
                                }
                            }
                        }
                        else { $emailAliases = "no Emails"}

                        $UserCreationdateOnly = $user.WhenCreated.ToString("yyyy-MM-dd") 

                        $usersInfoForCurrentOU += [PSCustomObject]@{
                            "Display Name" = $user.DisplayName
                            "Login Name" = $user.SamAccountName
                            "Enabled" = $user.enabled
                            "Last logon" = $user.LastLogonDate
                            "Password Last Set"       = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
                            "When created" = $UserCreationdateOnly  
                            "Exp. Date" = $UserExp
                            "Worker Type" = $user.workerType
                            "Employee Type" = $user.employeeType
                            "Manager" = $UserManager
                            "Office" = $user.Office
                            "Country" = $user.Country
                            "Job Title" = $user.Title
                            "Department" = $user.Department
                            "Company" = $user.Company
                            "First+Last names" = $user.GivenName +" "+ $user.Surname
                            "Email" = $user.Mail
                             "Non standart Email Aliases" = $emailAliases
                        }
                }

                Write-Host "  Collected information for $($usersInfoForCurrentOU.Count) users in $ouPath."
                $allCollectedData += $usersInfoForCurrentOU # Add users from current OU to the grand total array
            } 
            else {
                Write-Warning "No users found in the specified OU: $ouPath"
            }
      
    }
    catch {
        Write-Error "An error occurred while retrieving user information from $ouPath : $($_.Exception.Message)"
    }
    Write-Host "" # Add a blank line for readability between OUs
}

# --- Output Results ---

if ($allCollectedData.Count -gt 0) {
    Write-Host "----------------------------------------------------"
    Write-Host "Successfully collected information for $($allCollectedData.Count) users across all specified OUs."
    Write-Host "----------------------------------------------------"
    
    # You can now work with the $allCollectedData array.
    # For example, to display the first user's information (if any):
    # if ($allCollectedData.Count -gt 0) { $allCollectedData[0] | Format-List }

    # Or to export all collected data to a single CSV file:
    $csvPath = $OutputPath
    Write-Host "Exporting all collected data to CSV: $csvPath"
    $allCollectedData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # To display all collected data in the console (can be lengthy for many users):
    # $allCollectedData | Format-Table -AutoSize
} 
else {
    Write-Warning "No user data was collected from any of the specified OUs."
}