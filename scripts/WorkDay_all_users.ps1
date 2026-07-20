param(
    [string]$User = $env:WORKDAY_USER,
    [string]$Pass = $env:WORKDAY_PASS,
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$targetOU = "DC=yourcompany,DC=com",
    [string]$OutputPath = "workday.csv"
)


# CHECK MODULE WORKDAY
#region
$moduleName = "WorkdayApi"

# Check if the module is installed (any version)
$module = Get-Module -ListAvailable -Name $moduleName

if ($null -ne $module) {
    Write-Host "$moduleName is already installed."
} else {
    Write-Host "$moduleName is not installed. Installing..."
    try {
        Install-Module -Name $moduleName -RequiredVersion 2.3.2 -Scope AllUsers -Force
        Write-Host "$moduleName installed successfully."
    }
    catch {
        Write-Error "Failed to install $moduleName. Error: $_"
    }
}
#endregion

# Import the module
Write-Host "Import module"
Import-Module WorkdayApi -ErrorAction Stop

$Password = ConvertTo-SecureString -AsPlainText $Pass -Force
$Credential = New-Object System.Management.Automation.PSCredential $User, $Password

Write-Host "Set credentials"
Set-WorkdayCredential -Credential $Credential
Write-Host "set endpoint"
Set-WorkdayEndpoint -Endpoint Human_Resources -Uri 'https://wdX-services1.myworkday.com/ccx/service/yourtenant/Human_Resources/v30.2'
Write-Host "get user from workday"
$AllWordayWorkers = Get-WorkdayWorker -IncludeInactive -IncludeWork -IncludePersonal -Verbose

#     READ ALL Azure users
#region
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

Write-Host "Get users from AZ"
# Get all users with sign-in activity
$allAZUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,AccountEnabled,UserType" | Where-Object { $_.UserType -ne "Guest" }
#endregion

#    READ ALL AD users
#region
Write-Host "Get users from AD"
$adUsers = Get-ADUser -Filter * -SearchBase $targetOU -Properties DisplayName, GivenName, Surname, SamAccountName, Enabled
#endregion


$results = @()

foreach($workdayuser in $AllWordayWorkers){
Write-Host $workdayuser.PreferredName

    $results += [PSCustomObject]@{
            "Preferre dName" = $workdayuser.PreferredName
            "Full Name"      = "$($workdayuser.FirstName) $($workdayuser.LastName)"
            "UserID" = $workdayuser.UserId
            "WorkerID" = $workdayuser.WorkerId
            "Active" = $workdayuser.Active
            "AD enabled" = if ($workdayuser.UserId -in $adUsers.SamAccountName) { ($adUsers | Where-Object{$_.SamAccountName -like $workdayuser.UserId}).Enabled } else { "n/a" }
            "Azure enabled" = if (($workdayuser.UserId+"@yourcompany.com") -in $allAZUsers.UserPrincipalName) { ($allAZUsers | Where-Object{$_.UserPrincipalName -like "$($workdayuser.UserId)@yourcompany.com"}).AccountEnabled } else { "n/a" }
            "Active_Status_Date" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Active_Status_Date -replace '-\d{2}:\d{2}$'
            "Hire_Date" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Hire_Date -replace '-\d{2}:\d{2}$'
            "Original_Hire_Date" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Original_Hire_Date -replace '-\d{2}:\d{2}$'
            "FirstDayOfWork" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.First_Day_of_Work -replace '-\d{2}:\d{2}$'
            "TerminationDate" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Termination_Date -replace '-\d{2}:\d{2}$'
            "Termination_Last_Day_of_Work" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Termination_Last_Day_of_Work -replace '-\d{2}:\d{2}$'
            "Rehire" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Status_Data.Rehire
            "Title" = $workdayuser.XML.Worker.Worker_Data.Employment_Data.Worker_Job_Data.Position_Data.Position_Title


#            "Email" = ($workdayuser.Email | Where-Object { $_.UsageType -eq 'HOME' } | Select-Object -First 1).Email
    }

}



#
if ($results.Count -gt 0) {
    Write-Host "----------------------------------------------------"
    Write-Host "Successfully collected information for $($allCollectedData.Count) users across all Workday userss."
    Write-Host "----------------------------------------------------"
    
    # You can now work with the $allCollectedData array.
    # For example, to display the first user's information (if any):
    # if ($allCollectedData.Count -gt 0) { $allCollectedData[0] | Format-List }

    # Or to export all collected data to a single CSV file:
    $csvPath = $OutputPath
    Write-Host "Exporting all collected data to CSV: $csvPath"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    # To display all collected data in the console (can be lengthy for many users):
    # $allCollectedData | Format-Table -AutoSize
} 
else {
    Write-Warning "No user data was collected from any of the specified OUs."
}