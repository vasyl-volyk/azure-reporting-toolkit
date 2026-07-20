#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves all LastPass shared folders, their entries, and user access permissions.

.DESCRIPTION
    Uses the LastPass Enterprise API (cmd: getsfdata) to list all shared folders,
    the users who have access, and their permission type (read-only vs. read/write,
    admin or not). Outputs to console and optionally exports to CSV.

.PARAMETER CID
    Your LastPass Company ID (Account Number). Found in the Admin Console > Dashboard.

.PARAMETER ProvHash
    Your LastPass Provisioning Hash (API key). Found in Admin Console > Advanced > Enterprise API.

.PARAMETER ApiUser
    The admin username making the API call (your LastPass admin email).

.PARAMETER ExportCSV
    Optional. Path to export results as a CSV file. E.g. "C:\Reports\lastpass_folders.csv"

.EXAMPLE
    .\Get-LastPassSharedFolders.ps1 -CID "123456" -ProvHash "abc123..." -ApiUser "admin@company.com"

.EXAMPLE
    .\Get-LastPassSharedFolders.ps1 -CID "123456" -ProvHash "abc123..." -ApiUser "admin@company.com" -ExportCSV "C:\Reports\lastpass_folders.csv"

.NOTES
    Requires a LastPass Business/Enterprise account.
    The admin account used must have permissions to view shared folder data.
    API endpoint: https://lastpass.com/enterpriseapi.php
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "LastPass Company ID (CID)")]
    [string]$CID,

    [Parameter(Mandatory = $true, HelpMessage = "LastPass Provisioning Hash")]
    [string]$ProvHash,

    [Parameter(Mandatory = $true, HelpMessage = "Admin username (email) making the API call")]
    [string]$ApiUser,

    [Parameter(Mandatory = $false, HelpMessage = "Optional CSV export path")]
    [string]$ExportCSV
)

# --- Configuration ---
$ApiUrl = "https://lastpass.com/enterpriseapi.php"

# --- Helper: Decode permission flags ---
function Get-AccessType {
    param (
        [bool]$ReadOnly,
        [bool]$Admin,
        [bool]$HidePasswords
    )
    $accessParts = @()

    if ($Admin) {
        $accessParts += "Admin"
    }

    if ($ReadOnly) {
        $accessParts += "Read-Only"
    } else {
        $accessParts += "Read/Write"
    }

    if ($HidePasswords) {
        $accessParts += "Passwords Hidden"
    }

    return ($accessParts -join ", ")
}

# --- Step 1: Call the API ---
Write-Host "`n[LastPass] Connecting to API as '$ApiUser'..." -ForegroundColor Cyan

$payload = @{
    "cid"      = $CID
    "provhash" = $ProvHash
    "apiuser"  = $ApiUser
    "cmd"      = "getsfdata"
} | ConvertTo-Json

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method Post -Body $payload -ContentType "application/json" -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Error "Failed to reach the LastPass API: $_"
    exit 1
}

# --- Step 2: Parse the response ---
try {
    $data = $response.Content | ConvertFrom-Json -Depth 10
} catch {
    Write-Error "Failed to parse API response: $_"
    Write-Host "Raw response: $($response.Content)" -ForegroundColor Yellow
    exit 1
}

# Check for API-level errors (LastPass returns HTTP 200 even on errors)
if ($data.status -and $data.status -ne "OK") {
    Write-Error "API returned an error: $($data.status) - $($data.error)"
    exit 1
}

# Get all shared folder IDs (top-level properties of the response)
$folderIDs = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

if ($folderIDs.Count -eq 0) {
    Write-Warning "No shared folders found, or the account has no shared folders."
    exit 0
}

Write-Host "[LastPass] Found $($folderIDs.Count) shared folder(s).`n" -ForegroundColor Green

# --- Step 3: Build results ---
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($folderID in $folderIDs) {
    $folder        = $data.$folderID
    $folderName    = $folder.sharedfoldername
    $score         = $folder.score
    $users         = $folder.users
    $numSites      = if ($folder.sitecount) { $folder.sitecount } else { "N/A" }

    Write-Host "Folder: $folderName (ID: $folderID, Entries: $numSites)" -ForegroundColor White

    if ($null -eq $users -or ($users | Measure-Object).Count -eq 0) {
        Write-Host "  (No users found for this folder)" -ForegroundColor DarkGray

        $results.Add([PSCustomObject]@{
            FolderID       = $folderID
            FolderName     = $folderName
            NumEntries     = $numSites
            Username       = "(none)"
            UserFullName   = ""
            AccessType     = ""
            IsAdmin        = ""
            IsReadOnly     = ""
            PasswordHidden = ""
        })
        continue
    }

    # Users can be an array or a hashtable depending on API version
    $userList = if ($users -is [System.Array]) { $users } else {
        $users | Get-Member -MemberType NoteProperty | ForEach-Object { $users.$($_.Name) }
    }

    foreach ($user in $userList) {
        $username      = $user.username
        $fullname      = if ($user.fullname) { $user.fullname } else { "" }
        $readonly      = [bool]($user.readonly -eq 1 -or $user.readonly -eq "1" -or $user.readonly -eq $true)
        $isAdmin       = [bool]($user.can_administer -eq 1 -or $user.can_administer -eq "1" -or $user.can_administer -eq $true)
        $hidePasswords = [bool]($user.give -eq 0 -or $user.give -eq "0")
        $accessType    = Get-AccessType -ReadOnly $readonly -Admin $isAdmin -HidePasswords $hidePasswords

        $color = if ($isAdmin) { "Yellow" } elseif ($readonly) { "Gray" } else { "Green" }
        Write-Host ("  User: {0,-35} Access: {1}" -f $username, $accessType) -ForegroundColor $color

        $results.Add([PSCustomObject]@{
            FolderID       = $folderID
            FolderName     = $folderName
            NumEntries     = $numSites
            Username       = $username
            UserFullName   = $fullname
            AccessType     = $accessType
            IsAdmin        = $isAdmin
            IsReadOnly     = $readonly
            PasswordHidden = $hidePasswords
        })
    }

    Write-Host ""
}

# --- Step 4: Summary ---
Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Total shared folders : $($folderIDs.Count)"
Write-Host "Total user-folder ACLs: $($results.Count)"

# --- Step 5: Optional CSV Export ---
if ($ExportCSV) {
    try {
        $results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
        Write-Host "`n[Export] Results saved to: $ExportCSV" -ForegroundColor Green
    } catch {
        Write-Error "Failed to export CSV: $_"
    }
}

# --- Step 6: Return results as objects (pipeline-friendly) ---
return $results