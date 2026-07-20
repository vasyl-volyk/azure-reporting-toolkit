
param (
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$OutputPath = ".\CrossTenantAccessReport_Detailed.csv"
)



# 1. Validation & Connection
$SecuredPassword = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecuredPassword

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
# Added Organization.Read.All to help resolve names
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

# 2. Get Default Settings
$DefaultPolicy = Get-MgPolicyCrossTenantAccessPolicyDefault

# 3. Logic Function for Inheritance
function Get-EffectiveValue {
    param($PartnerValue, $DefaultValue)
    if ($null -ne $PartnerValue) { return $PartnerValue } else { return $DefaultValue }
}

# 4. Get Partner Settings
Write-Host "Fetching Partner Settings and resolving names..." -ForegroundColor Yellow
$Partners = Get-MgPolicyCrossTenantAccessPolicyPartner -All -Property "TenantId", "InboundTrust", "B2bCollaborationInbound", "B2bCollaborationOutbound"

$FinalReport = foreach ($Partner in $Partners) {
    
    # Resolve the Name manually since -Property 'DisplayName' fails
    Write-Host "Resolving name for Tenant: $($Partner.TenantId)..." -ForegroundColor Gray
    $ResolvedName = "Unknown/Hidden Tenant"
    try {
        # This attempts to find the name of the external organization
        $ExternalOrg = Get-MgTenantDetail -TenantId $Partner.TenantId -ErrorAction SilentlyContinue
        if ($ExternalOrg.DisplayName) { $ResolvedName = $ExternalOrg.DisplayName }
    } catch {
        $ResolvedName = "Tenant: $($Partner.TenantId)"
    }

    $HasCustomTrust = $null -ne $Partner.InboundTrust.IsMfaAccepted

    [PSCustomObject]@{
        PartnerName           = $ResolvedName
        PartnerTenantId       = $Partner.TenantId
        TrustMfa              = Get-EffectiveValue $Partner.InboundTrust.IsMfaAccepted $DefaultPolicy.InboundTrust.IsMfaAccepted
        TrustCompliantDevice  = Get-EffectiveValue $Partner.InboundTrust.IsCompliantDeviceAccepted $DefaultPolicy.InboundTrust.IsCompliantDeviceAccepted
        TrustHybridAzureAD    = Get-EffectiveValue $Partner.InboundTrust.IsHybridAzureAdJoinedDeviceAccepted $DefaultPolicy.InboundTrust.IsHybridAzureAdJoinedDeviceAccepted
        InboundAccessType     = Get-EffectiveValue $Partner.B2bCollaborationInbound.UsersAndGroups.AccessType $DefaultPolicy.B2bCollaborationInbound.UsersAndGroups.AccessType
        OutboundAccessType    = Get-EffectiveValue $Partner.B2bCollaborationOutbound.UsersAndGroups.AccessType $DefaultPolicy.B2bCollaborationOutbound.UsersAndGroups.AccessType
        PolicyStatus          = if ($HasCustomTrust) { "Custom Overrides" } else { "Inherited (Default)" }
    }
}

# 5. Add Global Default
$DefaultRow = [PSCustomObject]@{
    PartnerName           = "!!! ORGANIZATION_DEFAULT !!!"
    PartnerTenantId       = $TenantId
    TrustMfa              = $DefaultPolicy.InboundTrust.IsMfaAccepted
    TrustCompliantDevice  = $DefaultPolicy.InboundTrust.IsCompliantDeviceAccepted
    TrustHybridAzureAD    = $DefaultPolicy.InboundTrust.IsHybridAzureAdJoinedDeviceAccepted
    InboundAccessType     = $DefaultPolicy.B2bCollaborationInbound.UsersAndGroups.AccessType
    OutboundAccessType    = $DefaultPolicy.B2bCollaborationOutbound.UsersAndGroups.AccessType
    PolicyStatus          = "BASE DEFAULT"
}

$OutputData = @($DefaultRow) + $FinalReport
$OutputData | Export-Csv -Path $OutputPath -NoTypeInformation
$OutputData | Format-Table -AutoSize

Disconnect-MgGraph

