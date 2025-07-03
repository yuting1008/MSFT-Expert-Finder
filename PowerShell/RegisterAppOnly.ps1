param(
  [Parameter(Mandatory=$true,
  HelpMessage="The friendly name of the app registration")]
  [String]
  $AppName,

  [Parameter(Mandatory=$true,
  HelpMessage="The file path to your public key file")]
  [String]
  $CertPath,

  [Parameter(Mandatory=$false,
  HelpMessage="Your Azure Active Directory tenant ID")]
  [String]
  $TenantId,

  [Parameter(Mandatory=$false)]
  [Switch]
  $StayConnected = $false
)

# Graph permissions constants
$graphResourceId = "00000003-0000-0000-c000-000000000000"
$UserReadAll = @{
  Id="df021288-bdef-4463-88db-98f22de89214"
  Type="Role"
}
$UserReadWriteAll = @{
  Id = "741f803b-c850-494e-b5df-cde7c675a1ca"
  Type = "Role"
}
$SitesReadWriteAll = @{
  Id = "9492366f-7969-46a4-8d15-ed1a20078fff"
  Type = "Role"
}

# Requires an admin
if ($TenantId)
{
  Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -TenantId $TenantId
}
else
{
  Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read"
}

# Get context for access to tenant ID
$context = Get-MgContext

# Load cert
$resolvedCertPath = Resolve-Path -Path $CertPath
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($resolvedCertPath.Path)
Write-Host -ForegroundColor Cyan "Certificate loaded from $resolvedCertPath"

# Create app registration
$appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
 -Web @{ RedirectUris="http://localhost"; } `
 -RequiredResourceAccess @{ ResourceAppId=$graphResourceId; ResourceAccess=$UserReadAll, $UserReadWriteAll, $SitesReadWriteAll } `
 -AdditionalProperties @{} -KeyCredentials @(@{ Type="AsymmetricX509Cert"; Usage="Verify"; Key=$cert.RawData })
Write-Host -ForegroundColor Cyan "App registration created with app ID" $appRegistration.AppId

# Create corresponding service principal
New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
Write-Host -ForegroundColor Cyan "Service principal created"
Write-Host
Write-Host -ForegroundColor Green "Success"
Write-Host

# Generate Connect-MgGraph command
$connectGraph = "Connect-MgGraph -ClientId """ + $appRegistration.AppId + """ -TenantId """`
 + $context.TenantId + """ -CertificateName """ + $cert.SubjectName.Name + """"
Write-Host -ForeGroundColor Cyan "After providing admin consent, you can use the following values with Connect-MgGraph for app-only:"
Write-Host $connectGraph

if ($StayConnected -eq $false)
{
  Disconnect-MgGraph
  Write-Host "Disconnected from Microsoft Graph"
}
else
{
  Write-Host
  Write-Host -ForegroundColor Yellow "The connection to Microsoft Graph is still active. To disconnect, use Disconnect-MgGraph"
}

# Export ClientId and TenantId
$output = @{
    ClientId = $appRegistration.AppId
    TenantId = $context.TenantId
}
$output | ConvertTo-Json | Out-File -FilePath "./AppInfo.json" -Encoding UTF8
Write-Host -ForegroundColor Green "ClientId and TenantId are saved to AppInfo.json"