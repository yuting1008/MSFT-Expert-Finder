# Load application info
$appInfo = Get-Content "./AppInfo.json" | ConvertFrom-Json

# Connect to Microsoft Graph
try {
    Connect-MgGraph -ClientId $appInfo.ClientId -TenantId $appInfo.TenantId -CertificateName "CN=Certificate"
}
catch {
    Write-Host "Failed to connect to Microsoft Graph:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

# Load user data from CSV
$users = Import-Csv -Path "./UserSkills.csv"

foreach ($user in $users) {
    $upn = $user.UserPrincipalName

    # Update Skills
    if (![string]::IsNullOrEmpty($user.Skills)) {
        $skillsArray = $user.Skills -split "," | ForEach-Object { $_.Trim() }

        try {
            Write-Host "Updating skills for $upn..."
            Update-MgUser -UserId $upn -Skills $skillsArray
            Write-Host "Successfully updated skills for $upn." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to update skills for $upn." -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Update OfficeLocation
    if (![string]::IsNullOrEmpty($user.OfficeLocation)) {
        $officeLocation = $user.OfficeLocation.Trim()

        try {
            Write-Host "Updating OfficeLocation for $upn..."
            Update-MgUser -UserId $upn -OfficeLocation $officeLocation
            Write-Host "Successfully updated OfficeLocation for $upn." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to update OfficeLocation for $upn." -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "-------------------------"
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
