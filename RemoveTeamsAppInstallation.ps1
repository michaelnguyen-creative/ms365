# Two modules to manage Teams resources
# MicrosoftTeams powershell module OR MSGraph.Teams module
# Recommended approach: MSTeams module (easier to use)

param (
    [string]$AppInstallationName = ""  # Programming principle: Always set a default value
)

# Install latest MSTeams powershell module (prerelease)
# Overview: https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-overview
# Reference: https://learn.microsoft.com/en-us/powershell/module/teams/?view=teams-ps
# Check if the MicrosoftTeams module is already installed
$moduleName = "MicrosoftTeams"
$moduleInstalled = Get-Module -ListAvailable -Name $moduleName

if (-not $moduleInstalled) {
    Write-Host "MicrosoftTeams module not found. Installing..."
    Install-Module -Name $moduleName -Force -AllowClobber -AllowPrerelease -Scope CurrentUser
}
else {
    Write-Host "MicrosoftTeams module is already installed. Moving to next actions..."
}

# Exit 1 if AppInstallationName = ""
if ($AppInstallationName -eq "") {
    Write-Host "Error: -AppInstallationName is required"
    exit 1
}

# Import & connect to MicrosoftTeams
# Import-Module $moduleName
# Connect to MicrosoftTeams (interactive). This will open the browser
# Connect-MicrosoftTeams

# Test TeamsApp
# Get-TeamsAppInstallation -userId <userId>

# Read terminal param
# Load CSV file in data/ (get the first file in data/)
$users = Import-Csv -Path .\data\*.csv

# Data sample
# PS /home/luannvm/GMD/gmdcorp-ms365> foreach ($row in $Data) { write-host $row }
# @{userPrincipalName=...; displayName=...}
# @{userPrincipalName=...; displayName=...}
# ...


# Get viva connection appId
$AppId = Get-TeamsApp | Where-Object { $_.DisplayName -like "*$AppInstallationName*" } | Select-Object -ExpandProperty Id

# Exit with status code 1 if AppInstallationId not found, message (AppInstallation: $name not found)
if (!$AppId) {
    Write-Host "App Installation: $($AppInstallationName) not found"
    exit 1
}

# Loop through each row in CSV
foreach ($row in $users) {
    # break line
    Write-Host ""
    Write-Host "---START---"
    # Get user id by email
    $user = Get-CsOnlineUser -Identity $row.userPrincipalName | Select-Object -Property Identity
    # write-host "$user"
    if ($user) {
        try {
            # Attempt to remove Teams app installation for the user
            Remove-TeamsAppInstallation -UserId $($user.Identity) -AppId $AppId
            # Print message <email, appUninstalled>
            Write-Host "$($AppInstallationName), $($row.userPrincipalName)"
            Write-Host "SUCCESS"
        }
        catch {
            # Handle the error and print a detailed message
            Write-Host "$($AppInstallationName), $($row.userPrincipalName)"
            Write-Host "Error: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "$($row.userPrincipleName), UserId not found"
    }
    # break line
    Write-Host "---END---"
    Write-Host ""
}