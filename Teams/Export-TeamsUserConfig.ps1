<# 
The purpose of this script is to export all users in AAD and their corresponding Microsoft Teams configuration parameters to Excel

Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/05/powershell-script-for-exporting.html

Use: Get-Help Export-Excel for a full rundown on everything Export-Excel can do
#>

# Install and Import Teams and Excel Module
Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force
Import-Module -Name ImportExcel
Install-Module -Name MicrosoftTeams
Import-Module -Name MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Set the path to the Excel file that will store Microsoft Teams User configuration
$path = "C:\Scripts\"
$excelFileName = "MyCompanyTeamsUserConfig.xlsx"
$fullPathAndFile = $path + $excelFileName

# Export MS Teams configuration for every user into Excel file 
$excelFile = Export-Excel -Path $fullPathAndFile

# Prompt to ask user configuration to export
Write-Host "`nPlease enter the number corresponding to the selection for what configuration parameters to export:"
Write-Host "Export all configuration [1] `nExport only Enterprise Voice configuration [2] `n"

# Retrieve selection from user
$whatToExport = Read-Host "Selection"

# Keep looping until a valid selection has been selected (2 options)
Do 
    {
        if ($whatToExport -gt 2 -or $whatToExport -eq 0){
            $whatToExport = Read-Host "Please enter a valid selection"
        }
    }
    Until ($whatToExport -gt 0 -and $whatToExport -lt 3)

if ($whatToExport -eq 1){
    Get-CsOnlineUser | Export-Excel -Path $fullPathAndFile -AutoSize -TableName UserConfig
    Write-Host "Exported all user configuration to $($fullPathAndFile)"
} elseif ($whatToExport -eq 2) {
    Get-CsOnlineUser | Select-Object AccountEnabled, DisplayName, UserPrincipalName, EnterpriseVoiceEnabled, IsSipEnabled, LineUri, `
    OnPremEnterpriseVoiceEnabled, OnlineVoiceRoutingPolicy, TenantDialPlan, UsageLocation `
    | Export-Excel -Path $fullPathAndFile -AutoSize -TableName UserConfig
    Write-Host "Exported user Teams Enterprise Voice configuration to $($fullPathAndFile)"
}
