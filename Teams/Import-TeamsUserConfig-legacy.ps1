<# 
The purpose of this script is to import Microsoft Teams Enterprise Voice configuration from Excel

Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/05/powershell-script-for-exporting.html

Use: Get-Help Import-Excel for a full rundown on everything Import-Excel can do
#>

# Install and Import Teams and Excel Module
Install-Module -Name MicrosoftTeams
Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force # Use the following to get the commands available: Get-Command -Module ImportExcel | Select Name
Import-Module -Name ImportExcel
Import-Module -Name MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Set the path to the Excel file that will store Microsoft Teams User configuration
$path = "C:\Scripts\"
$excelFileName = "MyCompanyTeamsUserConfig.xlsx"
$fullPathAndFile = $path + $excelFileName

# Store Excel file in variable
$excelFile = Import-Excel -Path $fullPathAndFile

# GUI Grid View
# $excelFile | Out-GridView

# Loop through each record
foreach ($record in ($excelFile))
{
    $usernameUPN = $record."UserPrincipalName"
    $EnterpriseVoiceEnabled = $true
    $HostedVoiceMail = $true
    $extension = $record."LineURI"

    $PolicyName = $record."TenantDialPlan"

    $OnlineVoiceRoutingPolicy = $record."OnlineVoiceRoutingPolicy"

    <# For testing Excel import
    Write-Host "User Principal Name: $($usernameUPN)"
    Write-Host "Enterprise Voice is enabled: $($EnterpriseVoiceEnabled)"
    Write-Host "Hosted VoiceMail is enabled: $($HostedVoiceMail)"
    Write-Host "Extension is: $($extension)"
    Write-Host "Dial Plan is: $($PolicyName)"
    Write-Host "Routing Policy is: $($OnlineVoiceRoutingPolicy)"
    #>

# Enable user for Enterprise Voice with each row of data in the Excel spreadsheet
Set-CsUser `
-Identity $usernameUPN `
-EnterpriseVoiceEnabled $true `
-HostedVoiceMail $true `
-LineURI $extension

Grant-CsTenantDialPlan `
-PolicyName $PolicyName `
-Identity $usernameUPN

Grant-CsTenantDialPlan `
-PolicyName $PolicyName `
-Identity $usernameUPN

Grant-CsOnlineVoiceRoutingPolicy `
-Identity $usernameUPN `
-PolicyName $OnlineVoiceRoutingPolicy
}
# Get-CsOnlineUser -Identity $usernameUPN | FL *uri,ent*,hosted*,onlinevoicerout*,tenantdial*  <-- Use this to check configuration settings for individual users
