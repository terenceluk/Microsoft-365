<# 
The purpose of this script is to import Microsoft Teams Enterprise Voice configuration from Excel
Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/05/powershell-script-for-exporting.html
Use: Get-Help Import-Excel for a full rundown on everything Import-Excel can do

--Update July 8, 2022--
Set-CsUser has not been replaced with Set-CsPhoneNumberAssignment: https://docs.microsoft.com/en-us/powershell/module/teams/set-csphonenumberassignment?view=teams-ps

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
    $extension = $record."LineURI"
    # Remove any extensions that have "tel:" because Set-CsPhoneNumberAssignment does not accept it
    if ($extension -like "tel:*") {
        $extension = $extension -replace "tel:" -replace ""
    }

    $phoneNumberType = "DirectRouting"
    $PolicyName = $record."TenantDialPlan"

    $OnlineVoiceRoutingPolicy = $record."OnlineVoiceRoutingPolicy"

    <# For testing Excel import
    Write-Host "User Principal Name: $($usernameUPN)"
    Write-Host "Hosted VoiceMail is enabled: $($HostedVoiceMail)"
    Write-Host "Extension is: $($extension)"
    Write-Host "Dial Plan is: $($PolicyName)"
    Write-Host "Routing Policy is: $($OnlineVoiceRoutingPolicy)"
    #>

    # Test if extension is null and only process if it is NOT null
    if ($extension) {
        # Enable user for Enterprise Voice with each row of data in the Excel spreadsheet
        Set-CsPhoneNumberAssignment `
        -Identity $usernameUPN `
        -PhoneNumber $extension `
        -PhoneNumberType $phoneNumberType

        Grant-CsTenantDialPlan `
        -PolicyName $PolicyName `
        -Identity $usernameUPN

        Grant-CsTenantDialPlan `
        -PolicyName $PolicyName `
        -Identity $usernameUPN

        Grant-CsOnlineVoiceRoutingPolicy `
        -Identity $usernameUPN `
        -PolicyName $OnlineVoiceRoutingPolicy
    } else {
        # Do nothing and skip
    }   
}
# Get-CsOnlineUser -Identity $usernameUPN | FL *uri*,*voice*,*dial*,onlinevoicerout*,tenantdial*  <-- Use this to check configuration settings for individual users

# Check who is assigned an extension
# Get-CsOnlineUser | Where-Object LineUri -Like "*342*" | FL DisplayName,userPrincipalName,*uri,ent*,hosted*,onlinevoicerout*,tenantdial*
