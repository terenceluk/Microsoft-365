<#
    .DESCRIPTION
        Obtain email address and DID and enable user for Teams enterprise voice
    .NOTES
        AUTHOR: Terence Luk
        LASTEDIT: June 4, 2022
        Blog post: http://terenceluk.blogspot.com/2022/06/microsoft-teams-configuration.html
#>
Param
(
  [Parameter (Mandatory= $true)]
  [String] $EmailAddress,
  [String] $DID
)
# Connect to Azure with App Registration Service Principal Secret

# Retrieve the App Registration credential (App ID and secret)
$spCredential = Get-AutomationPSCredential -Name 'Teams Administrator Account'

# Retrieve the Azure AD tenant ID
$tenantID = Get-AutomationVariable -Name 'Tenant ID'

Connect-MicrosoftTeams -Credential $spCredential

# Declare variables with for dial plan and voice routing policy (can also be passed by form)
$dialPlan = "Toronto"
$voiceRoutingPolicy = "Toronto"

# Convert DID to e164 format with ;ext= extension format
$extension = $DID.SubString($DID.length - 3, 3)
$e164 = "tel:+1" + $DID + ";ext=" + $extension
Write-Host "Teams DID value:" $e164

Set-CsUser -Identity $EmailAddress -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -LineURI $e164
Grant-CsTenantDialPlan -PolicyName $dialPlan -Identity (Get-CsOnlineUser $EmailAddress).SipAddress
Grant-CsOnlineVoiceRoutingPolicy -Identity $EmailAddress -PolicyName $voiceRoutingPolicy
Get-CsOnlineUser -Identity $EmailAddress | FL *uri*,*voice*,*dial*
