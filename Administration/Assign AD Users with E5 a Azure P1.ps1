# install module
Install-Module AzureAD -Scope CurrentUser # if already installed you can skip this line
Import-Module AzureAD # if already imported you can skip this line

# connect
Connect-AzureAD # this will initiate a dialog to enter credentials

# Find the SkuID of the license we want to add - in this case the AAD P1 license, which is called "AAD_PREMIUM"
$AADP1SKUId = Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "AAD_PREMIUM"} | Select -ExpandProperty SkuId

# Find the SkuID of the license of the E5 license we want to check for because everyone with an E5 should have AAD P1
$O365E5SKUId = Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPREMIUM_NOPSTNCONF"} | Select -ExpandProperty SkuId

# Check how many license units are available
$ConsumedAADP1Licenses = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "AAD_PREMIUM"}).ConsumedUnits
$EnabledAADP1Licenses = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "AAD_PREMIUM"} | Select -ExpandProperty PrepaidUnits).Enabled

# Get all AD Users Accounts without an AAD P1 but has an E5
$UsersWithoutAADP1 = Get-AzureADUser -All $true | Where-Object {($_.AssignedLicenses).SkuId -ne $AADP1SKUId -and ($_.AssignedLicenses).SkuId -eq $O365E5SKUId -and $_.AccountEnabled -eq $true -and $_.UsageLocation -ne "" -and $_.UsageLocation -ne $null}
$UsersWithoutAADP1.Count #show count of users without an AAD P1 license but has an E5

# Check if there are enough licenses to cover the users without license, and stop if not the case
if ( $UsersWithoutAADP1.Count -gt ($EnabledAADP1Licenses-$ConsumedAADP1Licenses)){
    throw "Not enough licenses! There are $EnabledAADP1Licenses licenses in total, $ConsumedAADP1Licenses already consumed.. not enough for $($UsersWithoutAADP1.Count) users!"
}

# Create the objects we'll need
$Assignedlicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$Assignedlicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses

# Set the SkuId correctly
$Assignedlicense.SkuId = $AADP1SKUId

# Set the AADP1AAD P1 license as the license we want to add in the $licenses object
$Assignedlicenses.AddLicenses = $Assignedlicense

####
# THIS IS WHERE IT ALL HAPPENS :)
#
# For all the AD users without a license, call the Set-AzureADUserLicense cmdlet to set the license - note that this adds a license, not overwrite
$UsersWithoutAADP1 | Set-AzureADUserLicense -AssignedLicenses $Assignedlicenses
#
####

# Doublecheck: Get all AD Users Accounts without a AADP1AAD P1 Free License again
$UsersWithoutAADP1 = Get-AzureADUser -All $true | Where-Object {($_.AssignedLicenses).SkuId -ne $AADP1SKUId -and ($_.AssignedLicenses).SkuId -eq $O365E5SKUId}
$UsersWithoutAADP1.Count #show count of users without free license
