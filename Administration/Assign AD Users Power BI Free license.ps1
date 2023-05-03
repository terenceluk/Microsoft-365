# Install module
Install-Module AzureAD -Scope CurrentUser # if already installed you can skip this line
Import-Module AzureAD # if already imported you can skip this line

# connect
Connect-AzureAD # this will initiate a dialog to enter credentials

# Find the SkuID of the license we want to add - in this case the Power BI Free license, which is called "POWER_BI_STANDARD"
$PowerBIFreeSKUId = Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "POWER_BI_STANDARD"} | Select -ExpandProperty SkuId

# Check how many license units are available
$ConsumedPowerBIFreeLicenses = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "POWER_BI_STANDARD"}).ConsumedUnits
$EnabledPowerBIFreeLicenses = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "POWER_BI_STANDARD"} | Select -ExpandProperty PrepaidUnits).Enabled

# Get all AD Users Accounts without a Power BI Free License
$UsersWithoutPowerBIFree = Get-AzureADUser -All $true | Where-Object {($_.AssignedLicenses).SkuId -ne $PowerBIFreeSKUId -and $_.AccountEnabled -eq $true -and $_.UsageLocation -ne "" -and $_.UsageLocation -ne $null}
$UsersWithoutPowerBIFree.Count #show count of users without free license

# Check if there are enough licenses to cover the users without license, and stop if not the case
if ( $UsersWithoutPowerBIFree.Count -gt ($EnabledPowerBIFreeLicenses-$ConsumedPowerBIFreeLicenses)){
    throw "Not enough licenses! There are $EnabledPowerBIFreeLicenses licenses in total, $ConsumedPowerBIFreeLicenses already consumed.. not enough for $($UsersWithoutPowerBIFree.Count) users!"
}

# Create the objects we'll need
$Assignedlicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$Assignedlicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses

# Set the SkuId correctly
$Assignedlicense.SkuId = $PowerBIFreeSKUId

# Set the Power BI license as the license we want to add in the $licenses object
$Assignedlicenses.AddLicenses = $Assignedlicense

####
# THIS IS WHERE IT ALL HAPPENS :)
#
# For all the AD users without a license, call the Set-AzureADUserLicense cmdlet to set the license
$UsersWithoutPowerBIFree | Set-AzureADUserLicense -AssignedLicenses $Assignedlicenses
#
####

# Doublecheck: Get all AD Users Accounts without a Power BI Free License again
$UsersWithoutPowerBIFree = Get-AzureADUser -All $true | Where-Object {($_.AssignedLicenses).SkuId -ne $PowerBIFreeSKUId}
$UsersWithoutPowerBIFree.Count #show count of users without free license
