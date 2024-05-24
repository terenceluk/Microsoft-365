# Verify Microsoft.Graph is installed
Get-InstalledModule Microsoft.Graph

# Install Microsoft.Graph
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect and authenticate interactively with MgGraph
$RequiredScopes = @('Directory.AccessAsUser.All', 'Directory.ReadWrite.All') # Required scope permissions for retriving and assigning licenses
Connect-MgGraph -Scopes $RequiredScopes -NoWelcome

# Use device code for authenticating on another machine 
# Connect-MgGraph -Scopes $RequiredScopes -UseDeviceAuthentication

<#

The following block will generate a Microsoft 365 license report with GUI and console text. Use this to retrieve the skuID required for assigning licenses

#>

[array]$Skus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, @{Name = 'PrepaidUnits'; Expression = { $_.PrepaidUnits.enabled } }, ConsumedUnits
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Sku in $Skus) {
       $Command = 'Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq ' + $Sku.SkuId + ')" -All'
       [array]$SkuUsers = Invoke-Expression $Command
       $ReportLine = [PSCustomObject][Ordered]@{
              Sku                 = $Sku.SkuId
              Product             = $Sku.SkuPartNumber
              'Prepaid Units'     = $Sku.PrepaidUnits
              'Consumed Units'    = $Sku.ConsumedUnits
              'Calculated Units'  = $SkuUsers.Count
              'Assigned accounts' = $SkuUsers.UserPrincipalName -Join ", " 
       }
       $Report.Add($ReportLine) 
} # End ForEach

# Report in GUI Grid display
$Report | Sort-Object Product | Out-GridView

# Report in table format text displayed in console
$Report | Sort-Object Product | Format-Table

<# 

Use the following to swap out users assigned with an expired Microsoft 365 E5 with a valid E5. 
The skuID representing the licenses retrieved from the previous report will be used.

#>

# Get all users with assigned license SKU - E5 current
Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq cd2925a3-5076-4233-8931-638a8c94f773)" -All

# Get all users with assigned license SKU - E5 expired
Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq 26d45bd9-adf1-46cd-a9e1-51e9a5524128)" -All

# Assign users with expired E5 license with current E5 license - note that -AddLicenses expects a hashtable
[array]$userList = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq 26d45bd9-adf1-46cd-a9e1-51e9a5524128)" -All # Get all users who have expired E5
ForEach ($user in $userList) {
       Set-MgUserLicense -UserId $user.UserPrincipalName -Addlicenses @{SkuId = 'cd2925a3-5076-4233-8931-638a8c94f773' } -RemoveLicenses @() # Add current E5 to users
}

# Remove expired E5 licenses from users who now have current E5 - note that -RemoveLicenses expects an array and not hashtable
[array]$userList = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq cd2925a3-5076-4233-8931-638a8c94f773)" -All # Get all users who have current E5
ForEach ($user in $userList) {
       Set-MgUserLicense -UserId $user.UserPrincipalName -RemoveLicenses @('26d45bd9-adf1-46cd-a9e1-51e9a5524128') -AddLicenses @{} # Remove expired E5 from users
}

<# 

Use the following to retrieve all the users who have a Microsoft 365 E5 license assigned and assign them with a Copilot license.

#>

# Get users who have current assign Copilot for Microsoft 365
[array]$userList = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq cd2925a3-5076-4233-8931-638a8c94f773)" -All # Get all users who have current E5
ForEach ($user in $userList) {
       Set-MgUserLicense -UserId $user.UserPrincipalName -Addlicenses @{SkuId = '639dec6b-bb19-468b-871c-c5c441c4b0cb' } -RemoveLicenses @() # Add Copilot for Microsoft 365
}
