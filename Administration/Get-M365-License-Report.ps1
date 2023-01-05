<# 
The purpose of this script is to export Microsoft 365 licenses with Microsoft Graph PowerShell SDK as
Microsoft is deprecating the Azure AD PowerShell module and MS Online module in 2022

Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/07/using-microsoft-graph-powershell-sdk-to.html
#>

# Install-Module Microsoft.Graph -Scope CurrentUser

Install-Module Microsoft.Graph -Scope AllUsers 

# I felt that the cleanest and quickest way to merge two tables:
# 1. With the license details returned by Get-MgSubscribedSku
# 2. With the amount of licenses that have been purchased (this is stored in a multi-value table)
# ... was to use a SQL like join operation and a module I have used in the past that allows me to accomplish this will be used.
# Install JoinModule to join two arrays: https://www.powershellgallery.com/packages/Join/3.7.2
Install-Module -Name JoinModule
# Install-Script -Name Join

# Install the Excel module to export results to a spreadsheet
Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force
Import-Module -Name ImportExcel

# Set the path to the Excel file that will store license information
$path = "C:\Scripts\"
$excelFileName = "Microsoft 365 License Usage.xlsx"
$fullPathAndFile = $path + $excelFileName

# Export MS Teams configuration for every user into Excel file 
$excelFile = Export-Excel -Path $fullPathAndFile

# Use the following cmdlets to display the installed Microsoft.Graph module
## Find-Module Microsoft.Graph.* 
## Get-InstalledModule Microsoft.Graph 

# Connect to Microsoft Graph PowerShell, sign in and consent to the required scopes
# Note that the permissions required to use the cmdlet Get-MgSubscribedSku are outlined here: 
# https://docs.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0&tabs=powershell#permissions
Connect-MgGraph -Scopes "Organization.Read.All","Directory.Read.All","Organization.ReadWrite.All","Directory.ReadWrite.All"

# Use the following cmdlet to connect to a specific tenant
## Connect–MgGraph –TenantId <TenantId>

# Use the following cmdlets to display the cmdlets available
## Get-Command -Module Microsoft.Graph.Users  
## Get-Command -Module Microsoft.Graph.*
## Get-Command -Module Microsoft.Graph.* -Name *license*
## Get-Command -Module Microsoft.Graph.* | Where-Object Name -Like "*license*" 

# Use the following cmdlet to disconnect
## Disconnect-MgGraph 

# Use the following cmdlet to list users and list users who are enabled
## Get-MgUser -All
## Get-MgUser -Filter 'accountEnabled eq true' -All | Format-List UserPrincipalName,JobTitle,OfficeLocation

# Use the following cmdlet to list the available Microsoft 365 licenses
## Get-MgSubscribedSku

# Use the following cmdlets to expand the property PrepaidUnits to get the QTY of each license purchased
## Get-MgSubscribedSku | Select-Object -ExpandProperty PrepaidUnits

# Use the Get-MgSubscribedSku to get current license usage and store in an array
$licenseUsage = Get-MgSubscribedSku | Select-Object -Property SkuPartNumber,CapabilityStatus,@{Name="PrepaidUnits";expression={$_.PrepaidUnits.Enabled -join ";"}},ConsumedUnits,SkuId,AppliesTo 

# Import product friendly name from a csv https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference and remove duplicates
$productNames = Import-CSV C:\scripts\O365ProductNames.csv | Select-Object Product_Display_Name,String_Id | Sort-Object -Unique -Property String_Id

# Use JoinModule to perform a left join between 2 arrays
$joinedTable = $licenseUsage |LeftJoin $productNames -On SkuPartNumber -Equals String_Id -Property `
@{ SkuPartNumber = 'Left.SkuPartNumber' }, @{ FriendlyName = 'Right.Product_Display_Name' }, `
@{ PrepaidUnits = 'Left.PrepaidUnits' }, @{ ConsumedUnits = 'Left.ConsumedUnits' }, `
@{ CapabilityStatus = 'Left.CapabilityStatus' }, @{ AppliesTo = 'Left.AppliesTo' } 

# Add extra column to indicate whether licenses are fully utilized, under utilized, or needs to be reviewed
$joinedTable | ForEach-Object {
    if($_.ConsumedUnits -eq $_.PrepaidUnits) {
        $_ | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Fully Utilized'
    } elseif ($_.ConsumedUnits -lt $_.PrepaidUnits) {
        $_ | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Underutilized'
    } else {
		if($_.CapabilityStatus -eq "Suspended") {
			$_ | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Licenses are suspended'
		} else {
			$_ | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Needs Review'
		}
    }
}

# Export the joined table as an excel file
$joinedTable | Export-Excel -Path $fullPathAndFile -AutoSize -TableName LicenseUsage

# Display the joined table as a table
## $joinedTable | Format-Table

<# The following converts the joined table into an HTML file that can be emailed out 

$Header = @"
<style>
table {
	border-collapse: collapse;
    font-family: Calibri, sans-serif;
}
table td {
	padding: 5px;
}

table th {
	background-color: #4472C4;
	color: #ffffff;
	font-weight: bold;
	border: 1px solid #54585d;
	padding: 5px;
}
table tbody td {
	color: #636363;
	border: 1px solid #dddfe1;
}
table tbody tr {
	background-color: #f9fafb;
}
table tbody tr:nth-child(odd) {
	background-color: #ffffff;
}
</style>
"@

$joinedTable | Sort-Object -Property Status | ConvertTo-Html -Head $Header | Out-File -FilePath LicenseReport.html

Invoke-Expression C:\scripts\LicenseReport.html

#>
