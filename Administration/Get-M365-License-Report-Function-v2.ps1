<# 
The purpose of this script is to export Microsoft 365 licenses with Microsoft Graph PowerShell SDK as
Microsoft is deprecating the Azure AD PowerShell module and MS Online module in 2022
Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/12/updated-create-automated-report-for.html
#>

using namespace System.Net

<# Input bindings are passed in via param block - we'll be passing the following JSON code into the HTTP Body:
{
  "tenant": "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
#>
param($Request, $TriggerMetadata)

Function Get-LicenseUsage
{

# Create the HTML header
$Header = @"
<style>
BODY {font-family:Calibri;}
table {border-collapse: collapse; font-family: Calibri, sans-serif;}
table td {padding: 5px;}
table th {background-color: #4472C4; color: #ffffff; font-weight: bold;	border: 1px solid #54585d; padding: 5px;}
table tbody td {color: #636363;	border: 1px solid #dddfe1;}
table tbody tr {background-color: #f9fafb;}
table tbody tr:nth-child(odd) {background-color: #ffffff;}
</style>
"@

# Create the HTML Body
$Body = @"
<title>Office 365 / Microsoft 365 License Report</title>
<h1>License Usage Report</h1>
<h2>Client: Contoso Limited</h2>
"@

# Use the Get-MgSubscribedSku to get current license usage and store in an array
$licenseUsage = Get-MgSubscribedSku | Select-Object -Property SkuPartNumber,CapabilityStatus,@{Name="PrepaidUnits";expression={$_.PrepaidUnits.Enabled -join ";"}},ConsumedUnits,SkuId,AppliesTo 

# Import product friendly name from a csv from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference and remove duplicates
# and placed into a Storage Account for download, remove duplicates and store in an array
$productNames = @()
$FilePath= "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$productNames = Invoke-RestMethod $FilePath | ConvertFrom-Csv -Delimiter ',' -Header 'Product_Display_Name','String_Id','GUID','Service_Plan_Name','Service_Plan_Id','Service_Plans_Included_Friendly_Names' | Select-Object Product_Display_Name,String_Id | Sort-Object -Unique -Property String_Id

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

# Create the joined table and convert to HTML format
$joinedTableHTML = $joinedTable | Sort-Object -Property Status | ConvertTo-Html

# Combine Header, Body and HTML elements from $joinedTableHTML for output
$Header + $Body + $joinedTableHTML

}

# Function that will generate an access token that can be passed down to Connect-MGGraph -AccessToken for authentication
Function Get-AzToken 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $ResourceUri,
        [Switch]$AsHeader
    ) 
    $Context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
    $Token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, $ResourceUri).AccessToken
    if ($AsHeader) {
        return @{Headers = @{Authorization = "Bearer $Token" } }
    }
    return $Token
    
}

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request - we're retrieving the tenant ID passed in the body
$tenantID = $Request.Query.tenant
if (-not $tenantID) {
    $tenantID = $Request.Body.tenant
}

# Use Get-AzToken function to retrieve MSI token
$Token = Get-AzToken -ResourceUri 'https://graph.microsoft.com/'

# Use the Connect-MgGraph command to initiate a connection to Microsoft Graph
Connect-MgGraph -AccessToken $Token

# Use the function defined above to retrieve the license table in HTML format (includes header and body)
$HTML = Get-LicenseUsage

# Set the HTTP status code
$status = [HttpStatusCode]::OK

# Write the output data in HTML format
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    headers = @{'content-type' = 'text/html'}
    StatusCode = $status
    Body = $HTML
})
