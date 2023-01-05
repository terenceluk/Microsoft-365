<#

The purpose of this script is to create an inbound connector for Exchange Online that forces the defined incoming domains to require TLS.
This script will import a list of domains from an Excel spreadsheet with a column named domains

Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/02/using-powershell-to-configure-exchange.html

# The following is an example of how to use the New-InboundConnector Exchange Online cmdlet used to create an Inbound Connector to force incoming domains that belong to Company ABC to require TLS

New-InboundConnector `
-Name "Company ABC Domains to Company XYZ" `
-Enabled $true `
-ConnectorType Partner `
-SenderDomains "abc.com" `
-RequireTls $true

# The following are examples of how to get inbound connector domains configured

Get-InboundConnector "Company ABC Domains to Company XYZ" | FL
Get-InboundConnector | Where-Object {$_.Name  -Like "Company ABC Domains*"}

#>

# Install and import ExchangeOnline and Excel Modules
Install-Module -Name ExchangeOnlineManagement
Install-Module -Name ImportExcel # Use the following to get the commands available: Get-Command -Module ImportExcel | Select Name
Import-Module ImportExcel
Import-Module ExchangeOnline

# Connect to Exchange Online
Connect-ExchangeOnline

# Set the path to the Excel file with the domains
$path = "C:\Scripts\Domains-to-force-TLS.xlsx"

# Store Excel file in variable
$excelFile = Import-Excel -Path $path

# Store the imported Excel file as an array (note that the -SenderDomains accepts an array if a variable is passed and a string with domains separated with a comma if executed with domains typed out)
$senderDomainsToRequireTLS = @($excelFile)

# Create Inbound Connector to force incoming Aon domains to require TLS with the list of domains from the spreadsheet

New-InboundConnector `
-Name "Company ABC Domains to Company XYZ" `
-Enabled $true `
-ConnectorType Partner `
-SenderDomains $senderDomainsToRequireTLS."Domains" `
-RequireTls $true
