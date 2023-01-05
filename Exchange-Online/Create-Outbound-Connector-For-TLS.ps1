<# 

The purpose of this script is to create an outbound connector for Exchange Online that forces the defined destination domains to require TLS and use a defined smarthost
This script will import a list of domains from an Excel spreadsheet with a column named domains

Refer to my blog post for more information: http://terenceluk.blogspot.com/2022/02/using-powershell-to-configure-exchange.html

# The following is an example of how to use the New-OutboundConnector Exchange Online cmdlet used to create an Outbound Connector to force TLS and the use of a smart host for the domain

New-OutboundConnector `
-Name "Company ABC to contoso.com" -Enabled $true -UseMXRecord $false -ConnectorType Partner -SmartHosts "mail.messaging.microsoft.com" `
-TlsSettings "CertificateValidation" -RecipientDomains "contos.com" -RouteAllMessagesViaOnPremises $false

# The following is an example of using the Set-OutboundConnector to configure an already created Outbound Connector with multiple smarthosts for the destination domain

Set-OutboundConnector "Company ABC to contos.com" -SmartHosts "mx1.contoso.com","mx2.contoso.com","mx3.contoso.com","mx3.contoso.com"

#>

# Install and import ExchangeOnline and Excel Modules
Install-Module -Name ExchangeOnlineManagement
Install-Module -Name ImportExcel # Use the following to get the commands available: Get-Command -Module ImportExcel | Select Name
Import-Module ImportExcel
Import-Module ExchangeOnline

# Connect to Exchange Online
Connect-ExchangeOnline

# Set the path to the Excel file
$path = "C:\Scripts\Destination-Domains.xlsx"

# Store Excel file in variable
$excelFile = Import-Excel -Path $path

# Loop through each record
foreach ($record in ($excelFile))

{
    $recipientDomains = $record."Destination Domains"
    $mxRecords = $record."MX Records"

# Create new Outbound Connector with each row of data in the Excel spreadsheet
New-OutboundConnector `
-Name "Company ABC to $recipientDomains" `
-RecipientDomains $recipientDomains `
-SmartHosts $mxRecords `
-Enabled $true `
-UseMXRecord $false `
-ConnectorType Partner `
-TlsSettings "CertificateValidation" `
-RouteAllMessagesViaOnPremises $false

}
