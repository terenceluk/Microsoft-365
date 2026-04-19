# Exchange-Online

PowerShell examples for configuring Exchange Online mail flow connectors that enforce TLS for defined partner domains.

## Scripts

| Script | Status | Purpose | Expected Inputs | Example Use |
| --- | --- | --- | --- | --- |
| `Create-Inbound-Connector-For-TLS.ps1` | Current example | Imports sender domains from Excel and creates an inbound connector that requires TLS for those domains. | Excel file at `C:\Scripts\Domains-to-force-TLS.xlsx` with a `Domains` column, Exchange Online admin access. | Use when a partner’s inbound mail should only be accepted over TLS for a defined set of sender domains. |
| `Create-Outbound-Connector-For-TLS.ps1` | Current example | Reads destination domains and smart hosts from Excel and creates outbound connectors that require TLS. | Excel file at `C:\Scripts\Destination-Domains.xlsx` with `Destination Domains` and `MX Records` columns, Exchange Online admin access. | Use when outbound mail to partner domains must be routed through specific smart hosts with certificate validation enabled. |

## Typical Workflow

1. Prepare the Excel workbook with the expected column names.
2. Install `ExchangeOnlineManagement` and `ImportExcel`.
3. Update connector names, company names, and partner values in the script.
4. Connect to Exchange Online and create the connectors.
5. Verify the resulting connectors with `Get-InboundConnector` or `Get-OutboundConnector`.

## Module Requirements

- `ExchangeOnlineManagement`
- `ImportExcel`

## Input Notes

- The inbound connector script expects an array of domains and creates one connector covering all imported values.
- The outbound connector script loops row-by-row and creates one connector per record.
- Both scripts assume local Excel files and example naming that should be reviewed before use.

## Example Excel Layouts

### Inbound Connector Workbook

File expected by `Create-Inbound-Connector-For-TLS.ps1`:

| Domains |
| --- |
| partner1.example |
| partner2.example |
| subsidiary.partner1.example |

Notes:

- The column header should be `Domains`.
- Each row represents one sender domain to include in the inbound connector.

### Outbound Connector Workbook

File expected by `Create-Outbound-Connector-For-TLS.ps1`:

| Destination Domains | MX Records |
| --- | --- |
| contoso.com | mail.messaging.microsoft.com |
| fabrikam.com | mx1.fabrikam.com |
| adatum.com | smarthost.adatum.com |

Notes:

- The column headers should be `Destination Domains` and `MX Records`.
- Each row creates a separate outbound connector.
- If you need multiple smart hosts, verify whether the workbook value should be split before use or whether you want to adjust the script to support multiple values per row.

## Caution

- These scripts include example company names, connector names, file paths, and smart host values that should be updated for the target tenant.
- Validate connector settings in a test scenario before using them to enforce mail flow controls in production.