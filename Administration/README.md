# Administration

PowerShell scripts for license assignment, Microsoft Graph reporting, and Azure Function based administration workflows.

## Scope

This folder contains a mix of interactive scripts and HTTP-triggered Azure Function examples. Most files either assign licenses, export license consumption, or turn report data into HTML output.

## Scripts

| Script | Status | Purpose | Expected Inputs | Example Use |
| --- | --- | --- | --- | --- |
| `Assign AD Users Power BI Free license.ps1` | Current example | Finds Azure AD users who do not have the `POWER_BI_STANDARD` license and assigns it if sufficient units are available. | Azure AD connection, available Power BI Free capacity, users missing the target SKU. | Run after connecting to Azure AD to backfill Power BI Free licensing across unlicensed users. |
| `Assign AD Users with E5 a Azure P1.ps1` | Current example | Finds users with an E5 license who are missing Azure AD Premium P1 and adds the `AAD_PREMIUM` license. | Azure AD connection, E5 SKU present, available Azure AD Premium P1 capacity. | Run when auditing E5 users to ensure they also receive the expected Azure AD Premium P1 add-on. |
| `Assign-M365-Licenses.ps1` | Current example | Uses Microsoft Graph to report on assigned SKUs, replace an expired E5 SKU with a current one, and optionally add Copilot for Microsoft 365. | Interactive Graph sign-in with `Directory.AccessAsUser.All` and `Directory.ReadWrite.All`, target SKU IDs, optional Copilot SKU ID. | Use the built-in report first to identify SKU IDs, then run the assignment blocks after confirming the tenant’s current and expired license GUIDs. |
| `Get-CylanceDeviceReport.ps1` | Azure Function example | HTTP trigger that downloads a Cylance CSV device report and returns an HTML summary suitable for an email workflow. | Request body or query string containing `token`, Azure Function HTTP trigger context, Cylance report endpoint. | Trigger from Logic Apps or an HTTP client with `{ "token": "..." }` to generate an HTML device summary. |
| `Get-M365-License-Report-Function.ps1` | Azure Function example | HTTP trigger that connects to Microsoft Graph with app and certificate authentication and returns an HTML license usage report. | Request body or query string containing `tenant`, configured app registration, certificate thumbprint, access to product name CSV. | Call the function from an automation workflow to render an HTML Microsoft 365 license report for a tenant. |
| `Get-M365-License-Report-Function-v2.ps1` | Azure Function example | Updated HTTP trigger that retrieves a Graph access token programmatically and enriches license data with friendly product names. | Request body or query string containing `tenant`, managed identity or Azure context capable of retrieving a Graph token. | Use in Azure Automation or Function App scenarios where token-based Graph authentication is preferred over certificate setup. |
| `Get-M365-License-Report.ps1` | Current example | Interactive Microsoft Graph script that exports Microsoft 365 license usage data to Excel for manual review. | Graph sign-in, local output path such as `C:\Scripts`, access to a product name CSV file, `JoinModule`, `ImportExcel`. | Run locally to produce an Excel workbook that maps SKUs to friendly license names and utilization status. |

## Typical Workflow

1. Install the required modules.
2. Review the target tenant values, output paths, and SKU IDs.
3. Test the report-only sections first.
4. Run assignment or automation steps only after validating the report output.

## Module Requirements

- `AzureAD` for older Azure AD based license assignment scripts.
- `Microsoft.Graph` for Graph reporting and licensing.
- `ImportExcel` for Excel import or export workflows.
- `JoinModule` for license report data joining.

## Input Notes

- Azure Function scripts expect JSON in the request body or query string values.
- Report scripts assume supporting files such as product-name CSVs are already available.
- Assignment scripts use hardcoded SKU IDs or SKU names and should be reviewed against the target tenant.

## Caution

- Review hardcoded tenant names, file paths, app registration values, certificate thumbprints, SKU identifiers, and storage endpoints before running these scripts in another environment.
- Some scripts were written as blog-post examples and may need production hardening before operational use.