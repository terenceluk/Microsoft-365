# Teams

PowerShell scripts for exporting, importing, and assigning Microsoft Teams user configuration with a focus on Enterprise Voice.

## Scripts

| Script | Status | Purpose | Expected Inputs | Example Use |
| --- | --- | --- | --- | --- |
| `Enable-Teams-Users-For-Enterprise-Voice-Outdated.ps1` | Outdated reference | Enables a single Teams user for Enterprise Voice, assigns a Line URI, and applies dial plan and voice routing policies. | `-EmailAddress`, `-DID`, Azure Automation credential named `Teams Administrator Account`, Automation variable `Tenant ID`, hardcoded dial plan and voice routing policy values. | Use only as a reference for an older Azure Automation based Enterprise Voice provisioning flow. |
| `Export-TeamsUserConfig.ps1` | Current example | Exports Teams user configuration to Excel, either full configuration or Enterprise Voice focused fields only. | Teams admin connection, local output path such as `C:\Scripts\MyCompanyTeamsUserConfig.xlsx`, interactive selection `1` or `2`. | Run before bulk migration or audit work to capture the current Teams configuration of users into Excel. |
| `Import-TeamsUserConfig-legacy.ps1` | Legacy example | Reads Teams configuration from Excel and applies Enterprise Voice settings using the older `Set-CsUser` based command set. | Excel workbook with columns such as `UserPrincipalName`, `LineURI`, `TenantDialPlan`, and `OnlineVoiceRoutingPolicy`. | Use only when maintaining an older process that still depends on the legacy Enterprise Voice assignment approach. |
| `Import-TeamsUserConfig.ps1` | Current example | Reads Teams configuration from Excel and applies phone number, dial plan, and voice routing assignments with `Set-CsPhoneNumberAssignment`. | Excel workbook with `UserPrincipalName`, `LineURI`, `TenantDialPlan`, and `OnlineVoiceRoutingPolicy`, Teams admin connection. | Use after exporting or preparing a validated workbook for bulk Teams voice assignment changes. |

## Typical Workflow

1. Export the current Teams configuration to Excel.
2. Review and update the workbook.
3. Use the current import script to apply validated changes.
4. Verify user settings with `Get-CsOnlineUser` after import.

## Module Requirements

- `MicrosoftTeams`
- `ImportExcel`

## Input Notes

- The current import and export workflow expects a workbook named `MyCompanyTeamsUserConfig.xlsx` under `C:\Scripts` unless you change the path in the script.
- `Export-TeamsUserConfig.ps1` prompts for either full export or Enterprise Voice-only export.
- `Import-TeamsUserConfig.ps1` strips the `tel:` prefix before calling `Set-CsPhoneNumberAssignment`.

## Example Excel Layouts

### Recommended Enterprise Voice Import Layout

Minimum columns used by `Import-TeamsUserConfig.ps1`:

| UserPrincipalName | LineURI | TenantDialPlan | OnlineVoiceRoutingPolicy |
| --- | --- | --- | --- |
| alex.wilber@contoso.com | tel:+14165550100;ext=100 | Toronto | Toronto |
| megan.bowen@contoso.com | tel:+14165550101;ext=101 | Toronto | Toronto |
| diego.siciliani@contoso.com | tel:+14165550102;ext=102 | Montreal | Canada-DR |

Notes:

- The script reads `UserPrincipalName`, `LineURI`, `TenantDialPlan`, and `OnlineVoiceRoutingPolicy`.
- `Import-TeamsUserConfig.ps1` removes the `tel:` prefix before calling `Set-CsPhoneNumberAssignment`, so values exported directly from Teams are acceptable.
- Blank `LineURI` values are skipped by the current import script.

### Exported Enterprise Voice Workbook Layout

When `Export-TeamsUserConfig.ps1` is run with selection `2`, it exports these columns:

| AccountEnabled | DisplayName | UserPrincipalName | EnterpriseVoiceEnabled | IsSipEnabled | LineUri | OnPremEnterpriseVoiceEnabled | OnlineVoiceRoutingPolicy | TenantDialPlan | UsageLocation |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| True | Alex Wilber | alex.wilber@contoso.com | True | True | tel:+14165550100;ext=100 | False | Toronto | Toronto | CA |

Notes:

- This export is a good starting point for the current and legacy import workflows.
- PowerShell property names are case-insensitive, so `LineUri` from the export still maps correctly when the import script looks for `LineURI`.

### Legacy Import Layout

`Import-TeamsUserConfig-legacy.ps1` expects the same core workbook shape as the current import process:

| UserPrincipalName | LineURI | TenantDialPlan | OnlineVoiceRoutingPolicy |
| --- | --- | --- | --- |
| alex.wilber@contoso.com | tel:+14165550100;ext=100 | Toronto | Toronto |

Notes:

- The legacy script uses `Set-CsUser` rather than `Set-CsPhoneNumberAssignment`.
- Use the legacy path only when you intentionally need the older assignment model.

## Caution

- Review hardcoded values such as dial plans, routing policies, credential sources, and local paths before running these scripts.
- The outdated and legacy scripts are kept for historical reference and should not be treated as the default implementation path.