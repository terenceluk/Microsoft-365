# Microsoft-365

![PowerShell](https://img.shields.io/badge/PowerShell-Automation-5391FE?logo=powershell&logoColor=white)
![Microsoft 365](https://img.shields.io/badge/Microsoft%20365-Administration-107C10)
![Documentation](https://img.shields.io/badge/Docs-Expanded-0F6CBD)

Collection of Microsoft 365 administration scripts and HTML notification templates focused on licensing, Exchange Online mail flow, Microsoft Teams voice configuration, and HTML reporting workflows.

## Overview

This repository is a practical toolbox rather than a packaged module. Most scripts are standalone examples that assume you will review and update tenant-specific values before use.

## Quick Start

1. Review the folder README that matches the workload you want to work on.
2. Install the required PowerShell modules for that script.
3. Update hardcoded paths, tenant IDs, connector names, policy names, client names, or storage endpoints.
4. Test in a non-production tenant or with a limited dataset first.

## Repository Guide

| Folder | Focus | Documentation |
| --- | --- | --- |
| `Administration` | Licensing automation, Microsoft Graph reporting, Azure Function helpers | [Administration/README.md](Administration/README.md) |
| `Exchange-Online` | Inbound and outbound TLS connector examples | [Exchange-Online/README.md](Exchange-Online/README.md) |
| `HTML` | HTML email templates for approvals and file activity notifications | [HTML/README.md](HTML/README.md) |
| `Teams` | Microsoft Teams export/import automation for Enterprise Voice | [Teams/README.md](Teams/README.md) |

## What To Expect

- Most PowerShell scripts are designed to be run manually, not imported as functions.
- Several scripts reference blog posts from the repository author for background and deployment details.
- Some files are current operational examples, while others are clearly marked as legacy or outdated reference material.

## Common Dependencies

Depending on the script, this repository uses the following PowerShell modules:

- `Microsoft.Graph`
- `MicrosoftTeams`
- `ExchangeOnlineManagement`
- `ImportExcel`
- `JoinModule`
- `AzureAD`

## Configuration Notes

- Review local paths such as `C:\Scripts\...` before running any script.
- Review tenant-specific values such as client names, app IDs, certificate thumbprints, storage URLs, connector names, and voice policy names.
- For Azure Function examples, verify request body format, managed identity or certificate authentication, and output bindings.

## Naming Notes

- Original script filenames are preserved to match the source material and linked blog posts.
- Documentation uses consistent labels such as `current`, `legacy`, and `outdated` to clarify script status without renaming files in the repository.

## Contributing

Contributions should keep the repository practical and low-friction:

- Prefer small, focused changes.
- Preserve original script behavior unless the change explicitly fixes a bug.
- Update the relevant folder README when adding or significantly changing a script.
- Document tenant-specific assumptions and required inputs rather than hiding them in the code.
