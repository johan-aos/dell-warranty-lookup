# Dell Warranty Lookup (TechDirect API v5)

This PowerShell script retrieves Dell warranty information using Service Tags via the Dell TechDirect API (SBIL / EAPI v5).

## Features

- Cross-platform (Windows, Linux, macOS)
- CSV input (default)
- XLSX support on Windows only
- OAuth 2.0 authentication
- Outputs structured CSV report
- Works with PowerShell 5.1 and 7+

---

## Requirements

- Dell TechDirect API access
- Client ID and Client Secret
- PowerShell (recommended: PowerShell 7+)
- Excel (only required for XLSX input on Windows)

---

## Input Format

Create a file named `ServiceTags.csv`:

```csv
ServiceTag
6LSRKQ1
C15YLG1
