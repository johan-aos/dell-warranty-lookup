![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207+-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-green)
![API](https://img.shields.io/badge/API-Dell%20TechDirect-orange)

# Dell Warranty Lookup (TechDirect API v5)

PowerShell script to retrieve Dell warranty information using Service Tags via the Dell TechDirect API (SBIL / EAPI v5).

---

## Features

- Cross-platform (Windows, Linux, macOS)
- CSV input by default / XLSX support on Windows only
- OAuth 2.0 authentication
- Structured CSV output
- Compatible with PowerShell 5.1 and 7+

---

## Requirements

- Dell TechDirect API credentials
  - Client ID
  - Client Secret
- PowerShell (recommended: PowerShell 7+)
- Excel (COM required only for XLSX input on Windows OS)

---

## Input Format

⚠ Column header must be named `ServiceTag` (case-sensitive)

Create a file named `ServiceTags.csv`(This will contain all the Service Tags you want to verify against the Dell API): 

```csv
ServiceTag
6LSRKQ1
C15YLG1

```
---

## Output Example

ServiceTag, Product, EndDate, Status
```
6LSRKQ1, LATITUDE E5420, 2015-06-03, Expired
```
---

## Tested Environments

- Windows 10/11 (Powershell 5.1, 7)

---
