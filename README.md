![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207+-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-purple)
![API](https://img.shields.io/badge/API-Dell%20TechDirect-white)
![ZorinOS](https://img.shields.io/badge/tested-Zorin%20OS%2018.1-cyan)
![Windows11](https://img.shields.io/badge/tested-Windows%2011%20-cyan)
![Pop! OS](https://img.shields.io/badge/tested-Pop!_OS%20-cyan)
![Nobara GNOME](https://img.shields.io/badge/tested-Nobara_GNOME%20-cyan)

# Dell Warranty Lookup (TechDirect API v5)

PowerShell script to retrieve Dell warranty information using Service Tags via the Dell TechDirect API (SBIL / EAPI v5).

---

## Features

- Cross-platform (Windows, Linux, macOS)
- CSV input by default / XLSX support on Windows only
- OAuth 2.0 authentication
- Structured CSV output
- Compatible with PowerShell 5.1 and 7+
- Batch processing (up to 100 ServiceTags per request)
- One ServiceTag = one row (latest warranty only)
- Built-in logging file for execution tracking
- Linux-native execution support via `pwsh`

---

## Linux / macOS Notes

- Requires PowerShell (`pwsh`)
- Script can be executed natively using:

```bash
chmod +x Dell_Warranty_Lookup.ps1
./Dell_Warranty_Lookup.ps1
```

--- 

## Requirements

- Dell TechDirect API credentials
  - Client ID
  - Client Secret
- PowerShell (recommended: PowerShell 7+)
- Excel (COM required only for XLSX input on Windows OS)

---

## Input Format

⚠ It is recommended that the column header is named `ServiceTag` (case-sensitive), although I've added some flexibility to the the input file.

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

- PowerShell Core
- Windows 10/11 (Powershell 5.1, 7)
- Zorin OS 18.1 (Ubuntu based pwsh)
- Pop! OS 24.04 (Ubuntu based pwsh)
- Nobara GNOME (Fedora 44)

---
