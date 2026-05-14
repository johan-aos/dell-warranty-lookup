# Dell Warranty Lookup (TechDirect API v5)

PowerShell script to retrieve Dell warranty information using Service Tags via the Dell TechDirect API (SBIL / EAPI v5).

---

## 🚀 Features

- Cross-platform (Windows, Linux, macOS)
- CSV input by default / XLSX support on Windows only
- OAuth 2.0 authentication
- Structured CSV output
- Compatible with PowerShell 5.1 and 7+

---

## 📋 Requirements

- Dell TechDirect API credentials
  - Client ID
  - Client Secret
- PowerShell (recommended: PowerShell 7+)
- Excel (COM required only for XLSX input on Windows OS)

---

## 📥 Input Format

Create a file named `ServiceTags.csv`(This will contain all the Service Tags you want to verify against the Dell API): 

```csv
ServiceTag
6LSRKQ1
C15YLG1
