# Dell Warranty Lookup (TechDirect API v5)

PowerShell script to retrieve Dell warranty information using Service Tags via the Dell TechDirect API (SBIL / EAPI v5).

---

## 🚀 Features

- Cross-platform (Windows, Linux, macOS)
- CSV input (default)
- XLSX support (Windows only)
- OAuth 2.0 authentication
- Structured CSV output
- Compatible with PowerShell 5.1 and 7+

---

## 📋 Requirements

- Dell TechDirect API credentials
  - Client ID
  - Client Secret
- PowerShell (recommended: PowerShell 7+)

---

## 📥 Input Format

Create a file named `ServiceTags.csv`:

```csv
ServiceTag
6LSRKQ1
C15YLG1
