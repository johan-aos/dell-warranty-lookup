# ===============================================
# Dell Warranty API V5 - Cross-Platform Version
# ===============================================

# ---- CONFIGURATION ----
$ClientID     = "CLIENT_ID_HERE"
$ClientSecret = "CLIENT_SECRET_HERE"

# Input files
$InputCSV  = "INPUT_PATH_FROM\ServiceTags.csv"
$InputXLSX = "INPUT_PATH_FROM\ServiceTags.xlsx"

# Output file
$OutputCSV = "OUTPUT_PATH_TO\Dell_Warranty_Results.csv"

# API URLs (DO NOT MODIFY)
$TokenURL = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"
$BaseURL  = "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements?servicetags="

# ===============================================
# STEP 0 — OS DETECTION (PowerShell 5 SAFE)
# ===============================================
$IsWindows = $false

if ($PSVersionTable.PSEdition -eq "Desktop" -or $env:OS -eq "Windows_NT") {
    $IsWindows = $true
    $OSName = "Windows"
} else {
    $OSName = "Linux/macOS"
}

Write-Host "Detected OS: $OSName" -ForegroundColor Cyan


# ===============================================
# STEP 1 — LOAD SERVICE TAGS
# ===============================================
$Tags = @()

# ---- Try Excel input ONLY on Windows ----
if ($IsWindows -and (Test-Path $InputXLSX)) {

    Write-Host "Using XLSX input (Windows COM mode)..." -ForegroundColor Cyan

    try {
        $Excel    = New-Object -ComObject Excel.Application
        $Workbook = $Excel.Workbooks.Open($InputXLSX)
        $Sheet    = $Workbook.Sheets.Item(1)

        $row = 2
        while ($Sheet.Cells.Item($row,1).Text -ne "") {

            $tag = $Sheet.Cells.Item($row,1).Text.Trim()

            if ($tag -match '^[A-Za-z0-9]{7}$') {
                $Tags += $tag
            }

            $row++
        }

        $Workbook.Close($false)
        $Excel.Quit()
    }
    catch {
        Write-Host "⚠ Excel read failed. Falling back to CSV..." -ForegroundColor Yellow
    }
}

# ---- Default: CSV input ----
if ($Tags.Count -eq 0) {

    Write-Host "Using CSV input..." -ForegroundColor Cyan

    try {
        $Tags = Import-Csv $InputCSV | ForEach-Object {
            $_.ServiceTag.Trim()
        } | Where-Object {
            $_ -match '^[A-Za-z0-9]{7}$'
        }
    }
    catch {
        Write-Host "❌ Unable to read CSV file." -ForegroundColor Red
        return
    }
}

Write-Host "Loaded $($Tags.Count) valid Service Tags." -ForegroundColor Green


# ===============================================
# STEP 2 — AUTHENTICATE (OAuth 2.0)
# ===============================================
Write-Host "Requesting OAuth token..." -ForegroundColor Cyan

$pair = "${ClientID}:${ClientSecret}"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))

try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $TokenURL -Headers @{
        Authorization = "Basic $encodedCreds"
        "Content-Type" = "application/x-www-form-urlencoded"
    } -Body "grant_type=client_credentials"
}
catch {
    Write-Host ("❌ AUTH FAILED -> " + $_.ErrorDetails.Message) -ForegroundColor Red
    return
}

$AccessToken = $tokenResponse.access_token
Write-Host "Token acquired successfully." -ForegroundColor Green


# ===============================================
# STEP 3 — WARRANTY LOOKUPS
# ===============================================
Write-Host "Querying warranty information..." -ForegroundColor Cyan

$Results = @()

foreach ($tag in $Tags) {

    Write-Host ("Checking tag: " + $tag) -ForegroundColor Yellow

    $uri = $BaseURL + $tag

    try {
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{
            Authorization = "Bearer $AccessToken"
            Accept        = "application/json"
        }
    }
    catch {
        Write-Host ("Error querying tag " + $tag + " -> " + $_.Exception.Message) -ForegroundColor Red
        continue
    }

    foreach ($asset in $response) {
        foreach ($ent in $asset.entitlements) {

            $Results += [PSCustomObject]@{
                ServiceTag   = $asset.serviceTag
                Product      = $asset.productLineDescription
                ShipDate     = $asset.shipDate
                WarrantyType = $ent.entitlementType
                StartDate    = $ent.startDate
                EndDate      = $ent.endDate
                Level        = $ent.serviceLevelDescription
                Status       = if ((Get-Date $ent.endDate) -ge (Get-Date)) { "Active" } else { "Expired" }
            }
        }
    }
}

# ===============================================
# STEP 4 — EXPORT RESULTS
# ===============================================
$Results | Export-Csv -NoTypeInformation -Path $OutputCSV

Write-Host "DONE! Warranty results exported to: $OutputCSV" -ForegroundColor Green
