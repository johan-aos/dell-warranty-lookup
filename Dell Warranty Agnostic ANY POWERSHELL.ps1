# ===============================================
# Dell Warranty API V5 - v1.1
# ===============================================

# ---- CONFIGURATION ----
$ClientID     = "YOUR_CLIENT_ID_HERE"
$ClientSecret = "YOUR_CLIENT_SECRET_HERE"

$InputCSV  = "YOUR_INPUT_PATH_FROM\ServiceTags.csv"
$InputXLSX = "YOUR_INPUT_PATH_FROM\ServiceTags.xlsx"

$OutputCSV = "YOUR_OUTPUT_PATH_TO\Dell_Warranty_Results.csv"

# Logging file
$LogFile   = "YOUR_OUTPUT_PATH_TO\Dell_Warranty_Log.txt"

# API URLs
$TokenURL = "https://apigtwb2c.us.dell.com/auth/oauth/v2/token"
$BaseURL  = "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements?servicetags="

# Batch size (Dell supports up to 100)
$BatchSize = 100

# ===============================================
# STEP 0 — LOGGING FUNCTION (NEW)
# ===============================================
function Write-Log {
    param ($Message, $Color = "White")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"

    Write-Host $entry -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $entry
}

# ===============================================
# STEP 0 — OS DETECTION
# ===============================================
$IsWindows = $false

if ($PSVersionTable.PSEdition -eq "Desktop" -or $env:OS -eq "Windows_NT") {
    $IsWindows = $true
    $OSName = "Windows"
} else {
    $OSName = "Linux/macOS"
}

Write-Log "Detected OS: $OSName" "Cyan"


# ===============================================
# STEP 1 — LOAD SERVICE TAGS (IMPROVED CSV)
# ===============================================
$Tags = @()

if ($IsWindows -and (Test-Path $InputXLSX)) {

    Write-Log "Using XLSX input..." "Cyan"

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
        Write-Log "Excel read failed. Falling back to CSV." "Yellow"
    }
}

if ($Tags.Count -eq 0) {

    Write-Log "Using CSV input..." "Cyan"

    try {
        $Tags = Import-Csv $InputCSV | ForEach-Object {

            # ✅ Flexible column handling
            $value = $null

            if ($_.ServiceTag) {
                $value = $_.ServiceTag
            } else {
                # fallback to first column
                $value = $_.PSObject.Properties.Value[0]
            }

            $value = $value.Trim()

            if ($value -match '^[A-Za-z0-9]{7}$') {
                $value
            }
        }
    }
    catch {
        Write-Log "ERROR: Unable to read CSV file." "Red"
        return
    }
}

Write-Log "Loaded $($Tags.Count) valid Service Tags" "Green"


# ===============================================
# STEP 2 — AUTHENTICATE
# ===============================================
Write-Log "Requesting OAuth token..." "Cyan"

$pair = "${ClientID}:${ClientSecret}"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))

try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri $TokenURL -Headers @{
        Authorization = "Basic $encodedCreds"
        "Content-Type" = "application/x-www-form-urlencoded"
    } -Body "grant_type=client_credentials"
}
catch {
    Write-Log "AUTH FAILED" "Red"
    return
}

$AccessToken = $tokenResponse.access_token
Write-Log "Token acquired successfully." "Green"


# ===============================================
# STEP 3 — WARRANTY LOOKUPS (BATCHED)
# ===============================================
Write-Log "Querying warranty information..." "Cyan"

$Results = @()

# Create batches of max 100 tags
$Batches = for ($i=0; $i -lt $Tags.Count; $i += $BatchSize) {
    $Tags[$i..([math]::Min($i + $BatchSize - 1, $Tags.Count - 1))]
}

foreach ($batch in $Batches) {

    # ✅ Keep original per-tag visibility
    foreach ($tag in $batch) {
        Write-Log "Checking tag: $tag" "Yellow"
    }

    $uri = $BaseURL + ($batch -join ",")

    try {
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{
            Authorization = "Bearer $AccessToken"
            Accept        = "application/json"
        }
    }
    catch {
        Write-Log "ERROR: Failed request for batch" "Red"
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

Write-Log "DONE! Warranty results exported to: $OutputCSV" "Green"

