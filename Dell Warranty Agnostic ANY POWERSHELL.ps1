# ===============================================
# Dell Warranty API V5 - v1.2
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

# Batch + retry settings
$BatchSize  = 100
$MaxRetries = 3
$RetryDelay = 2

# ===============================================
# STEP 0 — LOGGING FUNCTION
# ===============================================
function Write-Log {
    param ($Message, $Color = "White")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"

    Write-Host $entry -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $entry
}

Write-Log "Script started"

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
# STEP 1 — LOAD SERVICE TAGS (ROBUST CSV)
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

            if ($_.ServiceTag) {
                $value = $_.ServiceTag
            } else {
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
# STEP 3 — WARRANTY LOOKUPS (BATCH + RETRY + 1 ROW PER TAG)
# ===============================================
Write-Log "Querying warranty information..." "Cyan"

$Results = @()
$TotalTags = $Tags.Count
$Processed = 0

# Create batches
$Batches = for ($i=0; $i -lt $Tags.Count; $i += $BatchSize) {
    $Tags[$i..([Math]::Min($i + $BatchSize - 1, $Tags.Count - 1))]
}

foreach ($batch in $Batches) {

    foreach ($tag in $batch) {
        $Processed++
        Write-Log "Checking tag [$Processed/$TotalTags]: $tag" "Yellow"
    }

    $uri = $BaseURL + ($batch -join ",")

    $attempt = 1
    $success = $false

    while (-not $success -and $attempt -le $MaxRetries) {

        try {
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{
                Authorization = "Bearer $AccessToken"
                Accept        = "application/json"
            }

            $success = $true
        }
        catch {
            if ($attempt -lt $MaxRetries) {
                Write-Log "Retrying batch..." "Yellow"
                Start-Sleep -Seconds $RetryDelay
            } else {
                Write-Log "ERROR: Failed batch after retries" "Red"
            }
        }

        $attempt++
    }

    if (-not $success) { continue }

    foreach ($asset in $response) {

        # ✅ FIX: Select SINGLE latest entitlement
        $latestEntitlement = $asset.entitlements | Sort-Object endDate -Descending | Select-Object -First 1

        if ($latestEntitlement) {

            $Results += [PSCustomObject]@{
                ServiceTag   = $asset.serviceTag
                Product      = $asset.productLineDescription
                ShipDate     = $asset.shipDate
                WarrantyType = $latestEntitlement.entitlementType
                StartDate    = $latestEntitlement.startDate
                EndDate      = $latestEntitlement.endDate
                Level        = $latestEntitlement.serviceLevelDescription
                Status       = if ((Get-Date $latestEntitlement.endDate) -ge (Get-Date)) { "Active" } else { "Expired" }
            }
        }
    }
}

# ===============================================
# STEP 4 — EXPORT
# ===============================================
$Results | Export-Csv -NoTypeInformation -Path $OutputCSV

# ===============================================
# SUMMARY
# ===============================================
$TotalRows = $Results.Count

Write-Log "Summary:" "Cyan"
Write-Log "Total Service Tags processed: $TotalTags" "Green"
Write-Log "Final rows exported: $TotalRows" "Green"

Write-Log "DONE! Warranty results exported to: $OutputCSV" "Green"
