<#
.SYNOPSIS
    Extract, export and synchronize user data from PeopleSpheres API to Azure SQL.

.DESCRIPTION
    This script performs a full data extraction from the PeopleSpheres API (both active and inactive users),
    flattens the relevant fields based on a predefined mapping, and exports the data into CSV format.

    It generates three output files:
        - Two timestamped CSVs (active and inactive users) for automation/archive purposes
        - One fixed-name CSV for Azure SQL ingestion (via Azure Blob + BULK INSERT)

    The script also handles:
        - Authentication with refreshable access tokens
        - Detection and alerting on unknown field IDs (excluding technical accounts)
        - Summary email reporting with export paths and key metrics
        - Upload to Azure Blob Storage and synchronization into SQL Server
        - Error logging and email alerting in case of failure

    Block structure overview:
        0. Environment Preparation & Module Loading
        1. Global Setup & Utility Functions
        2. Authentication & Token Management
        3. User Data Loading (active/inactive)
        4. Per-user Field Retrieval & Flattening
        5. Unknown Field Handling (formerly Alert for Unmapped Field IDs)
        6. CSV Export
        7. Summary Email Report
        8. Error Notification (with log)
        9. Upload to Azure Blob & SQL Bulk Insert

.PARAMETER IsTestMode
    Boolean. When enabled, only fetches 25 users (active + inactive) instead of the full dataset.

.OUTPUTS
    - UTF-16 CSV files for active and inactive users: usable in Excel.
    - UTF-8 CSV files with timestamp for automation ingestion.
    - Email notification with record counts and export locations.

.REQUIREMENTS
    - PowerShell 5.1+
    - Internet connectivity
    - SecureString file for support-itps@iadinternational.com credentials
    - IAD-Admin module available locally or globally

.MODULES_REQUIRED
    - ActiveDirectory
    - IAD-Admin

.NOTES
    Author  : Jérémie Poujol
    Company : iad Business Factory
    Version : 2.0
    Created : 2022-01 (v1), Refactored : 2025-04-30 (v2)

.LINK
    API Docs (legacy): https://rest.monportailrh.com/swagger/
    API Docs (modern): https://rest.monportailrh.com/docs/
    SSO Auth:          https://sso.monportailrh.com/auth/

.EXAMPLE
    PS> .\GenerateCSV-API-PeopleSpheres.ps1
#>

# =====================================================================
# 🧪 TEST MODE SWITCH – SET THIS VALUE BEFORE RUNNING THE SCRIPT
# =====================================================================
# This section controls whether the script runs in test or production mode.
#
#   - TEST MODE     → limits API calls to $MaxTestUsers users per status (active/inactive)
#   - PRODUCTION    → processes all available users from the API
#
# ✅ Change this setting before launching the script.
#    Do not edit it elsewhere in the script.

$IsTestMode   = $true         # ← ❗ Set to $false in production
$MaxTestUsers = 10            # Applies only in test mode

# Display warning if test mode is active
if ($IsTestMode) {
    $banner = @"
╔════════════════════════════════════════════════════════════════════╗
║                       ⚠ TEST MODE ACTIVE ⚠                        ║
╠════════════════════════════════════════════════════════════════════╣
║ Only the first $MaxTestUsers ACTIVE + $MaxTestUsers INACTIVE users       ║
║ will be retrieved from the PeopleSpheres API.                    ║
║                                                                  ║
║ ➤ This helps speed up testing and avoid unnecessary load         ║
║ ➤ To disable: set $IsTestMode = $false before running the script ║
╚════════════════════════════════════════════════════════════════════╝
"@
    Write-Host $banner -ForegroundColor Yellow
}

# -------------------------------------------------------
# Global Email Configuration
# -------------------------------------------------------
if ($IsTestMode) {
    $emailDetails = @{
        To  = @("jeremie.poujol@iadinternational.com")
        Cc  = @("pascal.raoult@iadinternational.com")
        Bcc = @("986c2ea3.iadgroup.onmicrosoft.com@fr.teams.ms")
    }
} else {
    $emailDetails = @{
        To  = @("exploitation.notify@iadinternational.com")
        Cc  = @("alexandre.kebaili-ext@iadinternational.com")
        Bcc = @("986c2ea3.iadgroup.onmicrosoft.com@fr.teams.ms")
    }
}
# -------------------------------------------------------
# BLOCK 0 : Environment Preparation & Module Loading
# -------------------------------------------------------

# Clear screen and enforce TLS 1.2 for all web requests
Clear-Host
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "🔐 TLS 1.2 enabled for secure web communication." -ForegroundColor Cyan

# Display ASCII banner
$logo = @"
.__            .___ ___.                .__                                _____               __                       
|__|____     __| _/ \_ |__  __ __  _____|__| ____   ____   ______ ______ _/ ____\____    _____/  |_  ___________ ___.__.
|  \__  \   / __ |   | __ \|  |  \/  ___/  |/    \_/ __ \ /  ___//  ___/ \   __\\__  \ _/ ___\   __\/  _ \_  __ <   |  |
|  |/ __ \_/ /_/ |   | \_\ \  |  /\___ \|  |   |  \  ___/ \___ \ \___ \   |  |   / __ \\  \___|  | (  <_> )  | \/\___  |
|__(____  /\____ |   |___  /____//____  >__|___|  /\___  >____  >____  >  |__|  (____  /\___  >__|  \____/|__|   / ____|
        \/      \/       \/           \/        \/     \/     \/     \/              \/     \/                   \/     
"@
Write-Host $logo -ForegroundColor Green
Write-Host "🔄 Starting PeopleSpheres API data extraction..." -ForegroundColor White
Start-Sleep -Seconds 1

# Timestamp
$script:Timestamp = Get-Date -Format "yyyyMMdd"

# === Export folders (simplified) ===
$script:CsvExportFolder = "E:\Powershell\03-FlatFilesStorage\GenerateCSV-API-PeopleSpheres"
$script:CsvAzureFolder  = "E:\Powershell\03-FlatFilesStorage\AzureSQLDatabase_csv"
$script:LogFolder       = "E:\Powershell\04-ScriptLogsAndOutputs\GenerateCSV-API-PeopleSpheres"

# === Final paths ===
$script:CsvActive_Timestamped   = Join-Path $script:CsvExportFolder "PeopleSpheres-active-$($script:Timestamp).csv"
$script:CsvInactive_Timestamped = Join-Path $script:CsvExportFolder "PeopleSpheres-inactive-$($script:Timestamp).csv"
$script:CsvAzure_NoTimestamp    = Join-Path $script:CsvAzureFolder  "GenerateCSV-API-PeopleSpheres.csv"
$script:LogPath                 = Join-Path $script:LogFolder       "PeopleSpheres-Export.log"

# === Create folders if missing ===
$folders = @($script:CsvExportFolder, $script:CsvAzureFolder, $script:LogFolder)
foreach ($folder in $folders) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
}

# === Import required modules ===
Write-Host "📦 Importing required modules..." -ForegroundColor White

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "✅ Module 'ActiveDirectory' loaded." -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to import 'ActiveDirectory' module." -ForegroundColor Red
    Exit 1
}

# Load IAD-Admin from default location or fallback path
if (-not (Get-Module -Name IAD-Admin -ListAvailable)) {
    $fallbackPath = "E:\Powershell\00_Modules\IAD-Admin"
    if (Test-Path $fallbackPath) {
        Import-Module $fallbackPath -ErrorAction Stop
        Write-Host "✅ Module 'IAD-Admin' loaded from fallback path." -ForegroundColor Green
    } else {
        Write-Host "❌ 'IAD-Admin' module not found in default or fallback path." -ForegroundColor Red
        Exit 1
    }
} else {
    Import-Module IAD-Admin -ErrorAction Stop
    Write-Host "✅ Module 'IAD-Admin' loaded." -ForegroundColor Green
}

Write-Host "=====================================n" -ForegroundColor White

# -------------------------------------------------------
# BLOCK 1 : Global Setup & Utility Functions
# -------------------------------------------------------

$scriptStartTime = Get-Date
$global:TokenRefreshCount = 0
$script:HadErrors = $false


function Show-TokenRefreshedBanner {
    param (
        [datetime]$startTime,
        [int]$refreshCount
    )

    $elapsed      = (Get-Date) - $startTime
    $elapsedStr   = $elapsed.ToString("hh\:mm\:ss")

    # Compose each line
    $title        = "TOKEN REFRESHED"
    $line1        = "⏱  Elapsed Time       : $elapsedStr"
    $line2        = "🔄  Total Refresh Count: $refreshCount"

    # Determine max line width
    $allLines     = @($title, $line1, $line2)
    $maxLength    = ($allLines | Measure-Object -Property Length -Maximum).Maximum
    $padding      = 4 # 2 spaces on each side of content inside borders
    $totalWidth   = $maxLength + $padding + 2 # +2 for the '#' at both ends

    # Top/Bottom border
    $borderLine   = "#" * $totalWidth

    # Helper to wrap lines
    function Wrap-Line($text) {
        $spaces = " " * ($maxLength - $text.Length)
        return "# $text$spaces #"
    }

    # Build full banner
    $banner = @(
        $borderLine
        Wrap-Line $title
        Wrap-Line ("-" * $maxLength)
        Wrap-Line $line1
        Wrap-Line $line2
        $borderLine
    ) -join "n"

    Write-Host $banner -ForegroundColor Cyan
}


function Normalize-Label {
    param ([string]$label)
    return ($label.Normalize('FormD') -replace '\p{Mn}', '' -replace '[^\w\s]', '' -replace '\s+', '_' -replace '_+', '_' | ForEach-Object { $_.ToLowerInvariant() })
}

function Refresh-AccessToken {
    Write-Host "🔁 Refreshing access token..." -ForegroundColor Yellow
    $refreshBody = @{
        "grant_type"    = "refresh_token"
        "client_id"     = $clientId
        "refresh_token" = $refreshToken
    }
    try {
        $refreshResponse = Invoke-RestMethod -Uri $authUrl -Method POST -Body $refreshBody -Headers $authHeaders
        $global:accessToken  = $refreshResponse.access_token
        $global:refreshToken = $refreshResponse.refresh_token
        $global:tokenTime    = Get-Date
        Write-Host "✅ Access token refreshed at $tokenTime" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to refresh access token." -ForegroundColor Red
        Exit 1
    }
}

# Add more reusable functions here if needed

# -------------------------------------------------------
# BLOCK 2 : AUTHENTICATION & TOKEN MANAGEMENT
# -------------------------------------------------------

# Global token values
$script:AccessToken  = $null
$script:RefreshToken = $null
$script:TokenTime    = $null

# Auth configuration
$script:ClientId     = "monportailrh-web-app"
$script:Username     = "support-itps@iadinternational.com"
$script:AuthUrl      = "https://sso.monportailrh.com/auth/realms/Internal-idp/protocol/openid-connect/token"
$script:SecurePath   = "C:\Scripts\SecureString\$Username.$env:USERNAME.securestring"

# Prompt for password if secure file doesn't exist
if (-not (Test-Path $SecurePath)) {
    Write-Host "🔐 No secure string found. Prompting for password..." -ForegroundColor Yellow
    Read-Host "Enter password for $Username" -AsSecureString |
        ConvertFrom-SecureString |
        Out-File -FilePath $SecurePath
}

# Load and decrypt password
$SecureString = Get-Content $SecurePath | ConvertTo-SecureString
$PlainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
)

# Define common headers and body
$script:AuthHeaders = @{
    "Content-Type" = "application/x-www-form-urlencoded"
    "Accept"       = "application/json"
}

# Function to request a fresh access token
function Request-AccessToken {
    $body = @{
        "grant_type" = "password"
        "client_id"  = $ClientId
        "username"   = $Username
        "password"   = $PlainPassword
    }

    try {
        $response = Invoke-RestMethod -Uri $AuthUrl -Method POST -Body $body -Headers $AuthHeaders
        $script:AccessToken  = $response.access_token
        $script:RefreshToken = $response.refresh_token
        $script:TokenTime    = Get-Date
        Write-Host "✅ Access token acquired at $TokenTime" -ForegroundColor Cyan
    } catch {
        Write-Host "❌ Failed to retrieve access token." -ForegroundColor Red
        Exit 1
    }
}

# Function to refresh token when needed (after 4 min+)
function Refresh-AccessToken {
    $refreshBody = @{
        "grant_type"    = "refresh_token"
        "client_id"     = $ClientId
        "refresh_token" = $RefreshToken
    }

    try {
        $response = Invoke-RestMethod -Uri $AuthUrl -Method POST -Body $refreshBody -Headers $AuthHeaders
        $script:AccessToken  = $response.access_token
        $script:RefreshToken = $response.refresh_token
        $script:TokenTime    = Get-Date
        Write-Host "🔁 Token refreshed at $TokenTime" -ForegroundColor Yellow
    } catch {
        Write-Host "❌ Failed to refresh access token." -ForegroundColor Red
        Exit 1
    }
}

# Trigger initial token request
Request-AccessToken

# -------------------------------------------------------
# BLOCK 3 : USER DATA LOADING (ACTIVE & INACTIVE)
# -------------------------------------------------------

# Base PeopleSpheres API URL
$BaseApiUrl = "https://rest.monportailrh.com"

# Build dynamic authorization headers
function Get-ApiHeaders {
    return @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "application/json"
    }
}

# Function to fetch user IDs based on active status
function Get-UserIds {
    param (
        [bool]$Active
    )

    $headers = Get-ApiHeaders
    $status  = if ($Active) { 1 } else { 0 }
    $label   = if ($Active) { "ACTIVE" } else { "INACTIVE" }

    if ($IsTestMode) {
        Write-Host "🧪 TEST MODE: Fetching first $MaxTestUsers $label users from PeopleSpheres..." -ForegroundColor Yellow
        $url = "$BaseApiUrl/search?include=data.quick_actions&name=&page=1&per-page=$MaxTestUsers&pso-type=usr&active=$status"
    } else {
        Write-Host "📦 Fetching $label user count..." -ForegroundColor Gray
        $metaUrl = "$BaseApiUrl/search?name=&pso-type=usr&page=1&per-page=1&active=$status"
        $metaResponse = Invoke-RestMethod -Uri $metaUrl -Headers $headers
        $totalUsers = $metaResponse.data.meta.pagination.total
        Write-Host "🔢 $label user count: $totalUsers" -ForegroundColor Cyan

        $url = "$BaseApiUrl/search?include=data.quick_actions&name=&page=1&per-page=$totalUsers&pso-type=usr&active=$status"
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers
        return $response.data.data.id
    } catch {
        Write-Host "❌ Failed to retrieve $label user list." -ForegroundColor Red
        return @()
    }
}

# Trigger both user lists
$UserIds_Active   = Get-UserIds -Active $true
$UserIds_Inactive = Get-UserIds -Active $false

# Display recap in test mode
if ($IsTestMode) {
    $totalTestFetched = $UserIds_Active.Count + $UserIds_Inactive.Count
    Write-Host "`n📥 Test mode total fetched: $($UserIds_Active.Count) active + $($UserIds_Inactive.Count) inactive = $totalTestFetched users" -ForegroundColor Magenta
}

# -------------------------------------------------------
# BLOCK 4 : PER-USER FIELD RETRIEVAL & FLATTENING
# -------------------------------------------------------
# Function to retrieve and flatten user data

# PeopleSpheres field label mapping
$FieldMapRaw = @{
    "711"  = "Date de début dans le poste"
    "2637" = "[YOUZER] Date d'embauche"
    "683"  = "Départ de la société"
    "2638" = "[YOUZER] Date de sortie"
    "1145" = "Type collaboration"
    "125"  = "Nom"
    "123"  = "Prénom"
    "423"  = "Image de profil"
    "550"  = "Civilité"
    "28"   = "Adresse e-mail professionnelle"
    "27"   = "Téléphone portable professionnel"
    "182"  = "Responsable"
    "344"  = "Service"
    "185"  = "Poste"
    "186"  = "Entité Légale"
    "348"  = "Site"
    "859"  = "Matricule"
}

# Normalize keys from labels
$FieldMap = @{ }
foreach ($kvp in $FieldMapRaw.GetEnumerator()) {
    $FieldMap[$kvp.Key] = Normalize-Label $kvp.Value
}

function Get-FlattenedUserData {
    param (
        [array]$UserIds,
        [string]$UserType  # "active" or "inactive"
    )

    $result = @()
    $unknownFieldIds = @{ }  # Initialize empty collection for unknown field IDs
    $counter = 1
    $total = $UserIds.Count
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($userId in $UserIds) {
        Write-Host "[$counter/$total] [$UserType] Fetching user ID $userId..." -ForegroundColor Cyan
        $counter++

        try {
            $stopwatch.Restart()
            $url = "$BaseApiUrl/psos/$userId/fields?active=true&include=type,items,options,settings,assignment_settings"
            $fields = (Invoke-RestMethod -Uri $url -Headers (Get-ApiHeaders)).data
            $stopwatch.Stop()

            $userData = [ordered]@{}
            $userData["dateexescript"] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

            # Prepare display values
            $prenom = ""
            $nom = ""
            $email = ""
            $jobtitle = ""
            $entite = ""
            $actif = ""
            $prenomCaractereSpeciaux = ""

            foreach ($item in $fields) {
                $idStr = "$($item.id)"
                $value = $item.value_details

                if ($FieldMap.ContainsKey($idStr)) {
                    $key = $FieldMap[$idStr]

                    # Null check before accessing value
                    if ($value -eq $null) {
                        $userData[$key] = ""
                    } elseif ($idStr -eq "182" -and $value -is [psobject]) {
                        $userData[$key] = $value.professional_email
                    } elseif ($idStr -eq "272") {
                        $userData[$key] = $value  # "Actif" value (Yes/No)
                    } elseif ($idStr -eq "3657") {
                        $userData[$key] = $value  # The first name with special characters
                    } elseif ($value -is [array]) {
                        $userData[$key] = ($value -join ", ")  # Join array values with commas
                    } elseif ($value -is [psobject]) {
                        $userData[$key] = $value.ToString()  # Convert psobject to string
                    } else {
                        $userData[$key] = $value  # Default case for other values
                    }

                    # Adding specific user data
                    switch ($key) {
                        "prenom"         { $prenom  = $userData[$key] }
                        "nom"            { $nom     = $userData[$key] }
                        "adresse_email_professionnelle" { $email = $userData[$key] }
                        "poste"          { $jobtitle = $userData[$key] }
                        "entite_legale"  { $entite  = $userData[$key] }
                        "actif"          { $actif = $userData[$key] }
                        "prenom_caracteres_speciaux" { $prenomCaractereSpeciaux = $userData[$key] }
                    }
                } else {
                    # Collect unknown field IDs
                    if (-not $unknownFieldIds.ContainsKey($idStr)) {
                        $unknownFieldIds[$idStr] = @{
                            label = $item.label
                            type  = $item.type
                            user  = $userId
                        }
                    }
                }
            }

            # Ensure every expected key is present in the final object
            foreach ($expectedKey in $FieldMap.Values) {
                if (-not $userData.Contains($expectedKey)) {
                    $userData[$expectedKey] = ""
                }
            }

            $result += [PSCustomObject]$userData

            # Debugging output
            $displayBlock = @"
------------------------------
👤 $prenom $nom
📧 $email
🏢 $entite
🏷  $jobtitle
------------------------------
"@
            Write-Host $displayBlock -ForegroundColor DarkCyan
        } catch {
            $message = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ❌ [$UserType] ERROR for user $userId : $($_.Exception.Message)"
            Write-Host $message -ForegroundColor Red
            Add-Content -Path $script:LogPath -Value $message
            $script:HadErrors = $true
        }

        # Refresh token if older than 4 minutes
        if ((New-TimeSpan -Start $TokenTime).TotalMinutes -ge 4) {
            $global:TokenRefreshCount++
            Refresh-AccessToken
            Show-TokenRefreshedBanner -startTime $scriptStartTime -refreshCount $global:TokenRefreshCount
        }
    }

    return $result
}

# Process active and inactive users separately
$ResultActive   = Get-FlattenedUserData -UserIds $UserIds_Active   -UserType "active"
$ResultInactive = Get-FlattenedUserData -UserIds $UserIds_Inactive -UserType "inactive"

# ---------------------------
# BLOCK 5 : UNKNOWN FIELD HANDLING
# ---------------------------

# List to store unmapped field IDs with associated user info
$unmappedFieldDetails = @()

# ID to ignore (service account with elevated access – sees extra internal fields)
$excludedUserIds = @(529)  # ← This account sees everything, ignore it on purpose

# Re-scan all users (active + inactive) for unmapped field IDs
foreach ($userId in ($UserIds_Active + $UserIds_Inactive)) {
    if ($excludedUserIds -contains $userId) {
        continue  # Skip known service/system accounts
    }

    try {
        $url = "$BaseApiUrl/psos/$userId/fields?active=true&include=type,items,options,settings,assignment_settings"
        $fields = (Invoke-RestMethod -Uri $url -Headers (Get-ApiHeaders)).data

        foreach ($item in $fields) {
            $fieldId = "$($item.id)"

            if (-not $FieldMap.ContainsKey($fieldId)) {
                $unmappedFieldDetails += [PSCustomObject]@{
                    UserId   = $userId
                    FieldId  = $fieldId
                    Label    = if ($item.label) { $item.label } else { "(empty)" }
                    Type     = if ($item.type.name) { $item.type.name } else { $item.type.alias }
                }
            }
        }
    } catch {
        Write-Warning "⚠️ Failed to scan unmapped fields for user ID $userId"
    }
}

# If unmapped fields found (excluding known service account), send a single alert email
if ($unmappedFieldDetails.Count -gt 0) {
    $userIdToCheck = $unmappedFieldDetails[0].UserId

    # If the user ID is 529, skip sending the email
    if ($userIdToCheck -ne 529) {
        # Start HTML body
        $alertBody = @"
<p>🚨 <strong>Unmapped PeopleSpheres Field IDs Detected</strong></p>
<p>The following field IDs were returned by the API but are not currently handled in the <code>FieldMap</code>.</p>
<p>Note: Fields returned by service account(s) such as User ID 529 have been intentionally excluded.</p>

<table border="1" cellpadding="5" cellspacing="0">
<tr><th>User ID</th><th>Field ID</th><th>Label</th><th>Type</th></tr>
"@

        foreach ($entry in $unmappedFieldDetails) {
            $alertBody += "<tr><td>$($entry.UserId)</td><td>$($entry.FieldId)</td><td>$($entry.Label)</td><td>$($entry.Type)</td></tr>`n"
        }

        $alertBody += "</table>
<p>Please review and update the field mappings if necessary.</p>"

        try {
            IADAdmin_SendMailMessage -Body $alertBody 
                                     -To "jeremie.poujol@iadinternational.com" 
                                     -Subject "[iadlife] 🔎 PeopleSpheres – Unmapped Field IDs Detected" 
                                     -BodyAsHtml
            Write-Host "📧 Global alert email sent for unmapped field IDs." -ForegroundColor Magenta
        } catch {
            Write-Host "❌ Failed to send unmapped fields alert email: $_" -ForegroundColor Red
            $script:HadErrors = $true
        }
    } else {
        Write-Host "Skipping email for User ID 529 (service account)." -ForegroundColor Yellow
    }
}

# -------------------------------------------------------
# BLOCK 6 : CSV EXPORT (UTF-8 WITH BOM to avoid encoding issues)
# -------------------------------------------------------

# 💡 WHY THIS CHANGE?
# Export-Csv with -Encoding UTF8 uses UTF-8 *without* BOM (Byte Order Mark),
# which can cause special characters (like í, ñ, ó, etc.) to appear incorrectly
# in Excel or downstream processing tools. Using .NET's UTF8Encoding($true)
# ensures a BOM is added to the CSV, making encoding explicit and Excel-safe.

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

# Active users - timestamped
try {
    [System.IO.File]::WriteAllLines($script:CsvActive_Timestamped, ($ResultActive | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
    Write-Host "✅ Active users exported (timestamped, UTF-8 BOM): $($script:CsvActive_Timestamped)" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to export active users: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

# Inactive users - timestamped
try {
    [System.IO.File]::WriteAllLines($script:CsvInactive_Timestamped, ($ResultInactive | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
    Write-Host "✅ Inactive users exported (timestamped, UTF-8 BOM): $($script:CsvInactive_Timestamped)" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to export inactive users: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

# Azure SQL version - fixed name
try {
    [System.IO.File]::WriteAllLines($script:CsvAzure_NoTimestamp, ($ResultActive | ConvertTo-Csv -NoTypeInformation), $utf8Bom)
    Write-Host "✅ Active users exported (fixed path, UTF-8 BOM): $($script:CsvAzure_NoTimestamp)" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to export fixed-named CSV: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

# -------------------------------------------------------
# BLOCK 7 : EMAIL SENDING (With Summary + ASCII Recap)
# -------------------------------------------------------

# Determine elapsed time
$scriptEndTime = Get-Date
$duration = $scriptEndTime - $scriptStartTime
$durationStr = $duration.ToString("hh\:mm\:ss")

# Compose HTML body with summary
$finalBody = @"
<p>Hello,</p>
<p>The PeopleSpheres export job completed successfully.</p>

<h3>Summary</h3>
<ul>
    <li><b>Total active users:</b> $($ResultActive.Count)</li>
    <li><b>Total inactive users:</b> $($ResultInactive.Count)</li>
    <li><b>Token refresh count:</b> $TokenRefreshCount</li>
    <li><b>Total duration:</b> $durationStr</li>
</ul>

<h3>Exported CSV files</h3>
<ul>
    <li><b>Active users (timestamped):</b> <a href="file:///$($script:CsvActive_Timestamped -replace '\\', '/')">$($script:CsvActive_Timestamped)</a></li>
    <li><b>Inactive users (timestamped):</b> <a href="file:///$($script:CsvInactive_Timestamped -replace '\\', '/')">$($script:CsvInactive_Timestamped)</a></li>
    <li><b>Active users (fixed path):</b> <a href="file:///$($script:CsvAzure_NoTimestamp -replace '\\', '/')">$($script:CsvAzure_NoTimestamp)</a></li>
</ul>
"@

# Optional ASCII recap in console (not in the email)
$asciiBanner = @"
╔═════════════════════════════════════════════════════════╗
║           ✅ PEOPLESPHERES EXPORT COMPLETED             ║
╠═════════════════════════════════════════════════════════╣
║ 📂 Active users exported   : $($ResultActive.Count.ToString().PadRight(25)) ║
║ 📂 Inactive users exported : $($ResultInactive.Count.ToString().PadRight(25)) ║
║ 🔁 Token refresh count     : $($TokenRefreshCount.ToString().PadRight(25)) ║
║ ⏱  Duration                : $durationStr                    ║
╚═════════════════════════════════════════════════════════╝
"@
Write-Host $asciiBanner -ForegroundColor Green

# Warn about too many token refreshes
if ($TokenRefreshCount -gt 3) {
    Write-Host "⚠ Warning: Token was refreshed $TokenRefreshCount times – check performance." -ForegroundColor Red
}

# Append script signature to final body
$finalBody += IADAdmin_AddScriptSignature

# Compose email subject
$emailSubject = "[iadlife] PeopleSpheres Export – $($scriptEndTime.ToString('yyyy-MM-dd'))"

# Set the recipients based on test mode
$emailRecipients = if ($IsTestMode) { 
    @{
        To  = @("jeremie.poujol@iadinternational.com")
        Cc  = @("jeremie.poujol@iadinternational.com")
        Bcc = "986c2ea3.iadgroup.onmicrosoft.com@fr.teams.ms"
    }
} else { 
    @{
        To  = @("exploitation.notify@iadinternational.com")
        Cc  = @("alexandre.kebaili-ext@iadinternational.com")
        Bcc = "986c2ea3.iadgroup.onmicrosoft.com@fr.teams.ms"
    }
}

# Send email
try {
    if ([string]::IsNullOrWhiteSpace($finalBody)) {
        Write-Warning "❗ Email not sent: email body is empty."
    } else {
        IADAdmin_SendMailMessage -Body $finalBody `
                                 -To $emailRecipients.To `
                                 -Cc $emailRecipients.Cc `
                                 -Bcc $emailRecipients.Bcc `
                                 -Subject $emailSubject `
                                 -BodyAsHtml
        Write-Host "📧 Email successfully sent to: $($emailRecipients.To -join ', ')" -ForegroundColor Green
    }
} catch {
    Write-Host "❌ Failed to send summary email: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

# -------------------------------------------------------
# BLOCK 8 : ERROR NOTIFICATION (if any error occurred)
# -------------------------------------------------------

if ($script:HadErrors -and (Test-Path $script:LogPath)) {
    $logLastWrite = (Get-Item $script:LogPath).LastWriteTime
    $logAgeHours  = (New-TimeSpan -Start $logLastWrite -End (Get-Date)).TotalHours

    if ($logAgeHours -le 12) {
        $subject = "[iadlife] ❌ Script Failure – PeopleSpheres Export"
        $body = @"
<p>⚠️ Hello Jérémie,</p>
<p>The PeopleSpheres export script encountered one or more errors during its execution.</p>
<p>Please review the log file located at the following path:</p>
<pre><code>$($script:LogPath)</code></pre>
<p>The log file has been attached to this email for your convenience.</p>
<p>This is an automated alert generated by the export script.</p>
"@

        try {
            # Sending the error alert to the recipients
            IADAdmin_SendMailMessage -To $emailDetails.To `
                                     -Cc $emailDetails.Cc `
                                     -Bcc $emailDetails.Bcc `
                                     -Subject $subject `
                                     -Body $body `
                                     -BodyAsHtml `
                                     -Attachments $script:LogPath
            Write-Host "📧 Error alert mail sent with attached log file." -ForegroundColor Magenta
        } catch {
            Write-Host "❌ Failed to send error alert email with log: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "⏱ Log file is older than 12 hours – no error alert sent." -ForegroundColor Yellow
    }
}

# -------------------------------------------------------
# BLOCK 9 : UPLOAD TO AZURE BLOB + BULK INSERT TO SQL
# -------------------------------------------------------

# Azure and SQL configuration
$tenant              = "e419a47d-b189-44f1-a28e-16be83c1f11e"
$subscription        = "f9d75155-d6d0-4867-b0c0-cec83ecea40c"
$userTenant          = "AzSce_PSScript@iadgroup.onmicrosoft.com"
$resourceGroupName   = "rg-frc-coreservices"
$storageAccountName  = "iadsamgmtfrccore"
$containerName       = "csvcontainer"

$serverInstance      = "sql-frc-coreservices-iad.database.windows.net"
$databaseName        = "iaddb"
$sqlLogin            = "PSScript"
$localCsvPath        = $script:CsvAzure_NoTimestamp  # should point to: E:\Powershell\03-FlatFilesStorage\AzureSQLDatabase_csv\GenerateCSV-API-PeopleSpheres.csv
$blobName            = "GenerateCSV-API-PeopleSpheres.csv"

# Load secure credentials for Azure authentication
$securePathAzure = "C:\Scripts\SecureString\$userTenant.$env:USERNAME.securestring"
if (-not (Test-Path $securePathAzure)) {
    Read-Host "Enter password for $userTenant" -AsSecureString |
        ConvertFrom-SecureString |
        Out-File -FilePath $securePathAzure
}
$securePasswordAzure = Get-Content $securePathAzure | ConvertTo-SecureString
$credAzure = New-Object System.Management.Automation.PSCredential ($userTenant, $securePasswordAzure)

# Connect to Azure
Connect-AzAccount -Credential $credAzure -Tenant $tenant -Subscription $subscription

# Get storage context
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -StorageAccountName $storageAccountName
$context = $storageAccount.Context

# Upload CSV to blob
try {
    Set-AzStorageBlobContent -Container $containerName -Blob $blobName -File $localCsvPath -Context $context -Force
    Write-Host "✅ CSV successfully uploaded to Azure Blob Storage: $blobName" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to upload CSV to blob: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

# Load secure SQL password
$securePathSql = "C:\Scripts\SecureString\$serverInstance.$env:USERNAME.securestring"
if (-not (Test-Path $securePathSql)) {
    Read-Host "Enter SQL SA password for $serverInstance" -AsSecureString |
        ConvertFrom-SecureString |
        Out-File -FilePath $securePathSql
}
$securePasswordSql = Get-Content $securePathSql | ConvertTo-SecureString
$sqlCredential = New-Object System.Management.Automation.PSCredential ($sqlLogin, $securePasswordSql)
$sqlPassword = $sqlCredential.GetNetworkCredential().Password

# BULK INSERT SQL command
$bulkInsertQuery = @"
TRUNCATE TABLE dbo.PEOPLESPHERE_Iad;
BULK INSERT dbo.PEOPLESPHERE_Iad
FROM 'csvcontainer/GenerateCSV-API-PeopleSpheres.csv'
WITH (
    DATA_SOURCE = 'blobcontainer',
    FIRSTROW = 2,
    FORMAT = 'CSV'
);
"@

# Execute query
$sqlParams = @{
    ServerInstance     = $serverInstance
    Database           = $databaseName
    Username           = $sqlLogin
    Password           = $sqlPassword
    Query              = $bulkInsertQuery
    EncryptConnection  = $true
    OutputSqlErrors    = $true
}
try {
    Invoke-Sqlcmd @sqlParams
    Write-Host "✅ BULK INSERT completed into dbo.PEOPLESPHERE_Iad" -ForegroundColor Green
} catch {
    Write-Host "❌ SQL BULK INSERT failed: $_" -ForegroundColor Red
    $script:HadErrors = $true
}

<#
-- SQL Server Management Studio (SSMS) - Table structure for dbo.PEOPLESPHERE_Iad

drop TABLE dbo.PEOPLESPHERE_Iad;
CREATE TABLE dbo.PEOPLESPHERE_Iad (
    Datededebutdansleposte                VARCHAR(255),
    Youzerdatedembauche                   VARCHAR(255),
    Departdelasociete                     VARCHAR(255),
    Youzerdatedesortie                    VARCHAR(255),
    Typecollaboration                     VARCHAR(255),
    Nom                                   VARCHAR(255),
    Prenom                                VARCHAR(255),
    Imagedeprofil                         VARCHAR(255),
    Civilite                              VARCHAR(255),
    Adresseemailprofessionnelle           VARCHAR(255),
    Telephoneportableprofessionnel        VARCHAR(255),
    Responsable                           VARCHAR(255),
    Service                               VARCHAR(255),
    Poste                                 VARCHAR(255),
    Entitelegale                          VARCHAR(255),
    Site                                  VARCHAR(255),
    Matricule                             VARCHAR(255),
    DateExeScript                         VARCHAR(255),
    Actif                                 VARCHAR(255),  -- New field for Actif (Yes/No)
    Prenom_caracteres_speciaux            VARCHAR(255)   -- New field for First Name with Special Characters
)

#>
