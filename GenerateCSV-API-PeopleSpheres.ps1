# -------------------------------------------------------
# DEBUG SCRIPT: Inspect one PeopleSpheres user in detail
# -------------------------------------------------------

# -------------------------------
# BLOCK 0: Token Loading Fallback
# -------------------------------

# If AccessToken is not already set, load it from secure file
if (-not $script:AccessToken) {
    $Username     = "support-itps@iadinternational.com"
    $SecurePath   = "C:\Scripts\SecureString\$Username.$env:USERNAME.securestring"
    $ClientId     = "monportailrh-web-app"
    $AuthUrl      = "https://sso.monportailrh.com/auth/realms/Internal-idp/protocol/openid-connect/token"

    if (-not (Test-Path $SecurePath)) {
        Write-Host "🔐 Secure string not found. Please re-save your password." -ForegroundColor Red
        return
    }

    $SecureString = Get-Content $SecurePath | ConvertTo-SecureString
    $PlainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    )

    $AuthHeaders = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "Accept"       = "application/json"
    }

    $Body = @{
        "grant_type" = "password"
        "client_id"  = $ClientId
        "username"   = $Username
        "password"   = $PlainPassword
    }

    try {
        $response = Invoke-RestMethod -Uri $AuthUrl -Method POST -Body $Body -Headers $AuthHeaders
        $script:AccessToken = $response.access_token
        Write-Host "✅ Access token retrieved successfully." -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to retrieve access token: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# ----------------------------
# BLOCK 1: Configuration
# ----------------------------
$UserLastName = "POUJOL"   # ← Set this to search by name (case insensitive)
$UserIdManual = ""         # ← Set this instead to inspect by ID directly

$BaseApiUrl = "https://rest.monportailrh.com"

# ----------------------------
# BLOCK 2: Helper - Headers
# ----------------------------
function Get-ApiHeaders {
    return @{
        "Authorization" = "Bearer $script:AccessToken"
        "Accept"        = "application/json"
    }
}

# ----------------------------
# BLOCK 3: Display Fields
# ----------------------------
function Inspect-UserFields($UserId, $FullName) {
    $FieldUrl = "$BaseApiUrl/psos/$UserId/fields?active=true&include=type,items,options,settings,assignment_settings"

    try {
        $Fields = (Invoke-RestMethod -Uri $FieldUrl -Headers (Get-ApiHeaders)).data

        Write-Host "`n===============================" -ForegroundColor Cyan
        Write-Host "👤 $FullName" -ForegroundColor Cyan
        Write-Host "===============================" -ForegroundColor Cyan

        foreach ($Item in $Fields) {
            $Id        = $Item.id
            $Alias     = $Item.alias
            $Label     = if ($Item.label) { $Item.label } elseif ($Item.name) { $Item.name } else { $Alias }
            $TypeName  = $Item.type.name
            $Value     = $Item.value_details

            $DisplayValue = if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
                ($Value -join ", ")
            } elseif ($Value -is [psobject]) {
                $Value | ConvertTo-Json -Compress
            } else {
                $Value
            }

            Write-Host "`n📌 ID    : $Id" -ForegroundColor DarkGray
            Write-Host "🔑 Alias : $Alias"
            Write-Host "🏷️ Label : $Label"
            Write-Host "📄 Type  : $TypeName"
            Write-Host "💬 Value : $DisplayValue"
        }
    } catch {
        Write-Host "❌ Error fetching fields for user ID ${UserId}:`n$($_.Exception.Message)" -ForegroundColor Red
    }
}

# ----------------------------
# BLOCK 4: Main Logic
# ----------------------------
if ($UserLastName) {
    Write-Host "🔎 Searching for users with last name $UserLastName..." -ForegroundColor Cyan
    $SearchUrl = "$BaseApiUrl/search?name=$UserLastName&pso-type=usr&active=1&include=data.quick_actions"

    try {
        $Response = Invoke-RestMethod -Uri $SearchUrl -Headers (Get-ApiHeaders)
        $UserMatches = $Response.data.data
    } catch {
        Write-Host "❌ Failed to search users: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    if (-not $UserMatches) {
        Write-Host "❌ No users found with name '$UserLastName'." -ForegroundColor Red
        exit 1
    }

    foreach ($User in $UserMatches) {
        $UserId = $User.id
        $FullName = "$($User.display_name) [ID: $UserId]"
        Inspect-UserFields -UserId $UserId -FullName $FullName
    }
} elseif ($UserIdManual) {
    $FullName = "User ID: $UserIdManual"
    Inspect-UserFields -UserId $UserIdManual -FullName $FullName
} else {
    Write-Host "⚠️ Please set either `$UserLastName or `$UserIdManual to begin." -ForegroundColor Yellow
}
