<#
.SYNOPSIS
    Uses ABM API to export device inventory of ABM
    
    Note: Currently only pulls non-released devices.
          However, release attributes are supported for single device calls.
          
.NOTES
    v1.0.0
        -Initial build
    v1.1.0
        -Added column selection during API
        -Added column export based on API columns
    v1.2.0
        -Resolved bug with column data exporting
    v2.0.0
        -Refactored into interactive
        -Added post-processing for platform reporting
    v2.1.0
        -Reorganized script blocks and added config variables
        -Added colorized UI + logging + menu helper
        -Fixed headers initialization + URL encoding + function name mismatch
    v2.2.0
        -Added menu coloring
        -Fixed filter call
        -Added logging
    v2.3.0
        -Removed calling url noise
        -Added static count progress for data retrieval indicator
    v2.3.1
        -Send 429 messages to log instead of output
        -Increase backoff rate
#>

#region == CONFIG ==
$scriptVersion = "2.3.1"

$clientId = "BUSINESSAPI.d709a28b-7f4d-49b7-89ef-160425c0d9fa"
$teamId   = "BUSINESSAPI.d709a28b-7f4d-49b7-89ef-160425c0d9fa"
$keyId    = "d0cd7906-cb39-42d3-a143-c6d4e8a97279"
$privateKeyPath = "C:\Scripting\Endpoint_Automation_Unencrypted.pem"

$tokenEndpoint = "https://account.apple.com/auth/oauth2/v2/token"
$baseDeviceAPI = "https://api-business.apple.com/v1"

$exportPath = "C:\Scripting"

# Logging
$LogFolder  = "C:\_davsupp\Log.Files\ABM_DeviceExport"
$LogFileTime = Get-Date -Format "yyyy.MM.ddTHH.mm.ss"
$LogPath    = Join-Path $LogFolder "ABM_DeviceExport_$LogFileTime.log"

# API paging / throttling
$pageLimit        = 1000
$baselineThrottle = 1500   # ms between successful calls
$maxRetries429    = 8      # per page

$global:AllFieldsList = @(
    "serialNumber"
    "addedToOrgDateTime"
    "releasedFromOrgDateTime"
    "releaserId"
    "releaserEntityType"
    "updatedDateTime"
    "deviceModel"
    "productFamily"
    "productType"
    "deviceCapacity"
    "partNumber"
    "orderNumber"
    "color"
    "status"
    "orderDateTime"
    "imei"
    "meid"
    "eid"
    "purchaseSourceId"
    "purchaseSourceType"
    "wifiMacAddress"
    "bluetoothMacAddress"
    "ethernetMacAddress"
)
#endregion

#region == FUNCTIONS ==

# ----------------------------
# UI / LOGGING (added)
# ----------------------------

function Confirm-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

Confirm-Folder -Path $exportPath
Confirm-Folder -Path $LogFolder

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","OK","STEP","TITLE")] [string]$Level = "INFO",
        [bool]$NoConsole = $false
    )

    $timestamp = [DateTime]::UtcNow.ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
    $line = "[$timestamp] [$Level] $Message"

    # Only write to console if not suppressed
    if (-not $NoConsole) {
        switch ($Level) {
            "TITLE" { Write-Host $line -ForegroundColor Cyan }
            "STEP"  { Write-Host $line -ForegroundColor Cyan }
            "OK"    { Write-Host $line -ForegroundColor Green }
            "WARN"  { Write-Host $line -ForegroundColor DarkYellow }
            "ERROR" { Write-Host $line -ForegroundColor Red }
            default { Write-Host $line -ForegroundColor Gray }
        }
    }

    try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch {}
}

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "============================================" -ForegroundColor DarkCyan
    Write-Host "   ABM Device Export Tool  v$scriptVersion"  -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor DarkCyan
    Write-Host ""
}

function Read-MenuChoice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][scriptblock]$LabelScript,
        [string]$Prompt = "Enter selection number"
    )

    if (-not $Items -or $Items.Count -lt 1) {
        throw "Menu has no items: $Title"
    }

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan

    for ($i = 0; $i -lt $Items.Count; $i++) {
        $label = & $LabelScript $Items[$i]
        Write-Host ("  [{0}] {1}" -f ($i + 1), $label) -ForegroundColor Yellow
    }

    while ($true) {
        $raw = Read-Host "$Prompt (1-$($Items.Count))"
        $num = 0
        if ([int]::TryParse($raw, [ref]$num) -and $num -ge 1 -and $num -le $Items.Count) {
            return $Items[$num - 1]
        }
        Write-Host "Invalid selection. Try again." -ForegroundColor DarkYellow
    }
}

function Convert-AbmValue {
    param([object]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return ($Value -join ";")
    }

    return $Value
}

function Get-PlatformSelection {

    $choice = Read-MenuChoice `
        -Title "=== ABM Device Export Tool ===" `
        -Items @("Export All Devices","Export iPads Only","Export macOS Only","Quit") `
        -LabelScript { param($x) $x } `
        -Prompt "Select option"

    switch ($choice) {
        "Export All Devices"  { return "All" }
        "Export iPads Only"   { return "iPad" }
        "Export macOS Only"   { return "Mac" }
        "Quit"                { return $null }
    }
}

function Select-Fields {
    param($availableFields)

    $selected = @()

    do {
        Write-Host ""
        Write-Host "=== Field Selection (toggle by number, Enter to finish) ===" -ForegroundColor Cyan

        for ($i = 0; $i -lt $availableFields.Count; $i++) {
            $field = $availableFields[$i]
            $mark = if ($selected -contains $field) { "[X]" } else { "[ ]" }
            Write-Host ("  [{0}] {1} {2}" -f ($i + 1), $mark, $field) -ForegroundColor Yellow
        }

        $userinput = Read-Host "Toggle field number"
        if ([string]::IsNullOrWhiteSpace($userinput)) { break }

        $index = 0
        if (-not [int]::TryParse($userinput, [ref]$index)) { continue }
        $index--

        if ($index -ge 0 -and $index -lt $availableFields.Count) {
            $field = $availableFields[$index]

            if ($selected -contains $field) {
                $selected = $selected | Where-Object { $_ -ne $field }
            } else {
                $selected += $field
            }
        }

    } while ($true)

    if (-not $selected -or $selected.Count -lt 1) {
        Write-Log "No fields selected; defaulting to Standard." "WARN"
        return @("serialNumber","productFamily","deviceModel","status","orderNumber","orderDateTime","purchaseSourceType")
    }

    return $selected
}

function Get-FieldProfile {

    $choice = Read-MenuChoice `
        -Title "=== Field Profiles ===" `
        -Items @("Minimal","Standard","Full","Custom") `
        -LabelScript { param($x) $x } `
        -Prompt "Select profile"

    switch ($choice) {
        "Minimal" {
            return @("serialNumber","productFamily","deviceModel","status")
        }
        "Standard" {
            return @("serialNumber","productFamily","deviceModel","status","orderNumber","orderDateTime","purchaseSourceType")
        }
        "Full" {
            return $global:AllFieldsList
        }
        "Custom" {
            return Select-Fields $global:AllFieldsList
        }
    }
}

function Resolve-Devices {
    param($devices, $platform)

    switch ($platform) {
        "iPad" { return $devices | Where-Object { $_.attributes.productFamily -eq "iPad" } }
        "Mac"  { return $devices | Where-Object { $_.attributes.productFamily -eq "Mac" } }
        default { return $devices }
    }
}

function Export-Devices {
    param($devices, $fieldsList)

    $rows = foreach ($d in $devices) {
        $row = [ordered]@{}
        foreach ($f in $fieldsList) {
            $row[$f] = Convert-AbmValue $d.attributes.$f
        }
        [pscustomobject]$row
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $exportPath "ABM_Devices_$timestamp.csv"

    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "✅ Export complete: $csvPath" -ForegroundColor Green
    Write-Log "Export complete: $csvPath" "OK"
}

function Confirm-Export {
    param($platform, $count, $fieldsList)

    Write-Host ""
    Write-Host "=== Export Summary ===" -ForegroundColor Cyan
    Write-Host ("Platform : {0}" -f $platform) -ForegroundColor Yellow
    Write-Host ("Devices  : {0}" -f $count) -ForegroundColor Yellow
    Write-Host ("Fields   : {0}" -f ($fieldsList -join ", ")) -ForegroundColor Yellow

    $confirm = Read-Host "Proceed with export? (Y/N)"
    return ($confirm -eq "Y")
}

function ConvertTo-Base64Url {
    param([byte[]]$bytes)
    $base64 = [Convert]::ToBase64String($bytes)
    $base64 = $base64.Split('=')[0].Replace('+', '-').Replace('/', '_')
    return $base64
}

function Initialize-API {
    # Client assertion
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $exp = $now + 3600

    $jwtHeader = @{
        alg = "ES256"
        kid = $keyId
    } | ConvertTo-Json -Compress

    $jwtPayload = @{
        iss = $teamId
        sub = $clientId
        aud = $tokenEndpoint
        iat = $now
        exp = $exp
        jti = [guid]::NewGuid().ToString()
    } | ConvertTo-Json -Compress

    $encodedHeader  = ConvertTo-Base64Url ([System.Text.Encoding]::UTF8.GetBytes($jwtHeader))
    $encodedPayload = ConvertTo-Base64Url ([System.Text.Encoding]::UTF8.GetBytes($jwtPayload))
    $unsignedToken  = "$encodedHeader.$encodedPayload"

    # Sign JWT ES256
    $privateKey = Get-Content $privateKeyPath -Raw
    $ecdsa = [System.Security.Cryptography.ECDsa]::Create()
    $ecdsa.ImportFromPem($privateKey)

    $signatureBytes = $ecdsa.SignData(
        [System.Text.Encoding]::UTF8.GetBytes($unsignedToken),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256
    )

    $encodedSignature = ConvertTo-Base64Url $signatureBytes
    $clientAssertion  = "$unsignedToken.$encodedSignature"

    # Request access token
    $tokenBody = @{
        grant_type            = "client_credentials"
        client_id             = $clientId
        client_assertion      = $clientAssertion
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = "business.api"
    }

    Write-Log "Requesting OAuth token..." "STEP"
    $response = Invoke-RestMethod -Method POST -Uri $tokenEndpoint `
        -Body $tokenBody `
        -ContentType "application/x-www-form-urlencoded"

    $accessToken = $response.access_token

    return @{
        Authorization = "Bearer $accessToken"
        Accept        = "application/json"
    }
}

function Get-ABMDevices {
    param($fieldsList)

    $headers = Initialize-API

    $fieldsParam = ($fieldsList -join ",")
    $fieldsParam = [uri]::EscapeDataString($fieldsParam)

    $nextUrl = "$baseDeviceAPI/orgDevices?limit=$pageLimit&fields[orgDevices]=$fieldsParam"

    $allDevices = @()
    
    if (-not $script:__progressInitialized) {
        Write-Host "Retrieving devices..." -NoNewline -ForegroundColor Cyan
        $script:__progressInitialized = $true
    }

    do {
        $retryCount = 0
        while ($true) {
            try {
                #Write-Host "Calling URL: $nextUrl" -ForegroundColor Cyan
                #Write-Log "Calling URL: $nextUrl" "STEP"

                $result = Invoke-RestMethod -Method GET -Uri $nextUrl -Headers $headers -ErrorAction Stop
                break
            }
            catch {
                $status = $null
                try { $status = [int]$_.Exception.Response.StatusCode } catch {}

                if ($status -eq 429) {
                    $retryCount++

                    if ($retryCount -gt $maxRetries429) {
                        Write-Log "429 max retries exceeded. Stopping pagination at: $nextUrl" "ERROR"

                        Write-Host "❌ API rate limit exceeded. Stopping retrieval." -ForegroundColor Red

                        $nextUrl = $null
                        break
                    }

                    $delay = [math]::Min(60, [math]::Pow(5, $retryCount))

                    Write-Log "429 rate limit. Retry $retryCount/$maxRetries429. Sleeping $delay seconds..." "WARN" -NoConsole $true

                    Start-Sleep -Seconds $delay
                    continue
                }

                Write-Log "API failure: $($_.Exception.Message)" "ERROR"
                throw
            }
        }

        if (-not $nextUrl) { break }

        if ($result.data) {
            $allDevices += $result.data            
            $displayCount = "{0:N0}" -f $allDevices.Count

            Write-Host ("`rRetrieving devices... {0}" -f $displayCount) -NoNewline -ForegroundColor Cyan
        }

        if ($result.links -and $result.links.next) {
            $nextUrl = $result.links.next
        } else {
            $nextUrl = $null
        }

        Start-Sleep -Milliseconds $baselineThrottle

    } while ($nextUrl)

    write-host ""
    Write-Log "Finished retrieval. Total devices collected: $($allDevices.Count)" "OK"
    return $allDevices
}

#endregion

#region == MAIN ==
Write-Banner
Write-Log "===== ABM Device Export Log Starting v$scriptVersion =====" "TITLE"

$platform = Get-PlatformSelection
if ($null -eq $platform) {
    Write-Log "User exited." "WARN"
    return
}

$fieldsList = Get-FieldProfile

$allDevices = Get-ABMDevices -fieldsList $fieldsList

$filteredDevices = Resolve-Devices -devices $allDevices -platform $platform

if (-not (Confirm-Export -platform $platform -count $filteredDevices.Count -fieldsList $fieldsList)) {
    Write-Host "❌ Export canceled" -ForegroundColor DarkYellow
    Write-Log "Export canceled by user." "WARN"
    return
}

# Export
Export-Devices -devices $filteredDevices -fieldsList $fieldsList

#endregion