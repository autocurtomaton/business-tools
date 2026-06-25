<#
.SYNOPSIS
    Purpose:    Report Entra/Intune status from consumed device input.

.DESCRIPTION
    The script is designed to be menu-driven and interactive with two input options:

    1. CSV file method
    Maintain list of DRA-exported CSV files for each OU that are refreshed periodically.
    Select from pre-determined markets at runtime to consume current CSV file as configured.
    Export desired lists at end of analysis.

    2. Paste method
    For custom selection, paste list of desired device names to lookup Entra properties.
    Cannot derive certain AD attributes for pasted objects, like OU, so this will show a placeholder.

.NOTES
    v3.2.5
        - Reorder summary lines
        - EA8 missing only counted if Entra record is healthy
        - testing
#>


#region -- Config --
[CmdletBinding()]
param()

$scriptVersion     = "3.2.6"
$LogFolder         = "C:\_support\Log.Files\Intune_IntlDeviceReadiness"
$BulkImportFolder  = "C:\_support\Intune_IntlDeviceReadiness"
$BulkImportPath    = Join-Path $BulkImportFolder "EntraIntuneReadiness-Input.csv"
$LogPath           = Join-Path $LogFolder "m365_EntraIntuneReadiness.log"
$ExportFolder      = $LogFolder
$LogFileTime       = Get-Date -Format "yyyy.MM.ddTHH.mm.ss"

$CsvMarketTable = @(
    [pscustomobject]@{ MarketName = 'Japan';          Alpha2 = 'JP'; Region = 'APAC'  }
    [pscustomobject]@{ MarketName = 'Malaysia';       Alpha2 = 'MY'; Region = 'APAC'  }
    [pscustomobject]@{ MarketName = 'Singapore';      Alpha2 = 'SG'; Region = 'APAC'  }
    [pscustomobject]@{ MarketName = 'United Kingdom'; Alpha2 = 'GB'; Region = 'EMEA'  }
    [pscustomobject]@{ MarketName = 'Poland';         Alpha2 = 'PL'; Region = 'EMEA'  }
    [pscustomobject]@{ MarketName = 'Portugal';       Alpha2 = 'PT'; Region = 'EMEA'  }
    [pscustomobject]@{ MarketName = 'Saudi Arabia';   Alpha2 = 'SA'; Region = 'EMEA'  }
    [pscustomobject]@{ MarketName = 'Brazil';         Alpha2 = 'BR'; Region = 'LATAM' }
    [pscustomobject]@{ MarketName = 'Chile';          Alpha2 = 'CL'; Region = 'LATAM' }
    [pscustomobject]@{ MarketName = 'Colombia';       Alpha2 = 'CO'; Region = 'LATAM' }
    [pscustomobject]@{ MarketName = 'Ecuador';        Alpha2 = 'EC'; Region = 'LATAM' }
    [pscustomobject]@{ MarketName = 'Panama';         Alpha2 = 'PA'; Region = 'LATAM' }
)

#endregion

#region -- Logging --
if (-not (Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
    )

    try {
        $timestamp = [DateTime]::UtcNow.ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
        $line = "[$timestamp] [$Level] $Message"
        Write-Host $line
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Host "Logging failure: $_" -ForegroundColor Yellow
    }
}

Write-Log "===== Entra / Intune Readiness Log Starting v$scriptVersion ====="
#endregion

#region -- Functions --
function Test-GraphConnection {
    [CmdletBinding()]
    param()

    if (-not (Get-MgContext)) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Device.Read.All","DeviceManagementManagedDevices.Read.All" | Out-Null
    }
}

function Read-MenuChoice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][scriptblock]$LabelScript,
        [string]$Prompt = "Enter selection number"
    )

    if (-not $Items -or $Items.Count -lt 1) { throw "Menu has no items: $Title" }

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan

    for ($i = 0; $i -lt $Items.Count; $i++) {
        $label = & $LabelScript $Items[$i]
        Write-Host ("  [{0}] {1}" -f ($i+1), $label) -ForegroundColor Yellow
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

function Get-MarketInputCsvPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Market
    )

    Join-Path $BulkImportFolder ("EntraIntuneReadiness-Input-{0}.csv" -f $Market.Alpha2)
}

function Get-CsvDeviceName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Row)

    # Supports AD exports like: Object Class, Name, Location, Email, Group Type, Group Scope
    # Also supports common alternate headers from device reports.
    $candidateColumns = @(
        'Name',
        'DisplayName',
        'DeviceName',
        'DeviceDisplayName',
        'ComputerName',
        'Hostname'
    )

    foreach ($col in $candidateColumns) {
        $p = $Row.PSObject.Properties[$col]
        if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
            return ([string]$p.Value).Trim()
        }
    }

    return $null
}

function Get-CsvDomainOU {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Row)

    $candidateColumns = @(
        'DomainOU',
        'Location',
        'DistinguishedName',
        'CanonicalName'
    )

    foreach ($col in $candidateColumns) {
        $p = $Row.PSObject.Properties[$col]
        if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
            return ([string]$p.Value).Trim()
        }
    }

    return $null
}

function Get-ExportMarketTag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$InputItems
    )

    $domainOUs = @(
        $InputItems |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.DomainOU) } |
            Select-Object -ExpandProperty DomainOU -Unique
    )

    if ($domainOUs.Count -lt 1) {
        return "Pasted"
    }

    $tags = foreach ($ou in $domainOUs) {
        if ($ou -match 'DVA-(APAC|EMEA|LATAM)[/\\]([A-Z]{2})') {
            "$($matches[1])-$($matches[2])"
        }
    }

    $tags = @($tags | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    if ($tags.Count -eq 1) {
        return $tags[0]
    }

    if ($tags.Count -gt 1) {
        return "MultiMarket"
    }

    return "UnknownMarket"
}

function Read-BulkInput {
    [CmdletBinding()]
    param()

    $mode = Read-MenuChoice -Title "Bulk input method:" -Items @("CSV file", "Paste device names") -LabelScript { param($x) $x }

    if ($mode -eq "CSV file") {
        $market = Read-MenuChoice `
            -Title "Select CSV market:" `
            -Items ($CsvMarketTable | Sort-Object Region, MarketName) `
            -LabelScript { param($m) "$($m.Region) - $($m.MarketName) ($($m.Alpha2))" }

        $path = Get-MarketInputCsvPath -Market $market

        Write-Host ""
        Write-Host "Selected input file: $path" -ForegroundColor Cyan

        while (-not (Test-Path $path)) {
            Write-Host ""
            Write-Host "⚠️ CSV file not found for selected market:" -ForegroundColor DarkYellow
            Write-Host "   $path" -ForegroundColor Yellow

            $missingAction = Read-MenuChoice `
                -Title "What would you like to do?" `
                -Items @(
                    "Select a different market",
                    "Enter a custom CSV path",
                    "Cancel"
                ) `
                -LabelScript { param($x) $x }

            switch ($missingAction) {
                "Select a different market" {
                    $market = Read-MenuChoice `
                        -Title "Select CSV market:" `
                        -Items ($CsvMarketTable | Sort-Object Region, MarketName) `
                        -LabelScript { param($m) "$($m.Region) - $($m.MarketName) ($($m.Alpha2))" }

                    $path = Get-MarketInputCsvPath -Market $market

                    Write-Host ""
                    Write-Host "Selected input file: $path" -ForegroundColor Cyan
                }

                "Enter a custom CSV path" {
                    $customPath = Read-Host "Enter full CSV path"
                    if (-not [string]::IsNullOrWhiteSpace($customPath)) {
                        $path = $customPath.Trim('"').Trim()
                    }
                }

                "Cancel" {
                    throw "Cancelled. CSV file not found: $path"
                }
            }
        }

        $rows = Import-Csv $path
        if (-not $rows) { throw "CSV contained no rows." }

        $items = foreach ($r in $rows) {
            $name = Get-CsvDeviceName -Row $r
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                [pscustomobject]@{
                    DeviceName = $name
                    DomainOU  = Get-CsvDomainOU -Row $r
                    Source    = $path
                    Market    = $market.Alpha2
                    Region    = $market.Region
                }
            }
        }

        if (-not $items) {
            throw "No usable device names found in CSV."
        }

        return @($items)
    }

    Write-Host ""
    Write-Host "Paste device names, one per line. When done, enter a blank line." -ForegroundColor Cyan
    $items = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $items.Add([pscustomobject]@{
            DeviceName    = $line.Trim()
            DomainOU      = $null
            Source        = "Pasted"
        })
    }
    if ($items.Count -lt 1) { throw "No device names entered." }
    return $items.ToArray()
}

function Get-DevicesByExactDisplayName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DeviceDisplayName)

    $escaped = $DeviceDisplayName -replace "'", "''"
    $select = "id,displayName,trustType,deviceTrustType,profileType,deviceId,operatingSystem,accountEnabled,registrationDateTime,approximateLastSignInDateTime,extensionAttributes"
    $devices = Get-MgDevice -All -Filter "displayName eq '$escaped'" -Select $select

    if (-not $devices) { return @() }
    return @($devices)
}

function Get-DeviceEA8Value {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Device)

    $ap = $Device.AdditionalProperties
    if ($null -eq $ap) { return $null }

    $ext = $ap['extensionAttributes']
    if ($null -eq $ext) { return $null }

    try {
        $v = $ext['extensionAttribute8']
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
    } catch { }

    $p = $ext.PSObject.Properties['extensionAttribute8']
    if ($p -and -not [string]::IsNullOrWhiteSpace($p.Value)) { return $p.Value }

    return $null
}

function Get-JoinStatusLabel {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Device)

    $tt = $Device.trustType
    if (-not $tt -and $Device.PSObject.Properties.Name -contains 'deviceTrustType') { $tt = $Device.deviceTrustType }

    switch ($tt) {
        'ServerAD'  { 'Hybrid AAD Joined' }
        'AzureAD'   { 'AAD Joined' }
        'Workplace' { 'AAD Registered' }
        default     { if ($tt) { $tt } else { 'Unknown' } }
    }
}

function Test-HybridPending {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Device)

    # Per Microsoft guidance for Microsoft Graph PowerShell:
    #   Healthy hybrid joined: TrustType = ServerAD and ProfileType = RegisteredDevice
    #   Pending hybrid joined: TrustType = ServerAD and ProfileType -ne RegisteredDevice
    #
    # The Entra admin center "Join type" can still show Microsoft Entra hybrid joined
    # because that label maps to TrustType/DeviceTrustType. The separate "Registered"
    # column is what shows Pending vs a completed registration timestamp.
    $tt = $Device.trustType
    if (-not $tt -and $Device.PSObject.Properties.Name -contains 'deviceTrustType') { $tt = $Device.deviceTrustType }

    if ($tt -ne 'ServerAD') { return $false }

    $entraProfile = [string]$Device.profileType
    if ($entraProfile -eq 'RegisteredDevice') { return $false }

    return $true
}

function Select-HybridDeviceMatch {
    [CmdletBinding()]
    param([object[]]$MatchResults)

    if (-not $MatchResults -or $MatchResults.Count -eq 0) { return $null }

    # Ignore AzureAD joined and Workplace/Registered duplicates. Only ServerAD matters here.
    $hybridMatches = @($MatchResults | Where-Object { $_.trustType -eq 'ServerAD' -or $_.deviceTrustType -eq 'ServerAD' })
    if ($hybridMatches.Count -lt 1) { return $null }

    # Prefer a fully joined hybrid object over a pending hybrid object.
    $fullyJoined = @($hybridMatches | Where-Object { -not (Test-HybridPending -Device $_) })
    if ($fullyJoined.Count -ge 1) {
        return ($fullyJoined | Sort-Object approximateLastSignInDateTime -Descending | Select-Object -First 1)
    }

    return ($hybridMatches | Sort-Object registrationDateTime -Descending | Select-Object -First 1)
}

function Get-IntuneManagedDeviceByName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DeviceName)

    $escaped = $DeviceName -replace "'", "''"
    $select = "id,deviceName,azureADDeviceId,managementAgent,managementState,complianceState,operatingSystem,lastSyncDateTime,userPrincipalName,isEncrypted"

    try {
        $managed = Get-MgDeviceManagementManagedDevice -All -Filter "deviceName eq '$escaped'" -Select $select
        if (-not $managed) { return @() }
        return @($managed)
    }
    catch {
        Write-Log "Managed device lookup failed for $DeviceName : $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Resolve-IntuneEnrollment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DeviceName,
        [object]$HybridDevice
    )

    $managedMatches = @(Get-IntuneManagedDeviceByName -DeviceName $DeviceName)
    if ($managedMatches.Count -lt 1) { return $null }

    if ($HybridDevice) {
        $hybridDeviceId = [string]$HybridDevice.DeviceId
        $idMatched = @($managedMatches | Where-Object {
            ([string]$_.AzureADDeviceId -eq $hybridDeviceId)
        })
        if ($idMatched.Count -ge 1) {
            return ($idMatched | Sort-Object lastSyncDateTime -Descending | Select-Object -First 1)
        }
    }

    return ($managedMatches | Sort-Object lastSyncDateTime -Descending | Select-Object -First 1)
}

function Get-PercentText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Count,
        [Parameter(Mandatory)][int]$Total
    )

    if ($Total -lt 1) { return "0.0%" }
    return ("{0:N1}%" -f (($Count / $Total) * 100))
}

function Show-ReadinessSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object]$Results)

    $normalizedResults = @()
    foreach ($r in $Results) { if ($null -ne $r) { $normalizedResults += $r } }

    $total = $normalizedResults.Count
    $summary = @(
        [pscustomobject]@{ Category = 'Devices enrolled in Intune'; Count = @($normalizedResults | Where-Object { $_.IntuneEnrollmentState -eq 'Enrolled' }).Count; Percent = $null }
        [pscustomobject]@{ Category = 'Devices not enrolled in Intune'; Count = @($normalizedResults | Where-Object { $_.IntuneEnrollmentState -eq 'NotEnrolled' }).Count; Percent = $null }
        [pscustomobject]@{ Category = 'Intune enrolled but not encrypted'; Count = @($normalizedResults | Where-Object { $_.IntuneEnrollmentState -eq 'Enrolled' -and $_.IntuneIsEncrypted -ne $true }).Count; Percent  = $null}
        [pscustomobject]@{ Category = 'Devices with Entra join issue'; Count = @($normalizedResults | Where-Object { $_.EntraHealth -ne 'Healthy' }).Count; Percent = $null }
        [pscustomobject]@{ Category = 'Devices with EA8 missing'; Count = @($normalizedResults | Where-Object { $_.EntraHealth -eq 'Healthy' -and $_.EA8Present -eq $false }).Count; Percent = $null }
    )

    foreach ($row in $summary) {
        $row.Percent = Get-PercentText -Count $row.Count -Total $total
    }

    Write-Host ""
    Write-Host "----- Intune Readiness Summary -----" -ForegroundColor Cyan
    Write-Host "Total input devices: $total" -ForegroundColor Cyan
    $summary | Format-Table -AutoSize
}

function Export-Results {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object]$Results,
        [Parameter(Mandatory)][ValidateSet('Full','MissingIntune','HybridIssues','EA8Issues')][string]$ExportType
    )

    $normalizedResults = @()
    foreach ($r in $Results) {
        if ($null -ne $r) { $normalizedResults += $r }
    }

    if (-not (Test-Path $ExportFolder)) {
        New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
    }

    switch ($ExportType) {
        'MissingIntune' {
            $data = @($normalizedResults | Where-Object { $_.IntuneEnrollmentState -eq 'NotEnrolled' })
            $path = Join-Path $ExportFolder "EntraIntuneReadiness-$ExportMarketTag-MissingIntune-$LogFileTime.csv"
        }
        'HybridIssues' {
            $data = @($normalizedResults | Where-Object { $_.EntraHealth -ne 'Healthy' })
            $path = Join-Path $ExportFolder "EntraIntuneReadiness-$ExportMarketTag-HybridIssues-$LogFileTime.csv"
        }
        'EA8Issues' {
            $data = @($normalizedResults | Where-Object { $_.EntraHealth -eq 'Healthy' -and $_.EA8Present -eq $false })
            $path = Join-Path $ExportFolder "EntraIntuneReadiness-$ExportMarketTag-EA8Issues-$LogFileTime.csv"
        }
        default {
            $data = @($normalizedResults)
            $path = Join-Path $ExportFolder "EntraIntuneReadiness-$ExportMarketTag-Full-$LogFileTime.csv"
        }
    }

    $data | Export-Csv -NoTypeInformation -Path $path -Encoding UTF8
    Write-Host "✅ Exported: $path" -ForegroundColor Green
}

function Export-SummaryResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object]$Results
    )

    $normalizedResults = @()
    foreach ($r in $Results) {
        if ($null -ne $r) { $normalizedResults += $r }
    }

    if (-not (Test-Path $ExportFolder)) {
        New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
    }

    $total = $normalizedResults.Count
    $runDate = Get-Date

    $summary = @(
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'Total Devices'
            Count      = $total
            Percent    = Get-PercentText -Count $total -Total $total
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'Already Enrolled'
            Count      = @($normalizedResults | Where-Object { $_.Readiness -eq 'AlreadyEnrolled' }).Count
            Percent    = $null
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category = 'Intune Enrolled, Not Encrypted'
            Count    = @($Results | Where-Object { $_.IntuneEnrollmentState -eq 'Enrolled' -and $_.IntuneIsEncrypted -ne $true }).Count
            Percent = $null
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'Not Intune Enrolled'
            Count      = @($normalizedResults | Where-Object { $_.IntuneEnrollmentState -eq 'NotEnrolled' }).Count
            Percent    = $null
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'Ready For Enrollment'
            Count      = @($normalizedResults | Where-Object { $_.Readiness -eq 'ReadyForEnrollment' }).Count
            Percent    = $null
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'Hybrid Issues'
            Count      = @($normalizedResults | Where-Object { $_.EntraHealth -ne 'Healthy' }).Count
            Percent    = $null
        }
        [pscustomobject]@{
            RunDate    = $runDate
            SourceFile = $BulkImportPath
            Category   = 'EA8 Issues'
            Count      = @($normalizedResults | Where-Object { $_.EntraHealth -eq 'Healthy' -and $_.EA8Present -eq $false }).Count
            Percent    = $null
        }

    )

    foreach ($row in $summary) {
        if ($null -eq $row.Percent) {
            $row.Percent = Get-PercentText -Count $row.Count -Total $total
        }
    }

    $path = Join-Path $ExportFolder "EntraIntuneReadiness-$ExportMarketTag-Summary-$LogFileTime.csv"
    $summary | Export-Csv -NoTypeInformation -Path $path -Encoding UTF8
    Write-Host "✅ Exported: $path" -ForegroundColor Green
}

#endregion

#region -- Main --
Clear-Host

try {
    Test-GraphConnection

    $inputItems = Read-BulkInput
    $ExportMarketTag = Get-ExportMarketTag -InputItems $inputItems

    Write-Host ""
    Write-Host "Bulk count: $($inputItems.Count)" -ForegroundColor Cyan
    $confirm = Read-Host "Type YES to continue"
    if ($confirm -ne "YES") { Write-Host "Cancelled." -ForegroundColor DarkYellow; return }

    $results = New-Object System.Collections.ArrayList
    $n = 0

    foreach ($item in $inputItems) {
        $n++
        $name = $item.DeviceName
        Write-Host ""
        Write-Host ("[{0}/{1}] Processing: {2}" -f $n, $inputItems.Count, $name) -ForegroundColor Cyan

        $entraMatches = @(Get-DevicesByExactDisplayName -DeviceDisplayName $name)
        $hybridDevice = Select-HybridDeviceMatch -MatchResults $entraMatches
        $registeredCount = @($entraMatches | Where-Object { $_.trustType -eq 'Workplace' -or $_.deviceTrustType -eq 'Workplace' }).Count
        $hybridCount = @($entraMatches | Where-Object { $_.trustType -eq 'ServerAD' -or $_.deviceTrustType -eq 'ServerAD' }).Count

        $entraHealth = $null
        $entraIssue  = $null
        $joinStatus  = $null
        $isPending   = $false
        $entraObjId  = $null
        $entraDevId  = $null
        $entraEA8    = $null
        $entraProfileType = $null
        $entraRegistrationDateTime = $null

        if (-not $hybridDevice) {
            $entraHealth = 'Unhealthy'
            if ($entraMatches.Count -gt 0) {
                $entraIssue = 'NoHybridObject_RegisteredOrOtherOnly'
            }
            else {
                $entraIssue = 'NoEntraObject'
            }
            $joinStatus = 'No Hybrid AAD Joined object'
        }
        else {
            $joinStatus = Get-JoinStatusLabel -Device $hybridDevice
            $isPending  = Test-HybridPending -Device $hybridDevice
            $entraObjId = $hybridDevice.Id
            $entraDevId = $hybridDevice.DeviceId
            $entraEA8   = Get-DeviceEA8Value -Device $hybridDevice
            $entraProfileType = $hybridDevice.profileType
            $entraRegistrationDateTime = $hybridDevice.registrationDateTime

            if ($isPending) {
                $entraHealth = 'Unhealthy'
                $entraIssue  = 'HybridPending'
            }
            else {
                $entraHealth = 'Healthy'
                $entraIssue  = $null
            }
        }

        # EA8 readiness must be based on the Entra hybrid device object's extensionAttribute8 only.
        # The source CSV DomainOU is retained for context/export, but it does not prove EA8 exists in Entra.
        $ea8Present = -not [string]::IsNullOrWhiteSpace($entraEA8)

        $managed = Resolve-IntuneEnrollment -DeviceName $name -HybridDevice $hybridDevice
        $intuneState = if ($managed) { 'Enrolled' } else { 'NotEnrolled' }

        $readiness = if ($entraHealth -eq 'Healthy' -and $ea8Present -eq $true) {
            if ($intuneState -eq 'Enrolled') { 'AlreadyEnrolled' } else { 'ReadyForEnrollment' }
        }
        else {
            'NotReady'
        }

        $resultRow = [pscustomobject]@{
            DeviceName               = $name
            IntuneEnrollmentState    = $intuneState
            Readiness                = $readiness
            EntraHealth              = $entraHealth
            EntraIssue               = $entraIssue
            EA8Present               = $ea8Present
            DomainOU                 = $item.DomainOU
            EntraEA8                 = $entraEA8
            JoinStatus               = $joinStatus
            HybridPending            = $isPending
            EntraProfileType         = $entraProfileType
            EntraRegistrationDateTime = $entraRegistrationDateTime
            EntraObjectId            = $entraObjId
            EntraDeviceId            = $entraDevId
            EntraMatchCount          = $entraMatches.Count
            HybridMatchCount         = $hybridCount
            RegisteredMatchCount     = $registeredCount
            IntuneManagedDeviceId    = if ($managed) { $managed.Id } else { $null }
            IntuneAzureADDeviceId    = if ($managed) { $managed.AzureADDeviceId } else { $null }
            IntuneManagementAgent    = if ($managed) { $managed.ManagementAgent } else { $null }
            IntuneManagementState    = if ($managed) { $managed.ManagementState } else { $null }
            IntuneComplianceState    = if ($managed) { $managed.ComplianceState } else { $null }
            IntuneIsEncrypted       = if ($managed) { $managed.IsEncrypted } else { $null }
            IntuneEncryptionIssue   = if ($managed -and $managed.IsEncrypted -ne $true) { $true } else { $false }
            IntuneLastSyncDateTime   = if ($managed) { $managed.LastSyncDateTime } else { $null }
            IntunePrimaryUser        = if ($managed) { $managed.UserPrincipalName } else { $null }
        }
        [void]$results.Add($resultRow)

        switch ($readiness) {
            'AlreadyEnrolled'    { Write-Host "  ✅ Enrolled in Intune" -ForegroundColor Green }
            'ReadyForEnrollment' { Write-Host "  ✅ Ready, not yet enrolled" -ForegroundColor Green }
            default              { Write-Host "  ⚠️ Not ready: $entraIssue; EA8Present=$ea8Present; Intune=$intuneState" -ForegroundColor DarkYellow }
        }
    }

    # Convert ArrayList to a plain PowerShell array once processing is complete.
    # Avoid @($results), which can throw "Argument types do not match" in some hosts/modules.
    $resultsArray = foreach ($r in $results) { $r }
    Show-ReadinessSummary -Results $resultsArray

    while ($true) {
        $choice = Read-MenuChoice -Title "Choose next action:" -Items @("Export full list", "Export devices missing Intune", "Export devices with Hybrid issue", "Export devices with EA8 issue", "Export summary CSV", "Exit") -LabelScript { param($x) $x }

        switch ($choice) {
            'Export full list'                 { Export-Results -Results $resultsArray -ExportType Full }
            'Export devices missing Intune'    { Export-Results -Results $resultsArray -ExportType MissingIntune }
            'Export devices with Hybrid issue' { Export-Results -Results $resultsArray -ExportType HybridIssues }
            'Export devices with EA8 issue'    { Export-Results -Results $resultsArray -ExportType EA8Issues }
            'Export summary CSV'               { Export-SummaryResults -Results $resultsArray }
            'Exit'                             { return }
        }
    }
}
catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($results -and $results.Count -gt 0) {
        Write-Host "Dumping partial results to output:" -ForegroundColor DarkYellow
        foreach ($r in $results) { $r }
    }
    throw
}
#endregion
