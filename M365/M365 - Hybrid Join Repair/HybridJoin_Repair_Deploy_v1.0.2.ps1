<#
.SYNOPSIS
    Author: Curt Harter    
    Purpose: Resolve Hybrid-join objects pending/missing

.NOTES
    v1.0.0
        - Initial build
    v1.0.1
        - Updating GPUpdate section to resolve script timeouts (WS1)
    v1.0.2
        - Improved pre-check logic for repair criteria to avoid Entra-only actions
        - Added variable for updateGPO control
#>


#region == CONFIG ==
$scriptVersion      = "1.0.2"
$LogPath            = "C:\_davsupp\Log.Files\HybridJoin_WS1_Repair.log"
$tenantId           = "d0746369-7df7-4138-87d2-a9b75386157f"
$performUpdateGPO   = $false

#endregion

#region == LOGGING ==
function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")] [string]$Level = "INFO"
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

Write-Log "Starting Hybrid Join Repair Deployment v$scriptVersion"

#endregion

#region == FUNCTIONS ==

function Write-Section {
    param([string]$Title)
    Add-Content -Path $LogPath -Value "" -Encoding UTF8
    Add-Content -Path $LogPath -Value ("=" * 70) -Encoding UTF8
    Add-Content -Path $LogPath -Value $Title -Encoding UTF8
    Add-Content -Path $LogPath -Value ("=" * 70) -Encoding UTF8
}

function Confirm-JoinState {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId
    )

    $dsreg = dsregcmd.exe /status 2>&1

    $azureAdJoined   = Get-DsRegStatusValue -DsRegOutput $dsreg -Key "AzureAdJoined"
    $domainJoined    = Get-DsRegStatusValue -DsRegOutput $dsreg -Key "DomainJoined"
    $currentTenantId = Get-DsRegStatusValue -DsRegOutput $dsreg -Key "TenantId"
    $deviceAuth      = Get-DsRegStatusValue -DsRegOutput $dsreg -Key "DeviceAuthStatus"

    Write-Log "Pre-check values: AzureAdJoined=$azureAdJoined, DomainJoined=$domainJoined, TenantId=$currentTenantId, DeviceAuthStatus=$deviceAuth"

    if ($azureAdJoined -eq "YES" -and
        $domainJoined -eq "YES" -and
        $currentTenantId -eq $TenantId -and
        $deviceAuth -eq "SUCCESS") {

        Write-Log "Device already appears healthy and joined to the expected tenant. No repair needed." "WARN"
        exit 0
    }

    if ($azureAdJoined -eq "YES" -and $currentTenantId -ne $TenantId) {
        Write-Log "Device is AzureAdJoined but tenant mismatch detected. Proceeding with repair." "WARN"
        return
    }

    if ($azureAdJoined -ne "YES") {
        Write-Log "Device is not fully AzureAdJoined. Proceeding with repair."
        return
    }

    if ($domainJoined -ne "YES" -and $azureAdJoined -eq "YES") {
        Write-Log "Device is AzureAdJoined but not domain joined. Repair is for hybrid join only and will not proceed." "WARN"
        exit 0
    }

    if ($deviceAuth -and $deviceAuth -ne "SUCCESS") {
        Write-Log "DeviceAuthStatus is not SUCCESS. Proceeding with repair." "WARN"
        return
    }

    Write-Log "Unable to classify join state. Cancelling repair."
    exit 1
}

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [string]$ActionName = $FilePath,
        [int[]]$SuccessExitCodes = @(0)
    )

    try {
        $output = & $FilePath @ArgumentList 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        Write-Log "$ActionName failed to launch: $($_.Exception.Message)" "ERROR"
        return [pscustomobject]@{
            Succeeded = $false
            ExitCode  = $null
            Output    = @($_.Exception.Message)
        }
    }

    if ($null -ne $output) {
        foreach ($line in @($output)) {
            Write-Log "$ActionName output: $line"
        }
    }

    if ($exitCode -notin $SuccessExitCodes) {
        Write-Log "$ActionName failed with exit code $exitCode" "ERROR"
        return [pscustomobject]@{
            Succeeded = $false
            ExitCode  = $exitCode
            Output    = @($output)
        }
    }

    Write-Log "$ActionName completed successfully with exit code $exitCode"
    return [pscustomobject]@{
        Succeeded = $true
        ExitCode  = $exitCode
        Output    = @($output)
    }
}

function Get-DsRegStatusValue {
    param(
        [Parameter(Mandatory=$true)][object[]]$DsRegOutput,
        [Parameter(Mandatory=$true)][string]$Key
    )

    $pattern = '^\s*' + [regex]::Escape($Key) + '\s*:\s*(.+?)\s*$'

    foreach ($line in $DsRegOutput) {
        $text = [string]$line
        if ($text -match $pattern) {
            return $Matches[1].Trim()
        }
    }

    return $null
}

function Invoke-AADLeave {
    $result = Invoke-NativeCommand -FilePath "dsregcmd.exe" -ArgumentList "/leave" -ActionName "dsregcmd /leave"
    if (-not $result.Succeeded) { exit 1 }
}

function Remove-AADKeys {
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ") {
        try {
            Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ" -Recurse -Force -ErrorAction Stop
            Write-Log "Removed registry path: HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ"
        }
        catch {
            Write-Log "Failed removing CDJ: $($_.Exception.Message)" "WARN"
        }
    }
    else {
        Write-Log "Registry path not present: HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ"
    }
}

function Update-GPO {
    $result = Invoke-NativeCommand -FilePath "gpupdate.exe" -ArgumentList "/target:computer","/force","/wait:0" -ActionName "gpupdate /target:computer /force /wait:0"
    if (-not $result.Succeeded) {
        Write-Log "gpupdate failed; continuing." "WARN"
    }
}

function Enable-JoinTasks {
    try {
        Write-Log "Checking Workplace Join scheduled tasks"
        $tasks = Get-ScheduledTask -TaskPath "\Microsoft\Windows\Workplace Join\" -ErrorAction Stop

        foreach ($task in $tasks) {
            try {
                Enable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop | Out-Null
                Write-Log "Enabled task: $($task.TaskName)"
            }
            catch {
                Write-Log "Could not enable task $($task.TaskName): $($_.Exception.Message)" "WARN"
            }
        }

        return $true
    }
    catch {
        Write-Log "Failed enumerating Workplace Join tasks: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Invoke-JoinTask {
    $result = Invoke-NativeCommand -FilePath "schtasks.exe" -ArgumentList "/run","/tn","\Microsoft\Windows\Workplace Join\Automatic-Device-Join" -ActionName "Start Automatic-Device-Join"
    if (-not $result.Succeeded) { exit 1 }
}

function Invoke-AADJoin {
    $result = Invoke-NativeCommand -FilePath "dsregcmd.exe" -ArgumentList "/join" -ActionName "dsregcmd /join"
    if (-not $result.Succeeded) {
        Write-Log "Explicit join failed; collecting final status anyway." "WARN"
    }
}
#endregion

#region == MAIN ==

# -- Log pre-repair state --
Write-Section "PRE-REPAIR DSREG STATUS"
$preStatus = dsregcmd.exe /status 2>&1
$preStatus | Out-File -FilePath $LogPath -Append -Encoding UTF8

# -- Check if repair needed --
Confirm-JoinState -TenantId $tenantId

# -- Leave AAD --
Invoke-AADLeave

# -- Remove AAD remnants --
Remove-AADKeys

# -- Update GPO (Optional)--
if ($performUpdateGPO) {Update-GPO}

# -- Enable join tasks --
if (-not (Enable-JoinTasks)) { exit 1 }

# -- Run join task --
Invoke-JoinTask
Start-Sleep -Seconds 20

# -- Manually attempt join (Optional) - Should not be needed when using autojoin schtsk --
#Invoke-AADJoin
#Start-Sleep -Seconds 10

# -- Log post-repair state --
Write-Section "POST-REPAIR DSREG STATUS"
$postStatus = dsregcmd.exe /status 2>&1
$postStatus | Out-File -FilePath $LogPath -Append -Encoding UTF8

Write-Log "Hybrid Join repair finished"
exit 0

#endregion
