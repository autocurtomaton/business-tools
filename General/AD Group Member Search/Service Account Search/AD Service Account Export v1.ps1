Import-Module ActiveDirectory

$OU1 = "OU=ServiceAccounts,OU=ENT-Management,DC=DAVITA,DC=Corp"
$OU2 = "OU=Svcaccts,DC=DAVITA,DC=Corp"
$Properties = @("SamAccountName", "pwdLastSet")

function Convert-FileTime {
    param([long]$FileTime)
    if ($FileTime -eq 0) {
        return "Never"
    }
    else {
        return [DateTime]::FromFileTime($FileTime).ToString("MM-dd-yyyy")
    }
}

# Function to get users from an OU with date conversions
function Get-UsersFromOU {
    param (
        [string]$OU
    )
    Get-ADUser -Filter * -SearchBase $OU -SearchScope Subtree -Properties $Properties | ForEach-Object {
        $PwdLastSet = Convert-FileTime $_.pwdLastSet

        [PSCustomObject]@{
            SamAccountName     = $_.SamAccountName
            PwdLastSet         = $PwdLastSet
        }
    }
}

$AllUsers = @(
    Get-UsersFromOU -OU $OU1
    Get-UsersFromOU -OU $OU2
)

$filetime = get-date -f "yyyy-MM-ddTHHmmss"
$AllUsers | Export-Csv -path ".\ServiceAccountsExport_$filetime.csv" -NoTypeInformation


