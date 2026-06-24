Import-Module ActiveDirectory
$seed = Get-Random


    
$orgUnits = @('OU=ServiceAccounts,OU=ENT-Management,DC=DAVITA,DC=Corp',
'OU=Svcaccts,DC=DAVITA,DC=Corp')

#Run LDAP query using search term
Write-Host -NoNewline "`n[OUs to Export]:"
                
#$GroupList = Get-ADGroup -Filter {Name -like $searchTerm} | Select-Object -Property Name | Format-Table
$orgUnits | Select-Object -Property Name,DistinguishedName | Format-Table

#Press any key to refine search terms...
#Press M to continue with Member Export...
Write-Host "Press any key to refine search terms..."
Write-Host "Press M if you are ready to export Members of all located groups"

 
Write-Host "`nExporting Members of OUs...."

$filters = 'samaccountname,pwdLastSet' 
    
#Interate through matching groups 
foreach ($unit in $orgUnits) {

    #-l filters column output for LDAP query, removing this parameter will output all possible LDAP attributes
    #samaccountname = output usernames


    $args = ' -d ' + $unit.DistinguishedName + ' -n -l ' + $filters + ' -f ' + '.\' + '"' + $unit.DistinguishedName + '_powershell_' + $seed + '.csv"'
        
    $args
    Invoke-Expression "csvde.exe $args"
    }

#Combine separate CSVDE exports into single file in working directory, while keeping single Header row
Get-ChildItem -Filter *powershell_$seed.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv .\CombinedOUExport_$seed.csv -NoTypeInformation -Append
