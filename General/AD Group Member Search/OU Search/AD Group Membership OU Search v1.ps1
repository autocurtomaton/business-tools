Import-Module ActiveDirectory
$seed = Get-Random

#Dev:
#   Remove individual .csvs or provide option to keep/remove
#   Add additional comments for other -l options
#   Switch statement for Columns

do{
    do{
        do{
            #Clear screen, show instructions
            Clear-Host
            Get-Content _readme.txt | Write-Host

            #Accept LDAP search term from user, repeat while input is null
            $searchTerm = Read-Host -Prompt 'Enter search term'
        } while ($searchTerm -eq "")

        $orgUnits = Get-ADOrganizationalUnit -Filter {Name -like $searchTerm}

        #Run LDAP query using search term
        Write-Host -NoNewline "`n[OUs Located]:"
                
        #$GroupList = Get-ADGroup -Filter {Name -like $searchTerm} | Select-Object -Property Name | Format-Table
        if ($orgUnits -eq $null) {Write-Host ' NO OUs LOCATED -- Press ENTER to try again'; Read-Host}
        else {
        $orgUnits | Select-Object -Property Name,DistinguishedName | Format-Table

        #Press any key to refine search terms...
        #Press M to continue with Member Export...
        Write-Host "Press any key to refine search terms..."
        Write-Host "Press M if you are ready to export Members of all located groups"

        $searchAgain = Read-Host
        }
    } while ($searchAgain -ne "M")

    Write-Host "`nExporting Members of discovered OUs...."

    #$Names = Get-ADGroup -Filter {Name -like $searchTerm}

    Write-Host "Enter attributes to export, separated by comma."
    Write-Host "Format is attribute1,attribute2,attribute3,..."
    $filters = Read-Host "(Example -- name,samaccountname,mail) "
    
    #Interate through matching groups 
    foreach ($unit in $orgUnits) {

        #-l filters column output for LDAP query, removing this parameter will output all possible LDAP attributes
        #samaccountname = output usernames


        $args = ' -d ' + $unit.DistinguishedName + ' -n -l ' + $filters + ' -f ' + '.\' + '"' + $unit.DistinguishedName + '_powershell_' + $seed + '.csv"'
        
        #$args
        Invoke-Expression "csvde.exe $args"
        }

    #Combine separate CSVDE exports into single file in working directory, while keeping single Header row
    Get-ChildItem -Filter *powershell_$seed.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv .\CombinedOUExport_$seed.csv -NoTypeInformation -Append
    

    #Prompt to continue
    Write-Host "`nPress any key to search again or Q to Quit..."
    $cont = Read-Host

} while ($cont -ne "Q")
