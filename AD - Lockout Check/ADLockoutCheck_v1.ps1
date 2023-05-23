Import-Module ActiveDirectory

do{
    do{
        
        #Clear screen, show instructions
        Clear-Host
        Get-Content _title.txt | Write-Host

        #Accept LDAP search term from user, repeat while input is null
        $searchTerm = Read-Host -Prompt 'Enter account name (ex. svc_account_name)'
        } while ($searchTerm -eq "")


   #Search for account name and return formatted list of relevant attributes
    Get-ADUser -Identity $searchTerm -Properties * | 
    Select-Object name,
    samaccountname,
    lockedout,
    accountlockouttime,
    lastbadpasswordattempt,
    badlogoncount,
    passwordlastset |
    Format-List
    

    #Prompt to continue
    Write-Host "`nPress any key to search again or Q to Quit..."
    $cont = Read-Host

} while ($cont -ne "Q")