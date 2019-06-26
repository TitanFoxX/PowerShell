# - Create infinte While-loop that will ask for username input on each cycle.

While($True){
    $Username = Read-Host "Please enter an username "
    $SAMuser = Get-ADUser -Filter {SamAccountName -like $Username}

    # - If the SamAccountName exists, set the appropriate attribute to "Office365".
    
    If($SAMuser){
        Write-Host "SAM Name Found : $SAMuser"
        Set-ADUser $SAMuser -add @{"extensionattribute7"="Office365"}
        
        # - Comment the previous rule and uncomment the next one to clear the attribute.
        
        #Set-ADUser $SAMuser -Clear "extensionattribute7"
        
        # - Display the ExtensionAttribute information after setting or clearing this parameter.
        
        Get-ADUser $SAMuser -Properties extensionAttribute7
    }
    
    # - If the SamAccountName can not be found, display an error and start over.

    Else{
        Write-Host "Error : Name not found !"
    }
}