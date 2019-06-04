while($true){
    $Username = Read-Host "Please enter a username "
    $SAMuser = Get-ADUser -Filter {SamAccountName -like $Username}

    If($SAMuser){
        Write-Host "SAM Name Found : $SAMuser"
        Set-ADUser $SAMuser -add @{"extensionattribute7"="Office365"}
        Get-ADUser $SAMuser -Properties extensionAttribute7
    }
    Else{
        Write-Host "Error : Name not found !"
    }
}