# - Query the date on which the password has last changed for selected SamAccountName.

Get-ADUser -Filter * -Properties passwordlastset | Where{$_.SamAccountName -like '**'}