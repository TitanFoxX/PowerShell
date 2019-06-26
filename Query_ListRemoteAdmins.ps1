# - Create new, empty array and fill it with all active servers which run Windows Server

$ADComputers = @()
$ADComputers += Get-ADComputer -Filter {(OperatingSystem -like "*Windows*Server*")} | Select-Object Name -ExpandProperty Name | Sort-Object Name

# - Start logging to defined path and filename.

Start-Transcript -Path "<your path>\Remote_Local_Admins.txt" -Append | Out-Null
    
    # - This Foreach will check every remote computer for members in the "Local Admins" group and report these to the log-file.
    
    Foreach($ADComputer in $ADComputers){
        
        # - Store all users from target computer in $LocalAdmins.

        $RemoteUsers = gwmi win32_groupuser -ComputerName $ADComputer
        
        # - Create a new array and store all users from $RemoteUsers that are in the Administrators group.
        
        $Admins = @()
        $Admins += $RemoteUsers | Where{$_.GroupComponent –like '*"Administrators"'}
        
        # - Create simple output that lists the administrators on the target computer.

        "Local administrators on $ADComputer :"
        Foreach($Admin in $Admins){    
            $Admin.PartComponent -replace '.*Name="','' -replace '"',''
        }
        ""
    }

# - Stop logging.
Stop-Transcript | Out-Null  