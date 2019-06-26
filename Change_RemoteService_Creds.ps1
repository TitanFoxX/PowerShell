# - Setting ExecutionPolicy and ErrorAction parameters.

Set-ExecutionPolicy -ExecutionPolicy Bypass
$ErrorActionPreference = 'SilentlyContinue'

# - Create credentials for remote service.

$ServiceUserName = Read-Host "Please enter username for remote service"
$ServiceUserPass = Read-Host "Please enter corresponding password"

# - Create a new array and fill it with all active AD-computers that match your input.
# - Sort the array.

$AD_Computers = @()
$AD_Computers += Get-ADComputer -Filter {Enabled -eq $True} | Where{$_.Name -like '*<your input>*'} | Select-Object Name -ExpandProperty Name
[Array]::Sort($AD_Computers)

# - Loop through the results stored in $AD_Computers

Foreach($AD_Computer in $AD_Computers){

    # - If service is present, store your defined service name into $ServicePresent

    $ServicePresent = Get-WmiObject Win32_Service -Filter "Name='<your service>'" -ComputerName $AD_Computer

    # - If ServicePresent contains a result, continue the script.

    If($ServicePresent.Name -eq '<your service>'){

        # - Check the current credentials used to manage this service and store it in $RunAs.

        $RunAs = (Get-WmiObject Win32_Service -Filter "Name='<your service>'" -ComputerName $AD_Computer).StartName

        # - If credentials match 'LocalSystem', continue the script.

        If($RunAs -eq 'LocalSystem'){

            # - Stop and disable the service.
            
            Get-Service "<your service>" -ComputerName $AD_Computer | Stop-Service -Force -WarningAction SilentlyContinue
            
            # - Import new credentials into the service you've selected.
            
            $ServicePresent.Change($null,$null,$null,$null,$null,$null,$ServiceUserName,$ServiceUserPass,$null,$null,$null) | Out-Null
            
            # - Recheck if the new credentials are active.
            
            $RunAs = (Get-WmiObject Win32_Service -Filter "<your service>'" -ComputerName $AD_Computer).StartName
            
            # - Start the service if this is the case.

            If($RunAs -eq '<your ServiceUserName>'){
                Get-Service "<your service>" -ComputerName $AD_Computer | Set-Service -StartupType Automatic -Status Running
                
                # - Pause script for 5 seconds.
                
                Start-Sleep -Seconds 5
                
                # - Output simple summary over selected service.
                
                'ServerName  : '+ $AD_Computer
                'Credentials : '+ $RunAs
                'Status      : '+ (Get-Service "<your service>" -ComputerName $AD_Computer | Select-Object Status -ExpandProperty Status)
                Write-Host
            }
            
            # - If new credentials are not active, output error message.

            Else{
                $AD_Computer + ' : Unknown error with <your service>'
            }
        }
    }
    
    # - If the selected service is not found, output error message.

    Else{
        $AD_Computer + ' : Your service is not installed on this system'
        Write-Host
    }
}