# - This script will obtain the Firewall-status from systems that run 
#   Windows Server 2016. It does not work on Server 2008R2 or older.

# - Suppress error messages.

$ErrorActionPreference = "SilentlyContinue"

# - Create new, empty array and fill it with the names of all enabled systems that
#   run a Windows Server operating system.

$Servers = @()
$Servers += Get-ADComputer -Filter {enabled -eq $True} -Properties Operatingsystem | Where-Object {$_.Operatingsystem -like "*Server*"} | Select-Object Name -ExpandProperty Name

# - Sort the content of this array.

[Array]::Sort($Servers)

# - Start logging to defined folder and file.

Start-Transcript -Path "<your path>\Firewall_Status.txt" -Append | Out-Null
    
    # - This Foreach-loop will check if the firewall profile is enabled on the
    #   remote computer and will output a corresponding message.
    
    Foreach($Server in $Servers){
        $FW_Status = Invoke-Command -ComputerName $Server {Get-NetFirewallProfile -Profile Domain | Select-Object -ExpandProperty Enabled} -ErrorAction SilentlyContinue
        If($FW_Status -eq $True){
            "Firewall status on $Server is enabled"
        }
        ElseIf($FW_Status -eq $False){
            "Firewall status on $Server is disabled"
        }
        Else{
            "Firewall status on $Server can not be determined"
        }
    }

# - Stop logging.
Stop-Transcript | Out-Null