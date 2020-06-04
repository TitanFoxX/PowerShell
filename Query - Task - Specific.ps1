$Servers = @()
$Servers += Get-ADComputer -Filter {enabled -eq $true} -Properties Name | Where-Object {($_.Name -like "*<enter remote computer>*")} | Select-Object Name -ExpandProperty Name

Foreach($Server in $Servers){
    $Result = Get-ScheduledTask -Taskname '*<enter scheduled taskname>*' -ErrorAction SilentlyContinue
    If($Result){
        "Task found on server $Server"
    }
}
