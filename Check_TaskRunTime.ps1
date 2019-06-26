# - Define the name of the Scheduled Task you want to monitor.

$TaskName = "Proxy" # - Can be changed to any preferred value.

# - Define the timeout (in minutes) on which the alarm should be triggered.

[int]$MaxTimeOut = "5" # - Can be changed to any preferred value.

# - Get the date and time on which the task last ran.

$LastRunTime = (Get-ScheduledTask -TaskName "$TaskName") | Get-ScheduledTaskInfo | Select-Object LastRunTime,LastTaskResult

# - Check if the task ran succesfully last time. If so, run the IF-statement.

If(($LastRunTime.LastTaskResult) -eq [int]0){
    
    # - Get the current date
    
    $Current_Date = Get-Date
    
    # - Calculate the difference in whole minutes between the LastRunTime and the current date.
    
    $Run_Difference = ((((New-TimeSpan -Start $LastRunTime.LastRunTime -End $Current_Date).TotalSeconds) / 60) | Out-String).Split(',')[0]
    
    # - If the value in $Run_Difference is larger than the value in $MaxTimeOut, give an error.
    
    If([int]$Run_Difference -ge [int]$MaxTimeOut){
        "Warning : The last time that $TaskName ran exceeds the MaxTimeOut value !"
    }
    
    # - If the task runs to specs, produce this notification.
    
    Else{
        "Note : The $TaskName task runs according to it's settings."
    }
}

# - If the status of the tasks LastRunResult is not (0x)0, 
#   produce this error on the end of the script:

Else{
    "Error : Task did not complete succesfully on last run."
}