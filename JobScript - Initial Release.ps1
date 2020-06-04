#[CmdletBinding()]
#Param(
#    [Parameter(Mandatory=$True,Position=0)]
#    [string]$ExecutablePath,
#    [Parameter(Mandatory=$True,Position=1)]
#    [string]$ExecutableName,
#    [Parameter(Mandatory=$True,Position=2)]
#    [string]$LogfileName,
#    [Parameter(Mandatory=$True,Position=3)]
#    [string]$LogfileEvaluator,
#    [Parameter(Mandatory=$True,Position=4)]
#    [string]$Task_Path,
#    [Parameter(Mandatory=$True,Position=5)]
#    [string]$NewTaskName,
#    [Parameter(Mandatory=$True,Position=6)]
#    [int]$Sleep_Timer,
#    [Parameter(Mandatory=$True,Position=7)]
#    [int]$MultiPass,
#    [Parameter(Mandatory=$True,Position=8)]
#    [int]$IsLastAction,
#    [Parameter(Mandatory=$True,Position=9)]
#    [string]$RebootDay,
#    [Parameter(Mandatory=$True,Position=10)]
#    [int]$RebootHour,
#    [Parameter(Mandatory=$True,Position=11)]
#    [int]$RebootMinutes
#)

# - Test Parameters ...

[string]$ExecutablePath = ""
[string]$ExecutableName = ""
[string]$LogfileName = ""
[string]$LogfileEvaluator = "**"
[string]$Task_Path = ""
[string]$NewTaskName = ""
[int]$Sleep_Timer = 1
[int]$MultiPass = 0
[int]$IsLastAction = 0
[string]$RebootDay = ""
[int]$RebootHour = 16
[int]$RebootMinutes = 37

# - If the $RebootHour or $RebootMinutes variables contain only one digit (first 9 hours or minutes on the hour),
#   add a leading '0' to prevent issues with the Do-Until loop.

[string]$RebootHour = ([int]$RebootHour)
[string]$RebootMinutes = ([int]$RebootMinutes)

If(($RebootHour.Length) -eq [string]1){
    $RebootHour = "0"+"$RebootHour"
}
If(($RebootMinutes.Length) -eq [string]1){
    $RebootMinutes = "0"+"$RebootMinutes"
}

# - These variables contain snippets of the string required by the Invoke-RestMethod at the 
#   end of the script. This is purely an esthetical functionality to keep the script organized.

$WH_URLSnip1 = ''
$WH_URLSnip2 = ''

# - Create an DO-loop which ends on the date specified with the values stored in
#   $RebootDay, $RebootHour and $RebootMinutes

Do{
    # - Disable check UNC Zone (supresses security dialog)
    
    $env:SEE_MASK_NOZONECHECKS = 1
    
    # - Run the Export executable and wait till it finishes
    
    Start-Process -WindowStyle Hidden -FilePath ("$ExecutablePath"+"\"+"$ExecutableName") -Wait -ErrorAction 'SilentlyContinue'
        
    # - Reset the check for UNC Zone to default settings
    
    Remove-Item env:SEE_MASK_NOZONECHECKS
    
    # - Read the last two lines of the executable's LOG-file
    
    $Status = Get-Content $ExecutablePath\$LogfileName -Tail 2 | Out-String

    # - If the last two lines do not match the predefined 'success' value, write out an error 
    #   message to Technical Support via Microsoft Teams and restart the scheduled task.

    If($Status -notlike $LogfileEvaluator){
        # - Define period of time since last successful job run.
        
        $Job_Run_Date = Get-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName | Get-ScheduledTaskInfo
        $Current_Date = Get-Date
        $Run_Difference = New-TimeSpan -Start $Job_Run_Date.LastRunTime -End $Current_Date
        
        # - Define the message shown within Microsoft Teams.
        
        $Message = ""
        $Message = $Message + ""
        $Message = $Message + ""
        $Message = $Message | ConvertTo-Json

        Invoke-RestMethod -Method Post -ContentType "Application/Json" -Body "{'text': '$Message'}" -Uri https://$WH_URLSnip1/$WH_URLSnip2 | Out-Null

        # - Stop the current task, then close this script process.

        Stop-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName
        Exit
    }

    # - If variable $MultiPass equals '1' and $IsLastAction equals '0', then this job is part of a task 
    #   with several actions which should only run once. When done, the script will end and continue to
    #   the next action in line.
    
    If(([int]$MultiPass -eq 1) -and ([int]$IsLastAction -eq 0)){
        Exit
    }
    
    # - If variable $MultiPass equals '1' and $IsLastAction equals '1', then this job is last in line of 
    #   a task with several actions which should only run once. When done, the parent task will be stopped,
    #   pause for 5 seconds, restart the parent scheduled task and end this script instance.
    
    If(([int]$MultiPass -eq 1) -and ([int]$IsLastAction -eq 1)){
        Stop-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName
        Start-Sleep -Seconds 5
        Start-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName
        Exit
    }
    
    # - If neither of the above conditions are met, this job will run indefinitely - either
    #   till the parent task is stopped or the job crashed.
    
    Else{
        # - Pause job to ensure optimal performance. Any $Sleep_Timer value below 5 minutes
        #   (including value missing all together) will specify a pause of 5 seconds to allow 
        #   the job to write to the log file before restarting the job. 
                
        If(!$Sleep_Timer){
            Start-Sleep -Seconds 5
        }
        ElseIf($Sleep_Timer -le 4){
            Start-Sleep -Seconds 5
        }
        Else{
            Start-Sleep -Seconds ($Sleep_Timer * 60)
        }
    }
    
    # - Check if the parent task is still running. If a status '0x2' is
    #   present in Task Scheduler, the task has stopped although the job might not
    #   have crashed at all. Once started, this script instance will end.
    
    If([int](Get-ScheduledTaskInfo -TaskPath $Task_Path -TaskName $NewTaskName).LastTaskResult -eq [int]2){
        Stop-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName
        Start-Sleep -Seconds 5
        Start-ScheduledTask -TaskPath $Task_Path -TaskName $NewTaskName
        Exit
    }
}
Until(
    ([String]$RebootDate = "$RebootDay".ToLower()+","+"$RebootHour"+":"+"$RebootMinutes") -eq
    ([String]$SystemDate = (((Get-Date).DayOfWeek).ToString()).ToLower()+','+((Get-Date).Hour).ToString()+':'+((Get-Date).Minute).ToString()) -eq
    $True
)