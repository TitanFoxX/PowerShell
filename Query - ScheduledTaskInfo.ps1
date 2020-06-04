# - Path for Powershell executable.
# - Location of Job-script file.
# - UNC Target Path for new tasks.
# - Silently continue on errors.
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

$PShell_Path = 'C:\Windows\System32\WindowsPowerShell\v1.0\'
$JobScript = ''
$ErrorActionPreference = 'SilentlyContinue'

# - Task Credentials.

$NewTaskUser = ''
$SecurePassword = $Password = Read-Host 'Please enter the password for  ' -AsSecureString
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $NewTaskUser, $SecurePassword
$Password = $Credentials.GetNetworkCredential().Password

# - Connect to the 'Task Schedule Service'.
# - Set ROOT-folder to '\'.
# - Check if folder 'Ruud_Test' exists.
# - If 'Ruud_Test' exists, set it as ROOT-folder. If not, create it first.

$scheduleObject = New-Object -ComObject schedule.service
$scheduleObject.connect()
$rootFolder = $scheduleObject.GetFolder('\')
$rootFolder = $scheduleObject.GetFolder('\Ruud_Bos-Test')
If($rootFolder.Name -notlike '*Ruud_Bos-Test*'){
    $rootFolder = $scheduleObject.GetFolder('\')
    $rootFolder.CreateFolder('Ruud_Bos-Test') | Out-Null
    $rootFolder = $scheduleObject.GetFolder('\Ruud_Bos-Test')
}

# - Reading existing tasks;
# - Exclude tasks located in the Task Scheduler's root ('\').
# - Exclude tasks located in folders containing the name 'Microsoft'.
# - Exclude tasks that are already in Powershell (PoSh) format.

$JobTasks = @()
$JobTasks = Get-ScheduledTask | Where{(((($_.TaskPath -ne '\') -and $_.TaskPath -ne '') -and ($_.TaskPath -notlike '*Microsoft*') -and $_.TaskName -notlike '*PoSh*'))} | Select-Object *

# - Foreach cycles through all available tasks.
Foreach($JobTask in $JobTasks){
    
    #$JobTask = Get-ScheduledTask -TaskPath '' -TaskName '' | Select-Object *

    # - Create three new empty arrays.
    $OldActions = @();$NewActions = @();$TaskActions = @()

    # - Fill $OldActions with all actions of the current scheduled task.
    $OldActions += $JobTask.Actions.Execute
    
    # - Get the relative path and scheduled task name to use as JobScript parameter.
    
    $Task_Path = $JobTask.TaskPath
    $Task_Name = $JobTask.TaskName
    
    # - Get date-triggers from task for use in new task and convert the
    #   whole minutes in ISO8601 Date-Time source format to an integer.
    $Task_Trigger_Interval = $JobTask.Triggers.Repetition.Interval
    $Sleep_Timer = [Xml.XmlConvert]::ToTimeSpan($Task_Trigger_Interval).Minutes
    
    # - $IsLastAction has a value of '0' by default. If a task has more than one action,
    #   the last action will get a value of '1' which will cause the last action to restart
    #   the entire task to ensure continuous. 
    [int]$IsLastAction = 0

    # - Generate new task name.
    $NewTaskName = (($JobTask.TaskName)+" - PoSh")

    # - The $MultiPass variable decides whether the new Scheduled Task contains either a
    #   single (0) or consecutive (1) action(s).
    # - The default value for this variable is '0'.
    [int]$MultiPass = 0

    # - If the source Scheduled Task has more then one (1) action, set $MultiPass to '1'.
    # - Setting the variable $MultiPass to '1' will cause the JobScript of each action to
    #   end after one run and continue with the next action.
    If(($OldActions.Count) -gt 1){
        [int]$MultiPass = 1
    }

    # - Foreach cycles through all available task actions.
    Foreach($OldAction in $OldActions){
        # - Store current path of task action into $InputPath.
        # - Select first index (drive letter) as string-separator.
        # - Remove drive letter from $InputPath.
        
        $InputPath = ($OldAction.Substring(0, $OldAction.LastIndexOf('\')))
        $Separator = $InputPath.Split('\')[0]
        $InputPath = $InputPath.Split($Separator,3)
        
        # - Set variables for new scheduled task.
        $ExecutablePath = $Path+$InputPath -replace '\s',''
        $ExecutableName = $OldAction.Split('\')[-1]
        
        # - Execute secondary script for TaskEvaluators
        . ''
        
        # - Translate new task variables to action arguments and store them in $NewAction.
        $NewActions += "$JobScript -ExecutablePath '$ExecutablePath' -ExecutableName '$ExecutableName' -LogfileName '$LogfileName' -LogfileEvaluator '$LogfileEvaluator' -Task_Path '$Task_Path' -NewTaskName '$NewTaskName' -Sleep_Timer '$Sleep_Timer' -MultiPass '$MultiPass' -IsLastAction '$IsLastAction'"
    }
    
    # - If there is more than one action available in this task, modify the string parameter
    #   'IsLastAction' of the last entry in #NewActions to '1', resulting in the last action
    #   (job) to restart the parent task. This ensures that the task itself keeps running automatically.
    If(($NewActions.Count) -gt 1){
        [int]$IsLastAction = 1
        ($NewActions[-1]) = ($NewActions[-1]) -replace "-IsLastAction '0'","-IsLastAction '$IsLastAction'"
    }

    # - Define target folder in Task Scheduler.
    $TargetTaskFolder = ($JobTask.TaskPath).TrimStart('\') -replace '.$'

    # - Creating corresponding folders in Task Scheduler.
    $Folders = $rootFolder.GetFolders(0) | Select-Object Name -ExpandProperty Name | Sort-Object Name
    If(!$Folders){
        $rootFolder.CreateFolder($TargetTaskFolder) | Out-Null
    }
    Foreach($Folder in $Folders){
        $NewTaskFolder = $scheduleObject.GetFolder($Folder)
        If($NewTaskFolder.Path -ne $TargetTaskFolder){
            $rootFolder.CreateFolder($TargetTaskFolder) | Out-Null
        }
    }
    
    # - Create a new, empty array to store new task actions.
    $NewTaskActions = @()

    # - This Foreach-loop will store the converted task actions.
    Foreach($NewAction in $NewActions){
        # - Create new task action.
        $NewTaskActions += New-ScheduledTaskAction -Execute ($PShell_Path+"powershell.exe  ") -Argument $NewAction
    }
    
    # - Specify trigger for starting task (at startup).
    $NewTaskTrigger = New-ScheduledTaskTrigger -AtStartup

    # - Disable Execution Time limit for this task.
    $NewTaskSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit '0' -MultipleInstances Queue -Compatibility Win8 -StartWhenAvailable
    
    # - Create new task.
    Register-ScheduledTask -TaskName $NewTaskName -TaskPath (($rootFolder.Path)+'\'+$TargetTaskFolder) -Action $NewTaskActions `
    -RunLevel Highest -Trigger $NewTaskTrigger -User $NewTaskUser -Password $Password -Settings $NewTaskSettings -Force

    # - Enabling and running the new task
    #Start-ScheduledTask -TaskName (($JobTask.TaskName)+" - PoSh")
}