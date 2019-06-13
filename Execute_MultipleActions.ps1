# - Keeping everything nice and quiet ...

$ErrorActionPreference = 'SilentlyContinue'

# - These variables contain snippets of the string required by the Invoke-RestMethod at the 
#   end of the script. This is purely an esthetical functionality to keep the script organized.

$WH_URLSnip1 = '<Part 1 of the Office365 URi>'
$WH_URLSnip2 = '<Part 1 of the Office365 URi>'

# - Basic unique parameters.

$NewTaskPath = "" # - Path within Microsoft Task Scheduler, including a '\' on the end.
$NewTaskName = "" # - Name of the task name in Microsoft Task Scheduler.
$SourceSRV = "" # - Name of the originated server. Will be used in the UNC path.
$MDSPath = "" # - UNC path including the name of the originating server.

# - Create an array to store job parameters.

$JobsArray = @()

# - Add all required apps with their unique parameters to the array.

# - App # 1
$JobsObject = New-Object -TypeName PSObject
$JobsObject | Add-Member -Name 'ExecutablePath' -MemberType NoteProperty -Value "<Path to executable>"; $JobsObject | Add-Member -Name 'ExecutableName' -MemberType NoteProperty -Value '<Name of executable>'
$JobsObject | Add-Member -Name 'LogfileName' -MemberType NoteProperty -Value '<Name of logfile>'; $JobsObject | Add-Member -Name 'LogfileEvaluator' -MemberType NoteProperty -Value '*<Evaluator>*'
$JobsArray += $JobsObject

# - App # 2
$JobsObject = New-Object -TypeName PSObject
$JobsObject | Add-Member -Name 'ExecutablePath' -MemberType NoteProperty -Value "<Path to executable>"; $JobsObject | Add-Member -Name 'ExecutableName' -MemberType NoteProperty -Value '<Name of executable>'
$JobsObject | Add-Member -Name 'LogfileName' -MemberType NoteProperty -Value '<Name of logfile>'; $JobsObject | Add-Member -Name 'LogfileEvaluator' -MemberType NoteProperty -Value '*<Evaluator>*'
$JobsArray += $JobsObject

# - App # 3
$JobsObject = New-Object -TypeName PSObject
$JobsObject | Add-Member -Name 'ExecutablePath' -MemberType NoteProperty -Value "<Path to executable>"; $JobsObject | Add-Member -Name 'ExecutableName' -MemberType NoteProperty -Value '<Name of executable>'
$JobsObject | Add-Member -Name 'LogfileName' -MemberType NoteProperty -Value '<Name of logfile>'; $JobsObject | Add-Member -Name 'LogfileEvaluator' -MemberType NoteProperty -Value '*<Evaluator>*'
$JobsArray += $JobsObject

# - Check if all required mapped drives are available. If not, mount any missing drive letters.

If(-not ($pathExists)){(New-Object -ComObject WScript.Network).MapNetworkDrive("<Driveletter>","UNC-Path")}

# - Disable check UNC Zone (supresses security dialog when executing app executable).

$env:SEE_MASK_NOZONECHECKS = 1

# - Create an infinite While-loop.

While($True){

    # - Loop through the $JobsArray and execute each job with it's own parameters.

    Foreach($UniqueJob in $JobsArray){

        # - Check if executable exists. If not, display error message in Teams via PRTG and end the script.

        If((Test-Path -Path (($UniqueJob).ExecutablePath+"\"+(($UniqueJob).ExecutableName))) -eq $False){
            
            # - Convert Job-parameters to variables.

            $ExecutableName = (($UniqueJob).ExecutableName)
            $ExecutablePath = (($UniqueJob).ExecutablePath)

            # - Define the message shown within Microsoft Teams.
            
            $Message = "$ExecutableName on server $env:COMPUTERNAME can not be found on specified location. "
            $Message = $Message + "Either the task parameters are incorrect or the folder / job executable "
            $Message = $Message + "has been moved, renamed or deleted. The scheduled task will now end."
            $Message = $Message | ConvertTo-Json

            # - Send the defined message through PRTG.

            Invoke-RestMethod -Method Post -ContentType "Application/Json" -Body "{'text': '$Message'}" -Uri https://$WH_URLSnip1/$WH_URLSnip2 | Out-Null

            # - Stop the current task, then close this script process.

            $NewTaskLocation = ($NewTaskPath)+($NewTaskName)

            schtasks.exe /End /tn $NewTaskLocation
            Stop-ScheduledTask -TaskPath -TaskName
            Exit
        }
        
        # - Create DO-loop to continue running the job until it runs succesfully.

        DO{

            # - Run the Export executable and wait till it finishes.
            
            Start-Process -WindowStyle Hidden -FilePath (($UniqueJob).ExecutablePath+"\"+(($UniqueJob).ExecutableName)) -Wait -ErrorAction 'SilentlyContinue'
            
            # - Read the last two lines of the executable's LOG-file
            
            $Status = Get-Content (($UniqueJob).ExecutablePath+"\"+(($UniqueJob).LogfileName)) -Tail 2 | Out-String

            # - If the last two lines do not match the predefined 'success' value, write out an error 
            #   message to Technical Support via Microsoft Teams and restart the scheduled task.
            
            If($Status -notlike (($UniqueJob).LogfileEvaluator)){
                
                # - Convert Job-parameters to variables.

                $ExecutableName = (($UniqueJob).ExecutableName)
                $ExecutablePath = (($UniqueJob).ExecutablePath)

                # - Define the message that will be sent through PRTG.
                
                $Message = "$ExecutableName in $ExecutablePath on server $env:COMPUTERNAME has crashed. "
                $Message = $Message + "Please check and resolve. The task will automatically be restarted."
                $Message = $Message | ConvertTo-Json

                # - Send error message to Microsoft Teams.

                Invoke-RestMethod -Method Post -ContentType "Application/Json" -Body "{'text': '$Message'}" -Uri https://$WH_URLSnip1/$WH_URLSnip2 | Out-Null
            }
        }
        Until($Status -notlike (($UniqueJob).LogfileEvaluator) -eq $False)
    }
    
    # - Will pause the script for 5 minutes after all jobs have been ran once.
    # - After time period has passed, the loop wil run again. This can be 
    #   modified to any desired value.

    Start-Sleep -Seconds 300
}