# - Function for creating a new Windows Scheduled Task session.

Function Get-ScheduleService{
    New-Object -ComObject schedule.service
}

# - Function for connecting to the Windows Scheduled Task service.
# - If no $Path value is specified, the root folder ("\") will be selected.

Function New-TaskObject($Path){
    $taskObject = Get-ScheduleService
    $taskObject.Connect()
    If(-not $Path){
        $Path = “\”
    }
    $taskObject.GetFolder($Path)
}

# - Function for displaying all (sub-)folders in specified location.

Function Get-TaskFolder($Folder,[Switch]$Recurse){
    If($Recurse){
        $colFolders = $Folder.GetFolders(0)
        Foreach($i in $colFolders){
            $i.path
            $subFolder = (New-taskObject -path $i.path)
            Get-taskFolder -Folder $subFolder -Recurse
        }
    }
    Else{
        $Folder.GetFolders(0) | ForEach-Object {$_.path}
    }
}

# - Function for creating a new folder.

Function New-TaskFolder($Folder,$Path){
    $Folder.CreateFolder($Path)
}

# - Function for removing a folder.

Function Remove-TaskFolder($Folder,$Path){
    $Folder.DeleteFolder($Path,$null)
}

# - List folders in root of given '-Path'
Get-TaskFolder -Folder (New-taskObject -Path “<Path>”)

# - List folders and subfolders in given '-Path'
Get-TaskFolder -Folder (New-taskObject -Path “<Path>”) -Recurse

# - Create a new folder in given '-Path'
New-TaskFolder -Folder (New-taskObject) -Path “<Path>”

# - Remove an empty folder in given '-Path'
Remove-TaskFolder -Folder (New-taskObject) -Path “<Path>”