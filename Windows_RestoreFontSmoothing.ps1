# - Edit settings for HKEY_Current_User (HKCU).
Function RegRestore{
    $RegLocation = "HKCU:"
    'Setting RegLocation to HKCU ...'
    RegSetUser;RegChange_NTUSER
}

# - Load default user hive.

Function LoadDefaultHive{
	'Loading default user hive ...'
    $Default_Hive = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' Default).Default
    Reg load "$RegLocation" $Default_Hive'\NTUSER.DAT'
    'Default user hive succesfully loaded ...'
}

# - Load preferred user hive.

Function LoadUserHives{
	'Loading user hives ...'
    $UserKeys = @()
    $RegLocation = 'HKLM\AllProfile'
    $ProfilePath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $Profile_Keys = Get-ChildItem -Path $ProfilePath -Recurse | Select-Object Name -ExpandProperty Name
    $Profile_Keys = $Profile_Keys -replace 'HKEY_LOCAL_MACHINE','HKLM:'
    
    Foreach($Profile_Key in $Profile_Keys){
        $UserKeys += Get-ItemProperty -Path $Profile_Key | Where{$_.ProfileImagePath -like '*<Enter Name>*'}
        Foreach($UserKey in $UserKeys){
            $Default_Hive = (Get-ItemProperty $UserKey.PSPath).ProfileImagePath
            Reg load "$RegLocation" $Default_Hive'\NTUSER.DAT'
            "$Default_Hive loaded successfully"
        }
    }
}

# - Unload default user hive.

Function UnloadDefaultHive{
    'Unloading default user hive ...'
    [gc]::collect()
    $RegLocation = 'HKLM\AllProfile'
    Reg unload "$RegLocation"
    'Default user hive succesfully unloaded ...'
}

# - Edit settings for new users (NTUSER.DAT).

Function RegChange_NTUSER{
    $RegLocation = 'HKLM\AllProfile'
	'Setting RegLocation to default NTUSER.DAT ...'
    LoadDefaultHive
    $RegLocation = 'HKLM:\AllProfile'
    RegSetUser; UnloadDefaultHive; LoadUserHives    
    'Setting RegLocation back to default value (empty) ...'
    $RegLocation = $null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -PropertyType DWord -Name 'EnableLinkedConnections' -Value '1' -Force | Out-Null
}

# - Actual registry keys that will be modified.

Function RegSetUser{
    $Windows_Current_Version = "$RegLocation" + "\Software\Microsoft\Windows\CurrentVersion"
    New-ItemProperty -Path $RegLocation'\Control Panel\Desktop' -PropertyType String -Name 'FontSmoothing' -Value '2' -Force | Out-Null
    New-ItemProperty -Path $RegLocation'\Control Panel\Desktop' -PropertyType DWord -Name 'FontSmoothingType' -Value '2' -Force | Out-Null
    New-ItemProperty -Path $Windows_Current_Version'\Policies\System' -PropertyType String -Name 'Wallpaper' -Force | Out-Null
}

# - Start with first function in the script and continue next steps automatically.

RegRestore