# - Specify servers to be queried.

$DBServers = @()
$DBServers = ""

# - Specify recipients.

$ToRecipients = ""
$CCRecipients = ""

# - Create empty array, used for storing file results.

$Files = @()

Foreach($DBServer in $DBServers){
    
    # - Query specified servers for logical drives that do not match either A, B or C-drives.
    
    $Drives = @()
    $Drives += Get-WmiObject Win32_LogicalDisk -ComputerName $DBServer -Filter DriveType=3 | Where{((($_.DeviceID -ne "A:") -and ($_.DeviceID -ne "B:") -and ($_.DeviceID -ne "C:")))} | Select-Object DeviceID -ExpandProperty DeviceID
    
    # - Query all found logical drives per server and search with specified parameters.
    # - Store all file results into the $Files-array.

    Foreach($Drive in $Drives){
        $RootDrive = $Drive -replace ':',''
        $Files += (Get-ChildItem -Path "\\$DBServer\$RootDrive$\" -Recurse) | Where-object {($_.Name -notmatch "") -and (($_.Name -like "") -or ($_.Name -like "**")) -and ($_.length -gt )}
    }
}

# - Create IF-statement that will only allow e-mails to be sent in case that files are found that match specified values.

If($Files){

    # - Create a new multi-array to store all results found in query.
    
    $FileArray = @()
    
    # - Loop through all results and transfer them to the multi-array.
    
    Foreach($File in $Files){
        $ServerName = ($File.FullName).Split('\')[2]
        $FileName = $File.FullName
        
        # - Convert the filesize (in bytes) to GB's, rounded with 2 decimals behind the comma.
        
        $FileSizeGB = [Math]::Round(($File.Length / 1024/1024/1024),2)
    
        # - Add converted values to the array.
    
        $FileObject = New-Object -TypeName PSObject
        $FileObject | Add-Member -Name 'ServerName' -MemberType NoteProperty -Value "$ServerName"; $FileObject | Add-Member -Name "FileName" -MemberType NoteProperty -Value "$FileName"
        $FileObject | Add-Member -Name 'FileSize (GB)' -MemberType NoteProperty -Value "$FileSizeGB"
        $FileArray += $FileObject
    }
    
    # - Convert $FileArray to HTML-table.
    
    $HTML_Table_DOCFILE = ([PSCustomobject]$FileArray | Where{$_.FileName -like ""} | ConvertTo-Html -Fragment -As Table) -replace '<table>','<table border="1">'
    $HTML_Table_MODLIB = ([PSCustomobject]$FileArray | Where{$_.FileName -like ""} | ConvertTo-Html -Fragment -As Table) -replace '<table>','<table border="1">'
    
    # - Set E-mail parameters.
    
    $NewMsg = new-object Net.Mail.MailMessage
    $SMTP = new-object Net.Mail.SmtpClient("")
    $NewMsg.IsBodyHTML = $True
    $NewMsg.From = ""
    $NewMsg.To.Add($ToRecipients)
    $NewMsg.CC.Add($CCRecipients)
    $NewMsg.Subject = ""
    
# - Define body and information to be sent out via E-mail.

$NewMsg.Body = @"
"@
$SMTP.Send($NewMsg)
}