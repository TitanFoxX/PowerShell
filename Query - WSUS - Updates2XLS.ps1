# - Get all enabled computers in AD that run a Windows Server OS

$ADName = Get-ADComputer -Filter {Enabled -eq $True} -Properties * | Where-Object {$_.OperatingSystem -like "*Server*"} | Select-Object Name -ExpandProperty Name | Sort-Object Name

# - Get the current date, the date one month ago, set the time format and supress errors in the console.

$DateNow = Get-Date
$DateThen = (Get-Date).AddMonths(-1)
$DateConv = Get-Date -Format "dd-MM-yyyy"
$ErrorActionPreference= 'SilentlyContinue'

# - Specify File Path for output.

$FilePath = ""

# - Setup basic Excel workbook.

$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = $DateConv
$sheet.Cells.Item(1,1) = 'ServerName'
$sheet.Cells.Item(1,2) = 'HotFixID'
$sheet.Cells.Item(1,3) = 'Description'
$sheet.Cells.Item(1,4) = 'InstalledOn'
$sheet.Cells.Item(1,1).Font.Bold=$True
$sheet.Cells.Item(1,2).Font.Bold=$True
$sheet.Cells.Item(1,3).Font.Bold=$True
$sheet.Cells.Item(1,4).Font.Bold=$True
$sheet.Cells(1,1).HorizontalAlignment = -4108
$sheet.Cells(1,2).HorizontalAlignment = -4108
$sheet.Cells(1,3).HorizontalAlignment = -4108
$sheet.Cells(1,4).HorizontalAlignment = -4108
$sheet.Cells(1,1).VerticalAlignment = -4108
$sheet.Cells(1,2).VerticalAlignment = -4108
$sheet.Cells(1,3).VerticalAlignment = -4108
$sheet.Cells(1,4).VerticalAlignment = -4108

$i = 2

# - Loop through servers and output the updates that were installed within the last month to Excel.

Foreach ($Name in $ADName.Name){
    $HotFix = Get-CimInstance -class Win32_QuickFixEngineering -ComputerName $Name | Where {$_.InstalledOn -gt "$DateThen" -AND $_.InstalledOn -lt "$DateNow"}
    Foreach ($Fixes in $HotFix){
        $excel.cells.item($i,1) = $Fixes.PSComputerName
        $excel.cells.item($i,2) = $Fixes.HotFixID
        $excel.cells.item($i,3) = $Fixes.Description
        $excel.cells.item($i,4) = $Fixes.InstalledOn
        $i++
    }
    $i++
}

# - Resize cells and export Excel-file.

$usedRange = $sheet.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null
$workbook.SaveAs($FilePath)
$excel.Quit()