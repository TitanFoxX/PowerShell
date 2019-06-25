# - Setting "culture" to "nl-NL" for correct date and time output.
# - Storing the current date in Dutch format into $Date variable.
# - Disabling any error notifications.
# - Defining a $Path variable.

Set-Culture -CultureInfo nl-NL
$Date = Get-Date -Format dd-MM-yyyy
$ErrorActionPreference = 'SilentlyContinue'
$Path = "<path for storing output file>"

# - Checking if "Test-Path" exists. If this is not the case, create this folder.

If(Test-Path -Path "$Path\$Date"){}
Else{New-Item -ItemType Directory -Path "$Path\$Date"}

# - Store current date into $Date variable.
# - The file path is defined in this variable.

$FilePath = "$Path\$Date\<file name without extension>"

# - Create a new Microsoft Excel file with some pre-formatting.
# - $Excel.Visual is temporarily set to $False to hide Excel window.
# - Disabling any visual alerts in Excel.
# - Create the first worksheet (tab).
# - Activate this worksheet.

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
$Excel.Workbooks.Add()
$Excel.DisplayAlerts = $False
$Workbook = $Excel.Workbooks.Item(1)
$Worksheet = $Workbook.Worksheets.Item("Sheet1") # use "Blad1" when using Dutch Excel version.
$Worksheet.Activate() | Out-Null

# - Changing the name of the new worksheet into '<your name>'.
# - Splitting the new worksheet below the first row and freezing first row for improved scrolling.

$Worksheet.Name = '<your name>'
$Worksheet.Application.ActiveWindow.SplitRow = 1
$Worksheet.Application.ActiveWindow.FreezePanes = $True

# - Naming the 'header'-cells 1,1 till 1,2 ('A1' till 'B1').

$Worksheet.Cells.Item(1,1) = 'ServerName'
$Worksheet.Cells.Item(1,2) = 'Logged in Users'

# - Define the 'header'-cells as range for graphical formatting.
# - 'Font.Bold = $True' will change font weight to 'bold'.
# - 'Interior.ColorIndex = 15' will change cell background to grey.
# - 'Font.ColorIndex = 1' will change the font color to black.
# - 'HorizontalAlignment = xlCenter' will center cells horizontally.
# - 'VerticalAlignment = xlVAlignCenter' will center cells vertically.

$Range = $Excel.Range('A1','B1')
$Range.Font.Bold = $True
$Range.Interior.ColorIndex = 15
$Range.Font.ColorIndex = 1
$Range.HorizontalAlignment = -4108
$Range.VerticalAlignment = xlVAlignCenter

# - Create a new array and fill it with the results of your input.
# - Sort the Array on the Name parameter.

$Servers = @()
$Servers = Get-ADComputer -Filter {(Enabled -eq $True)} -Properties Name | Where{$_.Name -like '*<your servers>*'} | Select-Object Name -ExpandProperty Name
[Array]::Sort($Servers)

# - Create a variable to use as counter within the Excel sheet.

$Var = 2

# - Loop through each server found with the Get-ADComputer cmdlet.

Foreach($Server in $Servers){
    
    # - Create new array $Results.
    # - Use Get-WmiObject to query the process Explorer.exe on target server and store logged in users into $Results.
    
    $Results = @()
    $Results += Get-WmiObject -class win32_process -Filter "name = 'Explorer.exe'" -ComputerName $Server -EA "Stop" | % {$_.GetOwner().User}
    
    # - Loop through $Results and store both the server name and users into Excel.
    
    Foreach($Result in $Results){
        $Excel.cells.item($Var,1) = $Server
        $Excel.cells.item($Var,2) = $Result
        $Var++
    }
}

# - Convert the rows containing date into a table.

$ListObject = $Excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.Name = 'TableData'
$ListObject.TableStyle = 'TableStyleMedium9'

# - Save the Excel file to specified Filepath.

$Workbook.SaveAs($FilePath)
$Excel.Quit()