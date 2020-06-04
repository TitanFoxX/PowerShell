# Setting "culture" to "nl-NL" for correct date and time output.
# Storing the current date in Dutch format into $Date variable.
# Disabling any error notifications.
# Defining a $Path variable (thanks Rens!)

Set-Culture -CultureInfo nl-NL
Get-Culture | Select -Expand DateTimeFormat | Select ShortDatePattern | Out-Null
$Date = Get-Date -Format "yyyy-MM-dd"
$ErrorActionPreference= 'SilentlyContinue'
$Path = ''

# Checking if "Test-Path" exists. If this is not the case, create this folder.

If(Test-Path -Path "$Path\$Date"){
}
Else{
    New-Item -ItemType Directory -Path "$Path\$Date"
}

# The file path is defined in this variable. The output file will be stored here.

$FilePath = "$Path\$Date\"+$Date+"__Zendesk_Users"

# Create a new Microsoft Excel file with some pre-formatting.
# $Excel.Visual is temporarily set to $True to display progress in filling in data.
# Disabling any visual alerts in Excel
# Splitting the new worksheet below the first row and freezes it for improved scrolling.
# Create the first worksheet (tab)
# Activate this worksheet

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel.Workbooks.Add()
$Excel.DisplayAlerts = $False
$Workbook = $Excel.Workbooks.Item(1)
$WorksheetZDACNTS = $Workbook.Worksheets.Item("Sheet1")
$WorksheetZDACNTS.Activate() | Out-Null

# Naming the 'header'-cells 1,1 till 1,3 ('A1' till 'C1').

$WorksheetZDACNTS.Cells.Item(1,1) = 'Username'
$WorksheetZDACNTS.Cells.Item(1,2) = 'Access Role'
$WorksheetZDACNTS.Cells.Item(1,3) = 'Account Type'

# Define the 'header'-cells as range for graphical formatting.
# 'Font.Bold = $True' will change font weight to 'bold'.
# 'Interior.ColorIndex = 15' will change cell background to grey.
# 'Font.ColorIndex = 1' will change the font color to black.
# 'HorizontalAlignment = xlCenter' will center cells horizontally.
# 'VerticalAlignment = xlVAlignCenter' will center cells vertically.

$ZDACNTS_Range = $Excel.Range('A1','C1')
$ZDACNTS_Range.Font.Bold = $True
$ZDACNTS_Range.Interior.ColorIndex = 15
$ZDACNTS_Range.Font.ColorIndex = 1
$ZDACNTS_Range.HorizontalAlignment = -4108
$ZDACNTS_Range.VerticalAlignment = xlVAlignCenter

# Set TLS protocol to v1.2 (required for Zendesk)
# Set starting page number to 1
# Set the destionation URL for API query.
# Set $ZDSK_VAR to 2, so that Excel will start input below header cells

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ZD_PageNumber = 1
$ZDBaseURL = ""
$ZD_QueryURL = $ZDBaseURL+$ZD_PageNumber
$ZDSK_Var = 2

# Credentials for accessing the REST API

$ZDCreds = Get-Content '' | ConvertTo-SecureString
$ZDCreds = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($ZDCreds))
$ZDHeaders = @{Authorization = "Basic $ZDCreds"}

# First query to REST API for determining the amount of users.
# The '[math]::ceiling'-line will devide the amount of users by 100 and then round it upwards to the next whole digit.
# Create the empty array $ConvertedJSON after the initial query has been completed.

# Example: (451 users / 100) = 4,51. Rounded upwards = 5
# Reason: Query contains 100 users per page. 451 users require 4 pages of 100 users and 1 page of 51 users = 5 pages.

$ResponseData = Invoke-WebRequest -Uri $ZD_QueryURL -Method Get -Headers $ZDHeaders -UseBasicParsing | Select-Object Content -ExpandProperty Content
$ConvertedJSON = $ResponseData | ConvertFrom-Json
$Total_ZDPages = [math]::ceiling($ConvertedJSON.count / 100)
$ConvertedJSON = @()

# This FOR-statement will start with $i = 1, which is equal to 'page=1' - the first
# page received from RESTAPI. The API query within this statement will run and count 
# 1 ($i++) towards the $i variable once done. This process will continue until $i 
# reaches the same value as the amount of API pages calculated from the initial query. 
# Every cycle of this loop will append the received data to $ConvertedJSON.

For($i = 1; $i -le $Total_ZDPages; $i++){
    $QueryURL = $ZDBaseURL+$ZD_PageNumber
    $ResponseData = Invoke-WebRequest -Uri $QueryURL -Method Get -Headers $ZDHeaders -UseBasicParsing | Select-Object Content -ExpandProperty Content
    $ConvertedJSON += $ResponseData | ConvertFrom-Json
    $ZD_PageNumber++
}

# Once the API query has cycled through all available pages and merged the received
# data into $ConvertedJSON, this Foreach-loop will run through all users reported within
# the JSON and will output the username and access role towards Excel.

Foreach($ZendeskUser in $ConvertedJSON.users){
    If($ZendeskUser.name){
        $Excel.cells.item($ZDSK_Var,1) = $ZendeskUser.name
        $Excel.cells.item($ZDSK_Var,2) = $ZendeskUser.role
        If($ZendeskUser.email -like ''){
            $Excel.cells.item($ZDSK_Var,3) = ''
        }
        Else{
            $Excel.cells.item($ZDSK_Var,3) = ''
        }        
        $ZDSK_Var++
    }
}

# Convert the rows containing date into a table.

$ListObject = $Excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.Name = 'TableData'
$ListObject.TableStyle = 'TableStyleMedium9'

# Define all cells in this worksheet, sort column 'A' (minus the headers) alphabetically 
# and resize cells so that content fits.

$ZDACNTS_Range = $WorksheetZDACNTS.UsedRange
$ZDACNTS_Range2 = $Excel.Range('A2')
[void]$ZDACNTS_Range.Sort($ZDACNTS_Range2,1,$null,$null,1,$null,1,1)
$ZDACNTS_Range.EntireColumn.AutoFit() | Out-Null

# Save the Excel file to specified Filepath.

$Workbook.SaveAs($FilePath)
$Excel.Quit()