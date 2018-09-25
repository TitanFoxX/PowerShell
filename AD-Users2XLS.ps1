# Setting 'Culture' to Dutch layout for date
Set-Culture -CultureInfo nl-NL
$Date = Get-Date -Format "dd-MM-yyyy"

# Creating a new empty array for $AD_User - practical if you re-run this script in the same PS session
# After this is done, retrieve all users from the Active Directory and sort on LastLogonDate
$AD_User = @()
$AD_User += Get-ADUser -Filter * -Property * | Sort LastLogonDate -Descending

# Please fill out your own file path between the quotes
# Script will continue without notification in case of a error
$FilePath = "<your file path comes here>"
$ErrorActionPreference= 'silentlycontinue'

# Create a new Microsoft Excel worksheet with some pre-formatted headers
$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$sheet= $workbook.Worksheets.Item(1)
$sheet.Name = $Date
$sheet.Cells.Item(1,1) = 'AD-Name'
$sheet.Cells.Item(1,2) = 'Username'
$sheet.Cells.Item(1,3) = 'Last Logon (DD-MM-YY)'
$sheet.Cells.Item(1,1).Font.Bold=$True
$sheet.Cells.Item(1,2).Font.Bold=$True
$sheet.Cells.Item(1,3).Font.Bold=$True
$sheet.Cells(1,1).HorizontalAlignment = -4108
$sheet.Cells(1,2).HorizontalAlignment = -4108
$sheet.Cells(1,3).HorizontalAlignment = -4108
$sheet.Cells(1,1).VerticalAlignment = -4108
$sheet.Cells(1,2).VerticalAlignment = -4108
$sheet.Cells(1,3).VerticalAlignment = -4108

# Started a variable with value 2, so that the first output will start on
# row 2, column 1 ($i,1). When the first three cells have been filled out,
# 1 is added to $i ($i++) resulting in $i = 3. Next output will start on
# row 3, column 1. This repeats itself till the Foreach-loop is complete.
$i = 2
Foreach ($Users in $AD_User){
    $excel.cells.item($i,1) = $Users.SamAccountName
    $excel.cells.item($i,2) = $Users.Name
    $excel.cells.item($i,3) = $Users.LastLogonDate
    $i++
}

# Format column to change width so that all content is displayed.
# After this, the workbook will be saved and Excel is closed.
$usedRange = $sheet.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null
$workbook.SaveAs($FilePath)
$excel.Quit()