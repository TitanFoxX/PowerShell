# - Keeping everything nice and quiet.

$ErrorActionPreference = 'SilentlyContinue'

# - Convert current date to NL format

$DateConv = Get-Date -Format 'dd-MM-yyyy'

# - Store computername from which this script is run into a variable.

$ComputerName = $env:COMPUTERNAME

# - Define file path to store CSV-results file.

$FilePath = '<fill in desired path>'

# - Test if defined file path exists.

$Result = Test-Path $FilePath'\'$ComputerName

# - If the file path does not exist yet, create it.

If($Result -eq $False){
    New-Item -ItemType Directory -Path $FilePath -Name $ComputerName -Force | Out-Null
}

# - Store all mapped logical disks into #Output.

$Output = Get-WmiObject Win32_MappedLogicalDisk | Select-Object PSComputerName,Name,Providername | FT -HideTableHeaders

# - Write content of $Output to CSV file in defined file path.

$Output | Out-File $FilePath'\'$ComputerName'\'$ComputerName'__'$DateConv.csv