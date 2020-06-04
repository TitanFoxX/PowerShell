# - Load required Snapins.

Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100

# - Define query for getting 

Function Test-SQLTableExists{ 
    Param ($Instance,$Database,$TableName)
    $Return = $SQL = $DataTable = $Null
    $SQL = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'$TableName'"
    $DataTable = Invoke-Sqlcmd -ServerInstance $Instance -Database $Database -Query $SQL
    If($DataTable){$Return = $True}
    Else{$Return = $False}
    $Return
}

# - Execute Query.

Test-SQLTableExists -Instance "" -Database "" -TableName TempTableSplunk