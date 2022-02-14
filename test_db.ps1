#test query
$query = 'SELECT * FROM Master'

#set up database connection
$connectString = 'Server=<SERVERNAME>;Database=<DATABASENAME>;User Id=<USERNAME>;Password=<PASSWORD>;'
$global:SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$global:SqlConnection.ConnectionString = $connectionString
#run query
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $query
$SqlCmd.CommandTimeout = 60
$SqlCmd.Connection = $global:SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet)
$DataTable = $DataSet.Tables[0]

$DataTable
	