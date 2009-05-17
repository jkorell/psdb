# psdb v0.01
# Copyright © 2009 Jorge Matos
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

function Invoke-DB 
{
<#
.Synopsis
    Execute a database command
.Description
    This function will execute a SQL command or a SQL Stored procedure against a database.
	It uses ADO.NET provider factories to allow a query to run against different databases.
.Parameter Sql
    The SQL block you want to execute or the name of a stored procedure
.Parameter SPROC
    Stored procedure name to execute
.Parameter ExecuteType 
    Determines how the command is executed.
	if "Query" then return an array of ADO.NET DataRow objects 
	if "NonQuery" then return the number of rows affected
	if "Scalar" then return a scalar value (string, int, etc...)
	if "Reader" then return an IDataReader (uses CommandBehavior.CloseConnection)
	"Query is the default
.Parameter Parameters 
	An array of DbParameter objects.  
	You can create a DbParameter object using either: Create-InParameter, Create-OutParameter, Create-ReturnParameter or Create-Parameter
.Parameter Connectionstring
	The connectionString you want to use.
	Default="data source=.;initial catalog=Northwind; Integrated Security=SSPI"
.Parameter Provider 
    The ADO.NET provider you want to use.
	Possible values are:
		System.Data.Odbc
		System.Data.OleDb
		System.Data.OracleClient
		System.Data.SqlClient
		Microsoft.SqlServerCe.Client
		System.Data.SqlServerCe
		System.Data.SqlServerCe.3.5
.Parameter CommandTimeout
	The number of seconds to allow a command to run. Default = 600
.Example	
$rows = Invoke-DB -Sql "SELECT TOP 10 * FROM Orders"

Return an array of DataRow objects using default values:
		ConnectionString = "data source=.;initial catalog=Northwind; Integrated Security=SSPI"
	    ExecuteType = "Query"
		Parameters = @()
		Provider = System.Data.SqlClient
		CommandTimeout = 600
.Example 
$sql = "SELECT TOP 10 * FROM Orders"
$rows = Invoke-DB -Sql $sql -Connectionstring "data source=.;initial catalog=Northwind; uid=test; pwd=test"

Return an array of DataRow objects with connectionstring 	
.Example
$sql = "UPDATE Orders SET EmployeeID = 6 WHERE OrderID = 10248"
$rowsAffected = Invoke-DB -Sql $sql -ExecuteType "NonQuery"   

Performs an update 
"NonQuery" is used to return the number of rows affected
.Example
$Parameters = @(
				(Create-InParameter -Name "@Country" -Value "USA"),
           		(Create-InParameter -Name "@Freight" -Value 100)
			    ) 
$sql = "SELECT * FROM Orders WHERE  (ShipCountry = @Country) AND (Freight > @Freight)" 
$rows = Invoke-DB -Sql $sql -Parameters $Parameters 	

Using Parameters
.Example
$rows = Invoke-DB -SPROC "Ten Most Expensive Products"

Calling a stored procedure 
.Example
$rows = Invoke-DB -SPROC "Ten Most Expensive Products" -Connectionstring "data source=.;initial catalog=Northwind; uid=test; pwd=test"

Calling a stored procedure with connectionstring
.Example
$Parameters = @( Create-InParameter -Name "@CustomerID" -Value "ALFKI" ) 
$rows = Invoke-DB -SPROC "CustOrderHist" -Parameters $Parameters

Calling a stored procedure with a parameter
.Example
$result = Invoke-DB -Sql "SELECT COUNT(*) FROM Orders" -ExecuteType "Scalar"

Returning a scalar value
.Example
$reader = Invoke-DB -Sql "SELECT TOP 2 * FROM Orders" -ExecuteType "Reader"
if ($reader.HasRows) {
	while($reader.Read()) {
		"{0} {1}" -f $reader[0],$reader[1]
	}
}

$reader.Close()

Return an IDataReader and access column values via index 
.Example
$reader = Invoke-DB -Sql "SELECT TOP 2 * FROM Orders" -ExecuteType "Reader"
if ($reader.HasRows) {
     $columns = $reader.GetSchemaTable() | % { $_.ColumnName }
     while($reader.Read()) {
        $columns | % { "$_ = " + $reader[$_] }       
     }     
}
$reader.Close()

Return an IDataReader and access column values via column name
.Example
$accessfile = Resolve-Path .\Database1.mdb
$provider = "System.Data.OleDb"
$connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessfile"
$sql = "SELECT * FROM users"
$rows = invoke-db -Sql $sql -Connectionstring $connectionString -Provider $provider

Query an Access Database
.Example
$excelfile = Resolve-Path .\book1.xls
$provider = "System.Data.OleDb"
$connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$excelfile;Extended Properties=Excel 8.0"
$sql = "SELECT * FROM [Sheet1$]"
$rows = invoke-db -Sql $sql -Connectionstring $connectionString -Provider $provider

Query an Excel file
.ReturnValue
    The return value varies based on the value of ExecuteType

	Query    -> Array of DataRow objects or an empty array if no results are returned   
	NonQuery -> Return number of rows affected (int)
	Scalar   -> Scalar Value (String, Int, DateTime, etc...)
	Reader   -> IDataReader (uses CommandBehavior.CloseConnection)	
.Link
    Create-Parameter
	Create-InParameter
	Create-OutParameter
	Create-ReturnParameter
.Notes
 NAME:      Invoke-DB
 AUTHOR:    Jorge Matos
 LASTEDIT:  05/13/2009 
#Requires -Version 2.0
#>

[CmdletBinding(
    SupportsShouldProcess=$False,
    SupportsTransactions=$False, 
    ConfirmImpact="None",
    DefaultParameterSetName="sql")]
param(
	[Parameter(Position=0,Mandatory=1,ParameterSetName="sql")]
	[Alias("q")]
	[string]$Sql,

	[Parameter(Position=0,Mandatory=1,ParameterSetName="sproc")]
	[Alias("sp")]
	[string]$SPROC,

	[Parameter(Position=1)]
	[Alias("et")]
	[string]$ExecuteType="Query",

	[Parameter(Position=2)]
	[Alias("p")]
	[System.Data.Common.DbParameter[]]$Parameters=@(),

	[Parameter(Position=3)]
	[Alias("cs")]
	[string]$Connectionstring="data source=.;initial catalog=Northwind; Integrated Security=SSPI",

	[Parameter(Position=4)]
	[string]$Provider="System.Data.SqlClient",

	[Parameter(Position=5)]
	[Alias("ct")]
	[int]$CommandTimeout=600
) #param

	Process 
	{
		$validExecuteTypes = "Query","NonQuery","Scalar","Reader"
		if (!($validExecuteTypes -contains $ExecuteType)) 
		{
			throw "Unknown ExecuteType: [$ExecuteType]. ExecuteType must be one of the following: $($validExecuteTypes -join "", "")"
		}
		
		$validProviders = "System.Data.Odbc", "System.Data.OleDb", "System.Data.OracleClient", 
		"System.Data.SqlClient", "Microsoft.SqlServerCe.Client", "System.Data.SqlServerCe",  
		"System.Data.SqlServerCe.3.5"
		
		if (!($validProviders -contains $Provider)) 
		{
			throw "Unknown Provider: [$Provider]. Provider must be one of the following: $($validProviders -join "", "")"
		}
		
		$providerFactory = [System.Data.Common.DBProviderFactories]::GetFactory($Provider) 
		$connection = $providerFactory.CreateConnection()
		$connection.ConnectionString = $connectionString
		
		$command = $providerFactory.CreateCommand()
		
		if ($PsCmdlet.ParameterSetName -eq "sql") 
		{
			$command.CommandText =  $Sql 
		} 
		else 
		{
			$command.CommandText =  $SPROC
			$command.CommandType = "StoredProcedure"
		}
		
		$command.CommandTimeOut = $Timeout
		$command.connection = $connection
		
		if ($Parameters.Length -gt 0) 
		{
			[void]$command.Parameters.AddRange($Parameters)
		}        
		
		try	
		{
			$connection.Open()		
			
			if ($ExecuteType -eq "NonQuery") 
			{			
				return $command.ExecuteNonQuery()
			}
			elseif ($ExecuteType -eq "Scalar") 
			{
				return $command.ExecuteScalar()
			}
			elseif ($ExecuteType -eq "Reader") 
			{
				$reader = $command.ExecuteReader("CloseConnection")
				return ,$reader
			}
			else 
			{
				$adapter = $ProviderFactory.CreateDataAdapter()
				$adapter.SelectCommand = $command
				$dataset = New-Object System.Data.DataSet
				[void]$adapter.Fill($dataSet)

				$rows = $dataSet.Tables | Select-Object -Expand Rows 
				
				if ($rows -eq $null) 
				{
					return ,@()                
				}
				else 
				{
					return ,@($rows)
				}      
			}  
		}
		finally 
		{
			if ($ExecuteType -ne "Reader") 
			{
				if ($connection -ne $null) 
				{
					if ($connection.state -eq [System.Data.ConnectionState]::Open) 
					{				
						$connection.Close()
					}
				}
			}
		} 
	} #Process
} #Invoke-DB

function Create-InParameter {
<#
.Synopsis
    Creates an input parameter that can be used with Invoke-DB
.Description
    Create an input parameter object of type System.Data.DBParameter that can be used to 
	pass parameters to a stored procedure call or a sql statement

.Parameter Name 
	The name of the input parameter
.Parameter DbType 
	The DbType of the parameter. Default="String"
	Possible values: Boolean, Int32, Double, Decimal, DateTime
	See "DbType Enumeration" in the MSDN documentation
.Parameter Value 
	The value that the parameter will contain
.Parameter Provider
	The ADO.NET provider you want to use. (Default="System.Data.SqlClient")
	Possible values are:
		System.Data.Odbc
		System.Data.OleDb
		System.Data.OracleClient
		System.Data.SqlClient
		Microsoft.SqlServerCe.Client
		System.Data.SqlServerCe
		System.Data.SqlServerCe.3.5 
.Example
    Create-InParameter -Name "@CustomerID" -Value "ALFKI"
.ReturnValue
    A DbParameter object
.Link
    Invoke-DB
	Create-Parameter
	Create-OutParameter
	Create-ReturnParameter
.Notes
 NAME:      Create-InParameter
 AUTHOR:    Jorge Matos
 LASTEDIT:  05/13/2009
#Requires -Version 2.0
#>

[CmdletBinding(
    SupportsShouldProcess=$False,
    SupportsTransactions=$False, 
    ConfirmImpact="None",
    DefaultParameterSetName="")]
param(
	[Parameter(Position=0,Mandatory=1)]
	[string]$Name,

	[Parameter(Position=1)]
	[System.Data.DbType]$DBType="String",

	[Parameter(Position=2,Mandatory=1)]
	[object]$Value,

	[Parameter(Position=3)]
	[string]$Provider="System.Data.SqlClient"
)

	Process	
	{
		return Create-Parameter -Name $Name -DbType $DbType -Value $Value -Direction "Input" -Provider $Provider
	}#Process
} # Create-InParameter

function Create-OutParameter 
{
<#
.Synopsis
    Creates an output parameter that can be used with Invoke-DB
.Description
    Create an output parameter object of type System.Data.DBParameter that can be used to 
	retrieve output parameters from a stored procedure call

.Parameter Name 
	The name of the output parameter
.Parameter DBType 
	The DbType of the parameter. Default="String"
	Possible values: Boolean, Int32, Double, Decimal, DateTime
	See "DbType Enumeration" in the MSDN documentation
.Parameter Size 
	The storage size in bytes of the parameter.
	For SQL Server 2005:
		char or varchar = 1 byte per character
		nchar or nvarchar = 2 bytes per character
		int = 4 bytes
		float = 8 bytes`
		datetime = 8 bytes
	For more info see "Data Types" in the SQL Server BOL	
.Parameter Provider
	The ADO.NET provider you want to use. (Default="System.Data.SqlClient")
	Possible values are:
		System.Data.Odbc
		System.Data.OleDb
		System.Data.OracleClient
		System.Data.SqlClient
		Microsoft.SqlServerCe.Client
		System.Data.SqlServerCe
		System.Data.SqlServerCe.3.5 
.Example
    Create-OutParameter -Name "@Email" -DbType "String" -Size 50
.ReturnValue
    A DbParameter object
.Link
    Invoke-DB
	Create-Parameter
	Create-InParameter
	Create-ReturnParameter
.Notes
 NAME:      Create-OutParameter
 AUTHOR:    Jorge Matos
 LASTEDIT:  05/13/2009
#Requires -Version 2.0
#>

[CmdletBinding(
    SupportsShouldProcess=$False,
    SupportsTransactions=$False, 
    ConfirmImpact="None",
    DefaultParameterSetName="")]
param(
	[Parameter(Position=0,Mandatory=1)]
	[string]$Name,

	[Parameter(Position=1)]
	[System.Data.DbType]$DBType = "String",

	[Parameter(Position=2,Mandatory=1)]
	[int]$Size,

	[Parameter(Position=3)]
	[string]$Provider = "System.Data.SqlClient"
)

	Process 
	{
		return Create-Parameter -Name $Name -DbType $DbType -Size $Size -Direction "Output" -Provider $Provider
	}#Process
} # Create-OutParameter

function Create-ReturnParameter {
<#
.Synopsis
    Creates a return parameter that can be used with Invoke-DB
.Description
    Create a return parameter object of type System.Data.DBParameter that can be used to 
	retrieve a value from a stored procedure call via a "return" statement in the stored procedure

.Parameter Name 
	The name of the return parameter
.Parameter Provider
	The ADO.NET provider you want to use. (Default="System.Data.SqlClient")
	Possible values are:
		System.Data.Odbc
		System.Data.OleDb
		System.Data.OracleClient
		System.Data.SqlClient
		Microsoft.SqlServerCe.Client
		System.Data.SqlServerCe
		System.Data.SqlServerCe.3.5 
.Example
    Create-ReturnParameter -Name "@My_Return_Value"
.ReturnValue
    A DbParameter object
.Link
    Invoke-DB
	Create-Parameter
	Create-InParameter
	Create-OutParameter
.Notes
 NAME:      Create-ReturnParameter
 AUTHOR:    Jorge Matos
 LASTEDIT:  05/13/2009
#Requires -Version 2.0
#>

[CmdletBinding(
    SupportsShouldProcess=$False,
    SupportsTransactions=$False, 
    ConfirmImpact="None",
    DefaultParameterSetName="")]
param(
	[Parameter(Position=0,Mandatory=1)]
	[string]$Name,

	[Parameter(Position=1)]
	[string]$Provider = "System.Data.SqlClient"
)

	Process 
	{
		return Create-Parameter -Name $Name -DbType "Int32" -Size 4 -Direction "ReturnValue" -Provider $Provider
	}#Process

} # Create-ReturnParameter

function Create-Parameter 
{
<#
.Synopsis
    Creates a parameter that can be used with Invoke-DB
.Description
    Creates an input/output/return parameter for a stored procedure

.Parameter Name 
	The name of the parameter
.Parameter DbType 
	The DbType of the parameter. Default="String"
	Possible values: Boolean, Int32, Double, Decimal, DateTime
	See "DbType Enumeration" in the MSDN documentation
.Parameter Size 
	The storage size in bytes of the parameter
	This is only required for output parameters
.Parameter Value 
	The value used for an input parameter
	This is not required for an output or return parameter
.Parameter Direction 
	The direction for the parameter.  
	Possible values are "Input", "Output", "InputOutput", or "ReturnValue"
.Parameter Provider
	The ADO.NET provider you want to use. (Default="System.Data.SqlClient")
	Possible values are:
		System.Data.Odbc
		System.Data.OleDb
		System.Data.OracleClient
		System.Data.SqlClient
		Microsoft.SqlServerCe.Client
		System.Data.SqlServerCe
		System.Data.SqlServerCe.3.5 
.Example
    Create-Parameter -Name "@p1" -Value 100 -DbType "Int32" -Direction "Input"
.ReturnValue
    DbParameter 
.Link
	Invoke-DB
	Create-InParameter
	Create-OutParameter
	Create-ReturnParameter
.Notes
 NAME:      Create-Parameter
 AUTHOR:    Jorge Matos
 LASTEDIT:  05/13/2009
#Requires -Version 2.0
#>

[CmdletBinding(
    SupportsShouldProcess=$False,
    SupportsTransactions=$False, 
    ConfirmImpact="None",
    DefaultParameterSetName="")]
param(
	[Parameter(Position=0,Mandatory=1)]
	[string]$Name,

	[Parameter(Position=1)]
	[System.Data.DbType]$DbType = "String",

	[Parameter(Position=2)]
	[int]$Size = $(if ($Direction -eq "Output") {throw "Please specify a parameter size."}),

	[Parameter(Position=3)]
	[object]$Value ,

	[Parameter(Position=4)]
	[System.Data.ParameterDirection]$Direction = "Input",

	[Parameter(Position=5)]
	[string]$Provider = "System.Data.SqlClient"
)

	Process 
	{
		$ProviderFactory = [System.Data.Common.DBProviderFactories]::GetFactory($Provider)
		$p = $ProviderFactory.CreateParameter()
		$p.ParameterName = $Name
		$p.Direction = $Direction
		$p.DbType = $DbType

		switch($direction) 
		{
			"Input" {$p.Value = $Value; break}
			"Output" {$p.Size = $Size; break}				
		}
		
		return $p
	}#Process
} # Create-Parameter
