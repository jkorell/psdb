#make sure that Invoke-DB (dbc) always uses the correct default connectionString
$PSDB_DefaultConnectionString = "data source=.\sqlexpress;initial catalog=Northwind; Integrated Security=SSPI"

Test Should_Return_Data_Given_Simple_Select_Query {
	$rows = dbc -Sql "SELECT TOP 10 * FROM Orders" 
	AssertAreEqual $rows.Length 10
}

Test Should_Return_Data_Given_Parameters {
	$Parameters = @(
					(dbinput -Name "@Country" -Value "USA"),
					(dbinput -Name "@Freight" -Value 100)
					) 
	$sql = "SELECT * FROM Orders WHERE  (ShipCountry = @Country) AND (Freight > @Freight)" 
	$rows = dbc -Sql $sql -Parameters $Parameters
	AssertGreaterThan $rows.Length 0
} 

Test Should_Return_NoRowsAffected_GreaterThanZero_For_Simple_Update_Query {	
	[int]$rowsAffected = dbc -Sql "UPDATE Orders SET EmployeeID = 6 WHERE OrderID = 10248" -ExecuteType "NonQuery"
	AssertGreaterThan $rowsAffected 0 
} -ShouldRollBack

Test Should_Return_Data_Given_StoredProcedure {
	$rows = dbc -SPROC "Ten Most Expensive Products"
	AssertGreaterThan $rows.Length 0 
}

Test Should_Return_Data_Given_StoredProcedure_And_ConnectionString {
	$rows = dbc -SPROC "Ten Most Expensive Products" -Connectionstring "data source=.\sqlexpress;initial catalog=Northwind; Integrated Security=SSPI"
	AssertAreEqual $rows.Length 10 "Number of Rows must match up"
}

Test Should_Return_Data_Given_StoredProcedure_And_Parameter {
	$Parameters = @( dbinput -Name "@CustomerID" -Value "ALFKI" ) 
	$rows = dbc -SPROC "CustOrderHist" -Parameters $Parameters
	AssertGreaterThan $rows.Length 0 
}

Test Should_Return_Scalar_Data_Given_Simple_Select_Query {
	$result = dbc -Sql "SELECT COUNT(*) FROM Orders" -et "Scalar"
	AssertGreaterThan $result 0 
}

Test Should_Return_DataReader_Given_Simple_Select_Query {	
	$reader = dbc -Sql "SELECT TOP 2 * FROM Orders" -ExecuteType "Reader"	
	try
	{	
		Assert $reader.HasRows "Reader was empty"	
	}
	finally
	{
		$reader.Dispose()
	}
}

Test Should_Return_Data_When_Querying_Access {
	$accessfile = Resolve-Path .\Database1.mdb
	$provider = "System.Data.OleDb"
	$connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessfile"
	$sql = "SELECT * FROM users"
	$rows = dbc -Sql $sql -cs $connectionString -Provider $provider
	AssertGreaterThan $rows.Length 0 
}

Test Should_Return_Data_When_Querying_Excel {
	$excelfile = Resolve-Path .\book1.xls
	$provider = "System.Data.OleDb"
	$connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$excelfile;Extended Properties=Excel 8.0"
	$sql = "SELECT * FROM [Sheet1$]"
	$rows = dbc -Sql $sql -cs $connectionString -Provider $provider
	AssertGreaterThan $rows.Length 0 
}