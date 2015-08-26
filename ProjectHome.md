**psdb** is a module written in PowerShell V2 CTP 3 that allows you to invoke database commands against a database.  It uses ADO.NET's DbProviderFactory class to enable access to different types of databases.

Installation Instructions:
  1. Create a folder called **psdb** in your powershell modules folder - this can be either in your **my documents\windowspowershell\modules** folder or in the powershell machine-level folder **$pshome\modules**.
  1. Copy the **psdb.psm1** file to the **psdb** folder you just created.
  1. Run the command **Import-Module psdb**.
  1. Type the PowerShell command **help Invoke-DBCommand** for more info.

You may want to add the **Import-Module psdb** command to your profile so that it gets loaded automatically.

---


Examples

---

```
$rows = Invoke-DBCommand -Sql "SELECT TOP 10 * FROM Orders"

/*
Return an array of DataRow objects using default values:
ConnectionString = "data source=.\sqlexpress;initial catalog=Northwind; Integrated Security=SSPI"
ExecuteType = "Query"
Parameters = @()
Provider = System.Data.SqlClient
CommandTimeout = 600
*/
```

---

```
$sql = "SELECT TOP 10 * FROM Orders"
$rows = Invoke-DBCommand -Sql $sql -Connectionstring "data source=.;initial catalog=Northwind; uid=test; pwd=test"

//Return an array of DataRow objects with connectionstring
```

---

```
$sql = "UPDATE Orders SET EmployeeID = 6 WHERE OrderID = 10248"
$rowsAffected = Invoke-DBCommand -Sql $sql -ExecuteType "NonQuery"   

/*
Performs an update 
"NonQuery" is used to return the number of rows affected
*/
```

---

```
$Parameters = @(
(New-DBInputParameter -Name "@Country" -Value "USA"),
(New-DBInputParameter -Name "@Freight" -Value 100)
) 
$sql = "SELECT * FROM Orders WHERE  (ShipCountry = @Country) AND (Freight > @Freight)" 
$rows = Invoke-DBCommand -Sql $sql -Parameters $Parameters 	

//Using Parameters
```

---

```
$rows = Invoke-DBCommand -SPROC "Ten Most Expensive Products"

//Calling a stored procedure
```

---

```
$result = Invoke-DBCommand -Sql "SELECT COUNT(*) FROM Orders" -ExecuteType "Scalar"

//Returning a scalar value
```

---

```
$reader = Invoke-DBCommand -Sql "SELECT TOP 2 * FROM Orders" -ExecuteType "Reader"
if ($reader.HasRows) {
	while($reader.Read()) {
		"{0} {1}" -f $reader[0],$reader[1]
	}
}

$reader.Close()

//Return an IDataReader and access column values via index
```

---

```
$accessfile = Resolve-Path .\Database1.mdb
$provider = "System.Data.OleDb"
$connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$accessfile"
$sql = "SELECT * FROM users"
$rows = Invoke-DBCommand -Sql $sql -Connectionstring $connectionString -Provider $provider

//Query an Access Database
```