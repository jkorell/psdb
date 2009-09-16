psdb files

Book1.xls and Database1.mdb are used by the examples in the help for the Invoke-DBCommand function
psdb.psm1 is the module containing the source 
psdb_tests.ps1 is a test script that exercises most of the features of psdb
pstest.psm1 is a module containing a simple unit testing framework

Tip: 
Instead of using the -ConnectionString or the -CommandTimeout parameters for Invoke-DBCommand 
you can instead set the following module Variables:

$PSDB_DefaultConnectionString = Default Connectionstring 
$PSDB_DefaultCommandTimeout = Default Command Timeout (600 secs)