<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<p>&lt;%</p>
<p>'' objDbConn is the object that is to be used in all ASP web pages that require access to the <br />
  '' database.  The class, clsDatabaseConnection, is responsible for creating the actual connection and is<br />
  '' designed to close the database connection once the ASP page using the connection goes out of scope.  This<br />
  '' design eliminates the need to call 'CloseConnection' and the possibility of leaving multiple open database<br />
  '' connections out there.<br />
  Public objDbConn</p>
<p>&nbsp;</p>
<p>'' Alter the connection string below as needed to establish a connection with the target database.<br />
  '' No other changes to this page are needed!<br />
  Private databaseConnection<br />
  Set databaseConnection = New clsDatabaseConnection<br />
  Call databaseConnection.OpenConnection(&quot;Provider=SQLOLEDB.1;Persist Security Info=False;User ID=daveo17;password=Reneea01*;Initial Catalog=oysters;Data Source=184.168.194.53&quot;)<br />
  <br />
  Set objDbConn = databaseConnection.Connection<br />
</p>
<p>'' ------------------------------------------------------------------------------<br />
  Class clsDatabaseConnection</p>
<p> Private mConnection<br />
  Private mConnectionString<br />
</p>
<p> '' ---------------------------------------------------------------------------<br />
  '' Constructor - Initializes class components<br />
  '' ---------------------------------------------------------------------------<br />
  Private Sub Class_Initialize()</p>
<p> End Sub</p>
<p> '' ---------------------------------------------------------------------------<br />
  '' Destructor - Terminates open database connection automatically<br />
  '' ---------------------------------------------------------------------------<br />
  Private Sub Class_Terminate()</p>
<p> If Not mConnection Is Nothing Then</p>
<p> ''	Conenction is open - close it!<br />
  mConnection.Close<br />
  Set mConnection = Nothing</p>
<p> End If</p>
<p> End Sub</p>
<p> '' ---------------------------------------------------------------------------<br />
  '' Interface to access open database connection<br />
  '' ---------------------------------------------------------------------------<br />
  Public Property Get Connection()<br />
  <br />
  Set Connection = mConnection<br />
  <br />
  End Property</p>
<p> '' ---------------------------------------------------------------------------<br />
  '' Returns connection string used to open database connection<br />
  '' ---------------------------------------------------------------------------<br />
  Public Property Get ConnectionString</p>
<p> ConnectionString = mConnectionString</p>
<p> End Property</p>
<p> '' ---------------------------------------------------------------------------<br />
  '' Opens connection to the database.  Connection string must be set first or<br />
  '' else an error will occur!<br />
  '' ---------------------------------------------------------------------------<br />
  Public Sub OpenConnection(ByRef rConnectionString)<br />
  <br />
  mConnectionString = rConnectionString</p>
<p> Set mConnection = Server.CreateObject(&quot;ADODB.Connection&quot;)<br />
  Call mConnection.Open(mConnectionString)</p>
<p> End Sub</p>
<p>End Class<br />
</p>
<p>%&gt;</p>
</body>
</html>
