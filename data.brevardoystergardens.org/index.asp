<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% option explicit %>
<!-- #include file="Connections/oysters.asp" -->
<%
' *** Validate request to log in to this site.
Dim MM_LoginAction
Dim MM_valUsername

MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("UserName"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  
  MM_fldUserAuthorization = "AccessLevel"
  MM_redirectLoginSuccess = "GardenerDataMenu.asp"
  MM_redirectLoginFailed = "http://brevardoystergardens.org/"

  MM_loginSQL = "SELECT UserName, Password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.accesscontrol WHERE UserName = ? AND Password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_oysters_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("Password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
	Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Brevard Oyster Garden Main</title>
<link href="twoColData.css" rel="stylesheet" type="text/css" />
</head>

<body class="twoColData">

<div id="container">
  <div id="header">
    <h1><img src="Images/oyster drawing.jpg" width="100" height="121" alt="Oyster" longdesc="Images/oyster drawing.jpg" />Brevard Oyster Gardens</h1>
  <!-- end #header --></div>
  <div id="sidebar1">
    <h3>Login</h3>
    <form id="login" name="login" method="POST" action="<%=MM_LoginAction%>">
      <p>User Name</p><input name="UserName" type="text" />
      <p>
        <label>Password
          <input type="password" name="Password" id="Password" />
        </label>
      </p>
      <p>
        <label>
          <input type="submit" name="login" id="login" value="Submit" />

               </label>
      </p>
    </form>
    <p>&nbsp;</p>
  <!-- end #sidebar1 --></div>
  <div id="mainContent">
    <h1>Data Access</h1>
    <p>This screen permits Oyster Gardeners to maintain their site data and enter observations.</p>
    <p>Staff can login to maintain the Oyster database and request reports.</p>
    <h2>Login Process</h2>
    <p>Oyster gardners login with their site id and PIN.</p>
    <p>If your are an Oyster Gardener and cannot login, please contact Sammy Anderson.</p>
    <p>Staff login with their user name and password.</p>
    <p>&nbsp;</p>
	<!-- end #mainContent --></div>
	<!-- This clearing element should immediately follow the #mainContent div in order to force the #container div to contain all child floats --><br class="clearfloat" />
  <div id="footer">
    <p> <a href="http://brevardoystergardens.org">Cancel</a></p>
  <!-- end #footer --></div>
<!-- end #container --></div>

</body>
</html>
