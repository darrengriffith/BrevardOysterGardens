<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="Connections/oysters.asp" -->

<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2"
MM_authFailedURL="GardenerDataMenu.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim MM_SiteID

MM_SiteID = Session("MM_Username")

response.Write("Site ID = " & MM_SiteID) 
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_oysters_STRING
    MM_editCmd.CommandText = "UPDATE dbo.people SET SiteID = ?, FirstName = ?, LastName = ?, BirthDate = ?, EMail = ?, ContactPhone = ?, AlternatePhone = ?, StreetAddress = ?, City = ?, County = ?, [State] = ?, Zipcode = ?, PermissionSiteAccess = ?, SiteAccess = ?, WorkshopDate = ?, SiteOwner = ?, OysterBuddy = ?, MasterOysterGardener = ?, PersonComment = ?, SiteLatitude = ?, SiteLongitude = ?, SiteComments = ?, SiteCharacteristicStructure = ?, SiteCharacteristicFlow = ?, SiteCharacteristicShore = ?, MainWaterBody = ?, SecondaryWaterBody = ?, CanalDistanceFromFreshWaterSource = ?, CanalDistanceFromStormDrain = ?, CanalDistanceFromDeadend = ? WHERE idPeople = ?" 
    MM_editCmd.Prepared = true
       MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("SiteID"), Request.Form("SiteID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 25, Request.Form("FirstName")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 25, Request.Form("LastName")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 10, Request.Form("BirthDate")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, 100, Request.Form("EMail")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, 25, Request.Form("ContactPhone")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 201, 1, 25, Request.Form("AlternatePhone")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 201, 1, 45, Request.Form("StreetAddress")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 201, 1, 45, Request.Form("City")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 201, 1, 50, Request.Form("County")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 201, 1, 9, Request.Form("State")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 201, 1, 9, Request.Form("Zipcode")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 201, 1, 50, Request.Form("PermissionSiteAccess")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 201, 1, 1000, Request.Form("SiteAccess")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 201, 1, 50, Request.Form("WorkshopDate")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 5, 1, -1, MM_IIF(Request.Form("SiteOwner"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 5, 1, -1, MM_IIF(Request.Form("OysterBuddy"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param19", 5, 1, -1, MM_IIF(Request.Form("MasterOysterGardener"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 201, 1, 2000, Request.Form("PersonComment")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 201, 1, 50, Request.Form("SiteLatitude")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 201, 1, 50, Request.Form("SiteLongitude")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param23", 201, 1, 1000, Request.Form("SiteComments")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param24", 201, 1, 50, Request.Form("SiteCharacteristicStructure")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param25", 201, 1, 50, Request.Form("SiteCharacteristicFlow")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param26", 201, 1, 50, Request.Form("SiteCharacteristicShore")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param27", 201, 1, 50, Request.Form("MainWaterBody")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param28", 201, 1, 50, Request.Form("SecondaryWaterBody")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param29", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromFreshWaterSource"), Request.Form("CanalDistanceFromFreshWaterSource"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param30", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromStormDrain"), Request.Form("CanalDistanceFromStormDrain"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param31", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromDeadend"), Request.Form("CanalDistanceFromDeadend"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param32", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "GardenerDataMenu.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim rsGardenerProfile
Dim rsGardenerProfile_cmd
Dim rsGardenerProfile_numRows

Set rsGardenerProfile_cmd = Server.CreateObject ("ADODB.Command")
rsGardenerProfile_cmd.ActiveConnection = MM_oysters_STRING

rsGardenerProfile_cmd.CommandText = "SELECT * FROM dbo.people WHERE idPeople =" & MM_SiteID 
rsGardenerProfile_cmd.Prepared = true

Set rsGardenerProfile = rsGardenerProfile_cmd.Execute
rsGardenerProfile_numRows = 0
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Gardener Menu</title>
<link href="/oneColData.css" rel="stylesheet" type="text/css" />
</head>

<body class="oneColElsCtrHdr">

<div id="container">
  <div id="header">
    <h1><img src="/oyster drawing.jpg" width="124" height="140" alt="Oyster" />Brevard Oyster Restoration</h1>
  <!-- end #header --></div>
  <div id="mainContent">
    <h1> Edit Gardener and Site Information</h1>
    <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
      <table align="center">
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Id:</td>
          <td><input type="text" name="idPeople" value="<%=(rsGardenerProfile.Fields.Item("idPeople").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Site ID:</td>
          <td><input type="text" name="SiteID" value="<%=(rsGardenerProfile.Fields.Item("SiteID").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">First Name:</td>
          <td><input type="text" name="FirstName" value="<%=(rsGardenerProfile.Fields.Item("FirstName").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Last Name:</td>
          <td><input type="text" name="LastName" value="<%=(rsGardenerProfile.Fields.Item("LastName").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Birth Date:</td>
          <td><input type="text" name="BirthDate" value="<%=(rsGardenerProfile.Fields.Item("BirthDate").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">E-Mail:</td>
          <td><input type="text" name="EMail" value="<%=(rsGardenerProfile.Fields.Item("EMail").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Contact Phone:</td>
          <td><input type="text" name="ContactPhone" value="<%=(rsGardenerProfile.Fields.Item("ContactPhone").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Alternate Phone:</td>
          <td><input type="text" name="AlternatePhone" value="<%=(rsGardenerProfile.Fields.Item("AlternatePhone").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Street Address:</td>
          <td><input type="text" name="StreetAddress" value="<%=(rsGardenerProfile.Fields.Item("StreetAddress").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">City:</td>
          <td><input type="text" name="City" value="<%=(rsGardenerProfile.Fields.Item("City").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">County:</td>
          <td><input type="text" name="County" value="<%=(rsGardenerProfile.Fields.Item("County").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">State:</td>
          <td><input type="text" name="State" value="<%=(rsGardenerProfile.Fields.Item("State").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Zipcode:</td>
          <td><input type="text" name="Zipcode" value="<%=(rsGardenerProfile.Fields.Item("Zipcode").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td height="68" align="right" nowrap="nowrap">May staff or volunteers
          <p> access site for additional</p>
          <p>data collection?:</p></td>
          <td><input type="text" name="PermissionSiteAccess" value="<%=(rsGardenerProfile.Fields.Item("PermissionSiteAccess").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Instructions for site access:</td>
          <td><input type="text" name="SiteAccess" value="<%=(rsGardenerProfile.Fields.Item("SiteAccess").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">WorkshopDate:</td>
          <td><input type="text" name="WorkshopDate" value="<%=(rsGardenerProfile.Fields.Item("WorkshopDate").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Site Owner:</td>
          <td><input type="checkbox" name="SiteOwner" value="1" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Oyster Buddy:</td>
          <td><input type="checkbox" name="OysterBuddy" value="1" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Master Oyster Gardener:</td>
          <td><input type="checkbox" name="MasterOysterGardener" value="1" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Comment:</td>
          <td><input type="text" name="PersonComment" value="<%=(rsGardenerProfile.Fields.Item("PersonComment").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">SiteLatitude:</td>
          <td><input type="text" name="SiteLatitude" value="<%=(rsGardenerProfile.Fields.Item("SiteLatitude").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">SiteLongitude:</td>
          <td><input type="text" name="SiteLongitude" value="<%=(rsGardenerProfile.Fields.Item("SiteLongitude").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Site Comments:</td>
          <td><input type="text" name="SiteComments" value="<%=(rsGardenerProfile.Fields.Item("SiteComments").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Type of dock:</td>
          <td><input type="text" name="SiteCharacteristicStructure" value="<%=(rsGardenerProfile.Fields.Item("SiteCharacteristicStructure").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Flow:
          <p>Enter flow as </p>
          <p>Low, Moderate, High</p></td>
          <td><input type="text" name="SiteCharacteristicFlow" value="<%=(rsGardenerProfile.Fields.Item("SiteCharacteristicFlow").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Shore:
          Enter Shore
          <p>as Natural, Rip Rap, Wall</p></td>
          <td><input type="text" name="SiteCharacteristicShore" value="<%=(rsGardenerProfile.Fields.Item("SiteCharacteristicShore").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Main Water Body:
          <p>Enter as Indian River or </p>
          <p>Banana River</p></td>
          <td><input type="text" name="MainWaterBody" value="<%=(rsGardenerProfile.Fields.Item("MainWaterBody").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">Site is located on
            <p>Main body:</p>
          <p> Grand/Arterial canal</p>
          <p>Secondary canal:</p></td>
          <td><input type="text" name="SecondaryWaterBody" value="<%=(rsGardenerProfile.Fields.Item("SecondaryWaterBody").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">What is the site's distance
          <p> from the main body? (feet)</p></td>
          <td><input type="text" name="CanalDistanceFromFreshWaterSource" value="<%=(rsGardenerProfile.Fields.Item("CanalDistanceFromFreshWaterSource").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">What is the site's distance
          <p> from storm drain? (feet):</p></td>
          <td><input type="text" name="CanalDistanceFromStormDrain" value="<%=(rsGardenerProfile.Fields.Item("CanalDistanceFromStormDrain").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">If on canal, what is distance 
          <p>from the deadend?: (feet)</p></td>
          <td><input type="text" name="CanalDistanceFromDeadend" value="<%=(rsGardenerProfile.Fields.Item("CanalDistanceFromDeadend").Value)%>" size="32" /></td>
        </tr>
        <tr valign="baseline">
          <td nowrap="nowrap" align="right">&nbsp;</td>
          <td><input type="submit" value="Update record" /></td>
        </tr>
      </table>
      <input type="hidden" name="MM_update" value="form1" />
      <input type="hidden" name="MM_recordId" value="<%= rsGardenerProfile.Fields.Item("idPeople").Value %>" />
    </form>
    <p>&nbsp;</p>
    
    
	<!-- end #mainContent --></div>
  <div id="footer">
    <p>    <p>Gardener Data Menu 4/16/2014	  SiteID = <%= Session("MM_Username") %></p></p>
  <!-- end #footer --></div>
<!-- end #container --></div>
</body>
</html>
<%

rsGardenerProfile.Close()
Set rsGardenerProfile = Nothing
%>
