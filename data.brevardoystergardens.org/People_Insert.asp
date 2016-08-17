<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/Connections/oysters.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="9"
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_oysters_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.people (SiteID, Status, FirstName, LastName, BirthDate, EMail, ContactPhone, AlternatePhone, StreetAddress, City, County, [State], Zipcode, PermissionSiteAccess, SiteAccess, WorkshopDate, SiteOwner, OysterBuddy, MasterOysterGardener, DateWithdrawn, PersonComment, SiteLatitude, SiteLongitude, SiteComments, SiteCharacteristicStructure, SiteCharacteristicFlow, SiteCharacteristicShore, MainWaterBody, SecondaryWaterBody, CanalDistanceFromFreshWaterSource, CanalDistanceFromStormDrain, CanalDistanceFromDeadend, DatePhotoOfSite) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("SiteID"), Request.Form("SiteID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("Status"), Request.Form("Status"), null)) ' adDouble
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
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 202, 1, 10, Request.Form("DateWithdrawn")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 201, 1, 2000, Request.Form("PersonComment")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 201, 1, 50, Request.Form("SiteLatitude")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param23", 201, 1, 50, Request.Form("SiteLongitude")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param24", 201, 1, 1000, Request.Form("SiteComments")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param25", 201, 1, 50, Request.Form("SiteCharacteristicStructure")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param26", 201, 1, 50, Request.Form("SiteCharacteristicFlow")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param27", 201, 1, 50, Request.Form("SiteCharacteristicShore")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param28", 201, 1, 50, Request.Form("MainWaterBody")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param29", 201, 1, 50, Request.Form("SecondaryWaterBody")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param30", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromFreshWaterSource"), Request.Form("CanalDistanceFromFreshWaterSource"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param31", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromStormDrain"), Request.Form("CanalDistanceFromStormDrain"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param32", 5, 1, -1, MM_IIF(Request.Form("CanalDistanceFromDeadend"), Request.Form("CanalDistanceFromDeadend"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param33", 202, 1, 10, Request.Form("DatePhotoOfSite")) ' adVarWChar
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
Dim rsPeople
Dim rsPeople_cmd
Dim rsPeople_numRows

Set rsPeople_cmd = Server.CreateObject ("ADODB.Command")
rsPeople_cmd.ActiveConnection = MM_oysters_STRING
rsPeople_cmd.CommandText = "SELECT SiteID, Status, FirstName, LastName, BirthDate, EMail, ContactPhone, AlternatePhone, StreetAddress, City, County, [State], Zipcode, PermissionSiteAccess, SiteAccess, WorkshopDate, SiteOwner, OysterBuddy, MasterOysterGardener, DateWithdrawn, PersonComment, SiteLatitude, SiteLongitude, SiteComments, SiteCharacteristicStructure, SiteCharacteristicFlow, SiteCharacteristicShore, MainWaterBody, SecondaryWaterBody, CanalDistanceFromFreshWaterSource, CanalDistanceFromStormDrain, CanalDistanceFromDeadend, DatePhotoOfSite FROM dbo.people" 
rsPeople_cmd.Prepared = true

Set rsPeople = rsPeople_cmd.Execute
rsPeople_numRows = 0
%>
<%
rsPeople.Close()
Set rsPeople = Nothing
%>

<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteID:</td>
      <td><input type="text" name="SiteID" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Status:</td>
      <td><input type="text" name="Status" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">FirstName:</td>
      <td><input type="text" name="FirstName" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">LastName:</td>
      <td><input type="text" name="LastName" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">BirthDate:</td>
      <td><input type="text" name="BirthDate" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">EMail:</td>
      <td><input type="text" name="EMail" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">ContactPhone:</td>
      <td><input type="text" name="ContactPhone" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">AlternatePhone:</td>
      <td><input type="text" name="AlternatePhone" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">StreetAddress:</td>
      <td><input type="text" name="StreetAddress" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">City:</td>
      <td><input type="text" name="City" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">County:</td>
      <td><input type="text" name="County" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">State:</td>
      <td><input type="text" name="State" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Zipcode:</td>
      <td><input type="text" name="Zipcode" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">PermissionSiteAccess:</td>
      <td><input type="text" name="PermissionSiteAccess" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteAccess:</td>
      <td><input type="text" name="SiteAccess" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WorkshopDate:</td>
      <td><input type="text" name="WorkshopDate" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteOwner:</td>
      <td><input type="checkbox" name="SiteOwner" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">OysterBuddy:</td>
      <td><input type="checkbox" name="OysterBuddy" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">MasterOysterGardener:</td>
      <td><input type="checkbox" name="MasterOysterGardener" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">DateWithdrawn:</td>
      <td><input type="text" name="DateWithdrawn" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">PersonComment:</td>
      <td><input type="text" name="PersonComment" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteLatitude:</td>
      <td><input type="text" name="SiteLatitude" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteLongitude:</td>
      <td><input type="text" name="SiteLongitude" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteComments:</td>
      <td><input type="text" name="SiteComments" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteCharacteristicStructure:</td>
      <td><input type="text" name="SiteCharacteristicStructure" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteCharacteristicFlow:</td>
      <td><input type="text" name="SiteCharacteristicFlow" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteCharacteristicShore:</td>
      <td><input type="text" name="SiteCharacteristicShore" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">MainWaterBody:</td>
      <td><input type="text" name="MainWaterBody" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SecondaryWaterBody:</td>
      <td><input type="text" name="SecondaryWaterBody" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">CanalDistanceFromFreshWaterSource:</td>
      <td><input type="text" name="CanalDistanceFromFreshWaterSource" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">CanalDistanceFromStormDrain:</td>
      <td><input type="text" name="CanalDistanceFromStormDrain" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">CanalDistanceFromDeadend:</td>
      <td><input type="text" name="CanalDistanceFromDeadend" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">DatePhotoOfSite:</td>
      <td><input type="text" name="DatePhotoOfSite" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
<p>&nbsp;</p>
