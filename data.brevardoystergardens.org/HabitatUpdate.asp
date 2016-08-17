<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="Connections/oysters.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2"
MM_authFailedURL="/GardenerDataMenu.asp"
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
Dim MM_SiteID

MM_SiteID = Session("MM_Username")

response.Write("Site ID = " & MM_SiteID) 
%>

<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_oysters_STRING
    MM_editCmd.CommandText = "UPDATE dbo.habitatmodules SET SiteID = ?, WaterDepthAtPlacementRed = ?, HabitatOffBottomRed = ?, DistanceFromNearestPilingRed = ?, DateHabitatInstalledRed = ?, DateOystersAddedRed = ?, WaterDepthAtPlacementYellow = ?, HabitatOffBottomYellow = ?, DistanceFromNearestPilingYellow = ?, DateHabitatInstalledYellow = ?, DateOystersAddedYellow = ?, WaterDepthAtModulePlacementGreen = ?, HabitatOffBottomGreen = ?, DistanceFromNearestPilingGreen = ?, DateHabitatInstalledGreen = ?, DateOystersAddedGreen = ?, WaterDepthAtModulePlacementBlue = ?, HabitatOffBottomBlue = ?, DistanceFromNearestPilingBlue = ?, DateHabitatInstalledBlue = ?, DateOystersAddedBlue = ?, HabitatModulesComment = ? WHERE SiteID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("SiteID"), Request.Form("SiteID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("WaterDepthAtPlacementRed"), Request.Form("WaterDepthAtPlacementRed"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("HabitatOffBottomRed"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("DistanceFromNearestPilingRed"), Request.Form("DistanceFromNearestPilingRed"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 10, Request.Form("DateHabitatInstalledRed")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 10, Request.Form("DateOystersAddedRed")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("WaterDepthAtPlacementYellow"), Request.Form("WaterDepthAtPlacementYellow"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("HabitatOffBottomYellow"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("DistanceFromNearestPilingYellow"), Request.Form("DistanceFromNearestPilingYellow"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 202, 1, 10, Request.Form("DateHabitatInstalledYellow")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 202, 1, 10, Request.Form("DateOystersAddedYellow")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 5, 1, -1, MM_IIF(Request.Form("WaterDepthAtModulePlacementGreen"), Request.Form("WaterDepthAtModulePlacementGreen"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 5, 1, -1, MM_IIF(Request.Form("HabitatOffBottomGreen"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 5, 1, -1, MM_IIF(Request.Form("DistanceFromNearestPilingGreen"), Request.Form("DistanceFromNearestPilingGreen"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 202, 1, 10, Request.Form("DateHabitatInstalledGreen")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 202, 1, 10, Request.Form("DateOystersAddedGreen")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 5, 1, -1, MM_IIF(Request.Form("WaterDepthAtModulePlacementBlue"), Request.Form("WaterDepthAtModulePlacementBlue"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 5, 1, -1, MM_IIF(Request.Form("HabitatOffBottomBlue"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param19", 5, 1, -1, MM_IIF(Request.Form("DistanceFromNearestPilingBlue"), Request.Form("DistanceFromNearestPilingBlue"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 202, 1, 10, Request.Form("DateHabitatInstalledBlue")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 202, 1, 10, Request.Form("DateOystersAddedBlue")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 201, 1, 2000, Request.Form("HabitatModulesComment")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param23", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "/GardenerDataMenu.asp"
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
Dim rsHabitats
Dim rsHabitats_cmd
Dim rsHabitats_numRows

Set rsHabitats_cmd = Server.CreateObject ("ADODB.Command")
rsHabitats_cmd.ActiveConnection = MM_oysters_STRING
rsHabitats_cmd.CommandText = "SELECT * FROM dbo.habitatmodules WHERE SiteID = "  & MM_SiteID
rsHabitats_cmd.Prepared = true

Set rsHabitats = rsHabitats_cmd.Execute
rsHabitats_numRows = 0
%>
<html xmlns="http://www.w3.org/1999/xhtml">


<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SiteID:</td>
      <td><input type="text" name="SiteID" value="<%=(rsHabitats.Fields.Item("SiteID").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Water Depth At Placement Red:</td>
      <td><input type="text" name="WaterDepthAtPlacementRed" value="<%=(rsHabitats.Fields.Item("WaterDepthAtPlacementRed").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Habitat Off Bottom Red:</td>
      <td><input type="checkbox" name="HabitatOffBottomRed" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Distance From Nearest Piling Red:</td>
      <td><input type="text" name="DistanceFromNearestPilingRed" value="<%=(rsHabitats.Fields.Item("DistanceFromNearestPilingRed").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Habitat Installed Red:</td>
      <td><input type="text" name="DateHabitatInstalledRed" value="<%=(rsHabitats.Fields.Item("DateHabitatInstalledRed").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Oysters Added Red:</td>
      <td><input type="text" name="DateOystersAddedRed" value="<%=(rsHabitats.Fields.Item("DateOystersAddedRed").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Water Depth At Placement Yellow:</td>
      <td><input type="text" name="WaterDepthAtPlacementYellow" value="<%=(rsHabitats.Fields.Item("WaterDepthAtPlacementYellow").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Habitat Off Bottom Yellow:</td>
      <td><input type="checkbox" name="HabitatOffBottomYellow" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Distance From Nearest Piling Yellow:</td>
      <td><input type="text" name="DistanceFromNearestPilingYellow" value="<%=(rsHabitats.Fields.Item("DistanceFromNearestPilingYellow").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Habitat Installed Yellow:</td>
      <td><input type="text" name="DateHabitatInstalledYellow" value="<%=(rsHabitats.Fields.Item("DateHabitatInstalledYellow").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Oysters Added Yellow:</td>
      <td><input type="text" name="DateOystersAddedYellow" value="<%=(rsHabitats.Fields.Item("DateOystersAddedYellow").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Water Depth At Placement Green:</td>
      <td><input type="text" name="WaterDepthAtModulePlacementGreen" value="<%=(rsHabitats.Fields.Item("WaterDepthAtModulePlacementGreen").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Habitat Off Bottom Green:</td>
      <td><input type="checkbox" name="HabitatOffBottomGreen" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Distance From Nearest Piling Green:</td>
      <td><input type="text" name="DistanceFromNearestPilingGreen" value="<%=(rsHabitats.Fields.Item("DistanceFromNearestPilingGreen").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Habitat Installed Green:</td>
      <td><input type="text" name="DateHabitatInstalledGreen" value="<%=(rsHabitats.Fields.Item("DateHabitatInstalledGreen").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Oysters Added Green:</td>
      <td><input type="text" name="DateOystersAddedGreen" value="<%=(rsHabitats.Fields.Item("DateOystersAddedGreen").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Water Depth At Placement Blue:</td>
      <td><input type="text" name="WaterDepthAtModulePlacementBlue" value="<%=(rsHabitats.Fields.Item("WaterDepthAtModulePlacementBlue").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Habitat Off Bottom Blue:</td>
      <td><input type="checkbox" name="HabitatOffBottomBlue" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Distance From Nearest Piling Blue:</td>
      <td><input type="text" name="DistanceFromNearestPilingBlue" value="<%=(rsHabitats.Fields.Item("DistanceFromNearestPilingBlue").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Habitat Installed Blue:</td>
      <td><input type="text" name="DateHabitatInstalledBlue" value="<%=(rsHabitats.Fields.Item("DateHabitatInstalledBlue").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Date Oysters Added Blue:</td>
      <td><input type="text" name="DateOystersAddedBlue" value="<%=(rsHabitats.Fields.Item("DateOystersAddedBlue").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Habitat Modules Comment:</td>
      <td><input type="text" name="HabitatModulesComment" value="<%=(rsHabitats.Fields.Item("HabitatModulesComment").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="Update record" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1" />
  <input type="hidden" name="MM_recordId" value="<%= rsHabitats.Fields.Item("SiteID").Value %>" />
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsHabitats.Close()
Set rsHabitats = Nothing
%>
