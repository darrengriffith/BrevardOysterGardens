<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Connections/oysters.asp" -->
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
Dim MM_Today

MM_Today = date()

MM_SiteID = Session("MM_Username")
response.Write("Site ID = " & MM_SiteID) 
response.Write("  Today = " & MM_Today)

%>
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_oysters_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.habitatobservations (SiteID, DateOfObservation, WtBeforeCleaningRed, WtAfterCleaningRed, WtBeforeCleaningYellow, WtAfterCleaningYellow, WtBeforeCleaningGreen, WtAfterCleaningGreen, WtBeforeCleaningBlue, WtAfterCleaningBlue, Barnacle, BoringSponge, BlueCrab, GrassShrimp, HermitCrab, LionFish, MudCrab, PinkShrimp, RibbedMussel, SeaSquirt, Sheepshead, SheepsheadMinnow, SlipperShell, SnappingShrimp, StoneCrab, OtherOrganisms, SpatOnRecruitmentShell, FoulingLoad, TimeSpentOnObservation, ObservationComment) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("SiteID"),MM_SiteID, MM_SiteID)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 10, Request.Form("DateOfObservation")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("WtBeforeCleaningRed"), Request.Form("WtBeforeCleaningRed"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("WtAfterCleaningRed"), Request.Form("WtAfterCleaningRed"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("WtBeforeCleaningYellow"), Request.Form("WtBeforeCleaningYellow"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("WtAfterCleaningYellow"), Request.Form("WtAfterCleaningYellow"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("WtBeforeCleaningGreen"), Request.Form("WtBeforeCleaningGreen"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("WtAfterCleaningGreen"), Request.Form("WtAfterCleaningGreen"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("WtBeforeCleaningBlue"), Request.Form("WtBeforeCleaningBlue"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 5, 1, -1, MM_IIF(Request.Form("WtAfterCleaningBlue"), Request.Form("WtAfterCleaningBlue"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 5, 1, -1, MM_IIF(Request.Form("Barnacle"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 5, 1, -1, MM_IIF(Request.Form("BoringSponge"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 5, 1, -1, MM_IIF(Request.Form("BlueCrab"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 5, 1, -1, MM_IIF(Request.Form("GrassShrimp"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 5, 1, -1, MM_IIF(Request.Form("HermitCrab"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 5, 1, -1, MM_IIF(Request.Form("LionFish"), 1, 0)) ' adDouble 
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 5, 1, -1, MM_IIF(Request.Form("MudCrab"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 5, 1, -1, MM_IIF(Request.Form("PinkShrimp"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param19", 5, 1, -1, MM_IIF(Request.Form("RibbedMussel"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 5, 1, -1, MM_IIF(Request.Form("SeaSquirt"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 5, 1, -1, MM_IIF(Request.Form("Sheepshead"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 5, 1, -1, MM_IIF(Request.Form("SheepsheadMinnow"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param23", 5, 1, -1, MM_IIF(Request.Form("SlipperShell"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param24", 5, 1, -1, MM_IIF(Request.Form("SnappingShrimp"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param25", 5, 1, -1, MM_IIF(Request.Form("StoneCrab"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param26", 201, 1, 2000, Request.Form("OtherOrganisms")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param27", 5, 1, -1, MM_IIF(Request.Form("SpatOnRecruitmentShell"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param28", 201, 1, 45, Request.Form("FoulingLoad")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param29", 5, 1, -1, MM_IIF(Request.Form("TimeSpentOnObservation"), Request.Form("TimeSpentOnObservation"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param30", 201, 1, 2000, Request.Form("ObservationComment")) ' adLongVarChar
   
   response.write(editCmd)
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
Dim rsObservations
Dim rsObservations_cmd
Dim rsObservations_numRows

Set rsObservations_cmd = Server.CreateObject ("ADODB.Command")
rsObservations_cmd.ActiveConnection = MM_oysters_STRING
rsObservations_cmd.CommandText = "SELECT * FROM dbo.habitatobservations" 
rsObservations_cmd.Prepared = true

Set rsObservations = rsObservations_cmd.Execute
rsObservations_numRows = 0
%>

<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">DateOfObservation:</td>
      <td><input type="text" name="DateOfObservation" value= "" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtBeforeCleaningRed:</td>
      <td><input type="text" name="WtBeforeCleaningRed" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtAfterCleaningRed:</td>
      <td><input type="text" name="WtAfterCleaningRed" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtBeforeCleaningYellow:</td>
      <td><input type="text" name="WtBeforeCleaningYellow" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtAfterCleaningYellow:</td>
      <td><input type="text" name="WtAfterCleaningYellow" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtBeforeCleaningGreen:</td>
      <td><input type="text" name="WtBeforeCleaningGreen" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtAfterCleaningGreen:</td>
      <td><input type="text" name="WtAfterCleaningGreen" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtBeforeCleaningBlue:</td>
      <td><input type="text" name="WtBeforeCleaningBlue" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">WtAfterCleaningBlue:</td>
      <td><input type="text" name="WtAfterCleaningBlue" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Barnacle:</td>
      <td><input type="checkbox" name="Barnacle" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">BoringSponge:</td>
      <td><input type="checkbox" name="BoringSponge" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">BlueCrab:</td>
      <td><input type="checkbox" name="BlueCrab" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">GrassShrimp:</td>
      <td><input type="checkbox" name="GrassShrimp" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">HermitCrab:</td>
      <td><input type="checkbox" name="HermitCrab" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">LionFish:</td>
      <td><input type="checkbox" name="LionFish" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">MudCrab:</td>
      <td><input type="checkbox" name="MudCrab" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">PinkShrimp:</td>
      <td><input type="checkbox" name="PinkShrimp" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">RibbedMussel:</td>
      <td><input type="checkbox" name="RibbedMussel" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SeaSquirt:</td>
      <td><input type="checkbox" name="SeaSquirt" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Sheepshead:</td>
      <td><input type="checkbox" name="Sheepshead" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SheepsheadMinnow:</td>
      <td><input type="checkbox" name="SheepsheadMinnow" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SlipperShell:</td>
      <td><input type="checkbox" name="SlipperShell" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SnappingShrimp:</td>
      <td><input type="checkbox" name="SnappingShrimp" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">StoneCrab:</td>
      <td><input type="checkbox" name="StoneCrab" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">OtherOrganisms:</td>
      <td><input type="text" name="OtherOrganisms" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">SpatOnRecruitmentShell:</td>
      <td><input type="checkbox" name="SpatOnRecruitmentShell" value="1" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">FoulingLoad:</td>
      <td><input type="text" name="FoulingLoad" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Minutes Spent On Observation:</td>
      <td><input type="text" name="TimeSpentOnObservation" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">ObservationComment:</td>
      <td><input type="text" name="ObservationComment" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
</table>
%>
<p>&nbsp;</p>
<%
rsObservations.Close()
Set rsObservations = Nothing
%>
