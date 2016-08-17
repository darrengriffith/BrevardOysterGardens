<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="Connections/oysters.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="2"
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
Dim MM_Today

MM_Today = date()

MM_SiteID = Session("MM_Username")


response.Write("Site ID = " & MM_SiteID) 
response.Write("  Today = " & MM_Today)

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
    MM_editCmd.CommandText = "INSERT INTO dbo.oystermeasurements (SiteID, ObservationDate, TimeSpentonObservation, LiveRed, RedWtRubbermaid, Red_1, Red_2, Red_3, Red_4, Red_5, Red_6, Red_7, Red_8, Red_9, Red_10, Red_11, Red_12, Red_13, Red_14, Red_15, Red_16, Red_17, Red_18, Red_19, Red_20, Red_21, Red_22, Red_23, Red_24, Red_25, LiveYellow, YellowWtRubbermaid, Yellow_1, Yellow_2, Yellow_3, Yellow_4, Yellow_5, Yellow_6, Yellow_7, Yellow_8, Yellow_9, Yellow_10, Yellow_11, Yellow_12, Yellow_13, Yellow_14, Yellow_15, Yellow_16, Yellow_17, Yellow_18, Yellow_19, Yellow_20, Yellow_21, Yellow_22, Yellow_23, Yellow_24, Yellow_25, LiveGreen, GreenRubbermaid, Green_1, Green_2, Green_3, Green_4, Green_5, Green_6, Green_7, Green_8, Green_9, Green_10, Green_11, Green_12, Green_13, Green_14, Green_15, Green_16, Green_17, Green_18, Green_19, Green_20, Green_21, Green_22, Green_23, Green_24, Green_25, LiveBlue, BlueWtRubbermaid, Blue_1, Blue_2, Blue_3, Blue_4, Blue_5, Blue_6, Blue_7, Blue_8, Blue_9, Blue_10, Blue_11, Blue12, Blue_13, Blue_14, Blue_15, Blue_16, Blue_17, Blue_18, Blue_19, Blue_20, Blue_21, Blue_22, Blue_23, Blue_24, Blue_25) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("SiteID"),MM_SiteID, MM_SiteID)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 10, Request.Form("ObservationDate")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("TimeSpentonObservation"), Request.Form("TimeSpentonObservation"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("LiveRed"), Request.Form("LiveRed"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("RedWtRubbermaid"), Request.Form("RedWtRubbermaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("Red_1"), Request.Form("Red_1"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("Red_2"), Request.Form("Red_2"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("Red_3"), Request.Form("Red_3"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("Red_4"), Request.Form("Red_4"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 5, 1, -1, MM_IIF(Request.Form("Red_5"), Request.Form("Red_5"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 5, 1, -1, MM_IIF(Request.Form("Red_6"), Request.Form("Red_6"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 5, 1, -1, MM_IIF(Request.Form("Red_7"), Request.Form("Red_7"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 5, 1, -1, MM_IIF(Request.Form("Red_8"), Request.Form("Red_8"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 5, 1, -1, MM_IIF(Request.Form("Red_9"), Request.Form("Red_9"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 5, 1, -1, MM_IIF(Request.Form("Red_10"), Request.Form("Red_10"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 5, 1, -1, MM_IIF(Request.Form("Red_11"), Request.Form("Red_11"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 5, 1, -1, MM_IIF(Request.Form("Red_12"), Request.Form("Red_12"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 5, 1, -1, MM_IIF(Request.Form("Red_13"), Request.Form("Red_13"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param19", 5, 1, -1, MM_IIF(Request.Form("Red_14"), Request.Form("Red_14"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param20", 5, 1, -1, MM_IIF(Request.Form("Red_15"), Request.Form("Red_15"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param21", 5, 1, -1, MM_IIF(Request.Form("Red_16"), Request.Form("Red_16"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param22", 5, 1, -1, MM_IIF(Request.Form("Red_17"), Request.Form("Red_17"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param23", 5, 1, -1, MM_IIF(Request.Form("Red_18"), Request.Form("Red_18"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param24", 5, 1, -1, MM_IIF(Request.Form("Red_19"), Request.Form("Red_19"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param25", 5, 1, -1, MM_IIF(Request.Form("Red_20"), Request.Form("Red_20"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param26", 5, 1, -1, MM_IIF(Request.Form("Red_21"), Request.Form("Red_21"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param27", 5, 1, -1, MM_IIF(Request.Form("Red_22"), Request.Form("Red_22"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param28", 5, 1, -1, MM_IIF(Request.Form("Red_23"), Request.Form("Red_23"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param29", 5, 1, -1, MM_IIF(Request.Form("Red_24"), Request.Form("Red_24"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param30", 5, 1, -1, MM_IIF(Request.Form("Red_25"), Request.Form("Red_25"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param31", 5, 1, -1, MM_IIF(Request.Form("LiveYellow"), Request.Form("LiveYellow"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param32", 5, 1, -1, MM_IIF(Request.Form("YellowWtRubbermaid"), Request.Form("YellowWtRubbermaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param33", 5, 1, -1, MM_IIF(Request.Form("Yellow_1"), Request.Form("Yellow_1"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param34", 5, 1, -1, MM_IIF(Request.Form("Yellow_2"), Request.Form("Yellow_2"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param35", 5, 1, -1, MM_IIF(Request.Form("Yellow_3"), Request.Form("Yellow_3"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param36", 5, 1, -1, MM_IIF(Request.Form("Yellow_4"), Request.Form("Yellow_4"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param37", 5, 1, -1, MM_IIF(Request.Form("Yellow_5"), Request.Form("Yellow_5"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param38", 5, 1, -1, MM_IIF(Request.Form("Yellow_6"), Request.Form("Yellow_6"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param39", 5, 1, -1, MM_IIF(Request.Form("Yellow_7"), Request.Form("Yellow_7"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param40", 5, 1, -1, MM_IIF(Request.Form("Yellow_8"), Request.Form("Yellow_8"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param41", 5, 1, -1, MM_IIF(Request.Form("Yellow_9"), Request.Form("Yellow_9"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param42", 5, 1, -1, MM_IIF(Request.Form("Yellow_10"), Request.Form("Yellow_10"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param43", 5, 1, -1, MM_IIF(Request.Form("Yellow_11"), Request.Form("Yellow_11"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param44", 5, 1, -1, MM_IIF(Request.Form("Yellow_12"), Request.Form("Yellow_12"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param45", 5, 1, -1, MM_IIF(Request.Form("Yellow_13"), Request.Form("Yellow_13"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param46", 5, 1, -1, MM_IIF(Request.Form("Yellow_14"), Request.Form("Yellow_14"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param47", 5, 1, -1, MM_IIF(Request.Form("Yellow_15"), Request.Form("Yellow_15"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param48", 5, 1, -1, MM_IIF(Request.Form("Yellow_16"), Request.Form("Yellow_16"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param49", 5, 1, -1, MM_IIF(Request.Form("Yellow_17"), Request.Form("Yellow_17"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param50", 5, 1, -1, MM_IIF(Request.Form("Yellow_18"), Request.Form("Yellow_18"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param51", 5, 1, -1, MM_IIF(Request.Form("Yellow_19"), Request.Form("Yellow_19"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param52", 5, 1, -1, MM_IIF(Request.Form("Yellow_20"), Request.Form("Yellow_20"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param53", 5, 1, -1, MM_IIF(Request.Form("Yellow_21"), Request.Form("Yellow_21"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param54", 5, 1, -1, MM_IIF(Request.Form("Yellow_22"), Request.Form("Yellow_22"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param55", 5, 1, -1, MM_IIF(Request.Form("Yellow_23"), Request.Form("Yellow_23"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param56", 5, 1, -1, MM_IIF(Request.Form("Yellow_24"), Request.Form("Yellow_24"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param57", 5, 1, -1, MM_IIF(Request.Form("Yellow_25"), Request.Form("Yellow_25"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param58", 5, 1, -1, MM_IIF(Request.Form("LiveGreen"), Request.Form("LiveGreen"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param59", 5, 1, -1, MM_IIF(Request.Form("GreenRubbermaid"), Request.Form("GreenRubbermaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param60", 5, 1, -1, MM_IIF(Request.Form("Green_1"), Request.Form("Green_1"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param61", 5, 1, -1, MM_IIF(Request.Form("Green_2"), Request.Form("Green_2"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param62", 5, 1, -1, MM_IIF(Request.Form("Green_3"), Request.Form("Green_3"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param63", 5, 1, -1, MM_IIF(Request.Form("Green_4"), Request.Form("Green_4"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param64", 5, 1, -1, MM_IIF(Request.Form("Green_5"), Request.Form("Green_5"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param65", 5, 1, -1, MM_IIF(Request.Form("Green_6"), Request.Form("Green_6"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param66", 5, 1, -1, MM_IIF(Request.Form("Green_7"), Request.Form("Green_7"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param67", 5, 1, -1, MM_IIF(Request.Form("Green_8"), Request.Form("Green_8"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param68", 5, 1, -1, MM_IIF(Request.Form("Green_9"), Request.Form("Green_9"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param69", 5, 1, -1, MM_IIF(Request.Form("Green_10"), Request.Form("Green_10"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param70", 5, 1, -1, MM_IIF(Request.Form("Green_11"), Request.Form("Green_11"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param71", 5, 1, -1, MM_IIF(Request.Form("Green_12"), Request.Form("Green_12"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param72", 5, 1, -1, MM_IIF(Request.Form("Green_13"), Request.Form("Green_13"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param73", 5, 1, -1, MM_IIF(Request.Form("Green_14"), Request.Form("Green_14"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param74", 5, 1, -1, MM_IIF(Request.Form("Green_15"), Request.Form("Green_15"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param75", 5, 1, -1, MM_IIF(Request.Form("Green_16"), Request.Form("Green_16"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param76", 5, 1, -1, MM_IIF(Request.Form("Green_17"), Request.Form("Green_17"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param77", 5, 1, -1, MM_IIF(Request.Form("Green_18"), Request.Form("Green_18"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param78", 5, 1, -1, MM_IIF(Request.Form("Green_19"), Request.Form("Green_19"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param79", 5, 1, -1, MM_IIF(Request.Form("Green_20"), Request.Form("Green_20"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param80", 5, 1, -1, MM_IIF(Request.Form("Green_21"), Request.Form("Green_21"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param81", 5, 1, -1, MM_IIF(Request.Form("Green_22"), Request.Form("Green_22"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param82", 5, 1, -1, MM_IIF(Request.Form("Green_23"), Request.Form("Green_23"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param83", 5, 1, -1, MM_IIF(Request.Form("Green_24"), Request.Form("Green_24"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param84", 5, 1, -1, MM_IIF(Request.Form("Green_25"), Request.Form("Green_25"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param85", 5, 1, -1, MM_IIF(Request.Form("LiveBlue"), Request.Form("LiveBlue"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param86", 5, 1, -1, MM_IIF(Request.Form("BlueWtRubbermaid"), Request.Form("BlueWtRubbermaid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param87", 5, 1, -1, MM_IIF(Request.Form("Blue_1"), Request.Form("Blue_1"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param88", 5, 1, -1, MM_IIF(Request.Form("Blue_2"), Request.Form("Blue_2"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param89", 5, 1, -1, MM_IIF(Request.Form("Blue_3"), Request.Form("Blue_3"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param90", 5, 1, -1, MM_IIF(Request.Form("Blue_4"), Request.Form("Blue_4"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param91", 5, 1, -1, MM_IIF(Request.Form("Blue_5"), Request.Form("Blue_5"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param92", 5, 1, -1, MM_IIF(Request.Form("Blue_6"), Request.Form("Blue_6"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param93", 5, 1, -1, MM_IIF(Request.Form("Blue_7"), Request.Form("Blue_7"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param94", 5, 1, -1, MM_IIF(Request.Form("Blue_8"), Request.Form("Blue_8"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param95", 5, 1, -1, MM_IIF(Request.Form("Blue_9"), Request.Form("Blue_9"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param96", 5, 1, -1, MM_IIF(Request.Form("Blue_10"), Request.Form("Blue_10"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param97", 5, 1, -1, MM_IIF(Request.Form("Blue_11"), Request.Form("Blue_11"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param98", 5, 1, -1, MM_IIF(Request.Form("Blue12"), Request.Form("Blue12"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param99", 5, 1, -1, MM_IIF(Request.Form("Blue_13"), Request.Form("Blue_13"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param100", 5, 1, -1, MM_IIF(Request.Form("Blue_14"), Request.Form("Blue_14"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param101", 5, 1, -1, MM_IIF(Request.Form("Blue_15"), Request.Form("Blue_15"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param102", 5, 1, -1, MM_IIF(Request.Form("Blue_16"), Request.Form("Blue_16"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param103", 5, 1, -1, MM_IIF(Request.Form("Blue_17"), Request.Form("Blue_17"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param104", 5, 1, -1, MM_IIF(Request.Form("Blue_18"), Request.Form("Blue_18"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param105", 5, 1, -1, MM_IIF(Request.Form("Blue_19"), Request.Form("Blue_19"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param106", 5, 1, -1, MM_IIF(Request.Form("Blue_20"), Request.Form("Blue_20"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param107", 5, 1, -1, MM_IIF(Request.Form("Blue_21"), Request.Form("Blue_21"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param108", 5, 1, -1, MM_IIF(Request.Form("Blue_22"), Request.Form("Blue_22"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param109", 5, 1, -1, MM_IIF(Request.Form("Blue_23"), Request.Form("Blue_23"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param110", 5, 1, -1, MM_IIF(Request.Form("Blue_24"), Request.Form("Blue_24"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param111", 5, 1, -1, MM_IIF(Request.Form("Blue_25"), Request.Form("Blue_25"), null)) ' adDouble
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
Dim rsOysterMeasurements
Dim rsOysterMeasurements_cmd
Dim rsOysterMeasurements_numRows

Set rsOysterMeasurements_cmd = Server.CreateObject ("ADODB.Command")
rsOysterMeasurements_cmd.ActiveConnection = MM_oysters_STRING
rsOysterMeasurements_cmd.CommandText = "SELECT SiteID, ObservationDate, TimeSpentonObservation, LiveRed, RedWtRubbermaid, Red_1, Red_2, Red_3, Red_4, Red_5, Red_6, Red_7, Red_8, Red_9, Red_10, Red_11, Red_12, Red_13, Red_14, Red_15, Red_16, Red_17, Red_18, Red_19, Red_20, Red_21, Red_22, Red_23, Red_24, Red_25, LiveYellow, YellowWtRubbermaid, Yellow_1, Yellow_2, Yellow_3, Yellow_4, Yellow_5, Yellow_6, Yellow_7, Yellow_8, Yellow_9, Yellow_10, Yellow_11, Yellow_12, Yellow_13, Yellow_14, Yellow_15, Yellow_16, Yellow_17, Yellow_18, Yellow_19, Yellow_20, Yellow_21, Yellow_22, Yellow_23, Yellow_24, Yellow_25, LiveGreen, GreenRubbermaid, Green_1, Green_2, Green_3, Green_4, Green_5, Green_6, Green_7, Green_8, Green_9, Green_10, Green_11, Green_12, Green_13, Green_14, Green_15, Green_16, Green_17, Green_18, Green_19, Green_20, Green_21, Green_22, Green_23, Green_24, Green_25, LiveBlue, BlueWtRubbermaid, Blue_1, Blue_2, Blue_3, Blue_4, Blue_5, Blue_6, Blue_7, Blue_8, Blue_9, Blue_10, Blue_11, Blue12, Blue_13, Blue_14, Blue_15, Blue_16, Blue_17, Blue_18, Blue_19, Blue_20, Blue_21, Blue_22, Blue_23, Blue_24, Blue_25 FROM dbo.oystermeasurements" 
rsOysterMeasurements_cmd.Prepared = true

Set rsOysterMeasurements = rsOysterMeasurements_cmd.Execute
rsOysterMeasurements_numRows = 0
%>
<form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">ObservationDate:</td>
      <td><input type="text" name="ObservationDate" value=  ""  size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Minutes Spent on Observation:</td>
      <td><input type="text" name="TimeSpentonObservation" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Live Red:</td>
      <td><input type="text" name="LiveRed" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red Wt Rubbermaid:</td>
      <td><input type="text" name="RedWtRubbermaid" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_1:</td>
      <td><input type="text" name="Red_1" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_2:</td>
      <td><input type="text" name="Red_2" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_3:</td>
      <td><input type="text" name="Red_3" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_4:</td>
      <td><input type="text" name="Red_4" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_5:</td>
      <td><input type="text" name="Red_5" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_6:</td>
      <td><input type="text" name="Red_6" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_7:</td>
      <td><input type="text" name="Red_7" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_8:</td>
      <td><input type="text" name="Red_8" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_9:</td>
      <td><input type="text" name="Red_9" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_10:</td>
      <td><input type="text" name="Red_10" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_11:</td>
      <td><input type="text" name="Red_11" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_12:</td>
      <td><input type="text" name="Red_12" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_13:</td>
      <td><input type="text" name="Red_13" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_14:</td>
      <td><input type="text" name="Red_14" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_15:</td>
      <td><input type="text" name="Red_15" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_16:</td>
      <td><input type="text" name="Red_16" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_17:</td>
      <td><input type="text" name="Red_17" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_18:</td>
      <td><input type="text" name="Red_18" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_19:</td>
      <td><input type="text" name="Red_19" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_20:</td>
      <td><input type="text" name="Red_20" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_21:</td>
      <td><input type="text" name="Red_21" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_22:</td>
      <td><input type="text" name="Red_22" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_23:</td>
      <td><input type="text" name="Red_23" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_24:</td>
      <td><input type="text" name="Red_24" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Red_25:</td>
      <td><input type="text" name="Red_25" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Live Yellow:</td>
      <td><input type="text" name="LiveYellow" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow Wt Rubbermaid:</td>
      <td><input type="text" name="YellowWtRubbermaid" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_1:</td>
      <td><input type="text" name="Yellow_1" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_2:</td>
      <td><input type="text" name="Yellow_2" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_3:</td>
      <td><input type="text" name="Yellow_3" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_4:</td>
      <td><input type="text" name="Yellow_4" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_5:</td>
      <td><input type="text" name="Yellow_5" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_6:</td>
      <td><input type="text" name="Yellow_6" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_7:</td>
      <td><input type="text" name="Yellow_7" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_8:</td>
      <td><input type="text" name="Yellow_8" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_9:</td>
      <td><input type="text" name="Yellow_9" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_10:</td>
      <td><input type="text" name="Yellow_10" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_11:</td>
      <td><input type="text" name="Yellow_11" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_12:</td>
      <td><input type="text" name="Yellow_12" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_13:</td>
      <td><input type="text" name="Yellow_13" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_14:</td>
      <td><input type="text" name="Yellow_14" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_15:</td>
      <td><input type="text" name="Yellow_15" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_16:</td>
      <td><input type="text" name="Yellow_16" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_17:</td>
      <td><input type="text" name="Yellow_17" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_18:</td>
      <td><input type="text" name="Yellow_18" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_19:</td>
      <td><input type="text" name="Yellow_19" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_20:</td>
      <td><input type="text" name="Yellow_20" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_21:</td>
      <td><input type="text" name="Yellow_21" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_22:</td>
      <td><input type="text" name="Yellow_22" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_23:</td>
      <td><input type="text" name="Yellow_23" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_24:</td>
      <td><input type="text" name="Yellow_24" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Yellow_25:</td>
      <td><input type="text" name="Yellow_25" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">LiveGreen:</td>
      <td><input type="text" name="LiveGreen" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green Wt Rubbermaid:</td>
      <td><input type="text" name="GreenRubbermaid" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_1:</td>
      <td><input type="text" name="Green_1" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_2:</td>
      <td><input type="text" name="Green_2" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_3:</td>
      <td><input type="text" name="Green_3" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_4:</td>
      <td><input type="text" name="Green_4" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_5:</td>
      <td><input type="text" name="Green_5" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_6:</td>
      <td><input type="text" name="Green_6" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_7:</td>
      <td><input type="text" name="Green_7" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_8:</td>
      <td><input type="text" name="Green_8" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_9:</td>
      <td><input type="text" name="Green_9" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_10:</td>
      <td><input type="text" name="Green_10" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_11:</td>
      <td><input type="text" name="Green_11" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_12:</td>
      <td><input type="text" name="Green_12" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_13:</td>
      <td><input type="text" name="Green_13" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_14:</td>
      <td><input type="text" name="Green_14" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_15:</td>
      <td><input type="text" name="Green_15" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_16:</td>
      <td><input type="text" name="Green_16" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_17:</td>
      <td><input type="text" name="Green_17" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_18:</td>
      <td><input type="text" name="Green_18" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_19:</td>
      <td><input type="text" name="Green_19" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_20:</td>
      <td><input type="text" name="Green_20" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_21:</td>
      <td><input type="text" name="Green_21" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_22:</td>
      <td><input type="text" name="Green_22" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_23:</td>
      <td><input type="text" name="Green_23" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_24:</td>
      <td><input type="text" name="Green_24" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Green_25:</td>
      <td><input type="text" name="Green_25" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">LiveBlue:</td>
      <td><input type="text" name="LiveBlue" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue Wt Rubbermaid:</td>
      <td><input type="text" name="BlueWtRubbermaid" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_1:</td>
      <td><input type="text" name="Blue_1" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_2:</td>
      <td><input type="text" name="Blue_2" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_3:</td>
      <td><input type="text" name="Blue_3" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_4:</td>
      <td><input type="text" name="Blue_4" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_5:</td>
      <td><input type="text" name="Blue_5" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_6:</td>
      <td><input type="text" name="Blue_6" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_7:</td>
      <td><input type="text" name="Blue_7" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_8:</td>
      <td><input type="text" name="Blue_8" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_9:</td>
      <td><input type="text" name="Blue_9" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_10:</td>
      <td><input type="text" name="Blue_10" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_11:</td>
      <td><input type="text" name="Blue_11" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue12:</td>
      <td><input type="text" name="Blue12" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_13:</td>
      <td><input type="text" name="Blue_13" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_14:</td>
      <td><input type="text" name="Blue_14" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_15:</td>
      <td><input type="text" name="Blue_15" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_16:</td>
      <td><input type="text" name="Blue_16" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_17:</td>
      <td><input type="text" name="Blue_17" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_18:</td>
      <td><input type="text" name="Blue_18" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_19:</td>
      <td><input type="text" name="Blue_19" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_20:</td>
      <td><input type="text" name="Blue_20" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_21:</td>
      <td><input type="text" name="Blue_21" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_22:</td>
      <td><input type="text" name="Blue_22" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_23:</td>
      <td><input type="text" name="Blue_23" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_24:</td>
      <td><input type="text" name="Blue_24" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">Blue_25:</td>
      <td><input type="text" name="Blue_25" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td nowrap="nowrap" align="right">&nbsp;</td>
      <td><input type="submit" value="Insert record" /></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
<p>&nbsp;</p>
<%
rsOysterMeasurements.Close()
Set rsOysterMeasurements = Nothing
%>
