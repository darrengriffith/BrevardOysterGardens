<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/oysters.asp" -->

<% @Import Namespace="System.Web.Script.Serialization" %>
<% @Import Namespace="System.Data" %>
<% @Import Namespace="System.Data.SqlClient" %>

<%
        Dim StartOfObservations = false
        Dim EndOfObservations = false
        Dim Error_Div = "none"
        Dim Error_Msg = "none"
        Dim Record_Div = "block"

	    Dim MyConnection As SqlConnection
	    Dim MyCommand As SqlCommand
	    Dim MyDataReader As SqlDataReader
	
	    Dim siteId As String
	siteId = Request.QueryString("siteId")
        Dim action As String
	action = Request.QueryString("action")
        Dim observationId As String
	observationId= Request.QueryString("observationId")
        Dim observationIdForPage As String
        observationIdForPage = Request.QueryString("observationId")

        Dim idHabitatObservations
        Dim lastEntryDate
        Dim wtBeforeCleaningRed
        Dim wtAfterCleaningRed
        Dim wtBeforeCleaningYellow
        Dim wtAfterCleaningYellow
        Dim wtBeforeCleaningGreen
        Dim wtAfterCleaningGreen
        Dim wtBeforeCleaningBlue
        Dim wtAfterCleaningBlue
        Dim Barnacle
        Dim BoringSponge
        Dim BlueCrab
        Dim GrassShrimp
        Dim HermitCrab 
        Dim LionFish
        Dim MudCrab
        Dim PinkShrimp
        Dim RibbedMussel
        Dim SeaSquirt
        Dim Sheepshead
        Dim SheepsheadMinnow
        Dim SlipperShell
        Dim SnappingShrimp
        Dim StoneCrab
        Dim OtherOrganisms
        Dim SpatOnRecruitmentShell
        Dim FoulingLoad
        Dim TimeSpentOnObservation
        Dim ObservationComment
	
        Dim Connection As New SqlClient.SqlConnectionStringBuilder

        If siteId Is Nothing Then
            Error_Div = "block"
            Record_Div = "none"
            Error_Msg = "ERROR NO SITE ID!  Log out and login again."
        Else
       
    
    Connection.DataSource = "184.168.194.53"
    Connection.InitialCatalog = "oysters"
    Connection.UserID = "daveo17"
    Connection.Password = "Reneea01*"
    Dim objSQlConnection = New SqlClient.SqlConnection(Connection.ConnectionString)
    
    MyConnection = New SqlConnection(Connection.ConnectionString)

    MyCommand = New SqlCommand("SELECT TOP 1 * FROM dbo.habitatobservations WHERE SiteID = " + siteId + " ORDER BY idHabitatObservations DESC;", MyConnection)
    If Not action Is Nothing Then
        If (action.equals("Previous")  AND IsNumeric(observationId))  Then
            MyCommand = New SqlCommand("SELECT TOP 1 * FROM dbo.habitatobservations WHERE SiteID = " + siteId + " AND idHabitatObservations < " + observationId +" ORDER BY idHabitatObservations DESC;", MyConnection)    
       ElseIf (action.equals("Next")  AND IsNumeric(observationId))  Then
             MyCommand = New SqlCommand("SELECT TOP 1 * FROM dbo.habitatobservations WHERE SiteID = " + siteId + " AND idHabitatObservations > " + observationId +" ORDER BY idHabitatObservations ASC;", MyConnection) 
        ElseIf (action.equals("Delete")  AND IsNumeric(observationId))  Then 
             Dim results As String  
             results = DeleteObservation(observationId)
             If results.equals("Success.") Then
                 ' Show previous record
                 MyCommand = New SqlCommand("SELECT TOP 1 * FROM dbo.habitatobservations WHERE SiteID = " + siteId + " AND idHabitatObservations < " + observationId +" ORDER BY idHabitatObservations DESC;", MyConnection)  
             Else
                 Error_Div = "block"
                 Error_Msg = results 
                 Record_Div = "none" 
             End If
        End If
    End If
   
    With MyCommand
            .CommandType = CommandType.Text

            .Connection.Open()

            MyDataReader = .ExecuteReader()

            If MyDataReader.HasRows Then

	        While MyDataReader.Read()

                idHabitatObservations = MyDataReader.item("idHabitatObservations").tostring()
                observationIdForPage =  MyDataReader.item("idHabitatObservations").tostring()


                lastEntryDate = Convert.ToDateTime(MyDataReader.item("DateOfObservation").tostring()).ToString("MM/dd/yyyy")
            
                wtBeforeCleaningRed = MyDataReader.item("wtBeforeCleaningRed").tostring()
                wtAfterCleaningRed = MyDataReader.item("wtAfterCleaningRed").tostring()
            
                wtBeforeCleaningYellow = MyDataReader.item("wtBeforeCleaningYellow").tostring()
                wtAfterCleaningYellow = MyDataReader.item("wtAfterCleaningYellow").tostring()
            
                wtBeforeCleaningGreen = MyDataReader.item("wtBeforeCleaningGreen").tostring()
                wtAfterCleaningGreen = MyDataReader.item("wtAfterCleaningGreen").tostring()
            
                wtBeforeCleaningBlue = MyDataReader.item("wtBeforeCleaningBlue").tostring()
                wtAfterCleaningBlue = MyDataReader.item("wtAfterCleaningBlue").tostring()
            
                Barnacle = MyDataReader.item("Barnacle").tostring()
                BoringSponge = MyDataReader.item("BoringSponge").tostring()
                BlueCrab = MyDataReader.item("BlueCrab").tostring()
                GrassShrimp = MyDataReader.item("GrassShrimp").tostring()
                HermitCrab = MyDataReader.item("HermitCrab").tostring()
                LionFish = MyDataReader.item("LionFish").tostring()
                MudCrab = MyDataReader.item("MudCrab").tostring()
                PinkShrimp = MyDataReader.item("PinkShrimp").tostring()
                RibbedMussel = MyDataReader.item("RibbedMussel").tostring()
                SeaSquirt = MyDataReader.item("SeaSquirt").tostring()
                Sheepshead = MyDataReader.item("Sheepshead").tostring()
                SheepsheadMinnow = MyDataReader.item("SheepsheadMinnow").tostring()
                SlipperShell = MyDataReader.item("SlipperShell").tostring()
                SnappingShrimp = MyDataReader.item("SnappingShrimp").tostring()
                StoneCrab = MyDataReader.item("StoneCrab").tostring()
	
                OtherOrganisms = MyDataReader.item("OtherOrganisms").tostring()
                SpatOnRecruitmentShell = MyDataReader.item("SpatOnRecruitmentShell").tostring()
                FoulingLoad = MyDataReader.item("FoulingLoad").tostring()
                TimeSpentOnObservation = MyDataReader.item("TimeSpentOnObservation").tostring()
                ObservationComment = MyDataReader.item("ObservationComment").tostring()	
            End While 
            ' show record.
          
        Else
            ' No data returned.
            Error_Div = "block"
            If action Is Nothing Then
                Error_Msg= "No Observations."
                StartOfObservations = true
                EndOfObservations = true
            Else If (action.equals("Previous") OR (action.equals("Delete") AND Error_Msg.equals("none"))) Then
                Error_Msg= "Start of observations.  Press 'Previous' or 'Next' buttons to continue from the latest record."
                StartOfObservations = true
            Else If (action.equals("Next")) Then
                Error_Msg = "End of observations.  Press 'Previous' or 'Next' buttons to continue from the latest record."
                EndOfObservations = true
            End If
            Record_Div = "none"
            
        End If

        .Dispose()
        MyConnection.Close()
    End With
 End If
 %>

<script runat="server" language="vb">

    Private Function GetConnection()
        Dim connection As New SqlClient.SqlConnectionStringBuilder
        connection.DataSource = "184.168.194.53"
        connection.InitialCatalog = "oysters"
        connection.UserID = "daveo17"
        connection.Password = "Reneea01*"
        Dim objSQlConnection = New SqlClient.SqlConnection(connection.ConnectionString)

        GetConnection = New SqlConnection(connection.ConnectionString)
    End Function

    Friend Function DeleteObservation(ByVal observationId As String)
         Dim cmd As SqlCommand
         Dim connection As SqlConnection
         Dim dataReader As SqlDataReader
         Dim strBuilder As New StringBuilder()
         Dim rowCount As Integer
        
         Try
             connection = GetConnection()
             ' Copy record into the deleted observations table.
             cmd = New SqlCommand("INSERT INTO dbo.habitatobservations_delrecs SELECT * FROM dbo.habitatobservations WHERE idHabitatObservations = " & observationId & ";", connection) 
             With cmd
                 .CommandType = CommandType.Text
                 .Prepare()
                 .Connection.Open()
                 .ExecuteNonQuery()
                  DeleteObservation = ("Success.")
                  .Dispose()             
             End With
         
             ' Remove record from the observations table.
             cmd = New SqlCommand("DELETE FROM dbo.habitatobservations WHERE idHabitatObservations = " & observationId & ";", connection) 
             With cmd
                 .CommandType = CommandType.Text
                 .Prepare()
                 .ExecuteNonQuery()
                 DeleteObservation = ("Success.")
                 .Dispose()
            End With
       Catch e As SQLException
           DeleteObservation = "Error: " & e.Message
       Finally
         connection.Close()
       End Try
     End Function   
</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
    <head runat="server">
        <title>Site Metrics</title>
    </head>
    <body>
       <span id="observationIdForPage" style="display: none"><%=observationIdForPage%></span>
       <span id="StartOfObservation" style="display: none"><%=StartOfObservations%></span>
       <span id="EndOfObservation" style="display: none"><%=EndOfObservations%></span>
       <div id="Error" style="display: <%=Error_Div%>">
           <%=Error_Msg%>    
       </div>
       <div id="Record" style="display: <%=Record_Div%>">
       		<table>
			<tr>
				<td><strong>Observation Date</strong></td>
				<td><%=lastEntryDate%></td>
				<td><strong>Site ID</strong></td>
				<td><%=siteId%></td>
                                <td><strong>Observation ID</strong></td>
				<td><span id="idHabitatObservations"><%=idHabitatObservations%></span></td>

			</tr>
			<tr><td><br></td></tr>
			<tr>
				<td></td>
				<td><strong>Red</strong></td>
				<td><strong>Yellow</strong></td>
				<td><strong>Green</strong></td>
				<td><strong>Blue</strong></td>
			</tr>
			<tr>
				<td><strong>Weight Before Cleaning<strong></td>
				<td><%=wtBeforeCleaningRed%></td>
				<td><%=wtBeforeCleaningYellow%></td>
				<td><%=wtBeforeCleaningGreen%></td>
				<td><%=wtBeforeCleaningBlue%></td>
			</tr>
			<tr>
				<td><strong>Weight After Cleaning</strong></td>
				<td><%=wtAfterCleaningRed%></td>
				<td><%=wtAfterCleaningYellow%></td>
				<td><%=wtAfterCleaningGreen%></td>
				<td><%=wtAfterCleaningBlue%></td>
			</tr>
			<tr><td><br></td></tr>
			<tr>
				<td><strong>Organisms</strong></td>
				<td><strong>Barnacle</strong></td>
				<td><%=Barnacle%></td>
            	<td><strong>Boring Sponge</strong></td>
				<td><%=BoringSponge%></td>
			</tr>
			<tr>
				<td></td>
            	<td><strong>Blue Crab</strong></td>
				<td><%=BlueCrab%></td>
            	<td><strong>Grass Shrimp</strong></td>
				<td><%=GrassShrimp%></td>
			</tr>
			<tr>
				<td></td>
            	<td><strong>Hermit Crab</strong></td>
				<td><%=HermitCrab%></td>
                <td><strong>Lion Fish</strong></td>
				<td><%=LionFish%></td>
                        </tr>
                        <tr>
                               <td></td>
            	<td><strong>Mud Crab</strong></td>
				<td><%=MudCrab%></td>
         	<td><strong>Pink Shrimp</strong></td>
				<td><%=PinkShrimp%></td>
                       </tr>
			<tr>
				<td></td>
            	<td><strong>Ribbed Mussel</strong></td>
				<td><%=RibbedMussel%></td>
	       	<td><strong>Sea Squirt</strong></td>
				<td><%=SeaSquirt%></td>
	                </tr>
			<tr>
				<td></td>
            	<td><strong>Sheepshead</strong></td>
				<td><%=Sheepshead%></td>
           	<td><strong>Sheepshead Minnow</strong></td>
				<td><%=SheepsheadMinnow%></td>
                       	</tr>
			<tr>
				<td></td>
            	<td><strong>Slipper Shell</strong></td>
				<td><%=SlipperShell%></td>
          	<td><strong>Snapping Shrimp</strong></td>
				<td><%=SnappingShrimp%></td>
                	</tr>
			<tr>
				<td></td>
            	<td><strong>Stone Crab</strong></td>
				<td><%=StoneCrab%></td>
           	<td><strong>Other</strong></td>
				<td colspan="3"><%=OtherOrganisms%></td>
			<tr><td><br></td></tr>
			<tr>
            	<td><strong>Spat On Recruitment Shell</strong></td>
				<td colspan="3"><%=SpatOnRecruitmentShell%></td>
			</tr>
			<tr>
            	<td><strong>Fouling Load</strong></td>
				<td colspan="3"><%=FoulingLoad%></td>
			</tr>
			<tr>
            	<td><strong>Time Spent On Observation</strong></td>
				<td colspan="3"><%=TimeSpentOnObservation%></td>
			</tr>
			<tr>
            	<td><strong>Observation Comment</strong></td>
				<td colspan="3"><%=ObservationComment%></td>
			</tr>
		</table>	
             </div>      	
         </body>
</html>
