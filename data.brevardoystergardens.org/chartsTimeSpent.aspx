<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/oysters.asp" -->

<% @Import Namespace="System.Web.Script.Serialization" %>
<% @Import Namespace="System.Data" %>
<% @Import Namespace="System.Data.SqlClient" %>

<%
	Dim MyConnection As SqlConnection
	Dim MyCommand As SqlCommand
	Dim MyDataReader As SqlDataReader
	
	Dim siteId As String
	siteId = Request.QueryString("siteId")
	
	Dim timeSpent as String
	Dim item as String
	Dim categories as String
	
	Dim firstTime as Boolean
	firstTime = true
	
	Dim Connection As New SqlClient.SqlConnectionStringBuilder
    Connection.DataSource = "184.168.194.53"
	Connection.InitialCatalog = "oysters"
    Connection.UserID = "daveo17"
    Connection.Password = "Reneea01*"
    Dim objSQlConnection = New SqlClient.SqlConnection(Connection.ConnectionString)
	
	MyConnection = New SqlConnection(Connection.ConnectionString)
	MyCommand = New SqlCommand("SELECT TOP 100 * FROM dbo.habitatobservations WHERE SiteID = " + siteId + " ORDER BY DateOfObservation ASC;", MyConnection)
	
    With MyCommand
        .CommandType = CommandType.Text

        .Connection.Open()

        MyDataReader = .ExecuteReader()

		timeSpent = "["
		categories = "["
	    While MyDataReader.Read()
			If Not firstTime
				timeSpent += ","
			End If

			if MyDataReader.IsDBNull(MyDataReader.GetOrdinal("TimeSpentOnObservation"))	
				timeSpent += "0"
			Else				
				timeSpent += MyDataReader.item("TimeSpentOnObservation").tostring()
			End If
	
	
			item = Convert.ToDateTime(MyDataReader.item("DateOfObservation").tostring()).ToString("MM/dd/yyyy")
		
			If Not firstTime
				categories += ","
			End If
	
			If Not String.IsNullOrEmpty(item)
				categories += "'"
				categories += item
				categories += "'"
			End If
	
			firstTime = false
	
        End While 
		timeSpent += "]"
		categories += "]"

        .Dispose()
        MyConnection.Close()
    End With
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
    <head runat="server">
        <title>Site Metrics</title>
    </head>
    <body>
		<div id="container" style="height: 100%; width: 100%"></div>
        
        <script type='text/javascript' src="./js/jquery-1.11.1.min.js"></script>
        <script type='text/javascript' src="./js/highcharts.js"></script>
        
        <script type='text/javascript'> //<![CDATA[
                $(document).ready(function () {
                    $('#container').highcharts({
                        chart: {
                            type: 'column',
							borderRadius: 7,
							backgroundColor: '#EBFAEB',
							inverted: false
                        },
						legend: {
							enabled: false
						},
						colors: ['#FF0000', '#E6E600', '#33CC33', '#0066FF'],
                        title: {
                            text: 'Time Spent'
                        },
                        xAxis: {
                            categories: <%=categories%>,
            				tickInterval: 5							
                        },
						yAxis: {
							title: {
								enabled: false
							}
						},
                        series: [
							{data: <%=timeSpent%>},
						],
                    });
                });       
            //]]>              
        </script>
    </body>
</html>