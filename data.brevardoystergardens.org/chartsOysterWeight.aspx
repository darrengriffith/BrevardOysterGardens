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
	
	Dim foulingRed as String
	Dim foulingYellow as String
	Dim foulingGreen as String
	Dim foulingBlue as String
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

		foulingRed = "["
		foulingYellow = "["
		foulingGreen = "["
		foulingBlue = "["
		categories = "["
	    While MyDataReader.Read()
			If Not firstTime
				foulingRed += ","
			End If

			if MyDataReader.IsDBNull(MyDataReader.GetOrdinal("WtAfterCleaningRed"))	
				foulingRed += "0"
			Else				
				foulingRed += MyDataReader.item("WtAfterCleaningRed").tostring()
			End If
	

			If Not firstTime
				foulingYellow += ","
			End If

			if MyDataReader.IsDBNull(MyDataReader.GetOrdinal("WtAfterCleaningYellow"))
				foulingYellow += "0"
			Else				
				foulingYellow += MyDataReader.item("WtAfterCleaningYellow").tostring()
			End If

	
			If Not firstTime
				foulingGreen += ","
			End If

			if MyDataReader.IsDBNull(MyDataReader.GetOrdinal("WtAfterCleaningGreen"))
				foulingGreen += "0"
			Else				
				foulingGreen += MyDataReader.item("WtAfterCleaningGreen").tostring()
			End If


			If Not firstTime
				foulingBlue += ","
			End If

			if MyDataReader.IsDBNull(MyDataReader.GetOrdinal("WtAfterCleaningBlue"))
				foulingBlue += "0"
			Else				
				foulingBlue += MyDataReader.item("WtAfterCleaningBlue").tostring()
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
		foulingRed += "]"
		foulingYellow += "]"
		foulingGreen += "]"
		foulingBlue += "]"
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
                            type: 'line',
							borderRadius: 7,
							backgroundColor: '#EBFAEB'
                        },
						legend: {
							enabled: false
						},
						colors: ['#FF0000', '#E6E600', '#33CC33', '#0066FF'],
                        title: {
                            text: 'Oyster Weight'
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
							{data: <%=foulingRed%>},
							{data: <%=foulingYellow%>},
							{data: <%=foulingGreen%>},
							{data: <%=foulingBlue%>}
						],
                    });
                });       
            //]]>              
        </script>
    </body>
</html>