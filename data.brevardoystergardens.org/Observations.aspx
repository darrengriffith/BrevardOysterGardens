<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%@ Import Namespace="System.Web.Script.Serialization" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Text.StringBuilder" %>
<%@ Import Namespace="System.Environment.NewLine" %>
<%
    Dim siteId As String
    Dim observationId As String
    Dim json As String
    Dim functionRequest As String
       
    siteId = Request.QueryString("siteId")

    observationId= Request.QueryString("observationId")

    functionRequest = Request.QueryString("Option")


    Select Case functionRequest

        Case "GetObservations"
            json = GetObservations(siteId)
        Case "GetAllObservations"
            json = GetAllObservations()
        Case "DeleteObservation"
            json = DeleteObservation(observationId)
        Case Else
            json = "Oops!  That was a bad request."

    End Select
    Response.Clear()    
    Response.ContentType = "application/json; charset=utf-8"
    Response.Write(json)
    Response.End
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
   
    Public Function GetObservations(ByVal siteId As String)
        Dim cmd As SqlCommand
        Dim connection As SqlConnection
        Dim dataReader As SqlDataReader
        Dim strBuilder As New StringBuilder()
        Dim rowCount As Integer
        
        connection = GetConnection()

        cmd = New SqlCommand("SELECT TOP 100" &
                             "idHabitatObservations, SiteId, DateOfObservation, WtBeforeCleaningRed, WtAfterCleaningRed, WtBeforeCleaningYellow, WtAfterCleaningYellow, " &
                             "WtBeforeCleaningGreen, WtAfterCleaningGreen, WtBeforeCleaningBlue, WtAfterCleaningBlue, Barnacle, BoringSponge, BlueCrab, " &
                             "GrassShrimp, HermitCrab, LionFish, MudCrab, PinkShrimp, RibbedMussel, SeaSquirt, Sheepshead, SheepsheadMinnow, SlipperShell, " &
                             "SnappingShrimp, StoneCrab, OtherOrganisms, SpatOnRecruitmentShell, FoulingLoad, TimeSpentOnObservation, ObservationComment " &
                             "FROM dbo.habitatobservations WHERE SiteID = " & siteId &" ORDER BY DateOfObservation DESC;", connection)


        With cmd
            .CommandType = CommandType.Text

            .Connection.Open()

            dataReader = .ExecuteReader()
            'note that escape char. for quotes is another set of quotes
            strBuilder.Append("{""Observations"": [" & Environment.NewLine)
            rowCount = 0

            While dataReader.Read()
                 ' Some field values contain characters that must be escaped to be valid JSON
                 Dim ser As New JavaScriptSerializer()                 
                 ' Max size of JSON output
                 ser.MaxJsonLength = Int32.MaxValue   

                 strBuilder.Append("{")
                 ' loop through the fields and build out the JSON 
                 For i = 0 To (dataReader.FieldCount -1) Step 1
                     strBuilder.Append("""")
                     strBuilder.Append(dataReader.GetName(i) & """:")
                     strBuilder.Append(ser.Serialize(dataReader.Item(i).ToString()))
                     strBuilder.Append(",")
                 Next
                 ' remove trailing comma on last item at this level.
                 If dataReader.FieldCount > 0
                     strBuilder.Remove(strBuilder.Length -1, 1)
                 End If
                 strBuilder.Append("},")
                 strBuilder.Append(Environment.NewLine)
                 rowCount += 1
            End While
            ' remove trailing comma on last item at this level.
            If rowCount > 0
                strBuilder.Remove(strBuilder.Length -3, 1)
            End If
            strBuilder.Append("]}")
            .Dispose()
            connection.Close()
        End With

        GetObservations = strBuilder.ToString()
    End Function

     Public Function GetAllObservations()
        Dim cmd As SqlCommand
        Dim connection As SqlConnection
        Dim dataReader As SqlDataReader
        Dim strBuilder As New StringBuilder()
        Dim rowCount As Integer
        
        connection = GetConnection()

        cmd = New SqlCommand("SELECT " &
                             "idHabitatObservations, SiteId, DateOfObservation, WtBeforeCleaningRed, WtAfterCleaningRed, WtBeforeCleaningYellow, WtAfterCleaningYellow, " &
                             "WtBeforeCleaningGreen, WtAfterCleaningGreen, WtBeforeCleaningBlue, WtAfterCleaningBlue, Barnacle, BoringSponge, BlueCrab, " &
                             "GrassShrimp, HermitCrab, LionFish, MudCrab, PinkShrimp, RibbedMussel, SeaSquirt, Sheepshead, SheepsheadMinnow, SlipperShell, " &
                             "SnappingShrimp, StoneCrab, OtherOrganisms, SpatOnRecruitmentShell, FoulingLoad, TimeSpentOnObservation, ObservationComment " &
                             "FROM dbo.habitatobservations ORDER BY SiteId, DateOfObservation ASC;", connection)


        With cmd
            .CommandType = CommandType.Text

            .Connection.Open()

            dataReader = .ExecuteReader()
            'note that escape char. for quotes is another set of quotes
            strBuilder.Append("{""Observations"": [" & Environment.NewLine)
            rowCount = 0

            While dataReader.Read()

                ' Some field values contain characters that must be escaped to be valid JSON
                 Dim ser As New JavaScriptSerializer()                 
                 ' Max size of JSON output
                 ser.MaxJsonLength = Int32.MaxValue   

                strBuilder.Append("{")
                ' loop through the fields and build out the JSON 
                For i = 0 To (dataReader.FieldCount -1) Step 1
                    strBuilder.Append("""")
                    strBuilder.Append(dataReader.GetName(i) & """:")
                    strBuilder.Append(ser.Serialize(dataReader.Item(i).ToString()))
                    strBuilder.Append(",")
                Next
                ' remove trailing comma on last item at this level.
                If dataReader.FieldCount > 0
                    strBuilder.Remove(strBuilder.Length -1, 1)
                End If
                strBuilder.Append("},")
                strBuilder.Append(Environment.NewLine)
                rowCount += 1
            End While
            ' remove trailing comma on last item at this level.
            If rowCount > 0
                strBuilder.Remove(strBuilder.Length -3, 1)
            End If
            strBuilder.Append("]}")
            .Dispose()
            connection.Close()
        End With

        GetAllObservations = strBuilder.ToString()
    End Function

    Public Function CreateObservation(ByVal siteId As String, ByVal observation As Object)
        Dim cmd As SqlCommand
        Dim connection As SqlConnection
        Dim dataReader As SqlDataReader
        Dim strBuilder As New StringBuilder()
        Dim rowCount As Integer
        
        connection = GetConnection()

        cmd = New SqlCommand("INSERT INTO dbo.habitatobservations (SiteID, DateOfObservation, WtBeforeCleaningRed, WtAfterCleaningRed, WtBeforeCleaningYellow, " &
                                                  "WtAfterCleaningYellow, WtBeforeCleaningGreen, WtAfterCleaningGreen, WtBeforeCleaningBlue, WtAfterCleaningBlue, Barnacle, BoringSponge,  " &
                                                  "BlueCrab, GrassShrimp, HermitCrab, LionFish, MudCrab, PinkShrimp, RibbedMussel, SeaSquirt, Sheepshead, SheepsheadMinnow, SlipperShell,  " &
                                                  "SnappingShrimp, StoneCrab, OtherOrganisms, SpatOnRecruitmentShell, FoulingLoad, TimeSpentOnObservation, ObservationComment)  " &
                                                  "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", connection) 
        cmd.Prepare()
        cmd.ExecuteNonQuery()
    End Function
</script>