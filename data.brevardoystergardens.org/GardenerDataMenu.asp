<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
	Dim siteId 
	siteId = Request.QueryString("siteId")
	If siteId = Null Or siteId = "" Then
	    siteId = Session("MM_Username")
	End If
%>
	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Brevard Oyster Garden Main</title>
<link href="GardenerDataMenu.css" rel="stylesheet" type="text/css" />
</head>

<body class="body">
	<div class="header">
		<h1>Brevard  Oyster Gardeners</h1>
	</div>
  
	<div>
		<div class="container" style="width: 400px; height: 100%; float: left">
			<div>
				<h2>Menu</h2>
				<ol>
					<strong>Gardeners</strong>
					<li><a href="GardenerProfileEdit.asp">Maintain Profile</a></li>
					<li><a href="HabitatUpdate.asp">Maintain Habitat Information</a></li>
					<li><a href="Observation_Insert.asp">Enter Habitat Observations</a></li>
					<br><strong>Master Gardeners</strong>
					<li><a href="OysterMeasurements_Insert.asp">Enter Oyster Measurements</a></li>
					<br><strong>Admin</strong>
					<li><a href="People_Insert.asp">Enter New Gardener</a></li>
					<li><a href="People_Update_Admin_List.asp">Maintain Gardener Profiles</a></li>
					<li>Maintain Habitat Information</li>
					<li><a href="/ManageAccess.asp">Maintain Access Control</a></li>
					<li>Reports</li>
				</ol>
			</div>
	  
			<div>
				<h3>Edit Gardener and Site Information</h3>
				<p>Oyster gardeners have access to their profile data and can enter details about the habitats on the site, enter weights and observations of their habitats.</p>
				<p>Master Gardeners can also enter their measurements of oysters.</p>
				<p>All gardeners should complete the details about their habitats as soon as convenient. Gardener profiles and habitat information can be updated at any time.</p>
				<p>Observation and measurement sets cannot be edited after they are saved.</p>
				<p>Please click the Update or Insert button at the bottom of screens to save your data.</p>
				<strong>Do not enter units with numbers - just the numbers</strong>
				<h3>Please enter data as:</h3>
				<ul>
					<li>Telephone numbers - with area code as: 321.555.1212</li>
					<li>Distances and depth - in feet</li>
					<li>Time in minutes</li>
					<li>Dates - mm/dd/yyyy as 04/14/2014 for April 14, 2014</li>
					<li>Measurements - in millimeters</li>
				</ul>
				<br>
				<p>If an error on a page occurs, use the back button on your browser to return to last page.</p>
			</div>
		
			<div id="footer">
				<p><a href="http://brevardoystergardens.org">Cancel</a></p>
			</div>
		</div>
		
		<div class="container" style="float: right; display: inline-block; width: calc(100% - 430px); border: none; overflow: hidden">
			<iframe src="gardenerOverview.aspx?siteId=<%=siteId%>" style="height: 450px; width: calc(100% - 5px); border: none"></iframe>
		</div>
		<div class="container" style="float: right; display: inline-block; width: calc(100% - 430px); border: none; overflow: hidden">
			<iframe src="chartsOysterWeight.aspx?siteId=<%=siteId%>" style="height: 450px; width: calc(100% - 5px); border: none"></iframe>
		</div>
		<div class="container" style="float: right; display: inline-block; width: calc(100% - 430px); border: none; overflow: hidden">
			<iframe src="chartsFoulingLoad.aspx?siteId=<%=siteId%>" style="height: 450px; width: calc(100% - 5px); border: none"></iframe>
		</div>
		<div class="container" style="float: right; display: inline-block; width: calc(100% - 430px); border: none; overflow: hidden">
			<iframe src="chartsTimeSpent.aspx?siteId=<%=siteId%>" style="height: 450px; width: calc(100% - 5px); border: none"></iframe>
		</div>
	</div>
</body>
</html>
