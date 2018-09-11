<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim Series()

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

conn.close
Set conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html lang="en">
<head>
	<title>GSE CCMeet Series</title>
	<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<title>GSE Cross-Country Meet Director Home</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta name="keywords" content="running, nordic skiing, cross-country, mountain biking, road races, snowshoe, race, timing, ">
	<meta name="description" content="A Fitness Event Timing Service specializing in road racing, nordic ski events, showshoe events, mountain bike events, and high school and college cross-country meet timing.">
	<meta name="postinfo" content="/scripts/postinfo.asp">
	<meta name="resource" content="document">
	<meta name="distribution" content="global">
	
	<script type="text/javascript" src="../../misc/vira.js"></script>
	<link rel="stylesheet" type="text/css" href="../../misc/vira.css">
</head>

<body>
<table style="background-color;#fff;">
	<tr>
		<td>
			<h3>GSE Cross-Country Running/Nordic Sking Series</h3>
		</td>
	</tr>
	<tr>
		<td>
			This page is currently under construction.  We appreciate your patience!
		</td>
	</tr>
</table>
</body>
</html>
