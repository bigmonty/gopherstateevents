<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim sSeriesName, sInfoSheet

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lSeriesID = Request.QueryString("series_id")
Session("series_id") = lSeriesID

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get series information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesName FROM Series WHERE SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
sSeriesName = rs(0).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT InfoSheet FROM Series WHERE SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sInfoSheet = rs(0).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sSeriesName%> Info Sheet Upload</title>
<meta name="description" content="Gopher State Events Series Logo page.">
<meta name="description" content="GSE race series for road races, nordic ski, showshoe events, mountain bike, duathlon, and cross-country meet management (timing).">
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div style="padding: 10px;margin: 10px;background-color: #fff;">
	<h3><%=sSeriesName%> Info Sheet Upload</h3>
		
	<div style="margin-top:10px;">
        <p>NOTE:  If there is already an info sheet uploaded for this series, this will over-write it.</p>

	    <form name="upload" method="Post" action="receive_info_sheet.asp" enctype="multipart/form-data">
	    <input type="file" name="file1" id="file1" size="50">
	    <br>
	    <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
	    <input type="submit" id="submit1" name="submit1" value="Upload!">
	    </form>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>