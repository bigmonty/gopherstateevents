<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim sSeriesName, sLogo

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lSeriesID = Request.QueryString("series_id")
Session("series_id") = lSeriesID

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get series information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesName, Logo FROM Series WHERE SeriesID = " & lSeriesID
rs.Open sql, conn, 1, 2
sSeriesName = rs(0).Value
sLogo = rs(1).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title><%=sSeriesName%> Logo Admin</title>
<meta name="description" content="Gopher State Events Series Logo page.">

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div style="padding: 10px;margin: 10px;background-color: #fff;">
	<h3><%=sSeriesName%> Logo Management</h3>
		
	<div style="margin-top:10px;">
        <%If Not sLogo & "" = "" Then%>
            <img src="/admin/series/logos/<%=sLogo%>" style="float: right;width: 150px;">
        <%End If%>

	    <form name="upload" method="Post" action="receive_logo.asp" enctype="multipart/form-data">
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