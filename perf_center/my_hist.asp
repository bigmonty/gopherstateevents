<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i

'Response.Redirect "/misc/taking_break.htm"

If CStr(Session("my_hist_id")) = vbNullString Then Response.Redirect "my_hist_login.asp"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"
	
Dim sRandPic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, PixName FROM RacePix ORDER BY NEWID()"
rs.Open sql, conn, 1, 2
sRandPic = "/gallery/" & rs(0).Value & "/" & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE&copy; History</title>
<meta name="description" content="My participant history for a Gopher State Events (GSE) timed event.">
<!--#include file = "../includes/js.asp" --> 
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My GSE History Portal">
    <h3 class="h3">My GSE History</h3>

    <div class="col-sm-6">
        &nbsp;
    </div>
	<div class="col-sm-6">
        <br>
		<a href="http://www.my-etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		    <img src="/graphics/my-etraxc_ad.gif" alt="My-eTRaXC" class="img-responsive">
        </a>

		<a href="<%=sRandPic%>" onclick="openThis2(this.href,1024,768);return false;">
            <img src="<%=sRandPic%>" alt="<%=sRandPic%>" class="img-responsive">
        </a>
	</div>
</div>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
