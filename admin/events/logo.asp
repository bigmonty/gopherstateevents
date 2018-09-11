<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sLogo
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
Session("event_id") = lEventID

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
sLogo = rs(2).Value
rs.Close
Set rs = Nothing
%>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Logo Admin</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div style="padding: 10px;margin: 10px;background-color: #fff;">
	<h3 class="h3"><%=sEventName%> Logo Management</h3>
		
	<div style="margin-top:10px;">
        <%If Not sLogo & "" = "" Then%>
            <img src="/events/logos/<%=sLogo%>" style="float: right;width: 150px;">
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