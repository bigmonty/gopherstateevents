<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sInfoSheet
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
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT InfoSheet FROM InfoSheet WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then sInfoSheet = rs(0).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sEventName%> Info Sheet Upload</title>
<!--#include file = "../../includes/meta2.asp" -->



</head>

<body>
<div style="padding: 10px;margin: 10px;background-color: #fff;">
	<h3 class="h3"><%=sEventName%> Info Sheet Uploaded</h3>
		
	<div style="margin-top:10px;">
        <p>NOTE:  If there is already an info sheet uploaded for this event, this will over-write it.</p>

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