<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lMeetID
Dim sFileSent

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sFileSent = Request.QueryString("file_sent")
lMeetID = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If sFileSent = "n" Then
	Session("this_meet") = lMeetID
Else
	Session("this_meet") = vbNullString
	Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>CCMeet Upload Run Order</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
	<h4 class="h4">Upload Run Order</h4>
	<form class="form" name="upload" method="Post" action="receive_run_order.asp" enctype="multipart/form-data">
	<input class="form-control" type="FILE" name="File1" id="File1" size="50">
	<br>
	<input class="form-control" type="submit" id="submit_1" name="submit_1" value="Upload!">
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
