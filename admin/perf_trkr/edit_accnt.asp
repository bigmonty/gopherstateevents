<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

If Not Session("role") = "admin" Then Response.Redirect "/index.html"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

%>
<<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Edit GSE&copy; Performance Tracker Account</title>
</head>

<body>
<div class="container">
	<div class="row">
		<div class="col-sm-10">
		    <h4 class="h4">Edit GSE Performance Tracker Account</h4>
        </div>
	</div>
</div>
<!--#include file = "../../includes/footer_js.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
