<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Sample Results Email</title>
<meta name="description" content="Sample GSE (Gopher State Events) results email, sent to participants within minutes of finishing.">
<!--#include file = "../includes/js.asp" --> 
</head>

<body>
<div>
    <h4 class="h4">Sample Gopher State Events Email</h4>
    
	<p>This is a sample of the email that individual participants receive after finishing their race, usually within a couple of minutes.</p>
    <img src="/graphics/my_results.png" alt="Individual Results Email">
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>