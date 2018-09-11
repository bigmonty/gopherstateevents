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
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Gopher State Events Sponsors</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{
<style type="text/css">
<!--
p, input{
	font-size:0.85em;
	}
-->
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    

			<h4 class="h4">Gopher State Events Sponsors</h4>
  		</div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>