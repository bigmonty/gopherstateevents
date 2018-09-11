<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "login.asp"


Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross=Country-Nordic Ski Performance Tracker Graphs</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <div class="bg-danger text-danger">
        This utility is currently under construction.  We appreciate your patience! Check back often and 
        monitor it's development.
    </div>
</div>
</body>
<%
conn.close
Set conn = Nothing

conn2.close
Set conn2 = Nothing
%>
</html>
