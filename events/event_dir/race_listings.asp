<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sLogo
Dim dEventDate

If Not Session("role") = "event_dir" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
Session("event_id") = lEventID

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Race Listing Sites</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive">
        <h3><%=sEventName%> Race Listing Sites</h3>
                
        <ul class="list-group">
                <li class="list-group-item"><a href="http://www.coolrunning.com/">Cool Running</a></li>
                <li class="list-group-item"><a href="http://www.exploreminnesota.com/index.aspx?gclid=CjwKEAiAg_CnBRDc1N_wuoCiwyESJABpBuMXE6e2l1z43wlZ7k8ztit738J6SkrrsUUODIS9MF4EShoCpaDw_wcB">Explore Minnesota</a></li>
                <li class="list-group-item"><a href="http://www.getsetusa.com/minnesota/calendar.php">Get Set USA</a></li>
                <li class="list-group-item"><a href="http://events.kare11.com/">Kare 11</a></li>
                <li class="list-group-item"><a href="http://www.raceberryjam.com/indexrr.html">Raceberry Jam</a></li>
                <li class="list-group-item"><a href="http://www.minnesotarunner.com/run?page=Calendar">Minnesota Runner</a></li>
                <li class="list-group-item"><a href="http://minneapolisrunning.com/calendar/">Minneapolis Running</a></li>
                <li class="list-group-item"><a href="http://onlineracecalendar.com/">Online Race Calendar</a></li>
                <li class="list-group-item"><a href="http://www.raceguide365.com/">RaceGuide 365</a></li>
                <li class="list-group-item"><a href="http://www.trifind.com/">Tri-Find</a></li>
                <li class="list-group-item"><a href="http://www.runningintheusa.com/">Running in the USA</a></li>
                <li class="list-group-item"><a href="http://www.racedirectorresource.com/">Race Director Resource</a></li>
        </ul>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>