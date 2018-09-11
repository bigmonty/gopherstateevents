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
<title>How a GSE&copy; CC/Nordic Series Works</title>
<meta name="description" content="How the Gopher State Events (GSE) Cross-Country/Nordic Ski Series works.">
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">

    <hr>

    <div class="row">
        <h4 class="h4">How our Cross-Country/Nordic Ski Series Works</h4>

        <p>GSE CC/Nordic series is just a virtual affiliation 
        between multiple events that we manage.  It's purpose is to rank individuals and teams based solely on performance in the meets included.  It can be purely 
        informational or used for All-Conference selection.  Coaches have input into how their conference's series is scored.</p>

        <p><span style="font-weight: bold;">About Our Algorithm-Scoring By Percentile Ranking:</span>  First place in each race is awarded 100 points.  Points for subsequent finishers are 
        awarded based on their place as compared to the size of the field in that gender/category.  Doing this allows races of different distances, modalities, 
        and sizes to be in the same series.</p>

        <p><span style="font-weight: bold;">About Our Algorithm-Scoring By Points:</span>  Conference coaches or bylaws can determine a points system for
        calculating rankings.  In this system, the first place finisher gets a certain number of points (decided by the conference) and everyone that 
        follows gets one less point than the person ahead of them, stopping at 0. Races in some events (like conference
        championship meets, for instance) may carry more "weight" than others.  Again, this is determined by the
        administration and/or coaches.</p>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
