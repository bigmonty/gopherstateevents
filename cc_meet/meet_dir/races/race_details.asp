<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lRaceID
Dim RaceDetails(11)
Dim i

lRaceID = Request.QueryString("race_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT RaceDesc, RaceTime, RaceDist, RaceUnits, Gender, ScoreMethod, NumAllow, NumScore, Comments, "
sql = sql & "TmAwds, IndAwds, StartType FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
For i = 0 to 11
	RaceDetails(i) = rs(i).Value
Next
Set rs = Nothing

If RaceDetails(6) = 0 Then RaceDetails(6) = "Unlimited"

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Upload Meet Info</title>
<!--#include file = "../../../includes/js.asp" -->
</head>
<body>

<div class="container">
    <h5 class="h5">Race Details For <%=RaceDetails(0)%></h5>
	<ul class="list-group">
		<li class="list-group-item">Race Time: <%=RaceDetails(1)%></li>
		<li class="list-group-item">Race Distance: <%=RaceDetails(2)%> <%=RaceDetails(3)%></li>
		<li class="list-group-item">Race Gender: <%=RaceDetails(4)%></li>
		<li class="list-group-item">Scoring Method: by <%=RaceDetails(5)%></li>
		<li class="list-group-item">Runners per Team: <%=RaceDetails(6)%></li>
		<li class="list-group-item">Number Scoring per Team: <%=RaceDetails(7)%></li>
		<li class="list-group-item">Start Type: <%=RaceDetails(11)%></li>
		<li class="list-group-item">Team Awards: <%=RaceDetails(9)%></li>
		<li class="list-group-item">Individual Awards: <%=RaceDetails(10)%></li>
		<li class="list-group-item">Comments: <%If Not RaceDetails(8) = vbNull Then Response.Write(Replace(RaceDetails(8), "''", "'"))%></li>
	</ul>
</div>
</body>
</html>
