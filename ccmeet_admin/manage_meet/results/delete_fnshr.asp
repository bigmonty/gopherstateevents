<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisMeet, lThisRace
Dim sMeetName, sRaceName
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
    Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If CStr(lThisRace) = vbNullString Then lThisRace = Races(0, 0)

If Not CLng(lThisRace) = 0 Then
	sql = "SELECT RaceName FROM Races WHERE RacesID = " & lThisRace
	Set rs = conn.Execute(sql)
	sRaceName = rs(0).Value
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE CC/Nordic Results Manager: Delete Finisher</title>
<!--#include file = "../../../includes/meta2.asp" -->



</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "../manage_meet_nav.asp" -->
			<%End If%>

			<h4 class="h4">Results Manager for <%=sMeetName%> on <%=dMeetDate%>:&nbsp;Delete Finisher</h4>

			<!--#include file = "results_nav.asp" -->
		</div>
    </div>	
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
