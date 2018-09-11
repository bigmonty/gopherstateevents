<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lThisMeet, lThisPart
Dim sMeetName
Dim MyRslts(10)
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
lThisPart = Request.QueryString("this_part")

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

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, t.TeamName, r.RosterID, r.Gender, ir.RaceTime, ra.RaceDist, "
sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.MeetsID = " & lThisMeet & " AND r.RosterID = " & lThisPart
rs.Open sql, conn, 1, 2
MyRslts(0) = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
MyRslts(1) = Replace(rs(2).Value, "''", "'")
MyRslts(2) = GetGrade()
MyRslts(3) = rs(4).Value
MyRslts(4) = rs(5).Value
MyRslts(5) = rs(6).Value
MyRslts(6) = rs(7).Value
MyRslts(7) = rs(8).Value
If CInt(rs(9).Value) = 0 Then
	MyRslts(8) = "---"
Else
	MyRslts(8) = rs(9).Value
End If
MyRslts(9) = rs(10).Value
MyRslts(10) = rs(3).Value
rs.Close
Set rs = Nothing
	
Private Function GetGrade()
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & " FROM Grades WHERE RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>Edidt GSE CC/Nordic Individual Results</title>
<!--#include file = "../../../includes/js.asp" -->
</head>
<body>
<div style="margin: 10px;padding: 10px;background-color: #fff;">
	<h4 class="h4">Edit Individual Results For <%=sMeetName%> on <%=dMeetDate%></h4>

</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
