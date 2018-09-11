<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lThisMeet
Dim MeetTeams(), RaceArr(), LineUp()
Dim sMeetName
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")

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

'get meet teams array
i = 0
ReDim MeetTeams(2, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	MeetTeams(2,  i) = rs(2).Value
	i = i + 1
	ReDim Preserve MeetTeams(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

'get races in this meet
i = 0
ReDim RaceArr(2, 0)
sql = "SELECT RacesID, RaceDesc, Gender FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY OrderBy"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RaceArr(0, i) = rs(0).Value
	RaceArr(1, i) = Replace(rs(1).Value, "''", "'")
	RaceArr(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve RaceArr(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetLineUp(lTeamID, lRaceID)
	Dim x
	
	x = 0
	ReDim LineUp(0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT r.LastName, r.FirstName, g.Grade" & Right(CStr(Year(Date)), 2) & ", ir.Bib FROM Roster r "
	sql = sql & "INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID INNER JOIN Grades g ON g.RosterID = r.RosterID "
	sql = sql & "WHERE ir.RacesID = " & lRaceID & " AND r.TeamsID = " & lTeamID & " ORDER BY r.LastName, r.FirstName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		LineUp(x) = rs(3).Value & " - " & Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ") "
		x = x + 1
		ReDim Preserve LineUp(x)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing 
End Sub

Private Function HasParticipants(lTeamID)
	HasParticipants = "n"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
	sql = sql & "WHERE t.TeamsID = " & lTeamID & " AND MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then HasParticipants = "y"
	rs.Close
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Meet Line-Ups</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<%For i = 0 To UBound(MeetTeams, 2) - 1%>
		<%If HasParticipants(MeetTeams(0, i)) = "y" Then%>
			<div style="margin:10px;page-break-after:always;">
				<h4 class="h4"><%=MeetTeams(1, i)%>&nbsp;Line-Up for <%=sMeetName%> on <%=dMeetDate%></h4>
				
				<%For j = 0 To UBound(RaceArr, 2) - 1%>
					<%If UCase(Left(MeetTeams(2, i), 1)) = UCase(Left(RaceArr(2, j), 1)) Then%>
						<%Call GetLineUp(MeetTeams(0, i), RaceArr(0, j))%>
						
						<h5 style="margin-top:10px;"><%=RaceArr(1, j)%></h5>
						
						<ul style="font-size:0.8em;list-style:none;">
							<li style="font-weight:bold;">Bib - Name (Gr)</li>
							<%For k = 0 To UBound(LineUp) - 1%>
								<li><%=LineUp(k)%></li>
							<%Next%>
						</ul>
					<%End If%>
				<%Next%>
			</div>
		<%End If%>
	<%Next%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
