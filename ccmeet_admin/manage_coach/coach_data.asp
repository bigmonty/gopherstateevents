<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim CoachArr(), TeamsArr()
Dim lCoachID

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim CoachArr(5, 0)
sql = "SELECT CoachesID,  FirstName, LastName, Email, Phone, UserID, Password FROM Coaches ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	CoachArr(0, i) = rs(0).Value
	CoachArr(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	CoachArr(2, i) = rs(3).Value
	CoachArr(3, i) = rs(4).Value
	CoachArr(4, i) = rs(5).Value
	CoachArr(5, i) = rs(6).Value
	i = i + 1
	ReDim Preserve CoachArr(5, i)
	rs.MoveNext
Loop
Set rs = Nothing

Function GetTeams(lCoachID)
	j = 0
	ReDim TeamsArr(1, 0)
	sql = "SELECT TeamsID, TeamName, Sport FROM Teams WHERE CoachesID = " & lCoachID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		TeamsArr(0, j) = rs(0).Value
		TeamsArr(1, j) = rs(1).Value & " (" & rs(2).Value & ")"
		j = j + 1
		ReDim Preserve TeamsArr(1, j)
		rs.MoveNext
	Loop
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>CCMeet Coach Data</title>
</head>
<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<h3 class="h4">Cross-Country/Nordic Ski Coach Data</h3>
			
            <ul class="nav">
                <li class="nav-item"><a class="nav-link" href="dwnld_data.asp">Download Coach Data</a></li>
                <li class="nav-item"><a class="nav-link" href="coach_data.asp">Refresh</a></li>
            </ul>

			<table class="table table-striped">
				<tr>
					<th>No.</th>
					<th>Name (click to edit)</th>
					<th>Email</th>
					<th>Phone</th>
					<th>User ID</th>
					<th>Password</th>
					<th>Team(s) (Sport)</th>
				</tr>
				<%For i = 0 to UBound(CoachArr, 2) - 1%>
					<tr>
						<td><%=i + 1%>)</td>
						<td><a href="javascript:pop('this_coach.asp?coach_id=<%=CoachArr(0, i)%>',100000,375)"><%=CoachArr(1, i)%></a></td>
						<td><a href="mailto:<%=CoachArr(2, i)%>">Send</a></td>
						<td><%=CoachArr(3, i)%></td>
						<td><%=CoachArr(4, i)%></td>
						<td><%=CoachArr(5, i)%></td>
						<td>
							<%Call GetTeams(CoachArr(0, i))%>
							<%For j = 0 to UBound(TeamsArr, 2) - 1%>
								<%If j <> UBound(TeamsArr, 2) - 1 Then%>
									<a href="javascript:pop('../manage_team/this_team.asp?team_id=<%=TeamsArr(0, j)%>',650,500)"><%=TeamsArr(1, j)%></a><br>
								<%Else%>
									<a href="javascript:pop('../manage_team/this_team.asp?team_id=<%=TeamsArr(0, j)%>',650,500)"><%=TeamsArr(1, j)%></a>
								<%End If%>
							<%Next%>
						</td>
					</tr>
				<%Next%>
			</table>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
