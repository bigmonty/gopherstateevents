<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lTeamID
Dim i
Dim TeamArr

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender, Comments, CoachesID, Sport FROM Teams ORDER BY Sport, TeamName"
rs.Open sql, conn, 1, 2
TeamArr = rs.GetRows()
rs.Close
Set rs = Nothing

Private Function GetCoach(lCoachID)
    GetCoach = "unknown"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstName, LastName FROM Coaches WHERE CoachesID = " & lCoachID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetCoach = Replace(rs(1).Value, "''", "'") & ", " &  Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Team Data</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h4 style="margin-left:10px;">CCMeet Team Data</h4>
			
			<table class="table table-striped">
				<tr>
					<th>No.</th>
					<th>Team Name (click to edit)</th>
					<th>M/F</th>
                    <th>Sport</th>
					<th>Comments</th>
					<th>Coach</th>
				</tr>
				<%For i = 0 to UBound(TeamArr, 2)%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td><a href="javascript:pop('edit_team.asp?team_id=<%=TeamArr(0, i)%>',600,300)"><%=TeamArr(1, i)%></a></td>
                        <td><%=TeamArr(2, i)%></td>
                        <td><%=TeamArr(5, i)%></td>
                        <td><%=TeamArr(3, i)%></td>
                        <td><%=GetCoach(TeamArr(4, i))%></td>
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
