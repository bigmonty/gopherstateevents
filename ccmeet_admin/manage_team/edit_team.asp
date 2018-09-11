<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lTeamID, lCoachID
Dim i
Dim sGender, sComments, sSport, sTeamName
Dim CoachArr

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sGender = Request.Form.Item("gender")
	lCoachID = Request.Form.Item("coaches")
	sComments = Request.Form.Item("comments")
	sSport = Request.Form.Item("sport")
    sTeamName = Request.Form.Item("team_name")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Gender, Comments, CoachesID, Sport, TeamName FROM Teams WHERE TeamsID = " & lTeamID
	rs.Open sql, conn, 1, 2
	rs(0).Value = sGender
	If sComments = vbNullString Then 
        rs(1).Value = Null
    Else
        rs(1).Value = Replace(sComments, "'", "''")
    End If
	rs(2).Value = lCoachID
	rs(3).Value = sSport
	If sTeamName = vbNullString Then 
        rs(4).Value = rs(4).OriginalValue
    Else
        rs(4).Value = Replace(sTeamName, "'", "''")
    End If
	rs.Update
	rs.Close
	Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Gender, Comments, CoachesID, Sport, TeamName FROM Teams WHERE TeamsID = " & lTeamID
rs.Open sql, conn, 1, 2
sGender = rs(0).Value
If Not rs(1).Value & "" = "" Then sComments = Replace(rs(1).Value, "''", "'")
lCoachID = rs(2).Value
sSport = rs(3).Value
sTeamName = Replace(rs(4).Value, "''", "'")
rs.Close
Set rs = Nothing

sql = "SELECT CoachesID, LastName, FirstName FROM Coaches ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
CoachArr = rs.GetRows()
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Edit Team Data</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body style="background: none;background-color: #ececd8;">
<div style="margin: 10px;padding: 10px;font-size: 0.85em;background-color: #fff;">
    <h4 style="margin-left:10px;">CCMeet Team Data: <%=sTeamName%></h4>
			
	<form name="team_data" method="post" action="edit_team.asp?team_id=<%=lTeamID%>">
	<table>
		<tr>
			<th>Team Name:</th>
			<td><input type="text" name="team_name" id="team_name" value="<%=sTeamName%>"></td>
			<th>M/F:</th>
			<td>
				<select name="gender" id="gender">
					<%If sGender = "M" Then%>
						<option value="M" selected>M</option>
						<option value="F">F</option>
					<%Else%>
						<option value="M">M</option>
						<option value="F" selected>F</option>
					<%End If%>
				</select>
			</td>
        </tr>
        <tr>
			<th>Sport:</th>
			<td><input type="text" name="sport" id="sport" value="<%=sSport%>"></td>
			<th>Coach:</th>
            <td>
				<select name="coaches" id="coaches">
					<%For i = 0 to UBound(CoachArr, 2)%>
						<%If CLng(lCoachID) = CLng(CoachArr(0, i)) Then%>
							<option value="<%=CoachArr(0, i)%>" selected><%=CoachArr(1, i)%>&nbsp;<%=CoachArr(2, i)%></option>
						<%Else%>
							<option value="<%=CoachArr(0, i)%>"><%=CoachArr(1, i)%>&nbsp;<%=CoachArr(2, i)%></option>
						<%End If%>
					<%Next%>
				</select>
			</td>
        </tr>
        <tr>
			<th valign="top">Comments:</th>
			<td colspan="3"><textarea name="comments" id="comments" rows="5" cols="70" style="font-size: 1.1em;"><%=sComments%></textarea></td>
		</tr>
		<tr>
			<td style="text-align:center;" colspan="4">
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit" id="submit" value="Save Changes">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
