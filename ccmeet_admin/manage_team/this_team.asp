<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lTeamID
Dim sTeamName
Dim RosterArray(), MeetsArr()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value
Set rs = Nothing

i = 0
ReDim RosterArr(2, 0)
sql = "SELECT FirstName, LastName, RosterID FROM Roster WHERE TeamsID = " & lTeamID & " ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RosterArr(0, i) = Replace(rs(0).Value, "''", "'")
	RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
	RosterArr(2, i) = GetGrade(rs(2).Value)
	i = i + 1
	ReDim Preserve RosterArr(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim MeetsArr(1, 0)
sql = "SELECT m.MeetName, m.MeetDate FROM Meets m INNER JOIN MeetTeams mt "
sql = sql & "ON m.MeetsID = mt.MeetsID WHERE mt.TeamsID = " & lTeamID & " ORDER BY m.MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetsArr(0, i) = Replace(rs(0).Value, "''", "'")
	MeetsArr(1, i) = rs(1).Value
	
	i = i + 1
	ReDim Preserve MeetsArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	If Month(Date) < 8 Then
        sql2 = "SELECT Grade" & Right(CStr(Year(Date) - 1), 2) & " FROM Grades WHERE RosterID = " & lMyID
    Else
        sql2 = "SELECT Grade" & Right(CStr(Year(Date)), 2) & " FROM Grades WHERE RosterID = " & lMyID
    End If
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then GetGrade = rs2(0).Value
	rs2.Close
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski This Team Data</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div style="margin:5px;">
	<table style="width:200px;font-size:0.9em;background-color:#fff;">
		<tr>
			<th colspan="2">
				Cross-Country Team Data For <%=sTeamName%>
			</th>
		</tr>
		<tr>
			<td valign="top">
				<fieldset>
					<legend>Roster</legend>
					<table>
						<tr>
							<th style="white-space:nowrap;text-align:left;">
								First Name:
							</th>
							<th style="white-space:nowrap;text-align:left;">
								Last Name:
							</th>
							<th style="white-space:nowrap;text-align:left;">
								Grade:
							</th>
						</tr>
						<%For i = 0 to UBound(RosterArr, 2) - 1%>
							<tr>
								<td style="white-space:nowrap;text-align:left;">
									<%=RosterArr(0, i)%>
								</td>
								<td style="white-space:nowrap;text-align:left;">
									<%=RosterArr(1, i)%>
								</td>
								<td style="white-space:nowrap;text-align:left;">
									<%=RosterArr(2, i)%>
								</td>
							</tr>
						<%Next%>
					</table>
				</fieldset>
			</td>
			<td valign="top">
				<fieldset>
					<legend>Meets</legend>
					<table>
						<tr>
							<th style="white-space:nowrap;text-align:left;">
								Meet Name
							</th>
							<th style="white-space:nowrap;text-align:left;">
								Meet Date
							</th>
						</tr>
						<%For i = 0 to UBound(MeetsArr, 2) - 1%>
							<tr>
								<td style="white-space:nowrap;text-align:left;">
									<%=MeetsArr(0, i)%>
								</td>
								<td style="white-space:nowrap;text-align:left;">
									<%=MeetsArr(1, i)%>
								</td>
							</tr>
						<%Next%>
					</table>
				</fieldset>
			</td>
		</tr>
	</table>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
