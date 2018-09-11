<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lTeamID, lCoachID
Dim RosterArr(), TeamsArr()
Dim sGradeYear

If Not (Session("role") = "coach" Or Session("role") = "team_staff") Then Response.Redirect "/default.asp?sign_out=y"

If Session("role") = "coach" Then
    lCoachID = Session("my_id")
Else
    lCoachID = Session("team_coach_id")
End If

lTeamID = Request.QueryString("team_id")
If CStr(lTeamID) = vbNullString Then lTeamID = 0
 
'get year for roster grades
If Month(Date) <=7 Then
	sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get teams
i = 0
ReDim TeamsArr(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender FROM Teams WHERE CoachesID = " & lCoachID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	TeamsArr(0, i) = rs(0).value 
	TeamsArr(1, i) = rs(1).Value & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve TeamsArr(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If UBound(TeamsArr, 2) = 1 Then lTeamID = TeamsArr(0, 0)

If Request.Form.Item("get_team") = "get_team" Then
	lTeamID = Request.Form.Item("teams")
ElseIf Request.Form.Item("edit_archives") = "edit_archives" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID, Archive FROM Roster WHERE TeamsID = " & lTeamID & " AND Archive = 'y'"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("archive_" & rs(0).Value) = "n" Then
            rs(1).Value = "n"
		    rs.Update
        End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

i = 0
ReDim RosterArr(4, 0)
If Not CLng(lTeamID) = 0 Then
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, g.Grade" & sGradeYear & ", r.Gender FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID "
    sql = sql & "WHERE TeamsID = " & lTeamID & " AND Archive = 'y' ORDER BY r.LastName, r.FirstName, g.Grade" & sGradeYear
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    RosterArr(0, i) = rs(0).Value
	    RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
	    RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
	    RosterArr(3, i) = rs(3).Value
	    RosterArr(4, i) = rs(4).Value
	    i = i + 1
	    ReDim Preserve RosterArr(4, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE CCMeet Archived Roster</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
 	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">GSE Cross-Country/Nordic Archived Roster Page</h4>
					
			<p>These are athletes that have been active on your roster but are no longer so.  They may have graduated or
			not come out for the previous season for instance.  To re-activate them, simply change their status to "Active".</p>
			
			<form role="form" class="form-inline" name="get_team" method="post" action="archived_roster.asp">
			<label for="teams">Select Team:</label>
			<select class="form-control" name="teams" id="teams" onchange="this.form.submit2.click();">
				<option value="0">&nbsp;</option>
				<%For i = 0 to UBound(TeamsArr, 2) - 1%>
					<%If CLng(TeamsArr(0, i)) = CLng(lTeamID) Then%>
						<option value="<%=TeamsArr(0, i)%>" selected><%=TeamsArr(1, i)%></option>
					<%Else%>
						<option value="<%=TeamsArr(0, i)%>"><%=TeamsArr(1, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="get_team" id="get_team" value="get_team">
			<input class="form-control" type="submit" name="submit2" id="submit2" value="Get This Team">
			</form>

			<%If Not CLng(lTeamID) = 0 Then%>
				<h5 class="h5">Archived Roster</h5>
				<form class="form" name="archives" method="post" action="archived_roster.asp?team_id=<%=lTeamID%>">
				<table class="table table-striped">
					<tr>
						<td style="text-align:center" colspan="6">
							<input class="form-control" type="hidden" name="edit_archives" id="edit_archives" value="edit_archives">
							<input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes">
						</td>
					</tr>
					<tr>
						<th>No.</th>
						<th>First</th>
						<th>Last</th>
						<th>Gr</th>
						<th>M/F</th>
						<th>Status</th>
					</tr>
					<%For i = 0 to UBound(RosterArr, 2) - 1%>
						<tr>
							<td><%=i +1%>)</td>
							<td><%=RosterArr(1, i)%></td>
							<td><%=RosterArr(2, i)%></td>
							<td><%=RosterArr(3, i)%></td>
							<td><%=RosterArr(4, i)%></td>
							<td>
								<select class="form-control" name="archive_<%=RosterArr(0, i)%>" id="archive_<%=RosterArr(0, i)%>"> 
									<option value="y">Archived</option>
									<option value="n">Active</option>
								</select>
							</td>
						</tr>
					<%Next%>
				</table>
				</form>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
