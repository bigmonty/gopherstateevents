<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet, lTeamID, lCoachID
Dim sTeamName, sGender, sMeetName, sFirstName, sLastName, sEmail, sPhone, sSport, sEdit
Dim MTeams(), FTeams(), Coaches(), Teams(), MeetArr()
Dim dMeetDate
Dim bFound

Dim sMapLink, sMeetInfoSheet, sCourseMap
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")
lTeamID = Request.QueryString("team_id")
sEdit = Request.QueryString("edit")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim MeetArr(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetArr(0, i) = rs(0).Value
	MeetArr(1, i) = rs(1).Value & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve MeetArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(MeetArr, 2) = 1 Then lThisMeet = MeetArr(0, 0)

If Request.Form.Item("submit_meet") = "submit_meet" Then 
    lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	If Request.Form.Item("remove") = "y" Then
		sql = "DELETE FROM MeetTeams WHERE TeamsID = " & lTeamID & " AND MeetsID = " & lThisMeet
		Set rs = conn.Execute(sql)
		Set rs = Nothing
		
		lTeamID = 0
	End If
ElseIf Request.Form.Item("submit_team") = "submit_team" Then
	sTeamName = Replace(Request.Form.Item("team_name"), "'", "''")
	sGender = Request.Form.Item("gender")
	lCoachID = Request.Form.Item("coaches")
	
	sql = "INSERT INTO Teams (CoachesID, TeamName, Gender, TeamYear, Sport) VALUES (" & lCoachID & ", '" 
	sql = sql & sTeamName & "', '" & sGender & "', '" & Year(Date) & "', '" & sSport & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
ElseIf Request.Form.Item("submit_coach") = "submit_coach" Then
	sFirstName = Replace(Request.Form.Item("first_name"), "'", "''")
	sLastName = Replace(Request.Form.Item("last_name"), "'", "''")
	sEmail = Request.Form.Item("email")
	sPhone = Request.Form.Item("phone")
	
	sql = "INSERT INTO Coaches (FirstName, LastName, Email, Phone) VALUES ('" & sFirstName & "', '" & sLastName & "', '" 
	sql = sql & sEmail & "', '" & sPhone & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
ElseIf Request.Form.Item("insert_team") = "insert_team" Then
	lTeamID = Request.Form.Item("new_team_id")
	
	sql = "INSERT INTO MeetTeams (TeamsID, MeetsID) VALUES (" & lTeamID & ", " & lThisMeet & ")"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

If Not CLng(lThisMeet) = 0 Then
    If sEdit = "y" Then
	    sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
	    Set rs = conn.Execute(sql)
	    sTeamName = Replace(rs(0).Value, "''", "'")
	    sGender = rs(1).Value
	    Set rs = Nothing
    End If

    sql = "SELECT MeetName, MeetDate, Sport, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    sSport = rs(2).Value
    dWhenShutdown = rs(3).Value
    Set rs = Nothing

	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing

    'get participating teams info	
    ReDim MTeams(1, 0)
    i = 0
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'M' ORDER BY t.TeamName"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    MTeams(0, i) = rs(0).Value
	    MTeams(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve MTeams(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    ReDim FTeams(1, 0)
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT t.TeamsID, t.TeamName FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND Gender = 'F' ORDER BY t.TeamName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    FTeams(0, i) = rs(0).Value
	    FTeams(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve FTeams(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
End If

ReDim Coaches(1, 0)
i = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CoachesID, FirstName, LastName FROM Coaches ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Coaches(0, i) = rs(0).Value
	Coaches(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Coaches(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

ReDim Teams(1, 0)
i = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender FROM Teams WHERE Sport = '" & sSport & "' ORDER BY TeamName, Gender"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	bFound = False
	
	Select Case rs(2).Value
		Case "M"
			For j = 0 to UBound(MTeams, 2) - 1
				If CLng(rs(0).value) = CLng(MTeams(0, j)) Then
					bFound = True
					Exit For
				End If
			Next
		Case "F"
			For j = 0 to UBound(FTeams, 2) - 1
				If CLng(rs(0).value) = CLng(FTeams(0, j)) Then
					bFound = True
					Exit For
				End If
			Next
	End Select
	
	If bFound = False Then
		Teams(0, i) = rs(0).Value
		Teams(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve Teams(1, i)
	End If
	rs.MoveNext
Loop
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>Cross-Country/Nordic Ski Edit Teams</title>
<!--#include file = "../../../includes/js.asp" -->

<script>
function chkTeam(){
	if (document.add_team.team_name.value==''){
		alert('You must supply a team name!');
		return false;
	}
	else
		if (document.add_team.gender.value==''){
			alert('You must supply a gender for this team!');
			return false;
		}
	else
		return true;
}

function chkInsertTeam(){
	if (document.insert_team.new_team_id.value==''){
		alert('You must select a team!');
		return false;
	}
	else
		return true;
}

function chkCoach(){
	if (document.add_coach.first_name.value==''){
		alert('You must supply a first name!');
		return false;
	}
	else
		if (document.add_coach.last_name.value==''){
			alert('You must supply a first name!');
			return false;
		}
	else
		return true;
}

function chkFields(){
	if (document.edit_team.team_name.value==''){
		alert('You must supply a race description!');
		return false;
	}
	else 
		if (document.edit_team.gender.value==''){
			alert('You must supply a race time!');
			return false;
		}
	else 
		return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<!--#include file = "../../../includes/meet_dir_menu.asp" -->

	<h4 class="h4">CC/Nordic Meet Director: Add/Remove Teams</h4>

	<form class="form-inline" name="get_meets" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>">
	<label for="meets">Select Meet:</label>
	<select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
		<option value="">&nbsp;</option>
		<%For i = 0 to UBound(MeetArr, 2) - 1%>
			<%If CLng(lThisMeet) = CLng(MeetArr(0, i)) Then%>
				<option value="<%=MeetArr(0, i)%>" selected><%=MeetArr(1, i)%></option>
			<%Else%>
				<option value="<%=MeetArr(0, i)%>"><%=MeetArr(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
	<input type="submit" class="form-control" name="submit1" id="submit1" value="Get This">
	</form>
			
    <%If Not CLng(lThisMeet) = 0 Then%>
		<!--#include file = "../meet_dir_nav.asp" -->
			
		<div class="col-xs-3">
			<h5 class="h5">Female Teams (<%=UBound(FTeams, 2)%>)</h5>
			<ol class="list-group">
				<%For i = 0 to UBound(FTeams, 2) - 1%>
					<li class="list-group-item"><a href="edit_teams.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=FTeams(0, i)%>&amp;edit=y"><%=FTeams(1, i)%></a></li>
				<%Next%>
			</ol>
		</div>
        <div class="col-xs-3">
			<h5 class="h5">Male Teams (<%=UBound(MTeams, 2)%>)</h5>
			<ol class="list-group">
				<%For i = 0 to UBound(MTeams, 2) - 1%>
					<li class="list-group-item"><a href="edit_teams.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=MTeams(0, i)%>&amp;edit=y"><%=MTeams(1, i)%></a></li>
				<%Next%>
			</ol>
		</div>
		<div class="col-xs-6">
			<%If CLng(lTeamID) = 0 Then%>
				&nbsp;
			<%Else%>
				<h4 class="h4">Remove This Team</h4>
							
				<form class="form" name="edit_team" method="post" action="edit_teams.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>" 
					onsubmit="return chkFields()">
				<table class="table">
					<tr>
						<th>Team:</th>
						<td><%=sTeamName%></td>
                    </tr>
                    <tr>
						<th>M/F:</th>
						<td><%=sGender%></td>
                    </tr>
                    <tr>
						<th>Remove?</th>
						<td>
							<select class="form-control" name="remove" id="remove" tabindex="3">
								<option value="n" selected>No</option>
								<option value="y">Yes</option>
							</select>
						</td>
                    </tr>
                    <tr>
						<td colspan="2">
							<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
							<input type="submit" class="form-control" name="submit" id="submit" tabindex="4" value="Save Changes">
						</td>
					</tr>
				</table>
				</form>
			<%End If%>

			<form class="form" name="insert_team" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkInsertTeam()">
			<h4 class="h4">Insert A Team Into This Meet:</h4>
            <div class="bg-danger">
                (NOTE: If a team is not listed, you must add them to the database using the form below!)</span>
            </div>
			<table class="table">
                <tr>
                    <th>Team Name:</th>
                    <td>
						<select class="form-control" name="new_team_id" id="new_team_id" tabindex="10">
							<option value="">&nbsp;</option>
							<%For i = 0 to UBound(Teams, 2) - 1%>
								<option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
							<%Next%>
						</select>
                    </td>
                    <td>
						<input type="hidden" name="insert_team" id="insert_team" value="insert_team">
						<input type="submit" class="form-control" name="submit3" id="submit3" tabindex="11" value="Insert This Team">
                    </td>
                </tr>
            </table>
			</form>

			<form class="form" name="add_team" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkTeam()">
			<h4 class="h4">Add A Team to Database:</h4>
            <div class="bg-danger">
                (NOTE: If a coach is not listed, you must add them to the database using the form below!)</span>
            </div>
			<table class="table">
				<tr>
					<th>Team Name:</th>
					<td><input type="text" class="form-control" name="team_name" id="team_name" maxlength="50" size="20"></td>
                </tr>
                <tr>
					<th>M/F:</th>
					<td>
						<select class="form-control" name="gender" id="gender">
							<option value="">&nbsp;</option>
							<option value="M">Male</option>
							<option value="F">Female</option>
						</select>
					</td>
                </tr>
                <tr>
					<th>Coach:</th>
					<td>
						<select class="form-control" name="coaches" id="coaches" tabindex="3">
							<option value="">&nbsp;</option>
							<%For i = 0 to UBound(Coaches, 2) - 1%>
								<option value="<%=Coaches(0, i)%>"><%=Coaches(1, i)%></option>
							<%Next%>
						</select>
					</td>
                </tr>
                <tr>
					<td colspan="2">
						<input type="hidden" name="submit_team" id="submit_team" value="submit_team">
						<input type="submit" class="form-control" name="submit1" id="submit1" tabindex="4" value="Add This Team">
					</td>
				</tr>
			</table>
			</form>

			<form class="form" name="add_coach" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkCoach()">
			<h4 class="h4">Add A Coach To The Database:</h4>
            <div class="bg-danger">
                (NOTE: First Name, Last Name, and Email are required.)
            </div>
			<table class="table">
				<tr>
					<th>First Name:</th>
					<td><input type="text" class="form-control" name="first_name" id="first_name" maxlength="10"></td>
					<th>Last Name:</th>
					<td><input type="text" class="form-control" name="last_name" id="last_name" maxlength="15"></td>
                </tr>
                <tr>
					<th>Email:</th>
					<td><input type="text" class="form-control" name="email" id="email" maxlength="100"></td>
					<th>Phone:</th>
					<td><input type="text" class="form-control" name="phone" id="phone" maxlength="20"></td>
				</tr>
				<tr>
					<td style="text-align:center;" colspan="4">
						<input type="hidden" name="submit_coach" id="submit_coach" value="submit_coach">
						<input type="submit" class="form-control" name="submit2" id="submit2" tabindex="9" value="Add This Coach">
					</td>
				</tr>
			</table>
			</form>
		</div>
    <%End If%>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
