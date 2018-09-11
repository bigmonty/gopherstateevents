<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet, lTeamID, lCoachID
Dim sTeamName, sGender, sMeetName, sSport, sFirstName, sLastName, sEmail, sPhone
Dim MTeams(), FTeams(), Coaches(), Teams(), MeetClasses()
Dim dMeetDate
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetName, MeetDate, Sport FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
sSport = rs(2).Value
Set rs = Nothing

If Request.Form.Item("submit_team") = "submit_team" Then
	sTeamName = Replace(Request.Form.Item("team_name"), "'", "''")
	sGender = Request.Form.Item("gender")
	lCoachID = Request.Form.Item("coaches")

	sql = "INSERT INTO Teams (CoachesID, TeamName, Gender, TeamYear, Sport) VALUES (" & lCoachID & ", '" 
	sql = sql & sTeamName & "', '" & sGender & "', '" & Year(Date) & "', '" & sSport & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
ElseIf Request.Form.Item("submit_coach") = "submit_coach" Then
    Dim lNewCoachID
    Dim sNewUserID, sNewPassword

	sFirstName = Replace(Request.Form.Item("first_name"), "'", "''")
	sLastName = Replace(Request.Form.Item("last_name"), "'", "''")
	sEmail = Request.Form.Item("email")
	sPhone = Request.Form.Item("phone")
	
	sql = "INSERT INTO Coaches (FirstName, LastName, Email, Phone) VALUES ('" & sFirstName & "', '" & sLastName & "', '"  & sEmail 
    sql = sql & "', '" & sPhone & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT CoachesID FROM Coaches WHERE FirstName = '" & sFirstName & "' AND LastName = '" & sLastName & "' ORDER BY CoachesID DESC"
    rs.Open sql, conn, 1, 2
    lNewCoachID = rs(0).Value
    sNewUserID = Left(sFirstName, 2) & "_" & Left(sLastName, 3) & "_" & rs(0).Value
    sNewPassword = Left(sLastName, 2) & "_" & Left(sFirstName, 3) & "_" & rs(0).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UserID, Password FROM Coaches WHERE CoachesID = " & lNewCoachID
    rs.Open sql, conn, 1, 2
    rs(0).Value = sNewUserID
    rs(1).Value = sNewPassword
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("insert_team") = "insert_team" Then
	lTeamID = Request.Form.Item("new_team_id")

	sql = "INSERT INTO MeetTeams (TeamsID, MeetsID) VALUES (" & lTeamID & ", " & lThisMeet & ")"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

'get participating teams info	
i = 0
ReDim MTeams(1, 0)
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

i = 0
ReDim FTeams(1, 0)
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

i = 0
ReDim Coaches(1, 0)
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

i = 0
ReDim Teams(1, 0)
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
		Teams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
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
<!--#include file = "../../includes/meta2.asp" -->
<title>CCMeet Edit Teams</title>

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
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h4 class="h4">Participating Teams/Coaches for <%=sMeetName%> on <%=dMeetDate%></h4>
			<ul class="nav">
				<li class="nav-item"><a class="nav-link" href="/ccmeet_admin/manage_meet/manage_teams.asp?meet_id=<%=lThisMeet%>">Back</a></li>
				<li class="nav-item"><a class="nav-link" href="edit_teams.asp?meet_id=<%=lThisMeet%>">Refresh</a></li>
			</ul>
				
            <h4 class="h4">Participating Teams (Click to Edit)</h4>
			
			<div class="row">
				<div class="col-sm-3">
					<h5 class="h5">Female Teams (<%=UBound(FTeams, 2)%>)</h5>
					
					<ol class="list-group">
						<%For i = 0 to UBound(FTeams, 2) - 1%>
							<li class="list-group-item"><a href="javascript:pop('edit_this_team.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=FTeams(0, i)%>',1000,600)"><%=FTeams(1, i)%></a></li>
						<%Next%>
					</ol>
				</div>
				<div class="col-sm-3">
					<h5 class="h5">Male Teams (<%=UBound(MTeams, 2)%>)</h5>
					
					<ol class="list-group">
						<%For i = 0 to UBound(MTeams, 2) - 1%>
							<li class="list-group-item"><a href="javascript:pop('edit_this_team.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=MTeams(0, i)%>',1000,600)"><%=MTeams(1, i)%></a></li>
						<%Next%>
					</ol>
				</div>
				<div class="col-sm-6">
					<div>
						<h5 class="h5">Insert A Team Into This Meet:</h5>
						<span>(NOTE: If a team is not listed, you must add them!)</span>
						<form role="form" class="form-inline" name="insert_team" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkInsertTeam()">
						<label for="new_team_id">Team Name:</label>
						<select class="form-control" name="new_team_id" id="new_team_id" tabindex="10">
							<option value="">&nbsp;</option>
							<%For i = 0 to UBound(Teams, 2) - 1%>
								<option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
							<%Next%>
						</select>
						<input type="hidden" name="insert_team" id="insert_team" value="insert_team">
						<input type="submit" class="form-control" name="submit3" id="submit3" tabindex="11" value="Insert">
						</form>
					</div>
					<hr>
					<div>
						<h5 class="h5">Add Team to Database:</h5>
						<span>(NOTE: If a coach is not listed, you must add them first!)</span>
						<form role="form" class="form-horizontal" name="add_team" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkTeam()">
						<div class="form-group">
							<label for="team_name" class="control-label col-xs-3">Team Name:</label>
							<div class="col-xs-6">
								<input type="text" class="form-control" name="team_name" id="team_name" maxlength="50" tabindex="1">
							</div>
							<label for="gender" class="control-label col-xs-1">M/F:</label>
							<div class="col-xs-2">
								<select class="form-control" name="gender" id="gender" tabindex="2">
									<option value="">&nbsp;</option>
									<option value="M">Male</option>
									<option value="F">Female</option>
								</select>
							</div>
						</div>
						<div class="form-group">
							<label for="coaches" class="control-label col-xs-3">Coach:</label>
							<div class="col-xs-6">
								<select class="form-control" name="coaches" id="coaches" tabindex="3">
									<option value="">&nbsp;</option>
									<%For i = 0 to UBound(Coaches, 2) - 1%>
										<option value="<%=Coaches(0, i)%>"><%=Coaches(1, i)%></option>
									<%Next%>
								</select>
							</div>
							<div class="col-xs-3">
								<input type="hidden" name="submit_team" id="submit_team" value="submit_team">
								<input type="submit" class="form-control" name="submit1" id="submit1" tabindex="4" value="Add">
							</div>
						</div>
						</form>
					</div>
					<hr>
					<div>
						<h5 class="h5">Add Coach To Database:</h5>
						<form role="form" class="form-horizontal" name="add_coach" method="post" action="edit_teams.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkCoach()">
						<div class="form-group">
							<label for="first_name" class="control-label col-xs-2">First Name:</label>
							<div class="col-xs-4">
								<input type="text" class="form-control" name="first_name" id="first_name" maxlength="10">
							</div>
							<label for="last_name" class="control-label col-xs-2">Last Name:</label>
							<div class="col-xs-4">
								<input type="text" class="form-control" name="last_name" id="last_name" maxlength="15">
							</div>
						</div>
						<div class="form-group">
							<label for="email" class="control-label col-xs-2">Email:</label>
							<div class="col-xs-4">
								<input type="text" class="form-control" name="email" id="email" maxlength="100">
							</div>
							<label for="phone" class="control-label col-xs-2">Phone:</label>
							<div class="col-xs-4">
								<input type="text" class="form-control" name="phone" id="phone" maxlength="20">
							</div>
						</div>
						<div class="form-group">
							<input type="hidden" name="submit_coach" id="submit_coach" value="submit_coach">
							<input type="submit" class="form-control" name="submit2" id="submit2" tabindex="9" value="Add This Coach">
						</div>
						</form>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
