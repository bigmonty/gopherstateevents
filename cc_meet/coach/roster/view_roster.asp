<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lTeamID, lMyID, lCoachID, lCellProvidersID
Dim RosterArr(), DeleteArr(), TeamsArr(), CellProviders
Dim i, j
Dim sFirstName, sLastName, sGender, iGrade, iMyGrade, sEmail, sCellPhone, sSport
Dim sNewFirst, sNewLast, sNewGender, iNewGrade
Dim sArchiveThis
Dim dShutdown, dMeetDate
Dim sErrMsg
Dim sGradeYear
Dim bInsertThis, bNullGrade

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

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

'get teams
i = 0
ReDim TeamsArr(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender, Sport FROM Teams WHERE CoachesID = " & lCoachID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	TeamsArr(0, i) = rs(0).value 
	TeamsArr(1, i) = rs(1).Value & " (" & rs(2).Value & ")"
    TeamsArr(2, i) = rs(3).Value
	i = i + 1
	ReDim Preserve TeamsArr(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If UBound(TeamsArr, 2) = 1 Then lTeamID = TeamsArr(0, 0)

If Request.Form.Item("get_team") = "get_team" Then
	lTeamID = Request.Form.Item("teams")
ElseIf Request.Form.Item("add_part") = "add_part" Then
	sNewFirst = Replace(Request.Form.Item("first_name"), "'", "''")
	sNewLast = Replace(Request.Form.Item("last_name"), "'", "''")
	sNewGender = Request.Form.Item("gender")
	iNewGrade = Request.Form.Item("grade")
    
    'see if they exist in the db
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstName, LastName, Gender FROM Roster WHERE FirstName = '" & sNewFirst & "' AND LastName = '" 
    sql = sql & sNewLast & "' AND Gender = '" & sNewGender & "' AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs.Close
        Set rs = Nothing
        
        sErrMsg = "An athlete with this information already exists in the database.  If you do not see them on your roster they may have been archived.  "
		sErrMsg = sErrMsg & "You can view your archived athletes by clicking the 'Show Archives' link above, and then change them to active if you wish.  "
		sErrmsg = sErrMsg & " Otherwise, please make a minor change in this athlete's name (add a middle initial for instance) and then re-enter them."
    Else
        rs.Close
        Set rs = Nothing
        
        'insert team member
        sql = "INSERT INTO Roster (TeamsID, FirstName, LastName, Gender) VALUES (" & lTeamID & ", '" & sNewFirst & "', '" & sNewLast & "', '" 
		sql = sql & sNewGender & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
        
        'get roster id
        sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & lTeamID & " AND FirstName = '" & sNewFirst & "' AND LastName = '"
        sql = sql & sNewLast & "' AND Gender = '" & sNewGender & "' ORDER BY RosterID DESC"
        Set rs = conn.Execute(sql)
        lMyID = rs(0).Value
        Set rs = Nothing
 
		'get year for roster grades
		If Month(Date) <=7 Then
			sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
		Else
			sGradeYear = Right(CStr(Year(Date)), 2)	
		End If
       
        'insert grade
        sql = "INSERT INTO Grades (RosterID, Grade" & sGradeYear & ") VALUES (" & lMyID & ", " & iNewGrade & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
		
		sNewFirst = vbNullString
		sNewLast = vbNullString
		iNewGrade = 0
    End If
End If

If Not CLng(lTeamID) = 0 Then
	'get the time to the next meet this team is participating in
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT m.MeetDate, m.WhenShutdown FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID "
	sql = sql & "WHERE mt.TeamsID = " & lTeamID & " AND m.MeetDate >= '" & Date & "' ORDER BY m.MeetDate"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		dMeetDate = rs(0).Value
		dShutdown = rs(1).Value
	End If
	rs.Close
	Set rs = Nothing

	i = 0
	ReDim RosterArr(6, 0)
	sql = "SELECT RosterID, FirstName, LastName, Gender, Email, CellPhone FROM Roster WHERE TeamsID = " & lTeamID
	sql = sql & " AND Archive = 'n' ORDER BY LastName, FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		RosterArr(0, i) = rs(0).Value
		RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
		RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
		RosterArr(3, i) = GetGrade(rs(0).Value)
		RosterArr(4, i) = rs(3).Value
        RosterArr(5, i) = rs(4).Value
        RosterArr(6, i) = rs(5).Value
		i = i + 1
		ReDim Preserve RosterArr(6, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
End If

Private Function IncrGrade(lThisPart)
	IncrGrade = 0
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & sGradeYear & ", Grade" & CInt(sGradeYear) - 1 & " FROM Grades WHERE RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If Not rs2(1).Value & "" = "" Then 
		IncrGrade = CInt(rs2(1).Value) + 1
		rs2(0).Value = CInt(rs2(1).Value) + 1
		rs2.Update
	End If
	rs2.Close
	Set rs2 = Nothing
End Function
	
Private Function UpdateGrade(lMyID, iCurrGrade)
    bInsertThis = False

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        rs2(0).Value = iCurrGrade
	    rs2.Update
    Else
        bInsertThis = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bInsertThis = True Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & sGradeYear & ") Values (" & lMyID & ", " & iCurrGrade & ")"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function
	
Private Function GetGrade(lMyID)
    GetGrade = "0"

    bNullGrade = False
	Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        GetGrade = rs2(0).Value
        bNullGrade = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bNullGrade = False Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & sGradeYear & ") Values (" & lMyID & ", 0)"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function

If CStr(iNewGrade) = vbNullString Then iNewGrade = 0
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE View Roster</title>
     
<script>
function chkFields(){
	if (document.add_part.first_name.value==''||
	document.add_part.last_name.value==''||
	document.add_part.gender.value==''||
	document.add_part.grade.value==''){
		alert('All fields are required!');
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
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">GSE Cross-Country/Nordic Roster Page</h4>

			<p>Your "roster" represents all of the participants on your team.  It is different 
			than your "line-up" in that not all of your roster members necessarily compete in a meet.  Your "line-up" represents the members of your 
			roster that are competing in a given meet and the race(s) that they are competing in.</p>
				
			<form role="form" class="form-inline" name="get_team" method="post" action="view_roster.asp">
			<label for="teams">Select Team:</label>
			<select class="form-control" name="teams" id="teams" onchange="this.form.submit2.click();">
				<option value="0">&nbsp;</option>
				<%For i = 0 to UBound(TeamsArr, 2) - 1%>
					<%If CLng(TeamsArr(0, i)) = CLng(lTeamID) Then%>
						<option value="<%=TeamsArr(0, i)%>" selected><%=TeamsArr(1, i)%>&nbsp;(<%=TeamsArr(2, i)%>)</option>
					<%Else%>
						<option value="<%=TeamsArr(0, i)%>"><%=TeamsArr(1, i)%>&nbsp;(<%=TeamsArr(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="get_team" id="get_team" value="get_team">
			<input class="form-control" type="submit" name="submit2" id="submit2" value="Get This Team">
			</form>

			<%If Not CLng(lTeamID) = 0 Then%>
				<%If Not sErrMsg = vbNullString Then%>
					<p class="bg-danger"><%=sErrMsg%></p>
				<%End If%>

				<h4 class="h4">Add Particpant</h4>
				<form role="form" class="form-inline" name="add_part" method="post" action="view_roster.asp?team_id=<%=lTeamID%>" onsubmit="return chkFields()">
				<div class="form-group">
					<label>First Name:</label>
					<input class="form-control" type="text" name="first_name" id="first_name" size="10" maxlength="15" value="<%=sNewFirst%>">
					<label>Last Name:</label>
					<input class="form-control" type="text" name="last_name" id="last_name" size="10" maxlength="25" value="<%=sNewLast%>">
					<label>Grade:</label>
					<select class="form-control" name="grade" id="grade"> 
						<option value="">&nbsp;</option>
						<%For i = 3 To 16%>
							<%If CInt(iNewGrade) = CInt(i) Then%>
								<option value="<%=i%>" selected><%=i%></option>
							<%Else%>
								<option value="<%=i%>"><%=i%></option>
							<%End If%>
						<%Next%>
					</select>
					<label>Gender:</label>
					<select class="form-control" name="gender" id="gender"> 
						<option value="">&nbsp;</option>
						<%If sNewGender = "M" Then%>
							<option value="M" selected>Male</option>
							<option value="F">Female</option>
						<%ElseIf sGender = "F" Then%>
							<option value="M">Male</option>
							<option value="F" selected>Female</option>
						<%Else%>
							<option value="M">Male</option>
							<option value="F">Female</option>
						<%End If%>
					</select>
					<input class="form-control" type="hidden" name="add_part" id="add_part" value="add_part">
					<input class="form-control" type="submit" name="submit2" id="submit2" value="Add Participant">
				</div>
				</form>

				<br>

				<h4 class="h4">Existing Roster</h4>

				<a href="javascript:pop('print_roster.asp?team_id=<%=lTeamID%>',800,700)">Print Roster</a>	

				<table class="table table-striped table-condensed table-responsive">
					<tr>
						<th>No.</th>
						<th>Roster ID</th>
						<th>Name (click to edit)</th>
						<th>Grade</th>
						<th>M/F</th>
						<th>Email</th>
						<th>Cell</th>
						<th>History</th>
					</tr>
					<%For i = 0 to UBound(RosterArr, 2) - 1%>
						<tr>
							<td><%=i +1%>)</td>
							<td><%=RosterArr(0, i)%></td>
							<td><a href="javascript:pop('edit_part.asp?roster_id=<%=RosterArr(0, i)%>',1000,600)"><%=RosterArr(2, i)%>,<%=RosterArr(1, i)%></a></td>
							<td><%=RosterArr(3, i)%></td>
							<td><%=RosterArr(4, i)%></td>
							<td><a href="mailto:<%=RosterArr(5, i)%>"><%=RosterArr(5, i)%></a></td>
							<td><%=RosterArr(6, i)%></td>
							<td><a href="javascript:pop('my_history.asp?roster_id=<%=RosterArr(0, i)%>',1000,700)">View</a></td>
						</tr>
					<%Next%>
				</table>
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
