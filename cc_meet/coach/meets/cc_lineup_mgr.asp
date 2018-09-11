<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, sql2, rs2
Dim i, j, k
Dim lTeamID, lThisMeet, lMyID, lCoachID, lMeetClass
Dim RosterArr(), DeleteArr(), RaceAssign(), TempArr(5), MeetsReg()
Dim sCoachName, sGender, sGradeYear, sOrderBy, sMeetName, sMeetInfoSheet, sMapLink, sCourseMap, sTeamName
Dim sDynamicRaceAssign, sLockClasses, sHasRelay
Dim iGrade, iNumRaces
Dim dShutdown, dMeetDate

If Not (Session("role") = "coach" Or Session("role") = "team_staff") Then Response.Redirect "/default.asp?sign_out=y"

If Session("role") = "coach" Then
    lCoachID = Session("my_id")
Else
    lCoachID = Session("team_coach_id")
End If

lTeamID = Request.QueryString("team_id")
lThisMeet = Request.QueryString("meet_id")
sOrderBy = Request.QueryString("order_by")
 
'get year for roster grades
If Month(Date) <=5 Then
	sGradeYear = Right(CStr(Year(Date) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_class") = "submit_class" Then
    lMeetClass = Request.Form.Item("meet_classes")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    If lMeetClass & "" = "" Then
        rs(0).Value = Null
    Else
        rs(0).Value = lMeetClass
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_lineup") = "submit_lineup" Then
	i = 0
	j = 0
	ReDim DeleteArr(0)
	ReDim RaceAssign(3, 0)
	sql = "SELECT r.RosterID FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE TeamsID = " & lTeamID & " AND r.Archive = 'n' "
	sql = sql & "ORDER BY r.LastName, g.Grade" & sGradeYear
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If Request.Form.Item("race_" & rs(0).Value) = "none" Then
			DeleteArr(j) = rs(0).value			'mark for deletion from this meet if they are entered
			j = j + 1
			ReDim Preserve DeleteArr(j)
		Else
			RaceAssign(0, i) = Request.Form.Item("race_" & rs(0).Value)
			RaceAssign(1, i) = rs(0).Value
			RaceAssign(2, i) = "n"		'use as a flag to indicate that this exists and does not need to be inserted
			RaceAssign(3, i) = Request.Form.Item("bib_" & rs(0).Value)
			i = i + 1
			ReDim Preserve RaceAssign(3, i)
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'overwrite existing race assignment for this meet/participant if they exist
	For i = 0 to UBound(RaceAssign, 2) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RacesID, Bib FROM IndRslts WHERE RosterID = " & RaceAssign(1, i) & " AND MeetsID = " & lThisMeet
		rs.Open sql, conn, 1, 2
		If rs.recordcount > 0 Then
			rs(0).Value = RaceAssign(0, i)
			RaceAssign(2, i) = "y"			'indicate that this was already applied
			If Not RaceAssign(3, i) = vbNullString Then rs(1).Value = RaceAssign(3, i)
			rs.Update
		End If		
		rs.Close
		Set rs = Nothing
	Next
	
	'enter race if they do not exist
	For i = 0 to UBound(RaceAssign, 2) - 1
		If RaceAssign(2, i) = "n" Then
			If RaceAssign(3, i) & "" = "" Then
				sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				sql = sql & ", " & RaceAssign(1, i) & ")"
				Set rs = conn.Execute(sql)
				Set rs = Nothing
			Else
				sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID, Bib) VALUES (" & lThisMeet & ", " & RaceAssign(0, i)
				sql = sql & ", " & RaceAssign(1, i) & ", " & RaceAssign(3, i) & ")"
				Set rs = conn.Execute(sql)
				Set rs = Nothing
			End If
		End If
	Next
	
	'now delete those that are so marked
	For i = 0 to UBound(DeleteArr) - 1
		sql = "DELETE FROM IndRslts WHERE RosterID = " & DeleteArr(i) & " AND MeetsID = " & lThisMeet
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	Next
ElseIf Request.Form.Item("get_meet") = "get_meet" Then
	lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("get_team") = "get_team" Then
	lTeamID = Request.Form.Item("teams")
End If

If CStr(lTeamID) = vbNullString Then lTeamID = 0
If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

'get coach info
sql = "SELECT FirstName, LastName FROM Coaches WHERE CoachesID = " & lCoachID
Set rs = conn.Execute(sql)
sCoachName = rs(0).Value & " " & rs(1).Value
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

'get team gender
i = 0
ReDim RosterArr(5, 0)
ReDim MeetsReg(1, 0)

If Not CLng(lTeamID) = 0 Then
	sql = "SELECT Gender FROM Teams WHERE TeamsID = " & lTeamID
	Set rs = conn.Execute(sql)
	sGender = rs(0).Value
	Set rs = Nothing
	
	'convert gender to full word
	Select Case sGender
		Case "M"
			sGender = "Male"
		Case "F"
			sGender = "Female"
	End Select

	sql = "SELECT RosterID, FirstName, LastName, Gender FROM Roster WHERE TeamsID = " & lTeamID & " AND Archive = 'n' ORDER BY LastName, FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		RosterArr(0, i) = rs(0).Value
		RosterArr(1, i) = Replace(rs(1).Value, "''", "'")
		RosterArr(2, i) = Replace(rs(2).Value, "''", "'")
		RosterArr(3, i) = GetGrade(rs(0).Value)
		RosterArr(4, i) = rs(3).Value
        RosterArr(5, i) = GetRace(rs(0).Value)
  		i = i + 1
		ReDim Preserve RosterArr(5, i)
		rs.MoveNext
	Loop
	Set rs = Nothing

    're-order if necessary
    If sOrderBy = "race-name" Then
	    For i = 0 to UBound(RosterArr, 2) - 2
		    For j = i + 1 to UBound(RosterArr, 2) - 1
			    If CLng(RosterArr(5, i)) > CLng(RosterArr(5, j)) Then
				    For k = 0 to 5
					    TempArr(k) = RosterArr(k, i)
					    RosterArr(k, i) = RosterArr(k, j)
					    RosterArr(k, j) = TempArr(k)
				    Next
			    End IF
		    Next
	    Next
    ElseIf sOrderBy = "grade-name" Then
	    For i = 0 to UBound(RosterArr, 2) - 2
		    For j = i + 1 to UBound(RosterArr, 2) - 1
			    If CLng(RosterArr(3, i)) < CLng(RosterArr(3, j)) Then
				    For k = 0 to 5
					    TempArr(k) = RosterArr(k, i)
					    RosterArr(k, i) = RosterArr(k, j)
					    RosterArr(k, j) = TempArr(k)
				    Next
			    End IF
		    Next
	    Next
    End If

	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT m.MeetsID, m.MeetName, m.MeetDate FROM Meets m INNER JOIN MeetTeams mt ON m.MeetsID = mt.MeetsID "
	sql = sql & "WHERE mt.TeamsID = " & lTeamID & " ORDER BY m.MeetDate DESC"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		MeetsReg(0, i) = rs(0).Value
		MeetsReg(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
		i = i + 1
		ReDim Preserve MeetsReg(1, i)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

ReDim RaceArr(1, 0)

If Not CLng(lThisMeet) = 0 Then
    i = 0
    ReDim MeetClasses(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClassesID, ClassName, Gender, Details FROM MeetClasses WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MeetClasses(0, i) = rs(0).Value
        MeetClasses(1, i) = Replace(rs(1).Value, "''", "'")
        MeetClasses(2, i) = rs(2).Value
        MeetClasses(3, i) = rs(3).Value
        i = i + 1
        ReDim Preserve MeetClasses(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & lTeamID
    rs.Open sql, conn, 1, 2
    lMeetClass = rs(0).Value
    rs.Close
    Set rs = Nothing

    If lMeetClass & "" = "" Then lMeetClass = 0

	'get meet name
	sql = "SELECT MeetName, MeetDate, WhenShutdown, LockClasses, DynamicRaceAssign FROM Meets WHERE MeetsID = " 
	sql = sql & lThisMeet
	Set rs = conn.Execute(sql)
	sMeetName = rs(0).Value & " on " & rs(1).Value 
	
	'get year for roster grades
	If Month(rs(1).Value) <=7 Then
		sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
	Else
		sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
	End If
	
	dShutdown = rs(2).Value
    sLockClasses = rs(3).Value
	sDynamicRaceAssign = rs(4).Value
	Set rs = Nothing

    sHasRelay = "n"
    i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RacesID, RaceDesc, IndivRelay FROM Races WHERE MeetsID = " & lThisMeet & " AND (Gender = '" & sGender & "' OR Gender = 'Open') "
    sql = sql & "AND (RaceClass IS NULL OR RaceClass = " & lMeetClass & ") ORDER BY ViewOrder"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
        If sHasRelay = "n" Then
            If rs(2).Value = "Relay" Then sHasRelay = "y"
        End If

		RaceArr(0, i) = rs(0).Value
		RaceArr(1, i) = rs(1).Value
		i = i + 1
		ReDim Preserve RaceArr(1, i)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

	iNumRaces = UBound(RaceArr, 2)

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
End If

'get races this part is entered for this meet
Function GetRace(lThisPart)	
	GetRace = 0
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RacesID FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RosterID = " & lThisPart
	rs2.Open sql2, conn, 1, 2
	If rs2.recordcount > 0 Then
		GetRace = rs2(0).Value
	Else
		GetRace = 0
	End If
	rs2.Close
	Set rs2 = Nothing
End Function
	
Private Function GetGrade(lMyID)
    GetGrade = 0

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID
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
<title>GSE Cross-Country Line-Up Manager</title>

$(document).ready(function(){
	$('#parent-wrapper')
	.mouseover( 
	function() {
		$('#box').show();
	} 
	);
});
</head>

<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
           <h4 class="h4">CC/Nordic Meet Line-Up Manager:  <%=sCoachName%>!</h4>

			<div class="row">
				<div class="col-sm-6">
					<form role="form" class="form-inline" name="get_team" method="post" action="cc_lineup_mgr.asp">
					<div class="form-group form-group-sm">
						<label>Select Team:</label>&nbsp;
						<select class="form-control input-sm" name="teams" id="teams" onchange="this.form.submit2.click();">
							<option value="0">&nbsp;</option>
							<%For i = 0 to UBound(TeamsArr, 2) - 1%>
								<%If CLng(TeamsArr(0, i)) = CLng(lTeamID) Then%>
									<option value="<%=TeamsArr(0, i)%>" selected><%=TeamsArr(1, i)%> - <%=TeamsArr(2, i)%></option>
								<%Else%>
									<option value="<%=TeamsArr(0, i)%>"><%=TeamsArr(1, i)%> - <%=TeamsArr(2, i)%></option>
								<%End If%>
							<%Next%>
						</select>
						<input class="form-control" type="hidden" name="get_team" id="get_team" value="get_team">
						<input class="form-control input-sm" type="submit" name="submit2" id="submit2" value="Go">
					</div>
					</form>
				</div>	
				<div class="col-sm-6">
					<%If CLng(lTeamID) > 0 Then%>
						<form role="form" class="form-inline" name="get_meet" method="post" action="cc_lineup_mgr.asp?team_id=<%=lTeamID%>">
						<div class="form-group form-group-sm">
							<label>Select Meet:</label>&nbsp;
							<select class="form-control input-sm" name="meets" id="meets" onchange="this.form.submit3.click();">
								<option value="0">&nbsp;</option>
								<%For i = 0 to UBound(MeetsReg, 2) - 1%>
									<%If CLng(MeetsReg(0, i)) = CLng(lThisMeet) Then%>
										<option value="<%=MeetsReg(0, i)%>" selected><%=MeetsReg(1, i)%></option>
									<%Else%>
										<option value="<%=MeetsReg(0, i)%>"><%=MeetsReg(1, i)%></option>
									<%End If%>
								<%Next%>
							</select>
							<input class="form-control" type="hidden" name="get_meet" id="get_meet" value="get_meet">
							<input class="form-control input-sm" type="submit" name="submit3" id="submit3" value="Go">
						</div>
						</form>
					<%Else%>
						&nbsp;
					<%End If%>
				</div>	
    		</div>

			<%If CLng(lThisMeet) > 0 Then%>
				<%If sDynamicRaceAssign="y" Then%>
					<div class="row">
						<div class="col-sm-12 small">
							<p style="color:red;">
								IMPORTANT:  This meet is using Dynamic Race Assignment, meaning <mark>YOU DO NOT HAVE TO DECLARE WHICH
								RACE AN ATHLETE IS COMPETING IN ahead of time</mark>.  If you expect an athlete to
								compete, <mark>SIMPLY CLICK THE "TBD" circle</mark> and we will assign them to a race based on when they cross 
								the finish line.  If you <mark>do not make this selection they will not be assigned a bib</mark> or entered 
								into the meet.
							</p>
							<p style="color:red;">
								<mark>PLEASE DO NOT ENTER AN ATHLETE UNLESS YOU EXPECT THEM TO COMPETE AND PLEASE DO NOT ENTER AN ATHLETE IN
								A SPECIFIC RACE UNLESS IT IS ABSOLUTELY CERTAIN THEY WILL COMPETE IN THAT RACE!</mark>  Entering athletes that you do NOT
								expect to compete requires more meet prep and a higher cost for timing services.  <mark>You may add a
								limited number of athletes on site</mark> if needed.
							</p>
						</div>
					</div>
				<%Else%>
					<br>
				<%End If%>
				<%If UBound(MeetClasses, 2) > 0 Then%>
					<div class="row">
						<h4 class="h4">This Meet Has Meet Class(es)!</h4>
										
						<p>
							This meet has classes for some or all of the races in the meet.  Each team in the meet must have their class designated in 
							order to enter participants in class-based races. PLEASE BE CAREFUL TO DO THIS ACCURATELY BASED ON THE CLASS DESCRIPTION.  
							Otherwise your team members could be entered into the incorrect race(s).
						</p>
						<div class="col-sm-6">
							<h5 class="h5">Meet Classes</h5>
							<ol>
								<%For i = 0 To UBound(MeetClasses, 2) -1%>
									<li><%=MeetClasses(1, i)%> (<%=MeetClasses(3, i)%>)</li>
								<%Next%>
							</ol>
						</div>
						<div class="col-sm-6">
							<h5 class="h5">This Team's Class</h5>
							<form role="form" class="form-inline" name ="designate_class" method="post" action="cc_lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
							<div class="form-group">
								<select class="form-control" name="meet_classes" id="meet_classes">
									<option value="">&nbsp;</option>
									<%For i = 0 To UBound(MeetClasses, 2) - 1%>
										<%If CLng(lMeetClass) = CLng(MeetClasses(0, i)) Then%>
											<option value="<%=MeetClasses(0, i)%>" selected><%=MeetClasses(1, i)%></option>
										<%Else%>
											<option value="<%=MeetClasses(0, i)%>"><%=MeetClasses(1, i)%></option>
										<%End If%>
									<%Next%>
								</select>
								<%If sLockClasses = "n" Then%>
									<input class="form-control" type="hidden" name="submit_class" id="submit_class" value="submit_class">
									<input class="form-control" type="submit" name="submit_team_class" id="submit_team_class" value="Submit This">
								<%End If%>
							</div>
							</form>
							<%If sLockClasses = "y" Then%>
								<p style="color:red;">
									Classes for this meet have been locked.  Your team's class, if designated, should appear above.  
									If it looks incorrect, or if it is not designated within a week of the meet, please notify 
									<a href="mailto:bob.schneider@gopherstateevents.com">bob.schneider@gopherstateevents.com</a>.
								</p>
							<%End If%>

							<p>Class designations determined by official MSHSL numbers as 
								found <a href="http://www.mshsl.org/mshsl/enrollments15.asp" onclick="openThis(this.href,1024,768);return false;">here.</a></p>
						</div>
					</div>
				<%End If%>


				<nav class="navbar navbar-expand-lg navbar-dark bg-dark">
					<button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup" aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
						<span class="navbar-toggler-icon"></span>
					</button>
					<div class="collapse navbar-collapse" id="navbarNavAltMarkup">
						<div class="navbar-nav">
							<%If sHasRelay = "y" Then%>
								<a class="nav-item nav-link" href="javascript:pop('relay_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',610,700)">Relay Team Manager</a>
							<%End If%>		
							<a class="nav-item nav-link" href="javascript:pop('http://www.gopherstateevents.com/events/ccmeet_info.asp?meet_id=<%=lThisMeet%>',900,700)">Info Link</a>
							<a class="nav-item nav-link" href="javascript:pop('meet_sheet.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>',750,700)">Meet Sheet</a>
							<a class="nav-item nav-link" href="javascript:pop('blnk_mt_sht.asp',750,700)">Blank Meet Sheet</a>
							<%If Not sMapLink = vbNullString Then%>
								<a class="nav-item nav-link" href="javascript:pop('<%=sMapLink%>',1024,768)">Link to Site</a>
							<%End If%>
							<%If Not sMeetInfoSheet = vbNullString Then%>
								<a class="nav-item nav-link" href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>',1024,768)">Info Sheet</a>
							<%End If%>
							<%If Not sCourseMap = vbNullString Then%>
								<a class="nav-item nav-link" href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>',1024,768)">Course Map</a>
							<%End If%>
							<a class="nav-item nav-link" href="cc_lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>">Sort By Name</a></li>
							<a class="nav-item nav-link" href="cc_lineup_mgr.asp?meet_id=<%=lThisMeet%>&amp;team_id=<%=lTeamID%>&amp;order_by=grade-name">Sort By Grade-Name</a>
							<%If Date >= CDate(dMeetDate) Then%>
								<a class="nav-item nav-link" href="our_rslts.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>">Our Results</a>
							<%End If%>
							<a class="nav-item nav-link" href="javascript:pop('../change_order.pdf',750,700)">Change Order</a>
							<a class="nav-item nav-link" href="javascript:pop('../read_rate.pdf',750,700)">Read Rate</a>    
						</div>
					</div>
				</nav>
				<br>
				<%If Now() >= dShutdown Then%>
					<p style="color:red;">
						We're sorry.  Line-up changes for this event have closed.  Subject to local meet management
						approval, you may make changes at the meet.
					</p>
				<%Else%>
					<p style="color:red;">
						Line-ups for this meet must be submitted by <%=dShutdown%>
					</p>
				<%End If%>

				<form class="form" name="assign_races" method="post" action="cc_lineup_mgr.asp?team_id=<%=lTeamID%>&amp;meet_id=<%=lThisMeet%>&amp;order_by=<%=sOrderBy%>">
				<table class="table table-striped">
					<tr>
						<td colspan="<%=iNumRaces + 4%>">
							<input type="hidden" name="submit_lineup" id="submit_lineup" value="submit_lineup">
							
							<%If Now() >= dShutdown Then%>
								<input type="submit" class="form-control" name="submit2" id="submit2" value="Click Here To Save Line-Up Changes" disabled>
							<%Else%>
								<input type="submit" class="form-control" name="submit2" id="submit2" value="Click Here To Save Line-Up Changes">
							<%End If%>
						</td>
					</tr>
					<tr>
						<th rowspan="2">No</th>
						<th rowspan="2">Name</th>
						<th rowspan="2">Gr</th>
						<th colspan="<%=iNumRaces + 1%>">Race(s)</th>
					</tr>
					<tr>
						<td>None</td>
						<%For i = 0 to UBound(RaceArr, 2) - 1%>
							<td><a href="javascript:pop('../../../meet_dir/races/race_details.asp?race_id=<%=RaceArr(0, i)%>',400,250)"><%=RaceArr(1, i)%></a></td>
						<%Next%>
					</tr>
					<%For i = 0 to UBound(RosterArr, 2) - 1%>
						<tr>
							<td><%=i + 1%>)</td>
							<td><%=RosterArr(2, i)%>,&nbsp;<%=RosterArr(1, i)%></td>
							<td><%=RosterArr(3, i)%></td>
							<td>
								<input type="radio" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>" value="none" 
											checked>
							</td>
							<%For j = 0 to UBound(RaceArr, 2) - 1%>
								<td>
									<input type="radio" name="race_<%=RosterArr(0, i)%>" id="race_<%=RosterArr(0, i)%>" 
												value="<%=RaceArr(0, j)%>"
										<%If CLng(GetRace(RosterArr(0, i))) = CLng(RaceArr(0, j)) Then%>
										checked
										<%End If%>
									>
								</td>
							<%Next%>
						</tr>
					<%Next%>
				</table>
				</form>
			<%End If%>
		</div>
	</div>
	<!--#include file = "../../../includes/footer.asp" -->
</div>

<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
