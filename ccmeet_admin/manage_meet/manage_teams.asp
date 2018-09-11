<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k, m, n, p
Dim lThisMeet, lThisTeam
Dim sMeetName, sTeamName, sGradeYear, sTeamGender, sAssign, sTeamRange, sCoachName, sCoachEmail, sCoachPhone, sRemove, sUserName, sPassword
Dim iFirstBib, iLastBib
Dim MeetTeams(), Roster(), Races(), LineUp(), AssgndBibs(), AvailBibs(), TeamParts(), BibRange()
Dim dMeetDate
Dim bFound, bChangeMade

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

lThisMeet = Request.QueryString("meet_id")
lThisTeam = Request.QueryString("this_team")
sAssign = Request.QueryString("assign")
sTeamRange = Request.QueryString("team_range")
sRemove = Request.QueryString("remove")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_these_bibs") = "submit_these_bibs" Then
	'get races in this meet
	i = 0
	ReDim Races(0)
	sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY ViewOrder"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Races(i) = rs(0).Value
		i = i + 1
		ReDim Preserve Races(i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	For i = 0 To UBound(Races) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RosterID, Bib FROM IndRslts WHERE RacesID = " & Races(i)
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
			If Request.Form.Item("avail_" & rs(0).Value & "_" & Races(i)) = "y" Then
				rs(1).Value = Request.Form.Item("bib_" & rs(0).Value & "_" & Races(i))
				rs.Update
			End If
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
	Next
ElseIf Request.Form.Item("submit_team") = "submit_team" Then
	lThisTeam = Request.Form.item("teams")
End If

If CStr(lThisTeam) = vbNullString Then lThisTeam = 0

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing
 
'get year for roster grades
If Month(dMeetDate) <= 5 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

'Set rs = Server.CreateObject("ADODB.Recordset")
'sql = "SELECT Grade16 FROM Grades"
'rs.Open sql, conn, 1, 2
'Do While Not rs.EOF
'    rs(0).Value = "0"
'    rs.Update
'    rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing

'increment grades if needed
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Grade" & CInt(sGradeYear) - 1 & ", Grade" & sGradeYear & " FROM Grades WHERE Grade" & sGradeYear & " = 0 AND (Grade" & CInt(sGradeYear) - 1 
sql = sql & " > 0 AND Grade" & CInt(sGradeYear) - 1 & " <= 12)"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    rs(1).Value = CInt(rs(0).Value) + 1
    rs.Update
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get meet teams array
i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Not CLng(lThisTeam) = 0 Then
	'get team name
	sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lThisTeam
	Set rs = conn.Execute(sql)
	sTeamName = Replace(rs(0).Value, "''", "'") & " (Gender: " & rs(1).Value & ")"
	If rs(1).Value = "F" Then
		sTeamGender = "Female"
	Else
		sTeamGender = "Male"
	End If
	Set rs = Nothing
	
	sql = "SELECT c.FirstName, c.LastName, c.Phone, c.Email, c.UserID, c.Password FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID "
    sql = sql & "WHERE t.TeamsID = " & lThisTeam
	Set rs = conn.Execute(sql)
	sCoachName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
	sCoachPhone = rs(2).Value
	sCoachEmail = rs(3).Value
    sUserName = rs(4).Value
    sPassword = rs(5).Value
	Set rs = Nothing

	i = 0
	ReDim Roster(1, 0)
	sql = "SELECT r.RosterID, r.FirstName, r.LastName, g.Grade" & sGradeYear & " FROM Roster r INNER JOIN Grades g "
	sql = sql & "ON r.RosterID = g.RosterID WHERE TeamsID = " & lThisTeam & " AND Archive = 'n' "
    sql = sql & "ORDER BY r.LastName, r.FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Roster(0, i) = rs(0).Value
		Roster(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (Gr: " & rs(3).Value & ")"
		i = i + 1
		ReDim Preserve Roster(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	'get races in this meet
	i = 0
	ReDim Races(1, 0)
	sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lThisMeet & " AND Gender = '" & sTeamGender & "' ORDER BY ViewOrder"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Races(0, i) = rs(0).Value
		Races(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve Races(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
	
    'get bib range
    i = 0
    ReDim BibRange(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamBibsID, FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lThisTeam
    rs.Open sql, conn, 1,  2
    Do While Not rs.EOF
        BibRange(0, i) = rs(0).Value
        BibRange(1, i) = rs(1).Value
        BibRange(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve BibRange(2, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

If sRemove = "team" Then
	For i = 0 To UBound(Roster, 2) - 1
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT Bib FROM IndRslts WHERE RosterID = " & Roster(0, i)
		rs.Open sql, conn, 1, 2
		Do While Not rs.EOF
			rs(0).Value = 0
			rs.Update
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
	Next
End If

Call GetAvailBibs

If sAssign = "team" Then
	If sTeamRange = "y" Then
	Else
		Call GetTeamParts(lThisTeam)
		
        If UBound(TeamParts) > 0 Then
		    For p = 0 To UBound(TeamParts) - 1
			    Set rs2 = Server.CreateObject("ADODB.Recordset")
			    sql2 = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RosterID = " & TeamParts(p)
			    rs2.Open sql2, conn, 1, 2
			    rs2(0).Value = AvailBibs(p)
			    rs2.Update
			    rs2.Close
			    Set rs2 = Nothing
		    Next
		
		    Call GetAvailBibs
        End If
	End If
End If

Private Sub GetLineUp(lThisRace)
	Dim x
	
	x = 0
	ReDim LineUp(3, 0)
    sql = "SELECT r.RosterID, r.FirstName, r.LastName, ir.Bib FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
    sql = sql & "WHERE TeamsID = " & lThisTeam & " AND ir.RacesID = " & lThisRace & " ORDER BY r.LastName, r.FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		LineUp(0, x) = rs(0).Value
		LineUp(1, x) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(0).Value & ")"
		LineUp(2, x) = GetGrade(rs(0).Value) 
		LineUp(3, x) = rs(3).Value
		x = x + 1
		ReDim Preserve LineUp(3, x)
		rs.MoveNext
	Loop
	Set rs = Nothing
End Sub

Private Sub GetAvailBibs()
	'get meet bib range
	sql = "SELECT BibStart, BibEnd FROM BibRange WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	iFirstBib = rs(0).Value
	iLastBib = rs(1).Value
	Set rs = Nothing
	
	i = 0
	ReDim AssgndBibs(0)
	sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND Bib > 0 ORDER BY Bib"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    AssgndBibs(i) = rs(0).Value
	    i = i + 1
	    ReDim Preserve AssgndBibs(i)
	    rs.MoveNext
	Loop
	Set rs = Nothing

	k = 0
	ReDim AvailBibs(0)
	For i = iFirstBib To iLastBib
		If UBound(AssgndBibs) = 0 Then
			AvailBibs(k) = i
			k = k + 1
			ReDim Preserve AvailBibs(k)
		Else
			For j = 0 To UBound(AssgndBibs) - 1
				If CInt(AssgndBibs(j)) = CInt(i) Then 
					Exit For
				Else
					If j = UBound(AssgndBibs) - 1 Then
						AvailBibs(k) = i
						k = k + 1
						ReDim Preserve AvailBibs(k)
					End If
				End If
		    Next
		End If
	Next
End Sub

Private Sub GetTeamParts(lTeamID)
	Dim x
	
	ReDim TeamParts(0)
	sql2 = "Select ir.RosterID FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE r.TeamsID = " & lTeamID 
	sql2 = sql2 & " AND MeetsID = " & lThisMeet & " AND ir.Bib = 0 ORDER BY r.LastName, r.FirstName"
	Set rs2 = conn.Execute(sql2)
	Do While Not rs2.EOF
		TeamParts(x) = rs2(0).Value
		x = x + 1
		ReDim Preserve TeamParts(x)
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Sub
	
Private Function GetGrade(lMyID)
 	sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lMyID & " AND Grade" & sGradeYear & " <= 13"
	Set rs2 = conn.Execute(sql2)
    If rs2.BOF and rs2.EOF Then
        GetGrade = 0
    Else
        GetGrade = rs2(0).Value
    End If
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country Team Meet Instructions Sheet</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
            <div class="row">
			    <div class="col-sm-6">		
			        <h4 class="h4">Manage Teams: <%=sMeetName%></h4>
                </div>
			    <div class="col-sm-6">		
			        <ul class="nav">
                        <li class="nav-item"><a class="nav-link" href="javascript:pop('rstr_lnup_upld.asp?meet_id=<%=lThisMeet%>',800,700)">Batch Upload Rosters & Line-Ups</a></li>
				        <li class="nav-item"><a class="nav-link" href="edit_teams.asp?meet_id=<%=lThisMeet%>">Add/Edit Teams</a></li>
                        <li class="nav-item"><a class="nav-link" href="manage_teams.asp?meet_id=<%=lThisMeet%>&amp;this_team=<%=lThisTeam%>">Refresh</a></li>
			        </ul>
                </div>
            </div>
			
            <div class="row">
			    <div class="col-sm-5">		
				    <form class="form-inline" name="mnge_teams" method="post" action="manage_teams.asp?meet_id=<%=lThisMeet%>">
				    <select class="form-control" name="teams" id="teams" onchange="this.form.submit1.click();" style="font-size: 0.9em;">
					    <option value="0">&nbsp;</option>
					    <%For i = 0 To UBound(MeetTeams, 2) - 1%>
						    <%If CLng(lThisTeam) = CLng(MeetTeams(0, i)) Then%>
							    <option value="<%=MeetTeams(0, i)%>" selected><%=MeetTeams(1, i)%></option>
						    <%Else%>
							    <option value="<%=MeetTeams(0, i)%>"><%=MeetTeams(1, i)%></option>
						    <%End If%>
					    <%Next%>
				    </select>
				    <input type="hidden" name="submit_team" id="submit_team" value="submit_team">
				    <input type="submit" class="form-control" name="submit1" id="submit1" value="View This Team">
				    </form>
			    </div>
			    <div class="col-sm-7">
			        <%If Not CLng(lThisTeam) = 0 Then%>
				        <h4 class="h4">Batch Assign Bibs</h4>
				
                    <ul class="nav">
                        <li class="nav-item">
                            <a class="nav-link" href="javascript:pop('/cc_meet/coach/meets/bib_list.asp?team_id=<%=lThisTeam%>&amp;meet_id=<%=lThisMeet%>',750,700)">Bib List</a>
                        </li>					    
                        <li class="nav-item">
                            <a class="nav-link" href="manage_teams.asp?meet_id=<%=lThisMeet%>&amp;this_team=<%=lThisTeam%>&amp;assign=team&amp;team_range=n" style="font-size: 0.9em;">This Team Missing</a>
					    </li>
                        <li class="nav-item">
                            <a class="nav-link" href="manage_teams.asp?meet_id=<%=lThisMeet%>&amp;this_team=<%=lThisTeam%>&amp;assign=team&amp;team_range=y" style="font-size: 0.9em;">This Team Missing (Use Team Bibs)</a>
					    </li>
                        <li class="nav-item">
                            <a class="nav-link" href="manage_teams.asp?meet_id=<%=lThisMeet%>&amp;this_team=<%=lThisTeam%>&amp;remove=team" style="font-size: 0.9em;">Remove All Team (NO UNDO!)</a>
			            </li>
                    </ul>
                    <%End If%>
                </div>
            </div>
			
			<%If Not CLng(lThisTeam) = 0 Then%>
                <ul class="nav">
                    <li class="nav-item">
                        <a class="nav-link" href="javascript:pop('../manage_team/edit_roster.asp?team_id=<%=lThisTeam%>',900,750)">Edit Roster</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="javascript:pop('../manage_meet/lineup_mgr.asp?team_id=<%=lThisTeam%>&amp;meet_id=<%=lThisMeet%>',900,750)">Edit Lineup</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" href="javascript:pop('bib_range.asp?this_team=<%=lThisTeam%>',600,300)">Edit Bib Range</a>
                    </li>
                </ul>

				<div class="row">
					<div class="col-sm-4">
						<h4 class="h4"><%=sTeamName%></h4>

						<h5 class="h5">Coach Info</h5>
								
						<ul class="list-group">
							<li class="list-group-item"><%=sCoachName%></li>
							<li class="list-group-item"><%=sCoachPhone%></li>
							<li class="list-group-item"><a href="mailto:<%=sCoachEmail%>"><%=sCoachEmail%></a></li>
							<li class="list-group-item">User Name:&nbsp;<%=sUserName%></li>
							<li class="list-group-item">Password:&nbsp;<%=sPassword%></li>
						</ul>
					
						<ol class="list-group">
							<%For i = 0 To UBound(Roster, 2) - 1%>
								<li class="list-group-item">
									<a href="javascript:pop('/cc_meet/coach/roster/my_history.asp?roster_id=<%=Roster(0, i)%>',700,700)"><%=Roster(1, i)%></a>
								</li>
							<%Next%>
						</ol>
					</div>
					<div class="col-sm-4">
						<h5 class="h5">Line-Up</h5>
								
						<form class="form" name="submit_bibs" method="post" action="manage_teams.asp?meet_id=<%=lThisMeet%>&amp;this_team=<%=lThisTeam%>">
						<%For i = 0 To UBound(Races, 2) - 1%>
							<%Call GetLineUp(Races(0, i))%>
							<h5 class="h5"><%=Races(1, i)%></h5>
							<table class="table table-striped">
								<tr>
									<th>Name</th>
									<th>Gr</th>
									<th>Bib</th>
								</tr>
								<%For j = 0 To UBound(LineUp, 2) - 1%>
									<tr>
										<td><%=LineUp(1, j)%></td>
										<td><%=LineUp(2, j)%></td>
										<td>
											<input type="text" class="form-control" name="bib_<%=LineUp(0, j)%>_<%=Races(0, i)%>" 
												id="bib_<%=LineUp(0, j)%>_<%=Races(0, i)%>" value="<%=LineUp(3, j)%>" 
												style="text-align:right;" maxlength="4" size="3">
											<input type="hidden" name="avail_<%=LineUp(0, j)%>_<%=Races(0, i)%>" 
												id="avail_<%=LineUp(0, j)%>_<%=Races(0, i)%>" value="y">
										</td>
									</tr>
								<%Next%>
							</table>
						<%Next%>
						<input type="hidden" name="submit_these_bibs" id="submit_these_bibs" value="submit_these_bibs">
						<input type="submit" name="submit1a" id="submit1a" value="Save Bibs">
						</form>
					</div>
					<div class="col-sm-2">
						<h5 class="h5">Team Bib Range(s):</h5>
								
						<ul class="list-group">
							<%For i = 0 To UBound(BibRange, 2) - 1%>
								<li class="list-group-item">From <%=BibRange(1, i)%> to <%=BibRange(2, i)%></li>
							<%Next%>
						</ul>
					</div>
					<div class="col-sm-2">
						<h5 class="h5">Avail Bibs</h5>
								
						<ul class="list-group">
							<%For i = 0 To UBound(AvailBibs) - 1%>
								<li class="list-group-item"><%=AvailBibs(i)%></li>
							<%Next%>
						</ul>
					</div>
				</div>
			<%End If%>
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
