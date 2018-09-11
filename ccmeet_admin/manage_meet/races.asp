<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, rs3, sql3
Dim i, j
Dim lThisMeet, lThisRace
Dim sMeetName, sDist, sStartTime, sScoreMethod, sTmAwds, sIndAwds, sRemoveInc, sStartType, sIndivRelay, sComments, sRaceGender, sGender, sNumAllow
Dim sGradeYear, sIsRelay, sTeamScores, sTechnique, sShowResults, sStageRace, sOrderBy
Dim iFieldSize, iNumScore, iTotalParts, iNumSplits, iNumLaps, iViewOrder
Dim Races(), Teams(), LineUp(), AvailBibs()
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_race") = "submit_race" Then
	lThisRace = Request.Form.Item("races")
End If

If CStr(lThisRace) = vbNullString Then lThisRace = 0

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing
 
'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

'get races in this meet
i = 0
ReDim Races(1, 0)
sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY ViewOrder, RaceTime"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Call GetRaceData(rs(0).Value)
	
	Races(0, i) = rs(0).Value
	Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

iTotalParts = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RosterID FROM IndRslts WHERE MeetsID = " & lThisMeet
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then iTotalParts = rs.RecordCount
rs.Close
Set rs = Nothing

ReDim Teams(1, 0)
If Not lThisRace = 0 Then
	'get race gender
    sIsRelay = "n"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Gender, IndivRelay FROM Races WHERE RacesID = " & lThisRace
	rs.Open sql, conn, 1, 2
	sRaceGender = rs(0).Value
    If rs(1).Value = "Relay" Then sIsRelay = "y"
	rs.Close
	Set rs = Nothing

	'get teams for this race
	i = 0
	sql = "SELECT mt.TeamsID, t.TeamName FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " AND t.Gender = '" & Left(sRaceGender, 1) & "' ORDER BY t.TeamName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		Teams(0,  i) = rs(0).Value
		Teams(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve Teams(1, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
End If

Private Sub GetRaceData(lRaceID)
	sTmAwds = vbNullString
	sIndAwds = vbNullString
	sComments = vbNullString
    sTechnique = vbNullString
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RaceDist, RaceUnits, RaceTime, ScoreMethod, NumAllow, NumScore, TmAwds, IndAwds, RemoveInc, Comments, StartType, IndivRelay, Gender, "
	sql2 = sql2 & "TeamScores, NumSplits, Technique, NumLaps, ShowResults, StageRace, OrderBy, ViewOrder FROM Races WHERE RacesID = " & lRaceID
	rs2.Open sql2, conn, 1, 2
	sDist = rs2(0).Value & " " & rs2(1).Value
	sStartTime = rs2(2).Value
	iFieldSize = FieldSize(lRaceID)
	sScoreMethod = rs2(3).Value
	If  rs2(4).Value = 0 Then
		sNumAllow = "Unlimited"
	Else
		sNumAllow = rs2(4).Value
	End If
	iNumScore = rs2(5).Value
	If Not rs2(6).Value & "" = "" Then sTmAwds = Replace(rs2(6).Value, "''", "'")
	If Not rs2(7).Value & "" = "" Then sIndAwds = Replace(rs2(7).Value, "''", "'")
	sRemoveInc = rs2(8).Value
	If Not rs2(9).Value & "" = "" Then sComments = Replace(rs2(9).Value, "''", "'")
	sStartType = rs2(10).Value
	sIndivRelay = rs2(11).Value
	sGender = rs2(12).Value
    sTeamScores = rs2(13).Value
    iNumSplits = rs2(14).Value
    sTechnique = rs2(15).Value
    iNumLaps = rs2(16).Value
    sShowResults = rs2(17).Value
    sStageRace = rs2(18).Value
    sOrderBy = rs2(19).Value
	iViewOrder = rs2(20).Value
	rs2.Close
	Set rs2 = Nothing
End Sub

Private Function FieldSize(lThisRaceID)
	FieldSize = 0

	Set rs3 = Server.CreateObject("ADODB.Recordset")
	sql3 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lThisRaceID
	rs3.Open sql3, conn, 1, 2
	If rs3.RecordCount > 0 Then FieldSize = rs3.RecordCount
	rs3.Close
	Set rs3 = Nothing
End Function

Private Sub GetLineUp(lThisTeam)
	Dim x
	
	x = 0
	ReDim LineUp(3, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT r.RosterID, r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.Bib FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
	sql = sql & "INNER JOIN Grades g ON g.RosterID = r.RosterID WHERE r.TeamsID = " & lThisTeam & " AND ir.RacesID = " & lThisRace 
	sql = sql & " ORDER BY r.LastName, r.FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		LineUp(0,  x) = rs(0).Value
		LineUp(1, x) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
		LineUp(2,  x) = rs(3).Value
		LineUp(3,  x) = rs(4).Value
		x = x + 1
		ReDim Preserve LineUp(3, x)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Races Sheet</title>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    

			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>
			
			<h4 class="h4">Manage Races: <%=sMeetName%>&nbsp;(<%=dMeetDate%>)</h4>
			
			<div style="text-align:right;background-color:#ececd8;font-size:0.85em;margin:0 0 10px 0;">
                <a href="/ccmeet_admin/manage_meet/race_specs.asp?meet_id=<%=lThisMeet%>">Race Specs</a>
                &nbsp;|&nbsp;
				<a href="races.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">Race Data</a>
				&nbsp;|&nbsp;
				<a href="add_race.asp?meet_id=<%=lThisMeet%>">Add Race</a>
			</div>
			
			<div style="width:250px;float:left;background-color:#ececec;">
				<h4 class="h4">Race Info</h4>
				<h5>Total Participants: <%=iTotalParts%></h5>
				
				<ol style="font-size:0.9em;">
					<%For i = 0 To UBound(Races, 2) - 1%>
						<%Call GetRaceData(Races(0, i))%>
						<li style="margin-top:10px;">
							<a href="javascript:pop('edit_this_race.asp?race_id=<%=Races(0, i)%>',800,600)">
								<span style="font-weight:bold;"><%=Races(1, i)%>&nbsp;(<%=sStartTime%>)</span>
							</a>
							<ul style="font-size:0.95em;">
								<li>Gender:&nbsp;<%=sGender%></li>
								<li>Distance:&nbsp;<%=sDist%></li>
                                <li>Technique:&nbsp;<%=sTechnique%></li>
								<li>Field Size:&nbsp;<%=iFieldSize%></li>
								<li>Score Method:&nbsp;<%=sScoreMethod%></li>
								<li>Partic./Team:&nbsp;<%=sNumAllow%></li>
								<li>Num Score:&nbsp;<%=iNumScore%></li>
                                <li>Num Splits:&nbsp;<%=iNumSplits%></li>
								<li>Team Awards:&nbsp;<%=sTmAwds%></li>
								<li>Indiv Awards:&nbsp;<%=sIndAwds%></li>
								<li>Remove Inc.:&nbsp;<%=sRemoveInc%></li>
								<li>Start Type:&nbsp;<%=sStartType%></li>
								<li>Indiv/Relay:&nbsp;<%=sIndivRelay%></li>
                                <li>Team Scores:&nbsp;<%=sTeamScores%></li>
								<li>Comments:&nbsp;<%=sComments%></li>
                                <li>Number of Laps:&nbsp;<%=iNumLaps%></li>
                                <li>Show Results:&nbsp;<%=sShowResults%></li>
                                <li>Stage Race:&nbsp;<%=sStageRace%></li>
                                <li>Order Results By:&nbsp;<%=sOrderBy%></li>
								<li>View Order:&nbsp;<%=iViewOrder%></li>
							</ul>
						</li>
					<%Next%>
				</ol>
			</div>
			
			<div style="margin-left:275px;">
				<form name="get_race" method="Post" action="races.asp?meet_id=<%=lThisMeet%>">
				<h4 class="h4">Race Entries</h4>
				<select name="races" id="races" onchange="this.form.submit1.click();">
					<option value="0">&nbsp;</option>
					<%For i = 0 To UBound(Races, 2) - 1%>
						<%If CLng(lThisRace) = CLng(Races(0, i)) Then%>
							<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
						<%Else%>
							<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>
				<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
				<input type="submit" name="submit1" id="submit1" value="View This Race">
				</form>
				
				<%If Not CLng(lThisRace) = 0 Then%>
                    <%If sIsRelay = "y" Then%>
                        <a href="javascript:pop('relay_teams.asp?race_id=<%=lThisRace%>',600,650)" style="font-size: 0.85em;">View Relay Teams</a>
                    <%End If%>

					<ul style="list-style:none;margin-top:10px;">
						<%For i = 0 To UBound(Teams, 2) - 1%>
							<%Call GetLineUp(Teams(0, i))%>
							<li>
								<h5><%=Teams(1, i)%></h5>
								<ol style="font-size:0.8em;">
									<%For j = 0 To UBound(LineUp, 2) - 1%>
										<li><%=LineUp(1, j)%> (Gr:<%=LineUp(2, j)%>; Bib:<%=LineUp(3, j)%>)</li>
									<%Next%>
								</ol>
							</li>
						<%Next%>
					</ul>
				<%End If%>
			</div>
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
