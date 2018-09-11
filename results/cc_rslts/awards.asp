
<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs
Dim i, j, k
Dim lMeetID, lRaceID, lRosterID
Dim sRaceName, sMeetName, sGradeYear, sOrderResultsBy, sScoreMethod, sTeamName, sUnits, sRaceDist, sSport
Dim iDist, iNumTeam, iNumInd
Dim RsltsArr(), TmRlsts(), MeetTms(), SortArr(8), Races
Dim dMeetDate
Dim bRsltsOfficial, bFound

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
If rs.BOF and rs.EOF Then
    bRsltsOfficial = False
Else
    bRsltsOfficial = True
End If
Set rs = Nothing

sql = "SELECT MeetName, MeetDate, Sport FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
sSport = rs(2).Value
Set rs = Nothing

'get year for roster grades
If Month(dMeetDate) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(dMeetDate) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(dMeetDate)), 2)
End If

If Len(sGradeYear) = 1 Then sGradeYear = "0" & sGradeYear

i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT t.TeamsID, t.TeamName, t.Gender FROM Teams t INNER JOIN MeetTeams mt ON t.TeamsID = mt.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lMeetID & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0, i) = rs(0).Value
	MeetTeams(1, i) = rs(1).Value & " (" & rs(2).Value & ")"

	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " AND ShowResults = 'y' ORDER BY ViewOrder"
Set rs = conn.Execute(sql)
Races = rs.GetRows()
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
	iNumTeam = Request.Form.Item("num_team")
	iNumInd = Request.Form.Item("num_ind")

	bFound = False
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT NumTeam, NumInd FROM Awards WHERE RacesID = " & lRaceID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		bFound = True
		rs(0).Value = iNumTeam
		rs(1).value = iNumInd
		rs.Update
	End If
	rs.Close
	Set rs = Nothing

	If bFound = False Then
		sql = "INSERT INTO Awards(RacesID, NumTeam, NumInd) VALUES (" & lRaceID & ", " & iNumTeam & ", " 
		sql = sql & iNumInd & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If
End If
	
If CLng(lRaceID) = 0 Then lRaceID = Races(0, 0)

If Not CLng(lRaceID) = 0 Then
    sql = "SELECT RaceDesc, RaceDist, RaceUnits, ScoreMethod, OrderBy FROM Races WHERE RacesID = " & lRaceID
    Set rs = conn.Execute(sql)
    sRaceName = Replace(rs(0).Value, "''", "'")
    iDist = rs(1).Value
    sUnits = rs(2).Value
    sScoreMethod = rs(3).Value
    sOrderResultsBy = rs(4).Value
    Set rs = Nothing

    sRaceDist = iDist & " " & sUnits

	iNumTeam = 0
	iNumInd = 0
	bFound = False
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT NumTeam, NumInd FROM Awards WHERE RacesID = " & lRaceID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
		bFound = True
		iNumTeam = rs(0).Value
		iNumInd = rs(1).value
	End If
	rs.Close
	Set rs = Nothing

	If bFound = False Then
		sql = "INSERT INTO Awards(RacesID, NumTeam, NumInd) VALUES (" & lRaceID & ", " & iNumTeam & ", " 
		sql = sql & iNumInd & ")"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
	End If
End If
	
i = 0
ReDim RsltsArr(8, 0)
If CInt(iNumInd) > 0 Then
	If sOrderResultsBy = "time" Then
		sql = "SELECT r.FirstName, r.LastName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ra.RaceDist, "
		sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
		sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
		sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
		sql = sql & "AND ir.Excludes = 'n' ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
	Else
		sql = "SELECT r.FirstName, r.LastName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ra.RaceDist, "
		sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
		sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
		sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
		sql = sql & "AND ir.Excludes = 'n' ORDER BY ir.Excludes, ir.Place"
	End If
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		'get gender for series
		RsltsArr(0,i) = rs(10).Value & "-" & Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
		RsltsArr(1,i) = Replace(rs(2).Value, "''", "'")
		RsltsArr(2,i) = rs(3).Value
		RsltsArr(3,i) = rs(4).Value
		RsltsArr(4,i) = rs(5).Value
		RsltsArr(5,i) = rs(6).Value
		RsltsArr(6,i) = rs(7).Value
		RsltsArr(7,i) = rs(8).Value
		If CInt(rs(9).Value) = 0 Then
			RsltsArr(8,i) = "---"
		Else
			RsltsArr(8,i) = rs(9).Value
		End If

		If CInt(i) = CInt(iNumInd) Then Exit Do

		i = i + 1
		ReDim Preserve RsltsArr(8, i)
		rs.MoveNext
	Loop
	Set rs = Nothing
End If

i = 0
ReDim TmRslts(8, 0)
If CInt(iNumInd) > 0 Then
	sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
	sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> ''"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		TmRslts(0, i) = rs(0).Value
		TmRslts(1, i) = rs(1).Value
		TmRslts(2, i) = Trim(rs(2).Value)
		TmRslts(3, i) = Trim(rs(3).Value)
		TmRslts(4, i) = Trim(rs(4).Value)
		TmRslts(5, i) = Trim(rs(5).Value)
		TmRslts(6, i) = Trim(rs(6).Value)
		TmRslts(7, i) = Trim(rs(7).Value)
		TmRslts(8, i) = Trim(rs(8).Value)

		If CInt(i) = CInt(iNumTeam) Then Exit Do

		i = i + 1
		ReDim Preserve TmRslts(8, i)
		rs.MoveNext
	Loop
	Set rs = Nothing

	If sSport = "Cross-Country" Then
		For i = 0 To UBound(TmRslts, 2) - 2
			For j = i + 1 To UBound(TmRslts, 2) - 1
				If CSng(TmRslts(1, i)) > CSng(TmRslts(1, j)) Then
					For k = 0 To 8
						SortArr(k) = TmRslts(k, i)
						TmRslts(k, i) = TmRslts(k, j)
						TmRslts(k, j) = SortArr(k)
					Next
				End If
			Next
		Next
	Else
		For i = 0 To UBound(TmRslts, 2) - 2
			For j = i + 1 To UBound(TmRslts, 2) - 1
				If CSng(TmRslts(1, i)) < CSng(TmRslts(1, j)) Then
					For k = 0 To 8
						SortArr(k) = TmRslts(k, i)
						TmRslts(k, i) = TmRslts(k, j)
						TmRslts(k, j) = SortArr(k)
					Next
				End If
			Next
		Next
	End If
End If

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<%
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Cross-Country/Nordic Ski Awards</title>
<meta name="description" content="Print GSE results for cross-country running and nordic skiing.">

<!--#include file = "../../includes/js.asp" --> 
</head>
<body>
<div class="container">
	    <h3 class="h3">Gopher State Events Awards for <%=sMeetName%> on <%=dMeetDate%></h3>

		<form role="form" class="form-inline" name="get_races" method="post" action="awards.asp?sport=<%=sSport%>&amp;meet_id=<%=lMeetID%>">
		<label for="races">Select Race:&nbsp;&nbsp;</label>
		<select class="form-control" name="races" id="races" onchange="this.form.get_race.click();">
			<option value="0">&nbsp;</option>
			<%For i = 0 to UBound(Races, 2)%>
				<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
					<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
				<%Else%>
					<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>&nbsp;&nbsp;
		<label for="num_ind">Individual Awards:&nbsp;&nbsp;</label>
		<select class="form-control" name="num_ind" id="num_ind" onchange="this.form.get_race.click();">
			<%For i = 0 To 100%>
				<%If CInt(iNumInd) = CInt(i) Then%>
					<option value="<%=i%>" selected><%=i%></option>
				<%Else%>
					<option value="<%=i%>"><%=i%></option>
				<%End If%>
			<%Next%>
		</select>&nbsp;&nbsp;
		<label for="num_team">Team Awards:&nbsp;&nbsp;</label>
		<select class="form-control" name="num_team" id="num_team" onchange="this.form.get_race.click();">
			<%For i = 0 To 25%>
				<%If CInt(iNumTeam) = CInt(i) Then%>
					<option value="<%=i%>" selected><%=i%></option>
				<%Else%>
					<option value="<%=i%>"><%=i%></option>
				<%End If%>
			<%Next%>
		</select>&nbsp;&nbsp;
		<input class="form-control" type="hidden" name="submit_race" id="submit_race" value="submit_race">
		<input class="form-control" type="submit" name="get_race" id="get_race" value="Get Results" style="font-size:0.8em;">
		</form>

		<h5 class="h4 bg-success">Team Awards</h5>
		
		<table class="table table-striped">
			<tr>
				<th>Pl</th>
				<th>Team</th>
				<th>Score</th>
				<th>R1</th>
				<th>R2</th>
				<th>R3</th>
				<th>R4</th>
				<th>R5</th>
				<th>R6</th>
				<th>R7</th>
			</tr>
			<%For i = 0 to UBound(TmRslts, 2) - 1%>
				<tr>
					<td><%=i + 1%></td>
					<td><%=TmRslts(0, i)%></td>
					<td><%=TmRslts(1, i)%></td>
					<td><%=TmRslts(2, i)%></td>
					<td><%=TmRslts(3, i)%></td>
					<td><%=TmRslts(4, i)%></td>
					<td><%=TmRslts(5, i)%></td>
					<td><%=TmRslts(6, i)%></td>
					<td><%=TmRslts(7, i)%></td>
					<td><%=TmRslts(8, i)%></td>
				</tr>
			<%Next%>
		</table>

		<h4 class="h4 bg-success">Individual Awards</h4>

		<table class="table table-striped">
			<tr>
				<th>Pl</th>
				<th>Tm</th>
				<th>Bib-Name</th>
				<th>Team</th>
				<th>Gr</th>
				<th>M/F</th>
				<th>Time</th>
				<th>Per Mi</th>
				<th>Per Km</th>
			</tr>
			<%k = 1%>
			<%For i = 0 to UBound(RsltsArr, 2) - 1%>
				<tr>
					<td>
						<%If RsltsArr(7, i) = "y" Then%>
							-
						<%Else%>
							<%=k%>
							<%k = k + 1%>
						<%End If%>
					</td>
					<td>
						<%=RsltsArr(8, i)%>
					</td>
					<td><%=RsltsArr(0, i)%></td>
					<td><%=RsltsArr(1, i)%></td>
					<td><%=RsltsArr(2, i)%></td>
					<td><%=RsltsArr(3, i)%></td>
					<td><%=RsltsArr(4, i)%></td>
					<td>
						<%=PacePerMile(RsltsArr(4, i), RsltsArr(5, i), RsltsArr(6, i))%>
					</td>
					<td>
						<%=PacePerKM(RsltsArr(4, i), RsltsArr(5, i), RsltsArr(6, i))%>
					</td>
				</tr>
			<%Next%>
		</table>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
