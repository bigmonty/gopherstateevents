
<%@ Language=VBScript%>
<%
Option Explicit

Dim sql, conn, rs, rs2, sql2
Dim i, j, k
Dim lMeetID, lRaceID, lSeriesID, lRosterID, lWhichTeam
Dim sRaceName, sSeriesName, sSeriesGender, sMeetName, sGradeYear, sOrderResultsBy, sScoreMethod, sRsltsPage, sTeamName, sUnits, sMeetSite, sWeather
Dim sRaceDist, sSport
Dim iDist
Dim RsltsArr(), TmRlsts(), MeetTms(), SortArr(8)
Dim dMeetDate
Dim bRsltsOfficial

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect("http://www.google.com")
If CLng(lRaceID) < 0 Then Response.Redirect("http://www.google.com")

lWhichTeam = Request.QueryString("which_team")
If CStr(lWhichTeam) = vbNullString Then lWhichTeam = 0
If Not IsNumeric(lWhichTeam) Then Response.Redirect("http://www.google.com")
If CLng(lWhichTeam) < 0 Then Response.Redirect("http://www.google.com")

sRsltsPage = Request.QueryString("rslts_page")
If sRsltsPage = vbNullString Then sRsltsPage = "overall_rslts.asp"

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

sql = "SELECT MeetName, MeetDate, MeetSite, Weather, Sport FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value
If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
If Not rs(3).Value & "" = "" Then sWeather = Replace(rs(3).Value, "''", "'")
sSport = rs(4).Value
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

sql = "SELECT RaceDesc, RaceDist, RaceUnits, ScoreMethod, OrderBy FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
iDist = rs(1).Value
sUnits = rs(2).Value
sScoreMethod = rs(3).Value
sOrderResultsBy = rs(4).Value
Set rs = Nothing

sRaceDist = iDist & " " & sUnits
	
'see if this race is in a series
sql = "SELECT s.SeriesID, s.SeriesName FROM Series s INNER JOIN SeriesMeets sm ON s.SeriesID = sm.SeriesID "
sql = sql & "WHERE sm.RacesID = " & lRaceID
Set rs = conn.Execute(sql)
If rs.BOF and rs.EOF Then
    lSeriesID = 0
Else
	lSeriesID = rs(0).Value
	sSeriesName = Replace(rs(1).Value, "''", "'")
End If
Set rs = Nothing
	
If sRsltsPage = "overall_rslts.asp" Then
	i = 0
	ReDim RsltsArr(8, 0)
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
		If i = 0 Then sSeriesGender = rs(4).Value
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
		i = i + 1
		ReDim Preserve RsltsArr(8, i)
		rs.MoveNext
	Loop
	Set rs = Nothing

	i = 0
	ReDim TmRslts(8, 0)
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
Else
	i = 0
	ReDim RsltsArr(4, 0)
	If Not CLng(lWhichTeam) = 0 Then
		sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lWhichTeam
		Set rs = conn.Execute(sql)
		sTeamName = rs(0).Value & " (" & rs(1).Value & ")"
		Set rs = Nothing
			
		Set rs = Server.CreateObject("ADODB.Recordset")
        If sOrderResultsBy = "time" Then
			sql = "SELECT r.RosterID, r.FirstName, r.LastName, r.Gender, ir.RaceTime, ir.Bib, g.Grade" & sGradeYear & " FROM IndRslts ir "
			sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
			sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
			sql = sql & lWhichTeam & " AND ir.Place > 0 ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
            Else
			sql = "SELECT r.RosterID, r.FirstName, r.LastName, r.Gender, ir.RaceTime, ir.Bib, g.Grade" & sGradeYear & " FROM IndRslts ir "
			sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Teams t ON r.TeamsID = t.TeamsID "
			sql = sql & "INNER JOIN Grades g ON r.RosterID = g.RosterID WHERE ir.RacesID = " & lRaceID & " AND t.TeamsID = " 
			sql = sql & lWhichTeam & " AND ir.Place > 0 ORDER BY ir.Excludes, ir.Place"
        End If
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			RsltsArr(0, i) = rs(5).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
			RsltsArr(1, i) = GetPlace(rs(0).Value)
			RsltsArr(2, i) = rs(6).Value
			RsltsArr(3, i) = Trim(rs(3).Value)
			RsltsArr(4, i) = Trim(rs(4).Value)
			i = i + 1
			ReDim Preserve RsltsArr(4, i)
			rs.MoveNext
		Loop
		Set rs = Nothing
	End If
End If

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<%

Function GetPlace(lRosterID)
	GetPlace = 0
	sql2 = "SELECT RosterID FROM IndRslts WHERE RacesID = " & lRaceID & " AND Place > 0 ORDER BY Place"
	Set rs2 = conn.Execute(sql2)
	Do While Not rs2.EOF
		GetPlace = GetPlace + 1
		If CLng(rs2(0).Value) = CLng(lRosterID) Then Exit Do
		rs2.MoveNext
	Loop
	Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Print GSE Cross-Country Results</title>
<meta name="description" content="Print GSE results for cross-country running and nordic skiing.">
<!--#include file = "../../includes/js.asp" --> 
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
	    <h3 class="h3">Results for <%=sMeetName%> on <%=dMeetDate%>
            <br><small><%=sRaceName%>&nbsp;(Dist:&nbsp;<%=sRaceDist%>)<br><%=sMeetSite%></small></h3>

	    <div class="bg-warning">
		    <a href="javascript:window.print();">Print</a>
	    </div>

	    <%If Not sWeather = vbNullString Then%>
		    <p>Weather:&nbsp;<%=sWeather%></p>
	    <%End If%>
	
	    <%If Not sRsltsPage = "overall_rslts.asp" Then%>
		    <h4 class="h4">Race:&nbsp; <%=sRaceName%>&nbsp;Team: <%=sTeamName%></h4>
	    <%End If%>

	    <%If sRsltsPage = "overall_rslts.asp" Then%>
		    <h5 class="h4 bg-success">Team Results</h5>
		
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

		    <h4 class="h4 bg-success">Individual Results</h4>

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
	    <%Else%>
		    <table class="table table-striped">
			    <tr>
				    <th>Rnr</th>
				    <th>Bib-Name</th>
				    <th>Pl</th>
				    <th>Gr</th>
				    <th>M/F</th>
				    <th>Time</th>
				    <th>Per Mi</th>
				    <th>Per Km</th>
			    </tr>
			    <%For i = 0 to UBound(RsltsArr, 2) - 1%>
				    <tr>
					    <td><%=i + 1%></td>
					    <td><%=RsltsArr(0, i)%></td>
					    <td><%=RsltsArr(1, i)%></td>
					    <td><%=RsltsArr(2, i)%></td>
					    <td><%=RsltsArr(3, i)%></td>
					    <td><%=RsltsArr(4, i)%></td>
					    <td>
							<%If RsltsArr(4, i) <> "00:00" Then%>
								<%=PacePerMile(RsltsArr(4, i), iDist, sUnits)%>
							<%End If%>
					    </td>
					    <td>
							<%If RsltsArr(4, i) <> "00:00" Then%>
								<%=PacePerKm(RsltsArr(4, i), iDist, sUnits)%>
							<%End If%>
					    </td>
				    </tr>
			    <%Next%>
		    </table>
	    <%End If%>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
