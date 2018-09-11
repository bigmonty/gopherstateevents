<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lMeetID, sMeetName, dMeetDate, sMeetSite, sWeather, sRaceDist, sSport, sTeamName, sPartName, sTime, sMilePace, sKmPace, sOrderResultsBy
Dim lRaceID, sRaceName
Dim i, j, k, x
Dim bRsltsOfficial
Dim RsltsArr(), TmRslts(), SortArr(8)
Dim sGradeYear
Dim fs, fname, sFileName

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

sql = "SELECT RaceDesc, RaceDist, RaceUnits, OrderBy FROM Races WHERE RacesID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = rs(0).Value
sRaceDist = rs(1).Value & " " & rs(2).Value
sOrderResultsBy = rs(3).Value
Set rs = Nothing

i = 0
ReDim RsltsArr(8, 0)
If sOrderResultsBy = "time" Then
    sql = "SELECT r.FirstName, r.LastName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ra.RaceDist, "
    sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
    sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
    sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.Excludes = 'n' ORDER BY ir.FnlScnds, ir.Place"
Else
    sql = "SELECT r.FirstName, r.LastName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ra.RaceDist, "
    sql = sql & "ra.RaceUnits, ir.Excludes, ir.TeamPlace, ir.Bib FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir "
    sql = sql & "ON r.RosterID = ir.RosterID INNER JOIN Teams t ON t.TeamsID = r.TeamsID INNER JOIN Races ra "
    sql = sql & "ON ra.RacesID = ir.RacesID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.Excludes = 'n' ORDER BY ir.Place"
End If
Set rs = conn.Execute(sql)
Do While Not rs.EOF
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

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<%

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\inetpub\h51web\gopherstateevents\dwnlds\" & sMeetName & "_" & sRaceName & "_" & Year(dMeetDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Race Results for " & sMeetName & " (" & sRaceName & ")  on " & dMeetDate)
fname.WriteLine("Meet Site: " & sMeetSite)
fname.WriteLine("Distance: " & sRaceDist)
If Not sWeather = vbNullString Then fname.WriteLine("Weather: " & sWeather)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)

fname.WriteLine("Team Results")
fname.WriteBlankLines(1)
fname.WriteLine("PL" & vbTab & "TEAM" & Space(14) & vbTab & "PTS" & vbTab & "R1" & vbTab & "R2" & vbTab & "R3" & vbTab & "R4" & vbTab & "R5"& vbTab & "R6"& vbTab & "R7")
fname.WriteLine("---" & vbTab & "------------------"             & vbTab & "---" & vbTab   & "--"   & vbTab & "--"   & vbTab & "--"   & vbTab & "--"   & vbTab & "--"  & vbTab & "--"  & vbTab & "--")
For i = 0 to UBound(TmRslts, 2) - 1
	sTeamName = TmRslts(0, i)
	If Len(sTeamName) < 18 Then
		sTeamName = sTeamName & Space(18 - Len(sTeamName))
	Else
		sTeamName = Left(sTeamName, 18)
	End If
	
	fname.WriteLine(i + 1 & vbTab & sTeamName & vbTab & TmRslts(1, i) & vbTab & TmRslts(2, i) & vbTab & TmRslts(3, i) & vbTab & TmRslts(4, i) & vbTab & TmRslts(5, i) & vbTab & TmRslts(6, i) & vbTab & TmRslts(7, i) & vbTab & TmRslts(8, i))
Next

fname.WriteBlankLines(2)

fname.WriteLine("Individual Results")
fname.WriteBlankLines(1)
fname.WriteLine("PL " & vbTab & "TM " & vbTab & "BIB-NAME" & Space(10) & vbTab & "TEAM" & Space(14) & vbTab & "GR" & vbTab & "M/F" & vbTab & "TIME    "  & vbTab & "PER MI  " & vbTab & "PER KM  ")
fname.WriteLine("---" & vbTab & "---" & vbTab & "------------------"   & vbTab & "------------------" & vbTab & "--"   & vbTab & "---" & vbTab & "--------" & vbTab & "--------" & vbTab & "--------")
x = 1
For i = 0 to UBound(RsltsArr, 2) - 1
	sPartName = RsltsArr(0, i)
	If Len(sPartName) < 18 Then
		sPartName = sPartName & Space(18 - Len(sPartName))
	Else
		sPartName = Left(sPartName, 18)
	End If
	
	sTeamName = RsltsArr(1, i)
	If Len(sTeamName) < 18 Then
		sTeamName = sTeamName & Space(18 - Len(sTeamName))
	Else
		sTeamName = Left(sTeamName, 18)
	End If
	
	sTime = ConvertToMinutes(Round(ConvertToSeconds(RsltsArr(4, i)), 3))
	
	sMilePace = CStr(PacePerMile(RsltsArr(4, i), RsltsArr(5, i), RsltsArr(6, i)))
	If Len(sMilePace) < 8 Then
		sMilePace = sMilePace & Space(8 - Len(sMilePace))
	Else
		sMilePace = Left(sMilePace, 8)
	End If
	
	sKmPace = CStr(PacePerKm(RsltsArr(4, i), RsltsArr(5, i), RsltsArr(6, i)))
	If Len(sKmPace) < 8 Then
		sKmPace = sKmPace & Space(8 - Len(sKmPace))
	Else
		sKmPace = Left(sKmPace, 8)
	End If
	
	If RsltsArr(7, i) = "y" Then
		fname.WriteLine("-" & vbTab & "-" & vbTab & sPartName & vbTab & sTeamName & vbTab & RsltsArr(2, i) & vbTab & RsltsArr(3, i) & vbTab & sTime & vbTab & sMilePace & vbTab & sKmPace)
	Else
		fname.WriteLine(x & vbTab & RsltsArr(8, i) & vbTab & sPartName & vbTab & sTeamName & vbTab & RsltsArr(2, i) & vbTab & RsltsArr(3, i) & vbTab & sTime & vbTab & sMilePace & vbTab & sKmPace)
		x = x + 1
	End If
Next

fname.Close
Set fname=nothing
Set fs=nothing

'begin download
Response.Redirect "../../dwnlds/" & sMeetName & "_" & sRaceName & "_" & Year(dMeetDate) & ".txt"
%>
<!DOCTYPE html>
<html>
<head>
<title>GSE Download Overall Cross-Country Results</title>
<meta name="description" content="GSE overall results for cross-country running and nordic skiing.">
</head>
<body>
<%
conn.close
Set conn = Nothing
%>
</body>
</html>
