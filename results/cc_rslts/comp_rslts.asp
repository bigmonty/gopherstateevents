<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lMeetID, lRaceID
Dim i, j, k, m, x
Dim sMeetName, dMeetDate, sMeetSite, sWeather, sSport, sTeamScores, sRaceName, sGradeYear, sUnits
Dim sTeamName, sPartName, sTime, sMilePace, sKmPace, sOrderResultsBy
Dim iDist
Dim RsltsArr, TmRslts, Races(), SortArr(8)
Dim fs, fname, sFileName
Dim bRsltsOfficial

lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
If Not IsNumeric(lMeetID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetID) < 0 Then Response.Redirect("http://www.google.com")

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

i = 0
ReDim Races(0)
sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(i) = rs(0).Value
	i = i + 1
	ReDim Preserve Races(i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub ThisRace(lRaceID)
	sql = "SELECT RaceDesc, RaceDist, RaceUnits, TeamScores, OrderBy FROM Races WHERE RacesID = " & lRaceID
	Set rs = conn.Execute(sql)
	sRaceName = rs(0).Value
	iDist = rs(1).Value 
    sUnits = rs(2).Value
    sTeamScores = rs(3).Value
    sOrderResultsBy = rs(4).Value
	Set rs = Nothing

    If sOrderResultsBy = "time" Then
		sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, ir.Bib "
        sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
        sql = sql & "INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
		sql = sql & "ORDER BY ir.Excludes, ir.FnlScnds, ir.Place"
    Else
		sql = "SELECT r.LastName, r.FirstName, t.TeamName, g.Grade" & sGradeYear & ", r.Gender, ir.RaceTime, ir.Excludes, ir.TeamPlace, ir.Bib "
        sql = sql & "FROM Roster r INNER JOIN Grades g ON r.RosterID = g.RosterID INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID "
        sql = sql & "INNER JOIN Teams t ON t.TeamsID = r.TeamsID WHERE ir.RacesID = " & lRaceID & " AND ir.Place > 0 AND ir.RaceTime > '00:00' "
		sql = sql & "ORDER BY ir.Excludes, ir.Place"
    End If
	Set rs = conn.Execute(sql)
    If rs.BOF and rs.EOF Then
       ReDim RsltsArr(8, 0)
    Else
        RsltsArr = rs.GetRows()
    End If
	Set rs = Nothing

    For i = 0 To UBound(RsltsArr, 2)
        RsltsArr(0, i) = Replace(RsltsArr(0, i), "''", "'")
        RsltsArr(1, i) = Replace(RsltsArr(1, i), "''", "'")
        RsltsArr(2, i) = Replace(RsltsArr(2, i), "''", "'")
'        If Not RsltsArr(3, i) & "" = "" Then RsltsArr(3, i) = GetGrade(RsltsArr(3,i))
        If RsltsArr(6, i) = "y" Then 
            RsltsArr(7,i) = "---"
        Else
            If CInt(RsltsArr(7, i)) = 0 Then RsltsArr(7,i) = "---"
        End If
    Next

    If sTeamScores = "y" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        If sSport = "Cross-Country" Then
			sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
			sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score DESC"
        Else
			sql = "SELECT t.TeamName, tr.Score, tr.R1, tr.R2, tr.R3, tr.R4, tr.R5, tr.R6, tr.R7 FROM Teams t INNER JOIN TmRslts tr "
			sql = sql & "ON t.TeamsID = tr.TeamsID WHERE tr.RacesID = " & lRaceID & " AND tr.Score <> '' ORDER BY Score"
        End If
		rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
           TmRslts = rs.GetRows()
        Else
            ReDim TmRslts(8, 0)
        End If
        rs.Close
		Set rs = Nothing

        If sSport = "Cross-Country" Then
            For i = 0 To UBound(TmRslts, 2) - 1
                For j = i + 1 To UBound(TmRslts, 2)
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
            For i = 0 To UBound(TmRslts, 2) - 1
                For j = i + 1 To UBound(TmRslts, 2)
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
        ReDim TmRslts(8, 0)
    End If
End Sub

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../includes/per_mile_cc.asp" -->
<!--#include file = "../../includes/per_km_cc.asp" -->
<%

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\inetpub\h51web\gopherstateevents\dwnlds\" & sMeetName & "_" & Year(dMeetDate) & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Comprehensive Results for " & sMeetName & " on " & dMeetDate)
fname.WriteLine("Meet Site: " & sMeetSite)
If Not sWeather = vbNullString Then fname.WriteLine("Weather: " & sWeather)
fname.WriteLine("Generated: " & Now())
fname.WriteBlankLines(2)

For m = 0 To UBound(Races) - 1
	Call ThisRace(Races(m))
	
	fname.WriteLine("Race: " & sRaceName)
	fname.WriteLine("Distance: " & iDist & " " & sUnits)
	
	fname.WriteBlankLines(1)

    If sTeamScores = "y" Then
	    fname.WriteLine("Team Results")
	    fname.WriteLine("PL" & vbTab & "TEAM" & Space(14) & vbTab & "PTS" & vbTab & "R1 " & vbTab & "R2 " & vbTab & "R3 " & vbTab & "R4 " & vbTab & "R5 "& vbTab & "R6 "& vbTab & "R7 ")
	    fname.WriteLine("---" & vbTab & "------------------" & vbTab & "---" & vbTab & "---"   & vbTab & "---" & vbTab & "---"   & vbTab & "---"   & vbTab & "---"  & vbTab & "---"  & vbTab & "---")
	    For i = 0 to UBound(TmRslts, 2)
		    sTeamName = TmRslts(0, i)
		    If Len(sTeamName) < 18 Then
			    sTeamName = sTeamName & Space(18 - Len(sTeamName))
		    Else
			    sTeamName = Left(sTeamName, 18)
		    End If
		
            For j = 1 To 8
                TmRslts(j, i) = Trim(TmRslts(j, i))
            Next

		    fname.WriteLine(i + 1 & vbTab & sTeamName & vbTab & TmRslts(1, i) & vbTab & TmRslts(2, i) & vbTab & TmRslts(3, i) & vbTab & TmRslts(4, i) & vbTab & TmRslts(5, i) & vbTab & TmRslts(6, i) & vbTab & TmRslts(7, i) & vbTab & TmRslts(8, i))
	    Next
	
	    fname.WriteBlankLines(1)
	End If

	fname.WriteLine("Individual Results")
	fname.WriteLine("PL " & vbTab & "TM " & vbTab & "BIB-NAME" & Space(16) & vbTab & "TEAM" & Space(14) & vbTab & "GR" & vbTab & "M/F" & vbTab & "TIME    "  & vbTab & "PER MI  " & vbTab & "PER KM  ")
	fname.WriteLine("---" & vbTab & "---" & vbTab & "------------------------" & vbTab & "------------------"           & vbTab & "--"   & vbTab & "---"    & vbTab & "--------"      & vbTab & "--------"        & vbTab & "--------")
	x = 1
	For i = 0 to UBound(RsltsArr, 2)
		sPartName = RsltsArr(8, i) & "-" & RsltsArr(1, i) & " " & RsltsArr(0, i)
		If Len(sPartName) < 24 Then
			sPartName = sPartName & Space(24 - Len(sPartName))
		Else
			sPartName = Left(sPartName, 24)
		End If
		
		sTeamName = RsltsArr(2, i)
		If Len(sTeamName) < 18 Then
			sTeamName = sTeamName & Space(18 - Len(sTeamName))
		Else
			sTeamName = Left(sTeamName, 18)
		End If
		
		sTime = CStr(RsltsArr(5, i))
		If Len(sTime) < 8 Then
			sTime = sTime & Space(8 - Len(sTime))
		Else
			sTime = Left(sTime, 8)
		End If
		
		sMilePace = CStr(PacePerMile(RsltsArr(5, i), iDist, sUnits))
		If Len(sMilePace) < 8 Then
			sMilePace = sMilePace & Space(8 - Len(sMilePace))
		Else
			sMilePace = Left(sMilePace, 8)
		End If
		
		sKmPace = CStr(PacePerKm(RsltsArr(5, i), iDist, sUnits))
		If Len(sKmPace) < 8 Then
			sKmPace = sKmPace & Space(8 - Len(sKmPace))
		Else
			sKmPace = Left(sKmPace, 8)
		End If
		
		If RsltsArr(6, i) = "y" Then
			fname.WriteLine("-" & vbTab & "-" & vbTab & sPartName & vbTab & sTeamName & vbTab & RsltsArr(3, i) & vbTab & RsltsArr(4, i) & vbTab & sTime & vbTab & sMilePace & vbTab & sKmPace)
		Else
			fname.WriteLine(i + 1 & vbTab & x & vbTab & sPartName & vbTab & sTeamName & vbTab & RsltsArr(3, i) & vbTab & RsltsArr(4, i) & vbTab & sTime & vbTab & sMilePace & vbTab & sKmPace)
			x = x + 1
		End If
	Next
	
	fname.WriteBlankLines(2)
Next

fname.Close
Set fname=nothing
Set fs=nothing

'begin download
Response.Redirect "../../dwnlds/" & sMeetName & "_" & Year(dMeetDate) & ".txt"
%>
<!DOCTYPE html>
<html>
<head>
<title>GSE Download Overall Cross-Country Results</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="GSE combined results download for cross-country running and nordic skiing.">

<link rel="icon" href="favicon.ico" type="image/x-icon"> 
<link rel="shortcut icon" href="favicon.ico" type="image/x-icon"> 

</head>
<body>
<%
conn.close
Set conn = Nothing
%>
</body>
</html>
