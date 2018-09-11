<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lMeetID, sMeetName, dMeetDate
Dim lRaceID
Dim i, j, k, x, m
Dim RsltsArr(), RacesArr(), TempArr(4)
Dim lTeamID, sTeamName, sGender
Dim lRosterID
Dim sGradeYear
Dim fs, fname, sFileName
Dim sPartName, sTime, sMilePace, sKmPace
Dim sOrderResultsBy

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")
lTeamID = Request.QueryString("team_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
   
'get order by
sql = "SELECT OrderBy FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sOrderResultsBy = rs(0).Value
Set rs = Nothing

sql = "SELECT TeamName, Gender FROM Teams WHERE TeamsID = " & lTeamID
Set rs = conn.Execute(sql)
sTeamName = rs(0).Value
sGender = rs(1).Value
Set rs = Nothing

sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = rs(0).Value
dMeetDate = rs(1).Value

If Month(rs(1).Value) <=7 Then
	sGradeYear = Right(CStr(Year(rs(1).Value) - 1), 2)
Else
	sGradeYear = Right(CStr(Year(rs(1).Value)), 2)	
End If

Set rs = Nothing

Select Case sGender
	Case "M"
		sGender = "Male"
	Case "F"
		sGender = "Female"
End Select
    
i = 0
ReDim RacesArr(3, 0)
sql = "SELECT RacesID, RaceDist, RaceUnits, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " AND Gender = '" & sGender & "'"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    RacesArr(0, i) = rs(0).Value
    RacesArr(1, i) = rs(1).Value
    RacesArr(2, i) = rs(2).Value
    RacesArr(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve RacesArr(3, i)
    rs.MoveNext
Loop
Set rs = Nothing

Private Sub RaceResults(lRaceID)   
	ReDim RsltsArr(4, 0)
	k = 0 
	sql = "SELECT r.FirstName, r.LastName, g.Grade" & sGradeYear & ", ir.RaceTime, t.TeamsID FROM IndRslts ir "
	sql = sql & "INNER JOIN Roster r ON ir.RosterID = r.RosterID INNER JOIN Grades g On g.RosterID = r.RosterID "
	sql = sql & "INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE ir.RacesID = " & lRaceID 
	sql = sql & " AND ir.Place > 0 ORDER BY ir.Place"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    RsltsArr(0, k) = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")	'name
		RsltsArr(1, k) = k + 1
	    RsltsArr(2, k) = rs(2).Value	'grade
	    RsltsArr(3, k) = rs(3).Value	'time
	    RsltsArr(4, k) = rs(4).Value	'team
		k = k + 1
	    ReDim Preserve RsltsArr(4, k)
	    rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing

	If sOrderResultsBy = "Time" Then	    
		're-order results if order by time
		For x = 0 To UBound(RsltsArr, 2) - 2
		    For m = x + 1 To UBound(RsltsArr, 2) - 1
		        If ConvertToSeconds(RsltsArr(3, x)) > ConvertToSeconds(RsltsArr(3, m)) Then
		            'swap places if first time is slower than last
		            For k = 0 To 4
		                TempArr(k) = RsltsArr(k, x)
		                RsltsArr(k, x) = RsltsArr(k, m)
		                RsltsArr(k, m) = TempArr(k)
		            Next
		        End If
		    Next
		Next
		
		'get race place
		For x = 0 To UBound(RsltsArr, 2) - 1
			RsltsArr(1, x) = x + 1
		Next
	End If
End Sub

%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->
<!--#include file = "../../../includes/per_mile_cc.asp" -->
<!--#include file = "../../../includes/per_km_cc.asp" -->
<%

Set fs=Server.CreateObject("Scripting.FileSystemObject")
sFileName = "C:\Inetpub\h51web\gopherstateevents\dwnlds\" & sMeetName & "_" & sTeamName & ".txt"
Set fname=fs.CreateTextFile(sFileName, True)

fname.WriteLine("GSE Race Results for " & sTeamName & " in the " & sMeetName & "  on " & dMeetDate)
fname.WriteLine("Generated: " & Now())

For j = 0 to UBound(RacesArr, 2) - 1
	Call RaceResults(RacesArr(0, j))
	
	fname.WriteBlankLines(2)
	fname.WriteLine("Race: " & RacesArr(3, j))
	fname.WriteBlankLines(1)

	fname.WriteLine("PL" & vbTab & "NAME" & Space(14) & vbTab & "GR" & vbTab & "TIME    "  & vbTab & "PER MI  " & vbTab & "PER KM  ")
	fname.WriteLine("---" & vbTab & "------------------"             & vbTab & "---"     & vbTab & "--------"      & vbTab & "--------"       & vbTab & "--------")
	For i = 0 to UBound(RsltsArr, 2) - 1
		If CLng(RsltsArr(4, i)) = CLng(lTeamID) Then
			sPartName = RsltsArr(0, i)
			If Len(sPartName) < 18 Then
				sPartName = sPartName & Space(18 - Len(sPartName))
			Else
				sPartName = Left(sPartName, 18)
			End If
		
			sTime = CStr(RsltsArr(3, i))
			If Len(sTime) < 8 Then
				sTime = sTime & Space(8 - Len(sTime))
			Else
				sTime = Left(sTime, 8)
			End If
		
			sMilePace = CStr(PacePerMile(RsltsArr(3, i), RacesArr(1, j), RacesArr(2, j)))
			If Len(sMilePace) < 8 Then
				sMilePace = sMilePace & Space(8 - Len(sMilePace))
			Else
				sMilePace = Left(sMilePace, 8)
			End If
		
			sKmPace = CStr(PacePerKm(RsltsArr(3, i), RacesArr(1, j), RacesArr(2, j)))
			If Len(sKmPace) < 8 Then
				sKmPace = sKmPace & Space(8 - Len(sKmPace))
			Else
				sKmPace = Left(sKmPace, 8)
			End If
		
			fname.WriteLine(RsltsArr(1, i) & vbTab & sPartName & vbTab & RsltsArr(2, i) & vbTab & sTime & vbTab & sMilePace & vbTab & sKmPace)
		End If
	Next
Next

fname.Close
Set fname=nothing
Set fs=nothing

'begin download
Response.Redirect "/dwnlds/" & sMeetName & "_" & sTeamName & ".txt"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Cross-Country/Nordic Download Our Results</title>
<!--#include file = "../../../includes/meta2.asp" -->
</head>

<body>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
